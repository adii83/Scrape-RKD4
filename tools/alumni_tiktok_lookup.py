import json
import random
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import quote


EXCEL_PATH = Path("Data Alumni.xlsx")
OUT_DIR = Path("outputs")
DEFAULT_UNIVERSITY = "Universitas Muhammadiyah Malang"

# Rentang dan batch (1-based, tanpa header)
START_ROW = int(__import__("os").environ.get("BH_START_ROW", "1"))
END_ROW = int(__import__("os").environ.get("BH_END_ROW", "10"))
BATCH_SIZE = int(__import__("os").environ.get("BH_BATCH_SIZE", "5"))

# Resume
RESUME_ENABLED = __import__("os").environ.get("BH_RESUME_ENABLED", "1") != "0"
PROGRESS_FILE = OUT_DIR / "alumni_tiktok_progress.json"

# Fallback source
USE_LINKEDIN_FALLBACK = True
LINKEDIN_ALL_FILE = OUT_DIR / f"alumni_all_{START_ROW}_{END_ROW}.json"
USE_INSTAGRAM_FALLBACK = True
INSTAGRAM_ALL_FILE = OUT_DIR / f"alumni_ig_all_{START_ROW}_{END_ROW}.json"
USE_FACEBOOK_FALLBACK = True
FACEBOOK_ALL_FILE = OUT_DIR / f"alumni_fb_all_{START_ROW}_{END_ROW}.json"

# Delay acak
SEARCH_DELAY_MIN = float(__import__("os").environ.get("BH_SEARCH_DELAY_MIN", "5.0"))
SEARCH_DELAY_MAX = float(__import__("os").environ.get("BH_SEARCH_DELAY_MAX", "15.0"))
PROFILE_DELAY_MIN = float(__import__("os").environ.get("BH_PROFILE_DELAY_MIN", "2.0"))
PROFILE_DELAY_MAX = float(__import__("os").environ.get("BH_PROFILE_DELAY_MAX", "4.5"))
BATCH_COOLDOWN_MIN = float(__import__("os").environ.get("BH_BATCH_COOLDOWN_MIN", "20.0"))
BATCH_COOLDOWN_MAX = float(__import__("os").environ.get("BH_BATCH_COOLDOWN_MAX", "90.0"))
BLOCK_RECHECK_MIN = float(__import__("os").environ.get("BH_BLOCK_RECHECK_MIN", "120.0"))
BLOCK_RECHECK_MAX = float(__import__("os").environ.get("BH_BLOCK_RECHECK_MAX", "240.0"))
ENABLE_BLOCK_WAIT = __import__("os").environ.get("BH_ENABLE_BLOCK_WAIT", "0") == "1"

BLOCKED_TIKTOK_URL_KEYWORDS = [
    "tiktok.com/@ummcampus",
]

SEARCH_ENGINES = [
    ("duckduckgo", "https://duckduckgo.com/?q={}"),
    ("yahoo", "https://search.yahoo.com/search?p={}"),
    ("yandex", "https://yandex.com/search/?text={}"),
    ("bing", "https://www.bing.com/search?q={}"),
]


def clean_text(value):
    return re.sub(r"\s+", " ", value or "").strip()


def search_url(query, engine_template):
    return engine_template.format(quote(query))


def detect_challenge_or_block():
    return bool(
        js(
            r"""
(() => {
  const t = (document.title || '').toLowerCase();
  const b = (document.body ? document.body.innerText : '').toLowerCase();
  const s = (t + "\n" + b).slice(0, 12000);
  const keys = [
    "captcha",
    "verify you are human",
    "are you human",
    "unusual traffic",
    "too many requests",
    "rate limit",
    "temporarily blocked",
    "access denied",
    "blocked",
    "detected unusual",
    "robot"
  ];
  return keys.some(k => s.includes(k));
})()
"""
        )
    )


def wait_until_unblocked(context):
    while True:
        if not detect_challenge_or_block():
            return
        wait_sec = random.uniform(BLOCK_RECHECK_MIN, BLOCK_RECHECK_MAX)
        print(f"[TT] challenge/rate-limit terdeteksi ({context}), tunggu {int(wait_sec)} detik lalu cek ulang...")
        wait(wait_sec)


def safe_navigate(url, context, min_delay, max_delay):
    while True:
        cdp("Page.navigate", url=url)
        wait_for_load()
        wait(random.uniform(min_delay, max_delay))
        if ENABLE_BLOCK_WAIT and detect_challenge_or_block():
            wait_until_unblocked(context)
            continue
        return


def first_email(text):
    if not text:
        return ""
    matches = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    for raw in matches:
        email = raw.strip(".,;:()[]{}<>\"'")
        parts = email.split("@")
        if len(parts) != 2:
            continue
        local, domain = parts
        if not local or not domain or "." not in domain:
            continue
        tld = domain.rsplit(".", 1)[-1]
        if not re.fullmatch(r"[A-Za-z]{2,}", tld):
            continue
        return email.lower()
    return ""


def first_phone(text):
    if not text:
        return ""
    candidates = re.findall(r"(\+?\d[\d\-\s()]{7,}\d)", text)
    for raw in candidates:
        value = clean_text(raw)
        if not (value.startswith("+") or value.startswith("0")):
            continue
        digits = re.sub(r"\D", "", value)
        if 9 <= len(digits) <= 15:
            return value
    return ""


def clean_position_text(value):
    text = clean_text(value)
    if re.search(r"^(follow|following|message|view profile)$", text, flags=re.IGNORECASE):
        return ""
    return text


def extract_work_from_text(text):
    t = clean_text(text)
    patterns = [
        r"\bat\s+([^|,\n]+)",
        r"\bdi\s+([^|,\n]+)",
        r"work(?:ing)?\s+at\s+([^|,\n]+)",
        r"founder\s+(?:of|@)\s*([^|,\n]+)",
        r"owner\s+(?:of|@)\s*([^|,\n]+)",
    ]
    for p in patterns:
        m = re.search(p, t, flags=re.IGNORECASE)
        if m:
            return clean_text(m.group(1))
    return ""


def classify_work_type(text):
    t = clean_text(text).lower()
    if any(k in t for k in ["pns", "pegawai negeri", "asn", "civil servant"]):
        return "PNS"
    if any(k in t for k in ["owner", "founder", "co-founder", "entrepreneur", "wirausaha", "business"]):
        return "Wirausaha"
    if t:
        return "Swasta"
    return ""


def _col_to_index(cell_ref):
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch.upper()) - ord("A") + 1)
    return value - 1


def _cell_text(cell, shared_strings, ns):
    value = cell.find("x:v", ns)
    inline = cell.find("x:is/x:t", ns)
    if inline is not None:
        return inline.text or ""
    if value is None:
        return ""
    text = value.text or ""
    if cell.attrib.get("t") == "s":
        return shared_strings[int(text)]
    return text


def read_xlsx_slice(path, start_row, end_row):
    if start_row < 1:
        raise ValueError("START_ROW minimal 1")
    if end_row < start_row:
        raise ValueError("END_ROW harus >= START_ROW")

    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("x:si", ns):
                parts = [t.text or "" for t in si.findall(".//x:t", ns)]
                shared_strings.append("".join(parts))

        root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        parsed_rows = []
        for row in root.findall(".//x:sheetData/x:row", ns):
            values = {}
            for cell in row.findall("x:c", ns):
                idx = _col_to_index(cell.attrib.get("r", "A1"))
                values[idx] = _cell_text(cell, shared_strings, ns)
            if values:
                parsed_rows.append([values.get(i, "") for i in range(max(values) + 1)])
            if len(parsed_rows) >= end_row + 1:
                break

    headers = parsed_rows[0]
    out = []
    for row in parsed_rows[start_row : end_row + 1]:
        out.append({headers[i]: row[i] if i < len(row) else "" for i in range(len(headers))})
    return out


def _xlsx_escape(value):
    return (
        str(value or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _col_name(index):
    s = ""
    n = index + 1
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def write_simple_xlsx(path, columns, rows):
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    lines.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    lines.append("<sheetData>")

    header_cells = []
    for i, col in enumerate(columns):
        cell_ref = f"{_col_name(i)}1"
        header_cells.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{_xlsx_escape(col)}</t></is></c>')
    lines.append(f'<row r="1">{"".join(header_cells)}</row>')

    for r_idx, row in enumerate(rows, start=2):
        row_cells = []
        for c_idx, col in enumerate(columns):
            cell_ref = f"{_col_name(c_idx)}{r_idx}"
            value = _xlsx_escape(row.get(col, ""))
            row_cells.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{value}</t></is></c>')
        lines.append(f'<row r="{r_idx}">{"".join(row_cells)}</row>')

    lines.append("</sheetData>")
    lines.append("</worksheet>")
    sheet1 = "\n".join(lines)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet1)


def write_csv(path, columns, rows):
    lines = [",".join(columns)]
    for row in rows:
        lines.append(",".join(json.dumps(str(row.get(col, "")), ensure_ascii=False) for col in columns))
    path.write_text("\ufeff" + "\n".join(lines), encoding="utf-8")


def save_progress(state):
    PROGRESS_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def load_progress():
    if not PROGRESS_FILE.exists():
        return None
    try:
        return json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None


def make_output_paths(batch_start, batch_end):
    base = f"alumni_tt_{batch_start}_{batch_end}"
    return {
        "json": OUT_DIR / f"{base}.json",
        "csv": OUT_DIR / f"{base}.csv",
        "xlsx": OUT_DIR / f"{base}.xlsx",
    }


def load_fallback_map(path, enabled):
    if not enabled or not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    out = {}
    for item in data:
        key = str(item.get("row_excel", "")).strip()
        if key:
            out[key] = item
    return out


def merge_base(*rows):
    merged = {}
    for row in rows:
        for k, v in (row or {}).items():
            if not clean_text(str(merged.get(k, ""))) and clean_text(str(v)):
                merged[k] = v
    return merged


def is_blocked_tiktok_url(url):
    u = clean_text(url).lower().rstrip("/")
    if not u:
        return False
    return any(k in u for k in BLOCKED_TIKTOK_URL_KEYWORDS)


def search_tiktok_candidates(query, engine_name, engine_template):
    safe_navigate(search_url(query, engine_template), f"search engine={engine_name}", SEARCH_DELAY_MIN, SEARCH_DELAY_MAX)
    return js(
        r"""
(() => {
  const links = Array.from(document.querySelectorAll('a[href]')).map(a => a.href);
  const out = [];
  const seen = new Set();
  for (const href of links) {
    if (!href.includes('tiktok.com/')) continue;
    const clean = href.split('?')[0].replace(/\/$/, '');
    if (/tiktok\.com\/(discover|tag|music|video|explore)\//i.test(clean)) continue;
    const m = clean.match(/tiktok\.com\/@([A-Za-z0-9._]+)/i);
    if (!m) continue;
    const username = m[1];
    const profile = "https://www.tiktok.com/@" + username;
    if (seen.has(profile)) continue;
    seen.add(profile);
    out.push({tiktok_url: profile, username});
    if (out.length >= 5) break;
  }
  return out;
})()
"""
    ) or []


def pick_valid_tiktok_candidate(candidates):
    for item in candidates or []:
        if not is_blocked_tiktok_url(item.get("tiktok_url", "")):
            return item
    return None


def read_tiktok_profile(tiktok_url):
    safe_navigate(tiktok_url, "open tiktok profile", PROFILE_DELAY_MIN, PROFILE_DELAY_MAX)
    return js(
        r"""
(() => {
  const bodyText = document.body ? document.body.innerText : '';
  const title = document.title || '';
  const descTag = document.querySelector('meta[name="description"]');
  const desc = descTag ? (descTag.content || '') : '';
  const links = Array.from(document.querySelectorAll('a[href]')).map(a => a.href);
  const extLinks = links.filter(h => !h.includes('tiktok.com')).slice(0, 10);
  return {
    body_text: bodyText,
    title,
    meta_description: desc,
    external_links: Array.from(new Set(extLinks))
  };
})()
"""
    ) or {}


def pick_first_non_empty(*values):
    for v in values:
        if clean_text(v):
            return clean_text(v)
    return ""


def merge_with_tiktok(base_row, tt_info):
    row = dict(base_row)
    body = clean_text(tt_info.get("body_text", ""))
    title = clean_text(tt_info.get("title", ""))
    desc = clean_text(tt_info.get("meta_description", ""))
    combined = clean_text(body + " " + title + " " + desc)

    company = extract_work_from_text(combined)
    email = first_email(combined)
    phone = first_phone(combined)
    position = clean_position_text(desc)
    ext_links = tt_info.get("external_links", []) or []
    work_social = ""
    for link in ext_links:
        l = link.lower()
        if "linkedin.com/company/" in l or "instagram.com/" in l or "facebook.com/" in l or "tiktok.com/" in l:
            work_social = link
            break

    row["alamat_sosial_media_tiktok"] = pick_first_non_empty(row.get("alamat_sosial_media_tiktok", ""), tt_info.get("tiktok_url", ""))
    if is_blocked_tiktok_url(row.get("alamat_sosial_media_tiktok", "")):
        row["alamat_sosial_media_tiktok"] = ""
    row["email"] = pick_first_non_empty(row.get("email", ""), email)
    row["no_hp"] = pick_first_non_empty(row.get("no_hp", ""), phone)
    row["tempat_bekerja"] = pick_first_non_empty(row.get("tempat_bekerja", ""), company)
    row["posisi"] = pick_first_non_empty(row.get("posisi", ""), position)
    row["kategori_pekerjaan"] = pick_first_non_empty(row.get("kategori_pekerjaan", ""), classify_work_type(row.get("posisi", "")))
    row["alamat_sosial_media_tempat_bekerja"] = pick_first_non_empty(row.get("alamat_sosial_media_tempat_bekerja", ""), work_social)
    return row


def run_batch(batch_start, batch_end, rows_done_before, total_target, all_rows, columns, linkedin_map, instagram_map, facebook_map):
    alumni_rows = read_xlsx_slice(EXCEL_PATH, batch_start, batch_end)
    if not alumni_rows:
        return []

    batch_rows = []
    for idx, alumni in enumerate(alumni_rows):
        global_idx = rows_done_before + idx + 1
        row_excel = str(batch_start + idx)
        name = clean_text(str(alumni.get("Nama Lulusan", "")))

        base = merge_base(linkedin_map.get(row_excel, {}), instagram_map.get(row_excel, {}), facebook_map.get(row_excel, {}))
        row = {
            "row_excel": row_excel,
            "nama_lulusan": name,
            "nim": str(alumni.get("NIM", "")),
            "tahun_masuk": str(alumni.get("Tahun Masuk", "")),
            "tanggal_lulus": str(alumni.get("Tanggal Lulus", "")),
            "fakultas": clean_text(str(alumni.get("Fakultas", ""))),
            "program_studi": clean_text(str(alumni.get("Program Studi", ""))),
            "query_dipakai": str(base.get("query_dipakai", "")),
            "mode_pencarian": str(base.get("mode_pencarian", "")),
            "alamat_sosial_media_linkedin": str(base.get("alamat_sosial_media_linkedin", "")),
            "alamat_sosial_media_instagram": str(base.get("alamat_sosial_media_instagram", "")),
            "alamat_sosial_media_facebook": str(base.get("alamat_sosial_media_facebook", "")),
            "alamat_sosial_media_tiktok": str(base.get("alamat_sosial_media_tiktok", "")),
            "email": str(base.get("email", "")),
            "no_hp": str(base.get("no_hp", "")),
            "tempat_bekerja": str(base.get("tempat_bekerja", "")),
            "alamat_bekerja": str(base.get("alamat_bekerja", "")),
            "posisi": str(base.get("posisi", "")),
            "kategori_pekerjaan": str(base.get("kategori_pekerjaan", "")),
            "alamat_sosial_media_tempat_bekerja": str(base.get("alamat_sosial_media_tempat_bekerja", "")),
            "lokasi_terlihat": str(base.get("lokasi_terlihat", "")),
        }

        queries = [
            f"{name} {DEFAULT_UNIVERSITY} tiktok",
            f"{name} tiktok",
        ]
        candidates = []
        top = None
        used_query = ""
        used_engine = ""
        for q in queries:
            for engine_name, engine_template in SEARCH_ENGINES:
                print(f"row={row_excel} [TT] coba query: {q} via {engine_name}")
                candidates = search_tiktok_candidates(q, engine_name, engine_template)
                top = pick_valid_tiktok_candidate(candidates)
                if top:
                    used_query = q
                    used_engine = engine_name
                    break
            if top:
                break
        if not used_query:
            used_query = queries[-1]
            top = None
            used_engine = ""
        row["query_dipakai_tiktok"] = used_query
        row["search_engine_tiktok"] = used_engine

        if top:
            info = read_tiktok_profile(top["tiktok_url"])
            info["tiktok_url"] = top["tiktok_url"]
            row = merge_with_tiktok(row, info)

        if is_blocked_tiktok_url(row.get("alamat_sosial_media_tiktok", "")):
            row["alamat_sosial_media_tiktok"] = ""

        batch_rows.append(row)
        all_rows.append(row)
        print(f"[{global_idx}/{total_target}] row={row_excel} {name}: {row.get('alamat_sosial_media_tiktok', '') or 'tidak ada kandidat'}")

        save_progress(
            {
                "status": "running",
                "start_row": START_ROW,
                "end_row": END_ROW,
                "batch_size": BATCH_SIZE,
                "next_row": int(row_excel) + 1,
                "last_completed_row": int(row_excel),
                "rows_done": global_idx,
            }
        )

    paths = make_output_paths(batch_start, batch_end)
    paths["json"].write_text(json.dumps(batch_rows, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(paths["csv"], columns, batch_rows)
    write_simple_xlsx(paths["xlsx"], columns, batch_rows)
    print(f"saved {paths['json']}")
    print(f"saved {paths['csv']}")
    print(f"saved {paths['xlsx']}")
    return batch_rows


def run_lookup():
    OUT_DIR.mkdir(exist_ok=True)
    linkedin_map = load_fallback_map(LINKEDIN_ALL_FILE, USE_LINKEDIN_FALLBACK)
    instagram_map = load_fallback_map(INSTAGRAM_ALL_FILE, USE_INSTAGRAM_FALLBACK)
    facebook_map = load_fallback_map(FACEBOOK_ALL_FILE, USE_FACEBOOK_FALLBACK)

    actual_start = START_ROW
    if RESUME_ENABLED:
        p = load_progress()
        if p and p.get("status") in {"running", "paused"}:
            same_scope = (
                int(p.get("start_row", -1)) == START_ROW
                and int(p.get("end_row", -1)) == END_ROW
                and int(p.get("batch_size", -1)) == BATCH_SIZE
            )
            if same_scope:
                actual_start = int(p.get("next_row", START_ROW))
                if actual_start < START_ROW:
                    actual_start = START_ROW
                if actual_start > END_ROW:
                    actual_start = END_ROW + 1
                print(f"resume aktif: lanjut dari row {actual_start}")

    if actual_start > END_ROW:
        print("semua row dalam rentang ini sudah selesai diproses")
        return

    total_target = END_ROW - START_ROW + 1
    rows_done_before = max(0, actual_start - START_ROW)
    all_rows = []
    columns = [
        "row_excel",
        "nama_lulusan",
        "nim",
        "tahun_masuk",
        "tanggal_lulus",
        "fakultas",
        "program_studi",
        "query_dipakai",
        "mode_pencarian",
        "query_dipakai_tiktok",
        "search_engine_tiktok",
        "alamat_sosial_media_linkedin",
        "alamat_sosial_media_instagram",
        "alamat_sosial_media_facebook",
        "alamat_sosial_media_tiktok",
        "email",
        "no_hp",
        "tempat_bekerja",
        "alamat_bekerja",
        "posisi",
        "kategori_pekerjaan",
        "alamat_sosial_media_tempat_bekerja",
        "lokasi_terlihat",
    ]

    new_tab("https://www.tiktok.com")
    wait_for_load()
    wait(2)

    batch_start = actual_start
    while batch_start <= END_ROW:
        batch_end = min(batch_start + BATCH_SIZE - 1, END_ROW)
        print(f"batch {batch_start}-{batch_end} mulai")
        batch_rows = run_batch(
            batch_start,
            batch_end,
            rows_done_before,
            total_target,
            all_rows,
            columns,
            linkedin_map,
            instagram_map,
            facebook_map,
        )
        rows_done_before += len(batch_rows)

        all_json = OUT_DIR / f"alumni_tt_all_{START_ROW}_{END_ROW}.json"
        all_csv = OUT_DIR / f"alumni_tt_all_{START_ROW}_{END_ROW}.csv"
        all_xlsx = OUT_DIR / f"alumni_tt_all_{START_ROW}_{END_ROW}.xlsx"
        all_json.write_text(json.dumps(all_rows, ensure_ascii=False, indent=2), encoding="utf-8")
        write_csv(all_csv, columns, all_rows)
        write_simple_xlsx(all_xlsx, columns, all_rows)
        print(f"saved {all_json}")
        print(f"saved {all_csv}")
        print(f"saved {all_xlsx}")

        save_progress(
            {
                "status": "running",
                "start_row": START_ROW,
                "end_row": END_ROW,
                "batch_size": BATCH_SIZE,
                "next_row": batch_end + 1,
                "last_completed_row": batch_end,
                "rows_done": rows_done_before,
            }
        )

        batch_start = batch_end + 1
        if batch_start <= END_ROW:
            wait(random.uniform(BATCH_COOLDOWN_MIN, BATCH_COOLDOWN_MAX))

    save_progress(
        {
            "status": "done",
            "start_row": START_ROW,
            "end_row": END_ROW,
            "batch_size": BATCH_SIZE,
            "next_row": END_ROW + 1,
            "last_completed_row": END_ROW,
            "rows_done": total_target,
        }
    )
    print("semua batch selesai")


run_lookup()

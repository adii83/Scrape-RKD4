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

# Konfigurasi rentang data (1-based, tanpa header)
START_ROW = int(__import__("os").environ.get("BH_START_ROW", "1"))
END_ROW = int(__import__("os").environ.get("BH_END_ROW", "200"))
BATCH_SIZE = int(__import__("os").environ.get("BH_BATCH_SIZE", "100"))

# Jika True: lanjut dari progress terakhir
RESUME_ENABLED = __import__("os").environ.get("BH_RESUME_ENABLED", "1") != "0"

# Jeda acak antar aksi (detik)
SEARCH_DELAY_MIN = float(__import__("os").environ.get("BH_SEARCH_DELAY_MIN", "5.0"))
SEARCH_DELAY_MAX = float(__import__("os").environ.get("BH_SEARCH_DELAY_MAX", "15.0"))
PROFILE_DELAY_MIN = float(__import__("os").environ.get("BH_PROFILE_DELAY_MIN", "2.0"))
PROFILE_DELAY_MAX = float(__import__("os").environ.get("BH_PROFILE_DELAY_MAX", "4.5"))
BATCH_COOLDOWN_MIN = float(__import__("os").environ.get("BH_BATCH_COOLDOWN_MIN", "20.0"))
BATCH_COOLDOWN_MAX = float(__import__("os").environ.get("BH_BATCH_COOLDOWN_MAX", "90.0"))
BLOCK_RECHECK_MIN = float(__import__("os").environ.get("BH_BLOCK_RECHECK_MIN", "120.0"))
BLOCK_RECHECK_MAX = float(__import__("os").environ.get("BH_BLOCK_RECHECK_MAX", "240.0"))

# Nama file state resume
PROGRESS_FILE = OUT_DIR / "alumni_scrape_progress.json"


def search_url(query):
    return "https://www.linkedin.com/search/results/people/?keywords=" + quote(query)


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
    "unusual activity",
    "unusual traffic",
    "too many requests",
    "temporarily restricted",
    "temporarily blocked",
    "access denied",
    "challenge"
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
        print(f"[LI] challenge/rate-limit terdeteksi ({context}), tunggu {int(wait_sec)} detik lalu cek ulang...")
        wait(wait_sec)


def safe_navigate(url, context, min_delay, max_delay):
    while True:
        cdp("Page.navigate", url=url)
        wait_for_load()
        wait(random.uniform(min_delay, max_delay))
        if detect_challenge_or_block():
            wait_until_unblocked(context)
            continue
        return


def clean_text(value):
    return re.sub(r"\s+", " ", value or "").strip()


def clean_linkedin_connection_suffix(value):
    text = clean_text(value)
    text = re.sub(r"\s*[•\-]?\s*\b(?:1st|2nd|3rd|\d+th)\+?\b\s*$", "", text, flags=re.IGNORECASE)
    return clean_text(text)


def clean_position_text(value):
    text = clean_linkedin_connection_suffix(value)
    if re.search(r"^(connect|follow|message|view profile)$", text, flags=re.IGNORECASE):
        return ""
    return text


def extract_work_from_headline(headline):
    text = clean_position_text(headline)
    patterns = [
        r"\bat\s+([^|,]+)",
        r"\bdi\s+([^|,]+)",
        r"@\s*([^|,]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return clean_text(match.group(1))
    return ""


def classify_work_type(headline):
    text = clean_position_text(headline).lower()
    if any(k in text for k in ["pns", "pegawai negeri", "asn", "civil servant"]):
        return "PNS"
    if any(k in text for k in ["owner", "founder", "co-founder", "entrepreneur", "wirausaha"]):
        return "Wirausaha"
    if text:
        return "Swasta"
    return ""


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


def extract_contact_info():
    return js(
        r"""
(() => {
  const textChunks = [];
  const links = [];
  const pushText = (s) => {
    const t = (s || '').trim();
    if (t) textChunks.push(t);
  };

  const bodyText = document.body ? document.body.innerText : '';
  pushText(bodyText);

  const contactBtn = Array.from(document.querySelectorAll('a, button'))
    .find(el => /contact info/i.test((el.getAttribute('aria-label') || '') + ' ' + (el.innerText || '')));

  if (contactBtn) {
    contactBtn.click();
    const start = Date.now();
    while (Date.now() - start < 2500) {}
    const dialog = document.querySelector('[role=\"dialog\"]');
    if (dialog) {
      pushText(dialog.innerText || '');
      Array.from(dialog.querySelectorAll('a[href]')).forEach(a => links.push(a.href));
    }
  }

  return {
    all_text: textChunks.join('\n'),
    links: Array.from(new Set(links)),
  };
})()
"""
    )


def close_contact_dialog():
    try:
        press_key("Escape")
    except Exception:
        pass


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
    csv_lines = [",".join(columns)]
    for row in rows:
        csv_lines.append(",".join(json.dumps(str(row.get(col, "")), ensure_ascii=False) for col in columns))
    path.write_text("\ufeff" + "\n".join(csv_lines), encoding="utf-8")


def linkedin_results():
    return js(
        r"""
(() => {
  const anchors = Array.from(document.querySelectorAll('a[href*="/in/"]'));
  const seen = new Set();
  const out = [];

  for (const a of anchors) {
    const href = (a.href || '').split('?')[0];
    if (!href || seen.has(href)) continue;
    seen.add(href);

    const card =
      a.closest('li') ||
      a.closest('.reusable-search__result-container') ||
      a.closest('[data-view-name]') ||
      a.parentElement;

    const lines = (card?.innerText || a.innerText || '')
      .split('\n')
      .map(s => s.trim())
      .filter(Boolean)
      .filter(s => !/^(Connect|Follow|Message|View|1st|2nd|3rd)$/i.test(s));

    const name = (a.innerText || lines[0] || '')
      .split('\n')[0]
      .replace(/\s*•.*$/, '')
      .trim();
    const headline = lines.find(s =>
      s !== name &&
      !/^Based on your profile/i.test(s) &&
      !/^People you may know/i.test(s) &&
      !/mutual connection/i.test(s)
    ) || '';
    const location = lines.find(s =>
      /(Indonesia|Malang|Jakarta|Surabaya|Sidoarjo|East Java|Jawa)/i.test(s)
    ) || '';

    if (name && !/LinkedIn|Search/i.test(name)) {
      out.push({ name, linkedin_url: href, headline, location, raw_text: lines.slice(0, 8).join(' | ') });
    }
    if (out.length >= 5) break;
  }

  return out;
})()
"""
    )


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
    base = f"alumni_{batch_start}_{batch_end}"
    return {
        "json": OUT_DIR / f"{base}.json",
        "csv": OUT_DIR / f"{base}.csv",
        "xlsx": OUT_DIR / f"{base}.xlsx",
    }


def run_batch(batch_start, batch_end, rows_done_before_batch, total_target, all_rows, columns):
    alumni_rows = read_xlsx_slice(EXCEL_PATH, batch_start, batch_end)
    if not alumni_rows:
        return []

    batch_rows = []
    for index, alumni in enumerate(alumni_rows):
        global_index = rows_done_before_batch + index + 1
        name = clean_text(str(alumni.get("Nama Lulusan", "")))
        attempts = [
            {"mode": "nama + universitas", "query": f"{name} {DEFAULT_UNIVERSITY}"},
            {"mode": "nama saja", "query": name},
        ]
        chosen = []
        used = attempts[0]

        for attempt in attempts:
            safe_navigate(search_url(attempt["query"]), "linkedin search", SEARCH_DELAY_MIN, SEARCH_DELAY_MAX)
            results = linkedin_results() or []
            if results:
                chosen = results
                used = attempt
                break

        first = chosen[0] if chosen else {}
        headline = clean_position_text(first.get("headline", ""))
        company = extract_work_from_headline(headline)
        email = ""
        phone = ""
        company_social = ""
        work_address = ""
        position = headline

        if first.get("linkedin_url"):
            safe_navigate(first["linkedin_url"], "open linkedin profile", PROFILE_DELAY_MIN, PROFILE_DELAY_MAX)
            profile_text = js("document.body ? document.body.innerText : ''") or ""
            if not position:
                position = clean_position_text(js("document.querySelector('h2') ? document.querySelector('h2').innerText : ''") or "")
            contact = extract_contact_info() or {}
            all_text = clean_text((profile_text or "") + "\n" + (contact.get("all_text") or ""))
            links = contact.get("links") or []
            email = first_email(all_text)
            phone = first_phone(all_text)
            company_social = next((u for u in links if "linkedin.com/company/" in u), "")
            close_contact_dialog()

        row = {
            "row_excel": str(batch_start + index),
            "nama_lulusan": name,
            "nim": str(alumni.get("NIM", "")),
            "tahun_masuk": str(alumni.get("Tahun Masuk", "")),
            "tanggal_lulus": str(alumni.get("Tanggal Lulus", "")),
            "fakultas": clean_text(str(alumni.get("Fakultas", ""))),
            "program_studi": clean_text(str(alumni.get("Program Studi", ""))),
            "query_dipakai": used["query"],
            "mode_pencarian": used["mode"],
            "alamat_sosial_media_linkedin": first.get("linkedin_url", ""),
            "email": email,
            "no_hp": phone,
            "tempat_bekerja": company,
            "alamat_bekerja": work_address,
            "posisi": position,
            "kategori_pekerjaan": classify_work_type(headline),
            "alamat_sosial_media_tempat_bekerja": company_social,
            "lokasi_terlihat": first.get("location", ""),
            "kandidat_linkedin": chosen,
        }
        batch_rows.append(row)
        all_rows.append(row)

        print(f"[{global_index}/{total_target}] row={batch_start + index} {name}: {first.get('linkedin_url', 'tidak ada kandidat')}")

        save_progress(
            {
                "status": "running",
                "start_row": START_ROW,
                "end_row": END_ROW,
                "batch_size": BATCH_SIZE,
                "next_row": batch_start + index + 1,
                "last_completed_row": batch_start + index,
                "rows_done": global_index,
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

    actual_start = START_ROW
    if RESUME_ENABLED:
        progress = load_progress()
        if progress and progress.get("status") in {"running", "paused"}:
            same_scope = (
                int(progress.get("start_row", -1)) == START_ROW
                and int(progress.get("end_row", -1)) == END_ROW
                and int(progress.get("batch_size", -1)) == BATCH_SIZE
            )
            if same_scope:
                actual_start = int(progress.get("next_row", START_ROW))
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
        "alamat_sosial_media_linkedin",
        "email",
        "no_hp",
        "tempat_bekerja",
        "alamat_bekerja",
        "posisi",
        "kategori_pekerjaan",
        "alamat_sosial_media_tempat_bekerja",
        "lokasi_terlihat",
    ]

    all_rows = []

    new_tab("https://www.linkedin.com")
    wait_for_load()
    wait(2)

    batch_start = actual_start
    while batch_start <= END_ROW:
        batch_end = min(batch_start + BATCH_SIZE - 1, END_ROW)
        print(f"batch {batch_start}-{batch_end} mulai")

        batch_rows = run_batch(batch_start, batch_end, rows_done_before, total_target, all_rows, columns)
        rows_done_before += len(batch_rows)

        all_json = OUT_DIR / f"alumni_all_{START_ROW}_{END_ROW}.json"
        all_csv = OUT_DIR / f"alumni_all_{START_ROW}_{END_ROW}.csv"
        all_xlsx = OUT_DIR / f"alumni_all_{START_ROW}_{END_ROW}.xlsx"
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

import json
import os
import zipfile
from pathlib import Path


OUT_DIR = Path("outputs")
START_ROW = int(os.environ.get("BH_START_ROW", "1"))
END_ROW = int(os.environ.get("BH_END_ROW", "10"))


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


def load_map(path):
    if not path.exists():
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


def pick_first_non_empty(*values):
    for v in values:
        if str(v or "").strip():
            return str(v).strip()
    return ""


def main():
    li = load_map(OUT_DIR / f"alumni_all_{START_ROW}_{END_ROW}.json")
    ig = load_map(OUT_DIR / f"alumni_ig_all_{START_ROW}_{END_ROW}.json")
    fb = load_map(OUT_DIR / f"alumni_fb_all_{START_ROW}_{END_ROW}.json")
    tt = load_map(OUT_DIR / f"alumni_tt_all_{START_ROW}_{END_ROW}.json")

    rows = []
    for n in range(START_ROW, END_ROW + 1):
        k = str(n)
        a = li.get(k, {})
        b = ig.get(k, {})
        c = fb.get(k, {})
        d = tt.get(k, {})

        row = {
            "row_excel": k,
            "nama_lulusan": pick_first_non_empty(a.get("nama_lulusan"), b.get("nama_lulusan"), c.get("nama_lulusan"), d.get("nama_lulusan")),
            "nim": pick_first_non_empty(a.get("nim"), b.get("nim"), c.get("nim"), d.get("nim")),
            "tahun_masuk": pick_first_non_empty(a.get("tahun_masuk"), b.get("tahun_masuk"), c.get("tahun_masuk"), d.get("tahun_masuk")),
            "tanggal_lulus": pick_first_non_empty(a.get("tanggal_lulus"), b.get("tanggal_lulus"), c.get("tanggal_lulus"), d.get("tanggal_lulus")),
            "fakultas": pick_first_non_empty(a.get("fakultas"), b.get("fakultas"), c.get("fakultas"), d.get("fakultas")),
            "program_studi": pick_first_non_empty(a.get("program_studi"), b.get("program_studi"), c.get("program_studi"), d.get("program_studi")),
            "alamat_sosial_media_linkedin": pick_first_non_empty(a.get("alamat_sosial_media_linkedin"), b.get("alamat_sosial_media_linkedin"), c.get("alamat_sosial_media_linkedin"), d.get("alamat_sosial_media_linkedin")),
            "alamat_sosial_media_instagram": pick_first_non_empty(a.get("alamat_sosial_media_instagram"), b.get("alamat_sosial_media_instagram"), c.get("alamat_sosial_media_instagram"), d.get("alamat_sosial_media_instagram")),
            "alamat_sosial_media_facebook": pick_first_non_empty(a.get("alamat_sosial_media_facebook"), b.get("alamat_sosial_media_facebook"), c.get("alamat_sosial_media_facebook"), d.get("alamat_sosial_media_facebook")),
            "alamat_sosial_media_tiktok": pick_first_non_empty(a.get("alamat_sosial_media_tiktok"), b.get("alamat_sosial_media_tiktok"), c.get("alamat_sosial_media_tiktok"), d.get("alamat_sosial_media_tiktok")),
            "email": pick_first_non_empty(a.get("email"), b.get("email"), c.get("email"), d.get("email")),
            "no_hp": pick_first_non_empty(a.get("no_hp"), b.get("no_hp"), c.get("no_hp"), d.get("no_hp")),
            "tempat_bekerja": pick_first_non_empty(a.get("tempat_bekerja"), b.get("tempat_bekerja"), c.get("tempat_bekerja"), d.get("tempat_bekerja")),
            "alamat_bekerja": pick_first_non_empty(a.get("alamat_bekerja"), b.get("alamat_bekerja"), c.get("alamat_bekerja"), d.get("alamat_bekerja")),
            "posisi": pick_first_non_empty(a.get("posisi"), b.get("posisi"), c.get("posisi"), d.get("posisi")),
            "kategori_pekerjaan": pick_first_non_empty(a.get("kategori_pekerjaan"), b.get("kategori_pekerjaan"), c.get("kategori_pekerjaan"), d.get("kategori_pekerjaan")),
            "alamat_sosial_media_tempat_bekerja": pick_first_non_empty(a.get("alamat_sosial_media_tempat_bekerja"), b.get("alamat_sosial_media_tempat_bekerja"), c.get("alamat_sosial_media_tempat_bekerja"), d.get("alamat_sosial_media_tempat_bekerja")),
            "lokasi_terlihat": pick_first_non_empty(a.get("lokasi_terlihat"), b.get("lokasi_terlihat"), c.get("lokasi_terlihat"), d.get("lokasi_terlihat")),
        }
        rows.append(row)

    columns = [
        "row_excel",
        "nama_lulusan",
        "nim",
        "tahun_masuk",
        "tanggal_lulus",
        "fakultas",
        "program_studi",
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

    out_json = OUT_DIR / f"alumni_master_{START_ROW}_{END_ROW}.json"
    out_xlsx = OUT_DIR / f"alumni_master_{START_ROW}_{END_ROW}.xlsx"
    out_json.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    write_simple_xlsx(out_xlsx, columns, rows)
    print(f"saved {out_json}")
    print(f"saved {out_xlsx}")


if __name__ == "__main__":
    main()

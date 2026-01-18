from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET


_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _col_letters_to_index(col: str) -> int:
    col = col.strip().upper()
    if not col or not re.fullmatch(r"[A-Z]+", col):
        raise ValueError(f"Invalid column letters: {col!r}")
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _index_to_col_letters(idx0: int) -> str:
    if idx0 < 0:
        raise ValueError("idx0 must be >= 0")
    n = idx0 + 1
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def _cell_ref(col0: int, row1: int) -> str:
    return f"{_index_to_col_letters(col0)}{row1}"


def write_simple_xlsx(path: Path, rows: List[List[Any]], sheet_name: str = "Sheet1") -> None:
    """
    Write a minimal .xlsx with inline strings (no styles/sharedStrings) that Excel can open.

    This is intentionally tiny to avoid extra dependencies (e.g. openpyxl).
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    def esc(s: str) -> str:
        return html.escape(s, quote=False)

    # sheet1.xml
    sheet_rows: List[str] = []
    for r_idx0, row in enumerate(rows):
        r1 = r_idx0 + 1
        cells: List[str] = []
        for c_idx0, value in enumerate(row):
            ref = _cell_ref(c_idx0, r1)
            if value is None or value == "":
                continue
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                cells.append(f'<c r="{ref}"><v>{value}</v></c>')
            else:
                text = esc(str(value))
                cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
        if cells:
            sheet_rows.append(f'<row r="{r1}">{"".join(cells)}</row>')
        else:
            sheet_rows.append(f'<row r="{r1}"/>')

    sheet_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{_NS_MAIN}"><sheetData>'
        f'{"".join(sheet_rows)}'
        f"</sheetData></worksheet>"
    )

    # workbook.xml
    sheet_name = (sheet_name or "Sheet1").strip() or "Sheet1"
    sheet_name = sheet_name[:31]  # Excel sheet name limit
    workbook_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{_NS_MAIN}" xmlns:r="{_NS_REL}">'
        f"<sheets>"
        f'<sheet name="{esc(sheet_name)}" sheetId="1" r:id="rId1"/>'
        f"</sheets>"
        f"</workbook>"
    )

    # workbook.xml.rels
    workbook_rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_NS_PKG_REL}">'
        f'<Relationship Id="rId1" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet1.xml"/>'
        f"</Relationships>"
    )

    # package .rels
    rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_NS_PKG_REL}">'
        f'<Relationship Id="rId1" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        f'Target="xl/workbook.xml"/>'
        f"</Relationships>"
    )

    # [Content_Types].xml
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )

    with zipfile.ZipFile(path, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def _xml_text(el: Optional[ET.Element]) -> str:
    if el is None:
        return ""
    return "".join(el.itertext())


def _read_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    name = "xl/sharedStrings.xml"
    if name not in set(zf.namelist()):
        return []
    xml_bytes = zf.read(name)
    root = ET.fromstring(xml_bytes)
    ns = {"m": _NS_MAIN}
    out: List[str] = []
    for si in root.findall("m:si", ns):
        # Either <t> (plain) or rich text <r><t>
        t = si.find("m:t", ns)
        if t is not None:
            out.append(_xml_text(t))
            continue
        parts = []
        for r in si.findall("m:r", ns):
            parts.append(_xml_text(r.find("m:t", ns)))
        out.append("".join(parts))
    return out


def _pick_first_sheet_path(zf: zipfile.ZipFile) -> str:
    candidates = sorted([n for n in zf.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml")])
    if candidates:
        return candidates[0]
    raise ValueError("No worksheets found in xlsx")


def read_simple_xlsx_table(path: Path, max_rows: int = 5000, max_cols: int = 200) -> List[List[str]]:
    """
    Read the first worksheet into a 2D table of strings.

    Supports common Excel outputs (sharedStrings / inlineStr / numeric).
    """
    path = Path(path)
    with zipfile.ZipFile(path, mode="r") as zf:
        shared = _read_shared_strings(zf)
        sheet_path = _pick_first_sheet_path(zf)
        xml_bytes = zf.read(sheet_path)

    root = ET.fromstring(xml_bytes)
    ns = {"m": _NS_MAIN}

    cells: Dict[Tuple[int, int], str] = {}
    max_r = 0
    max_c = 0

    for row in root.findall(".//m:sheetData/m:row", ns):
        for c in row.findall("m:c", ns):
            ref = c.get("r") or ""
            m = re.fullmatch(r"([A-Za-z]+)([0-9]+)", ref)
            if not m:
                continue
            col_letters, row_digits = m.group(1), m.group(2)
            r1 = int(row_digits)
            c0 = _col_letters_to_index(col_letters)
            if r1 <= 0 or c0 < 0:
                continue
            if r1 > max_rows or c0 >= max_cols:
                continue

            t = (c.get("t") or "").strip()
            if t == "s":
                v = _xml_text(c.find("m:v", ns)).strip()
                try:
                    idx = int(v)
                except ValueError:
                    value = ""
                else:
                    value = shared[idx] if 0 <= idx < len(shared) else ""
            elif t == "inlineStr":
                value = _xml_text(c.find("m:is", ns)).strip()
            else:
                value = _xml_text(c.find("m:v", ns)).strip()

            cells[(r1, c0)] = value
            max_r = max(max_r, r1)
            max_c = max(max_c, c0 + 1)

    if max_r <= 0 or max_c <= 0:
        return []

    table: List[List[str]] = []
    for r1 in range(1, min(max_r, max_rows) + 1):
        row_out = []
        for c0 in range(0, min(max_c, max_cols)):
            row_out.append(cells.get((r1, c0), ""))
        table.append(row_out)
    return table


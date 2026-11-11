"""Ubah nilai sel langsung di XML sheet.

Menyimpan ulang IM MASTER dengan openpyxl terbukti merusak gambar dan conditional formatting,
jadi hanya XML sheet tujuan (dan satu atribut di workbook.xml) yang ditulis ulang; entri zip lain
disalin apa adanya.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING, Iterable
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from openpyxl.utils import column_index_from_string, get_column_letter

from .errors import Matrix2ImError
from .xlsx_styles import ATTR_RE, StyleSheet, set_attr

if TYPE_CHECKING:   # hanya anotasi; impor nyata akan melingkar lewat im_layout
    from .write_plan import CellChange, FillChange

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_CELL_OPEN = re.compile(rb"<c\b[^>]*>")
_REF_ATTR = re.compile(rb'(?<![\w:])r="([A-Z]+)(\d+)"')
_S_ATTR = re.compile(rb'(?<![\w:])s="([^"]*)"')
_FORMULA = re.compile(rb"<f[\s>/]")
_SPANS_ATTR = re.compile(rb'\sspans="[^"]*"')
_COLS = re.compile(rb"<cols>(.*?)</cols>", re.S)
_COL = re.compile(rb"<col\b[^>]*>")
_CALC_PR = re.compile(rb"<calcPr\b[^>]*?/?>")
# Elemen yang menurut skema berada sesudah calcPr.
_AFTER_CALC = (b"<oleSize", b"<customWorkbookViews", b"<pivotCaches", b"<smartTagPr",
               b"<smartTagTypes", b"<webPublishing", b"<fileRecoveryPr", b"<webPublishObjects", b"<extLst")


def sheet_xml_path(zf: zipfile.ZipFile, sheet_name: str) -> str:
    """Path XML sheet di dalam zip, lewat workbook.xml dan workbook.xml.rels."""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rid = None
    for sh in wb.iter(f"{_MAIN}sheet"):
        if sh.get("name") == sheet_name:
            rid = sh.get(f"{_REL}id")
            break
    if rid is None:
        raise Matrix2ImError(f"Sheet '{sheet_name}' tidak ada di IM MASTER")
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.iter(f"{_PKG_REL}Relationship"):
        if rel.get("Id") == rid:
            target = rel.get("Target", "")
            return target.lstrip("/") if target.startswith("/") else "xl/" + target
    raise Matrix2ImError(f"Relasi sheet '{sheet_name}' tidak ditemukan")


def styles_xml_path(zf: zipfile.ZipFile) -> str:
    """Path styles.xml di dalam zip, lewat relasi workbook (biasanya xl/styles.xml)."""
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.iter(f"{_PKG_REL}Relationship"):
        if rel.get("Type", "").endswith("/styles"):
            target = rel.get("Target", "")
            return target.lstrip("/") if target.startswith("/") else "xl/" + target
    raise Matrix2ImError("styles.xml tidak ditemukan di IM MASTER")


def _col_styles(data: bytes) -> list[tuple[int, int, int]]:
    """(min, max, style) dari <cols>: format bawaan kolom untuk sel yang tidak ada di XML."""
    end = data.find(b"<sheetData")
    m = _COLS.search(data, 0, end if end >= 0 else len(data))
    if m is None:
        return []
    out = []
    for tag in _COL.findall(m.group(1)):
        a = dict(ATTR_RE.findall(tag))
        if b"style" in a:
            out.append((int(a[b"min"]), int(a[b"max"]), int(a[b"style"])))
    return out


def _cell_style(cell: bytes | None, row_tag: bytes, cols: list[tuple[int, int, int]], col: int) -> int:
    """Indeks format yang berlaku: atribut s sel; sel tak ada di XML ikut gaya baris lalu gaya kolom."""
    if cell is not None:
        m = _S_ATTR.search(_CELL_OPEN.match(cell).group(0))
        return int(m.group(1)) if m else 0
    a = dict(ATTR_RE.findall(row_tag))
    if a.get(b"customFormat") in (b"1", b"true") and b"s" in a:
        return int(a[b"s"])
    return next((st for lo, hi, st in cols if lo <= col <= hi), 0)


def read_fills(path: Path, sheet_name: str, cells: Iterable[tuple[int, int]]) -> dict[tuple[int, int], str | None]:
    """Kunci warna latar (xlsx_styles.fill_key) untuk tiap (baris, kolom) yang diminta."""
    by_row: dict[int, list[int]] = {}
    for r, c in cells:
        by_row.setdefault(r, []).append(c)
    if not by_row:
        return {}
    with zipfile.ZipFile(path) as zf:
        styles = StyleSheet(zf.read(styles_xml_path(zf)))
        data = zf.read(sheet_xml_path(zf, sheet_name))
    cols = _col_styles(data)
    out: dict[tuple[int, int], str | None] = {}
    cursor = 0
    for r in sorted(by_row):
        try:
            start, open_end, selfclosing = _find_row(data, r, cursor)
        except Matrix2ImError:          # baris tak ada di XML: sel ikut gaya kolom
            out.update({(r, c): styles.xf_fill_key(_cell_style(None, b"", cols, c)) for c in by_row[r]})
            continue
        tag = data[start:open_end]
        if selfclosing:
            row_cells, cursor = {}, open_end
        else:
            cursor = data.index(b"</row>", open_end)
            row_cells = dict(_split_cells(data[open_end:cursor]))
        for c in by_row[r]:
            out[(r, c)] = styles.xf_fill_key(_cell_style(row_cells.get(c), tag, cols, c))
    return out


def _col_ref(col: int) -> bytes:
    return get_column_letter(col).encode()


def _cell_xml(ref: bytes, style: bytes | None, text: str | None) -> bytes:
    s = b' s="' + style + b'"' if style is not None else b""
    if text is None:
        return b'<c r="' + ref + b'"' + s + b"/>"
    sp = b' xml:space="preserve"' if (text != text.strip() or "\n" in text or "  " in text) else b""
    return (b'<c r="' + ref + b'"' + s + b' t="inlineStr"><is><t' + sp + b">"
            + escape(text).encode("utf-8") + b"</t></is></c>")


def _find_row(data: bytes, row: int, frm: int = 0) -> tuple[int, int, bool]:
    """(awal tag, akhir tag pembuka, self-closing) untuk <row r="row">; atribut r tak harus pertama.

    Pencarian mulai dari frm (akhir baris sebelumnya) agar total biaya linear; bila tak ketemu
    di sana, ulang dari awal dan terakhir lewat regex.
    """
    needle = b'<row r="%d"' % row
    pos = data.find(needle, frm)
    if pos < 0 and frm:
        pos = data.find(needle)
    if pos < 0:
        m = re.search(rb'<row\b[^>]*?(?<![\w:])r="%d"' % row, data)
        if m is None:
            raise Matrix2ImError(f"Baris {row} tidak ada di sheet; penulisan dibatalkan")
        pos = m.start()
    end = data.index(b">", pos) + 1
    return pos, end, data[end - 2:end] == b"/>"


def _split_cells(inner: bytes) -> list[tuple[int, bytes]]:
    cells = []
    pos = 0
    while True:
        m = _CELL_OPEN.search(inner, pos)
        if m is None:
            return cells
        end = m.end()
        if inner[end - 2:end] != b"/>":
            end = inner.index(b"</c>", end) + 4
        ref = _REF_ATTR.search(m.group(0))
        if ref is None:
            raise Matrix2ImError("Sel tanpa atribut r di XML sheet")
        cells.append((column_index_from_string(ref.group(1).decode()), inner[m.start():end]))
        pos = end


def _rewrite_row(data: bytes, row: int, edits: dict[int, str | None], frm: int = 0,
                 fills: dict[int, str | None] | None = None, styles: StyleSheet | None = None,
                 cols: list[tuple[int, int, int]] | None = None) -> tuple[int, int, bytes]:
    start, open_end, selfclosing = _find_row(data, row, frm)
    tag = data[start:open_end]
    row_tag = tag
    if selfclosing:
        end, inner, tag = open_end, b"", tag[:-2] + b">"
    else:
        close = data.index(b"</row>", open_end)
        end, inner = close + 6, data[open_end:close]
    cells = dict(_split_cells(inner))
    original = dict(cells)
    for col, text in edits.items():
        old = cells.get(col)
        ref = _col_ref(col) + b"%d" % row
        if old is not None:
            head = _CELL_OPEN.match(old).group(0)
            if _FORMULA.search(old):
                raise Matrix2ImError(f"Sel {ref.decode()} berisi rumus; penulisan dibatalkan")
            style = _S_ATTR.search(head)
            cells[col] = _cell_xml(ref, style.group(1) if style else None, text)
        elif text is not None:
            # Sel baru mewarisi gaya baris/kolom (mis. warna satu baris utuh); gaya 0 tanpa atribut s.
            inherited = _cell_style(None, row_tag, cols or [], col)
            cells[col] = _cell_xml(ref, b"%d" % inherited if inherited else None, text)
            tag = _SPANS_ATTR.sub(b"", tag)         # spans lama tak mencakup sel baru
    for col, key in (fills or {}).items():
        # Hanya atribut s yang diganti (juga di sel rumus); format baru = format lama dengan warna lain.
        s = b"%d" % styles.restyle(_cell_style(original.get(col), row_tag, cols or [], col), key)
        cell = cells.get(col)
        if cell is None:
            cells[col] = b'<c r="' + _col_ref(col) + b"%d" % row + b'" s="' + s + b'"/>'
            tag = _SPANS_ATTR.sub(b"", tag)
        else:
            head = _CELL_OPEN.match(cell).group(0)
            cells[col] = set_attr(head, b"s", s) + cell[len(head):]
    return start, end, tag + b"".join(cells[c] for c in sorted(cells)) + b"</row>"


def _set_full_calc(xml: bytes) -> bytes:
    m = _CALC_PR.search(xml)
    if m:
        tag = m.group(0)
        if b"fullCalcOnLoad=" in tag:
            tag = re.sub(rb'fullCalcOnLoad="[^"]*"', b'fullCalcOnLoad="1"', tag)
        else:
            tag = tag.replace(b"<calcPr", b'<calcPr fullCalcOnLoad="1"', 1)
        return xml[:m.start()] + tag + xml[m.end():]
    new = b'<calcPr fullCalcOnLoad="1"/>'
    base = max(xml.find(b"</sheets>"), xml.rfind(b"</definedNames>"), 0)   # abaikan extLst di bookViews
    hits = [i for i in (xml.find(t, base) for t in _AFTER_CALC) if i >= 0]
    at = min(hits) if hits else xml.rindex(b"</workbook>")
    return xml[:at] + new + xml[at:]


def apply_cell_values(src: Path, dst: Path, sheet_name: str, changes: Iterable[CellChange],
                      fills: Iterable[FillChange] = ()) -> None:
    """Salin src ke dst dengan nilai (dan warna latar) sel diganti; dst tidak dibuat bila ada kesalahan.

    Warna menambah format baru di styles.xml (xlsx_styles); tanpa perubahan warna styles.xml tidak disentuh.
    """
    by_row: dict[int, dict[int, str | None]] = {}
    for ch in changes:
        by_row.setdefault(ch.row, {})[ch.col] = ch.new
    fill_rows: dict[int, dict[int, str | None]] = {}
    for f in fills:
        fill_rows.setdefault(f.row, {})[f.col] = f.new
    with zipfile.ZipFile(src) as zin:
        sheet_path = sheet_xml_path(zin, sheet_name)
        data = zin.read(sheet_path)
        styles_path = styles_xml_path(zin) if fill_rows else None
        styles = StyleSheet(zin.read(styles_path)) if styles_path else None
        cols = _col_styles(data)
        patches, cursor = [], 0
        for r in sorted(set(by_row) | set(fill_rows)):
            patch = _rewrite_row(data, r, by_row.get(r, {}), cursor, fill_rows.get(r), styles, cols)
            patches.append(patch)
            cursor = patch[1]
        patches.sort()
        parts, pos = [], 0
        for start, end, new in patches:
            parts += [data[pos:start], new]
            pos = end
        parts.append(data[pos:])
        new_sheet = b"".join(parts)
        new_wb = _set_full_calc(zin.read("xl/workbook.xml"))
        with zipfile.ZipFile(dst, "w") as zout:
            for item in zin.infolist():
                if item.filename == sheet_path:
                    payload = new_sheet
                elif item.filename == "xl/workbook.xml":
                    payload = new_wb
                elif styles is not None and item.filename == styles_path:
                    payload = styles.to_bytes()
                else:
                    payload = zin.read(item.filename)
                zi = zipfile.ZipInfo(item.filename, item.date_time)
                zi.compress_type = item.compress_type
                zi.external_attr = item.external_attr
                zout.writestr(zi, payload)

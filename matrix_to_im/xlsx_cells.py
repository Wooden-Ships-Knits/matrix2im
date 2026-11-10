"""Ubah nilai sel langsung di XML sheet.

Menyimpan ulang IM MASTER dengan openpyxl terbukti merusak gambar dan conditional formatting,
jadi hanya XML sheet tujuan (dan satu atribut di workbook.xml) yang ditulis ulang; entri zip lain
disalin apa adanya.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from openpyxl.utils import column_index_from_string, get_column_letter

from .errors import Matrix2ImError
from .write_plan import CellChange

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_CELL_OPEN = re.compile(rb"<c\b[^>]*>")
_REF_ATTR = re.compile(rb'(?<![\w:])r="([A-Z]+)(\d+)"')
_S_ATTR = re.compile(rb'(?<![\w:])s="([^"]*)"')
_FORMULA = re.compile(rb"<f[\s>/]")
_SPANS_ATTR = re.compile(rb'\sspans="[^"]*"')
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


def _rewrite_row(data: bytes, row: int, edits: dict[int, str | None], frm: int = 0) -> tuple[int, int, bytes]:
    start, open_end, selfclosing = _find_row(data, row, frm)
    tag = data[start:open_end]
    if selfclosing:
        end, inner, tag = open_end, b"", tag[:-2] + b">"
    else:
        close = data.index(b"</row>", open_end)
        end, inner = close + 6, data[open_end:close]
    cells = dict(_split_cells(inner))
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
            cells[col] = _cell_xml(ref, None, text)
            tag = _SPANS_ATTR.sub(b"", tag)         # spans lama tak mencakup sel baru
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


def apply_cell_values(src: Path, dst: Path, sheet_name: str, changes: Iterable[CellChange]) -> None:
    """Salin src ke dst dengan nilai sel diganti; dst tidak dibuat bila ada kesalahan."""
    by_row: dict[int, dict[int, str | None]] = {}
    for ch in changes:
        by_row.setdefault(ch.row, {})[ch.col] = ch.new
    with zipfile.ZipFile(src) as zin:
        sheet_path = sheet_xml_path(zin, sheet_name)
        data = zin.read(sheet_path)
        patches, cursor = [], 0
        for r in sorted(by_row):
            patch = _rewrite_row(data, r, by_row[r], cursor)
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
                else:
                    payload = zin.read(item.filename)
                zi = zipfile.ZipInfo(item.filename, item.date_time)
                zi.compress_type = item.compress_type
                zi.external_attr = item.external_attr
                zout.writestr(zi, payload)

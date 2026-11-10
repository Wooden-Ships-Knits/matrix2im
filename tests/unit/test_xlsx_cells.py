"""Penulisan sel langsung di XML sheet: bagian lain workbook tidak boleh tersentuh."""
from __future__ import annotations

import re
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.im_lock import check_modify_password
from matrix_to_im.legend_config import IM_HEADER_ROW, IM_SHEET
from matrix_to_im.write_plan import CellChange
from matrix_to_im.xlsx_cells import apply_cell_values, sheet_xml_path
from tests.helpers.im_factory import TEST_PASSWORD, GroupSpec, add_file_sharing, build_im_sheet

R1 = IM_HEADER_ROW + 1


@pytest.fixture
def im(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("S1", [["RED", "BLUE"], ["GREEN"]])])
    add_file_sharing(p)
    return p


def _sheet_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        return zf.read(sheet_xml_path(zf, IM_SHEET)).decode("utf-8")


def test_values_written_and_readable(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", "NAVY"), CellChange(R1, 7, "BLUE", None)])
    ws = load_workbook(dst)[IM_SHEET]
    assert ws.cell(R1, 5).value == "NAVY"
    assert ws.cell(R1, 7).value is None
    assert ws.cell(R1 + 1, 5).value == "GREEN"


def test_other_parts_byte_identical(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", "NAVY")])
    with zipfile.ZipFile(im) as a, zipfile.ZipFile(dst) as b:
        skip = {sheet_xml_path(a, IM_SHEET), "xl/workbook.xml"}
        assert [i.filename for i in a.infolist()] == [i.filename for i in b.infolist()]
        for ia, ib in zip(a.infolist(), b.infolist()):
            assert ia.compress_type == ib.compress_type
            if ia.filename not in skip:
                assert ia.CRC == ib.CRC


def test_style_preserved(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    before = load_workbook(im)[IM_SHEET].cell(R1, 5)._style
    apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", "NAVY")])
    xml = _sheet_xml(dst)
    s = re.search(rf'<c [^>]*r="E{R1}"[^>]*>', xml).group(0)
    assert re.search(r'\bs="\d+"', s)
    ws = load_workbook(dst)[IM_SHEET]
    assert ws.cell(R1, 5).fill.fgColor.rgb == "FFFFFF00"
    assert ws.cell(R1, 5)._style == before


def test_missing_cell_inserted_in_column_order(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = IM_SHEET
    ws["A5"] = "a"
    ws["Z5"] = "z"
    p = tmp_path / "x.xlsx"
    wb.save(p)
    dst = tmp_path / "out.xlsx"
    apply_cell_values(p, dst, IM_SHEET, [CellChange(5, 13, None, "m")])
    row = re.search(r'<row r="5".*?</row>', _sheet_xml(dst), re.S).group(0)
    assert re.findall(r'<c [^>]*?r="([A-Z]+)5"', row) == ["A", "M", "Z"]
    assert load_workbook(dst)[IM_SHEET]["M5"].value == "m"


def test_clear_keeps_style_removes_value(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", None)])
    xml = _sheet_xml(dst)
    m = re.search(rf'<c [^>]*r="E{R1}"[^>]*?(/>|>.*?</c>)', xml)
    assert m.group(1) == "/>"
    assert re.search(r'\bs="\d+"', m.group(0))
    assert load_workbook(dst)[IM_SHEET].cell(R1, 5).value is None


def test_formula_cell_refused_and_no_output(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    with pytest.raises(Matrix2ImError, match=r"Sel D%d berisi rumus; penulisan dibatalkan" % R1):
        apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", "NAVY"), CellChange(R1, 4, "x", "y")])
    assert not dst.exists()


def test_special_chars_and_unicode(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    text = 'A & B <C> "D" APRÈS'
    apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", text), CellChange(R1, 7, "BLUE", " pad ")])
    ws = load_workbook(dst)[IM_SHEET]
    assert ws.cell(R1, 5).value == text
    assert ws.cell(R1, 7).value == " pad "


def test_000_stays_text(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", "000")])
    v = load_workbook(dst)[IM_SHEET].cell(R1, 5).value
    assert v == "000" and isinstance(v, str)


def test_full_calc_on_load_set(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", "NAVY")])
    with zipfile.ZipFile(dst) as zf:
        wbxml = zf.read("xl/workbook.xml").decode()
    assert re.search(r'<calcPr\b[^>]*fullCalcOnLoad="1"', wbxml)


def test_full_calc_on_load_created_when_no_calcpr(tmp_path):
    p = tmp_path / "x.xlsx"
    wb = Workbook()
    wb.active.title = IM_SHEET
    wb.active["A1"] = "a"
    wb.save(p)
    stripped = tmp_path / "s.xlsx"
    with zipfile.ZipFile(p) as a, zipfile.ZipFile(stripped, "w", zipfile.ZIP_DEFLATED) as b:
        for i in a.infolist():
            d = a.read(i.filename)
            if i.filename == "xl/workbook.xml":
                d = re.sub(rb"<calcPr[^>]*/>", b"", d)
                assert b"calcPr" not in d
            b.writestr(i, d)
    dst = tmp_path / "out.xlsx"
    apply_cell_values(stripped, dst, IM_SHEET, [CellChange(1, 1, "a", "b")])
    with zipfile.ZipFile(dst) as zf:
        assert b'<calcPr fullCalcOnLoad="1"/>' in zf.read("xl/workbook.xml")
    assert load_workbook(dst)[IM_SHEET]["A1"].value == "b"


def test_file_sharing_preserved_and_password_still_valid(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", "NAVY")])
    check_modify_password(dst, TEST_PASSWORD)


def test_missing_row_raises(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    with pytest.raises(Matrix2ImError):
        apply_cell_values(im, dst, IM_SHEET, [CellChange(999, 5, None, "x")])
    assert not dst.exists()


def test_shared_string_and_empty_styled_and_attr_order(tmp_path):
    """Sel t="s", sel kosong bergaya, dan atribut r bukan yang pertama."""
    p = tmp_path / "x.xlsx"
    wb = Workbook()
    wb.active.title = IM_SHEET
    wb.active["A1"] = "shared"
    wb.save(p)
    hacked = tmp_path / "h.xlsx"
    with zipfile.ZipFile(p) as a, zipfile.ZipFile(hacked, "w", zipfile.ZIP_DEFLATED) as b:
        for i in a.infolist():
            d = a.read(i.filename)
            if i.filename == "xl/worksheets/sheet1.xml":
                d = re.sub(rb"<row [^>]*>.*?</row>",
                           b'<row r="1"><c s="0" t="s" r="A1"><v>0</v></c><c s="0" r="B1"/></row>', d, flags=re.S)
            b.writestr(i, d)
    dst = tmp_path / "out.xlsx"
    apply_cell_values(hacked, dst, IM_SHEET, [CellChange(1, 1, "shared", "n1"), CellChange(1, 2, None, "n2")])
    ws = load_workbook(dst)[IM_SHEET]
    assert (ws["A1"].value, ws["B1"].value) == ("n1", "n2")


def test_full_calc_not_placed_inside_bookviews_extlst(tmp_path):
    p = tmp_path / "x.xlsx"
    wb = Workbook()
    wb.active.title = IM_SHEET
    wb.active["A1"] = "a"
    wb.save(p)
    hacked = tmp_path / "h.xlsx"
    with zipfile.ZipFile(p) as a, zipfile.ZipFile(hacked, "w", zipfile.ZIP_DEFLATED) as b:
        for i in a.infolist():
            d = a.read(i.filename)
            if i.filename == "xl/workbook.xml":
                d = re.sub(rb"<calcPr[^>]*/>", b"", d)
                d = re.sub(rb"<workbookView([^>]*)/>", rb"<workbookView\1><extLst><ext uri='x'/></extLst></workbookView>", d)
                assert b"<extLst>" in d
            b.writestr(i, d)
    dst = tmp_path / "out.xlsx"
    apply_cell_values(hacked, dst, IM_SHEET, [CellChange(1, 1, "a", "b")])
    with zipfile.ZipFile(dst) as zf:
        wbxml = zf.read("xl/workbook.xml").decode()
    assert wbxml.index("</bookViews>") < wbxml.index("<calcPr")
    ET.fromstring(wbxml)


def test_row_search_is_linear(tmp_path):
    """Banyak baris berubah: tiap pencarian <row> mulai dari baris sebelumnya, bukan dari awal."""
    n = 5000
    wb = Workbook()
    ws = wb.active
    ws.title = IM_SHEET
    for r in range(1, n + 1):
        ws.cell(r, 1, "x")
    p = tmp_path / "x.xlsx"
    wb.save(p)
    import time
    t0 = time.perf_counter()
    apply_cell_values(p, tmp_path / "o.xlsx", IM_SHEET,
                      [CellChange(r, c, None, "v") for r in range(1, n + 1) for c in (2, 3, 4)])
    assert time.perf_counter() - t0 < 10
    assert load_workbook(tmp_path / "o.xlsx")[IM_SHEET].cell(n, 4).value == "v"

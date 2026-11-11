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


def test_row_search_is_linear(tmp_path, monkeypatch):
    """Tiap pencarian <row> mulai di/after akhir baris sebelumnya, bukan dari offset 0."""
    from matrix_to_im import xlsx_cells
    n = 300
    wb = Workbook()
    ws = wb.active
    ws.title = IM_SHEET
    for r in range(1, n + 1):
        ws.cell(r, 1, "x")
    p = tmp_path / "x.xlsx"
    wb.save(p)
    calls = []
    real = xlsx_cells._find_row

    def spy(data, row, frm=0):
        res = real(data, row, frm)
        calls.append((row, frm, res[1]))
        return res

    monkeypatch.setattr(xlsx_cells, "_find_row", spy)
    apply_cell_values(p, tmp_path / "o.xlsx", IM_SHEET,
                      [CellChange(r, c, None, "v") for r in range(n, 0, -1) for c in (2, 3)])
    assert [c[0] for c in calls] == list(range(1, n + 1))
    prev_end = 0
    for row, frm, end in calls:
        assert frm >= prev_end, f"baris {row}: pencarian mulai di {frm}, sebelum akhir baris sebelumnya {prev_end}"
        prev_end = end
    assert load_workbook(tmp_path / "o.xlsx")[IM_SHEET].cell(n, 3).value == "v"


# ---- warna latar (fill) ----

from openpyxl.styles import Font, PatternFill  # noqa: E402

from matrix_to_im.write_plan import FillChange  # noqa: E402
from matrix_to_im.xlsx_cells import read_fills  # noqa: E402
from matrix_to_im.xlsx_styles import check_appended  # noqa: E402


def _fill_wb(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = IM_SHEET
    ws["X5"] = "T"
    ws["X5"].fill = PatternFill("solid", fgColor="FFFF99FF")
    ws["X6"] = "T"
    ws["X7"] = "T"
    ws["X7"].fill = PatternFill("solid", fgColor="FFFFCC99")
    ws["X7"].font = Font(bold=True)
    ws["A8"] = "a"                                   # X8 tidak ada di XML
    ws["A9"] = "a"
    ws.row_dimensions[9].fill = PatternFill("solid", fgColor="FF969696")   # gaya baris (customFormat)
    ws["X11"] = "=1+1"
    p = tmp_path / "f.xlsx"
    wb.save(p)
    return p


def test_read_fills_cell_row_style_and_missing(tmp_path):
    p = _fill_wb(tmp_path)
    cells = [(5, 24), (6, 24), (7, 24), (8, 24), (9, 24), (10, 24)]
    assert read_fills(p, IM_SHEET, cells) == {(5, 24): "FF99FF", (6, 24): None, (7, 24): "FFCC99",
                                              (8, 24): None, (9, 24): "969696", (10, 24): None}
    assert read_fills(p, IM_SHEET, []) == {}


def test_fills_written_values_kept_other_parts_untouched(tmp_path):
    p = _fill_wb(tmp_path)
    dst = tmp_path / "out.xlsx"
    fills = [FillChange(5, 24, "FF99FF", "FFCC99"), FillChange(6, 24, None, "969696"),
             FillChange(7, 24, "FFCC99", None), FillChange(8, 24, None, "FF99FF"),
             FillChange(9, 24, "969696", None), FillChange(11, 24, None, "FF99FF")]
    apply_cell_values(p, dst, IM_SHEET, [CellChange(6, 24, "T", "U")], fills)
    assert read_fills(dst, IM_SHEET, [(f.row, f.col) for f in fills]) == {(f.row, f.col): f.new for f in fills}
    ws = load_workbook(dst)[IM_SHEET]
    assert (ws["X5"].value, ws["X6"].value, ws["X7"].value, ws["X8"].value) == ("T", "U", "T", None)
    assert ws["X7"].font.b and ws["X11"].value == "=1+1"
    with zipfile.ZipFile(p) as a, zipfile.ZipFile(dst) as b:
        check_appended(a.read("xl/styles.xml"), b.read("xl/styles.xml"))
        changed = {n for n in a.namelist() if a.read(n) != b.read(n)}
        assert {sheet_xml_path(a, IM_SHEET), "xl/styles.xml"} <= changed
        assert changed <= {sheet_xml_path(a, IM_SHEET), "xl/workbook.xml", "xl/styles.xml"}


def test_fill_on_cell_with_same_style_reuses_format(tmp_path):
    p = _fill_wb(tmp_path)
    one, two = tmp_path / "1.xlsx", tmp_path / "2.xlsx"
    apply_cell_values(p, one, IM_SHEET, [], [FillChange(5, 24, "FF99FF", "FFCC99")])
    apply_cell_values(one, two, IM_SHEET, [], [FillChange(6, 24, None, "FF99FF"),
                                               FillChange(5, 24, "FFCC99", "FF99FF")])
    with zipfile.ZipFile(one) as a, zipfile.ZipFile(two) as b:
        xfs = [re.search(rb'<cellXfs count="(\d+)"', z.read("xl/styles.xml")).group(1) for z in (a, b)]
    assert xfs[0] == xfs[1]                          # format X5 asli dipakai lagi, tidak menumpuk


def test_no_fills_leaves_styles_untouched(im, tmp_path):
    dst = tmp_path / "out.xlsx"
    apply_cell_values(im, dst, IM_SHEET, [CellChange(R1, 5, "RED", "NAVY")], [])
    with zipfile.ZipFile(im) as a, zipfile.ZipFile(dst) as b:
        assert a.read("xl/styles.xml") == b.read("xl/styles.xml")


def test_value_into_absent_cell_keeps_row_style(tmp_path):
    """Sel baru (tak ada di XML) mewarisi gaya baris/kolom; tanpa itu warna baris hilang diam-diam."""
    p = _fill_wb(tmp_path)
    dst = tmp_path / "out.xlsx"
    apply_cell_values(p, dst, IM_SHEET, [CellChange(9, 24, None, "BARU"), CellChange(8, 24, None, "LAIN")])
    assert read_fills(dst, IM_SHEET, [(9, 24), (8, 24)]) == {(9, 24): "969696", (8, 24): None}
    ws = load_workbook(dst)[IM_SHEET]
    assert (ws["X9"].value, ws["X8"].value) == ("BARU", "LAIN")
    assert re.search(r'<c r="X8" t="inlineStr">', _sheet_xml(dst))      # gaya 0 tetap tanpa atribut s

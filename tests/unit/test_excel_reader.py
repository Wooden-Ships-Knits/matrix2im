import pytest
from openpyxl import Workbook
from openpyxl.styles import Border, Color, Font, PatternFill, Side

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.excel_reader import (fill_key, font_rgb, is_black_line, is_red_line,
                                       line_above, norm, open_workbooks)


def _cell():
    return Workbook().active["A1"]


def test_fill_key_rgb_theme_tint_none():
    c = _cell()
    assert fill_key(c) is None
    c.fill = PatternFill("solid", fgColor="FFFF66FF")
    assert fill_key(c) == "FFFF66FF"
    c.fill = PatternFill("solid", fgColor=Color(theme=0, tint=-0.1499984740745262))
    assert fill_key(c) == "theme0-0.15"
    c.fill = PatternFill("solid", fgColor=Color(indexed=64))
    assert fill_key(c) == "indexed64"


def test_font_rgb_only_for_rgb_type():
    c = _cell()
    assert font_rgb(c) is None
    c.font = Font(color="FF00B050")
    assert font_rgb(c) == "FF00B050"
    c.font = Font(color=Color(theme=1))
    assert font_rgb(c) is None
    c.font = Font(color=Color(indexed=8))
    assert font_rgb(c) is None


def test_is_red_line():
    for style, rgb in (("medium", "FFFF0000"), ("thick", "FFE40500"), ("double", "FFC00000")):
        assert is_red_line(Side(style=style, color=rgb))
    assert not is_red_line(Side(style="thin", color="FFFF0000"))
    assert not is_red_line(Side(style="medium", color="FF00B050"))
    assert not is_red_line(None)


def test_is_black_line():
    assert is_black_line(Side(style="medium"))
    assert is_black_line(Side(style="medium", color=Color(auto=True)))
    assert is_black_line(Side(style="medium", color=Color(indexed=64)))
    assert is_black_line(Side(style="medium", color=Color(indexed=8)))
    assert is_black_line(Side(style="medium", color=Color(theme=1)))
    assert is_black_line(Side(style="medium", color="FF000000"))
    assert not is_black_line(Side(style="thin", color="FF000000"))
    assert not is_black_line(Side(style="medium", color="FFFF0000"))
    assert not is_black_line(None)


def test_line_above_top_or_bottom_of_previous_row():
    ws = Workbook().active
    red = Side(style="medium", color="FFFF0000")
    assert not line_above(ws, 5, 1, is_red_line)
    ws.cell(5, 1).border = Border(top=red)
    assert line_above(ws, 5, 1, is_red_line)
    ws2 = Workbook().active
    ws2.cell(4, 1).border = Border(bottom=red)
    assert line_above(ws2, 5, 1, is_red_line)
    assert not line_above(ws2, 1, 1, is_red_line)


def test_open_workbooks_missing_file_raises(tmp_path):
    with pytest.raises(Matrix2ImError, match="tidak bisa dibaca"):
        open_workbooks(tmp_path / "tidak_ada.xlsx")


def test_open_workbooks_not_xlsx_raises(tmp_path):
    p = tmp_path / "palsu.xlsx"
    p.write_text("bukan excel")
    with pytest.raises(Matrix2ImError, match="tidak bisa dibaca"):
        open_workbooks(p)


def test_open_workbooks_returns_normal_and_data_only(tmp_path):
    p = tmp_path / "a.xlsx"
    wb = Workbook()
    wb.active["A1"] = 1
    wb.active["A2"] = "=A1+1"
    wb.save(p)
    normal, data_only = open_workbooks(p)
    assert normal.active["A2"].value == "=A1+1"
    assert data_only.active["A2"].value is None  # tidak ada cache nilai


def test_norm_collapses_spaces_keeps_newline():
    assert norm("TAG  for\xa0Heathered:\nC1 ") == "TAG for Heathered:\nC1"

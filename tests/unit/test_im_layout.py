import pytest

from helpers.im_factory import GroupSpec, build_im_sheet
from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.im_layout import normalize, read_layout


def _cw(n):
    return [[f"C{i}"] for i in range(n)]


def test_columns_by_header_name(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("A", _cw(1))])
    cols = read_layout(p).columns
    assert cols.description == 1
    assert cols.colors[0] == 5
    assert cols.colors[3] == 9
    assert len(cols.colors) == 18
    assert cols.ws_tag == 24
    assert cols.whs == 25
    assert cols.last == 26


def test_multiline_header_matched(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("A", _cw(1))], extra_header={"X": "WS TAG COLOR\n(ket)"})
    assert read_layout(p).columns.ws_tag == 24


def test_missing_header_raises(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("A", _cw(1))], extra_header={"X": "LAIN"})
    with pytest.raises(Matrix2ImError, match="WS TAG COLOR"):
        read_layout(p)


def test_whs_not_right_of_ws_tag_raises(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("A", _cw(1))], extra_header={"Y": "LAIN"})
    with pytest.raises(Matrix2ImError, match="tepat di kanan WS TAG COLOR"):
        read_layout(p)


def test_groups_and_subtotals(tmp_path):
    groups = [
        GroupSpec("AAA", _cw(3)),
        GroupSpec("BBB", _cw(1)),
        GroupSpec("CCC", _cw(2), subtotal=None),
        GroupSpec("S - M/L", _cw(2), subtotal="BOTH WHS - FAF"),
        GroupSpec("S - M/L", _cw(2), subtotal="BOTH WHS - FAH"),
    ]
    lay = read_layout(build_im_sheet(tmp_path / "im.xlsx", groups))
    got = [(g.description, g.rows, g.subtotal_row) for g in lay.groups]
    assert got == [
        ("AAA", (55, 56, 57), 58),
        ("BBB", (59,), 60),
        ("CCC", (61, 62), None),
        ("S - M/L", (63, 64), 65),
        ("S - M/L", (66, 67), 68),
    ]


def test_normalize():
    assert normalize("Reegan  cable - x/s ") == "REEGAN CABLE - X/S"
    assert normalize(None) == ""


def test_formula_in_target_kept_as_text(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("A", [["x", "=E55"]])])
    lay = read_layout(p)
    assert lay.values[(55, 7)] == "=E55"
    assert lay.values[(55, 8)] is None
    assert (55, 2) not in lay.values


def test_subtotal_formulas_captured(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("A", _cw(3))])
    lay = read_layout(p)
    assert lay.subtotal_formulas[58][2] == "=SUBTOTAL(9,B55:B57)"

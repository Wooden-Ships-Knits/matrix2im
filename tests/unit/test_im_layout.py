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


def test_subtotal_detected_by_formula_not_label(tmp_path):
    """R14: label subtotal di file nyata bervariasi; hanya rumus SUBTOTAL yang menentukan."""
    groups = [
        GroupSpec("AAA - K", _cw(2), subtotal_label="AAA - K -BOTH WHS"),
        GroupSpec("AAA - L", _cw(1), subtotal_label="AAA - L - SY ONLY"),
        GroupSpec("AAA - M", _cw(1), subtotal_label="AAA - M-  SHOPIFY ONLY - FAH\n(New design)"),
    ]
    lay = read_layout(build_im_sheet(tmp_path / "im.xlsx", groups))
    assert [(g.description, g.rows, g.subtotal_row) for g in lay.groups] == [
        ("AAA - K", (55, 56), 57), ("AAA - L", (58,), 59), ("AAA - M", (60,), 61)]
    assert set(lay.subtotal_formulas) == {57, 59, 61}


def test_label_without_subtotal_formula_is_data_row(tmp_path):
    groups = [GroupSpec("BBB - K - BOTH WHS - ALT 1", _cw(2), subtotal=None)]
    lay = read_layout(build_im_sheet(tmp_path / "im.xlsx", groups))
    assert [(g.description, g.rows, g.subtotal_row) for g in lay.groups] == [
        ("BBB - K", (55, 56), None)]    # R15: kunci grup = STYLE - SIZE
    assert lay.subtotal_formulas == {}


def test_lowercase_subtotal_function_counts(tmp_path):
    groups = [GroupSpec("CCC - K", _cw(1), subtotal_func="subtotal")]
    lay = read_layout(build_im_sheet(tmp_path / "im.xlsx", groups))
    assert lay.groups[0].subtotal_row == 56


def _summary(path):
    return [(g.description, g.rows, g.subtotal_row) for g in read_layout(path).groups]


def test_annotated_first_row_stays_in_group(tmp_path):
    """R15: kunci grup = 'STYLE - SIZE'; teks sesudahnya hanya catatan."""
    g = GroupSpec("AAA - K", _cw(3), row_descriptions=["AAA - K - BOTH WHS - ALT 1", None, None])
    assert _summary(build_im_sheet(tmp_path / "im.xlsx", [g])) == [("AAA - K", (55, 56, 57), 58)]


def test_two_different_annotations_one_group(tmp_path):
    g = GroupSpec("AAA - K", _cw(4), row_descriptions=["AAA - K - FAF - ALT 1", "AAA - K - FAF - ALT 2 (INSET)", None, None])
    assert _summary(build_im_sheet(tmp_path / "im.xlsx", [g])) == [("AAA - K", (55, 56, 57, 58), 59)]


def test_size_change_without_subtotal_splits(tmp_path):
    groups = [GroupSpec("AAA - K", _cw(2), subtotal=None), GroupSpec("AAA - L", _cw(2), subtotal=None)]
    assert _summary(build_im_sheet(tmp_path / "im.xlsx", groups)) == [("AAA - K", (55, 56), None),
                                                                      ("AAA - L", (57, 58), None)]


def test_one_segment_label_is_own_group(tmp_path):
    groups = [GroupSpec("GRASSY", _cw(1), subtotal=None), GroupSpec("AAA - K", _cw(1), subtotal=None)]
    assert _summary(build_im_sheet(tmp_path / "im.xlsx", groups)) == [("GRASSY", (55,), None),
                                                                      ("AAA - K", (56,), None)]


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


def _layout_with_formula(tmp_path, formula, label):
    from openpyxl import load_workbook
    path = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("AAA - K", _cw(2), subtotal=None)])
    wb = load_workbook(path)
    wb.active.cell(57, 1, label)
    wb.active.cell(57, 2, formula)
    wb.save(path)       # fixture sintetis (R2): boleh memakai openpyxl
    return read_layout(path)


def test_subtotal_formula_with_whitespace_counts(tmp_path):
    assert _layout_with_formula(tmp_path, "= SUBTOTAL(9,B55:B56)", "apa saja").groups[0].subtotal_row == 57


def test_blank_label_subtotal_still_closes_group(tmp_path):
    lay = _layout_with_formula(tmp_path, "=SUBTOTAL(9,B55:B56)", None)
    assert lay.groups[0].subtotal_row == 57 and 57 in lay.subtotal_formulas


def test_wrapped_subtotal_is_not_a_subtotal_row(tmp_path):
    lay = _layout_with_formula(tmp_path, "=ROUND(SUBTOTAL(9,B55:B56),0)", "AAA - K - BOTH WHS")
    assert lay.subtotal_formulas == {} and lay.groups[0].subtotal_row is None


def test_ws_tag_and_whs_fills_read_for_group_rows(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("S - X/S", [["A"], ["B"], ["C"]],
                                                        ws_tag_fills=["FF99FF", None, "969696"],
                                                        whs_fills=[None, "FFCC99", "969696"])])
    lay = read_layout(p)
    assert {k: v for k, v in lay.fills.items() if k[1] in (24, 25)} == {
        (55, 24): "FF99FF", (56, 24): None, (57, 24): "969696",
        (55, 25): None, (56, 25): "FFCC99", (57, 25): "969696"}


def test_color_fills_read_for_group_rows(tmp_path):
    """Warna COLOR1..18 dibaca untuk tanda benang final (7F7F7F); pabrik IM memberi COLOR1 (kolom E) latar kuning."""
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("S - X/S", [["A", "B"]])])
    lay = read_layout(p)
    assert lay.fills.get((55, 5)) == "FFFF00"
    assert (55, 7) in lay.fills and lay.fills[(55, 7)] is None   # COLOR2 tanpa warna
    assert all((55, c) in lay.fills for c in lay.columns.colors)

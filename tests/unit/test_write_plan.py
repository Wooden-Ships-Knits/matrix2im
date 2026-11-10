import pytest

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.im_layout import ImColumns, ImGroup, ImLayout
from matrix_to_im.write_plan import (
    Colorway, InsertOp, Skip, match_style, parse_ref, plan_inserts, project_layout, style_colorways,
)


def _rec(style, colors, ref, ws=None, whs=None):
    rec = {"DESCRIPTION": style}
    for i, c in enumerate(colors, start=1):
        rec[f"COLOR{i}"] = c
    rec["WS TAG COLOR"] = ws
    rec["WHS"] = whs
    rec["_ref"] = ref
    return rec


def _final(*records):
    return {"records": list(records)}


def test_parse_ref_with_and_without_section():
    assert parse_ref("ESSENTIALS GRASSY No.20 | B15 | above_black") == (
        "ESSENTIALS GRASSY", "ESSENTIALS GRASSY No.20")
    assert parse_ref("LUXE (WINTER) No.S9 - A | C16 | below_red") == (
        "LUXE", "LUXE (WINTER) No.S9 - A")


def test_parse_ref_invalid_raises():
    with pytest.raises(Matrix2ImError):
        parse_ref("sembarang")


def test_split_cards_same_sheet_concatenated():
    final = _final(
        _rec("HUTTON", ["A"], "SH No.S2 - A | B1 | g"),
        _rec("HUTTON", ["B"], "SH No.S2 - A | B2 | g"),
        _rec("HUTTON", ["C"], "SH No.S2 - B | B3 | g"),
    )
    out, skips = style_colorways(final)
    assert skips == []
    assert [c.colors for c in out["HUTTON"]] == [("A",), ("B",), ("C",)]


def test_cross_sheet_same_set_uses_first_order_and_unions_whs():
    final = _final(
        _rec("HUTTON", ["N"], "S1 No.1 | B1 | g", whs="000"),
        _rec("HUTTON", ["P"], "S1 No.1 | B2 | g", whs="SY"),
        _rec("HUTTON", ["G"], "S1 No.1 | B3 | g", whs="000"),
        _rec("HUTTON", ["G"], "S2 No.2 | B1 | g", whs="000 & SY"),
        _rec("HUTTON", ["N"], "S2 No.2 | B2 | g", whs="000"),
        _rec("HUTTON", ["P"], "S2 No.2 | B3 | g", whs="SY"),
    )
    out, skips = style_colorways(final)
    assert skips == []
    assert [c.colors[0] for c in out["HUTTON"]] == ["N", "P", "G"]
    assert [c.whs for c in out["HUTTON"]] == ["000", "SY", "000 & SY"]


def test_cross_sheet_different_set_skipped():
    final = _final(
        _rec("HUTTON", ["A"], "S1 No.1 | B1 | g"),
        _rec("HUTTON", ["B"], "S2 No.2 | B1 | g"),
        _rec("HUTTON", ["C"], "S2 No.2 | B2 | g"),
    )
    out, skips = style_colorways(final)
    assert "HUTTON" not in out
    assert len(skips) == 1 and skips[0].style == "HUTTON"
    assert "berbeda" in skips[0].reason
    assert "S1 1 colorway" in skips[0].reason and "S2 2 colorway" in skips[0].reason


def test_cross_sheet_ws_tag_differs_skipped():
    final = _final(
        _rec("HUTTON", ["A"], "S1 No.1 | B1 | g", ws="X"),
        _rec("HUTTON", ["A"], "S2 No.2 | B1 | g", ws="Y"),
    )
    out, skips = style_colorways(final)
    assert "HUTTON" not in out
    assert "WS TAG COLOR berbeda" in skips[0].reason


def test_whs_none_wins():
    final = _final(
        _rec("HUTTON", ["A"], "S1 No.1 | B1 | g", whs=None),
        _rec("HUTTON", ["A"], "S2 No.2 | B1 | g", whs="000"),
    )
    out, _ = style_colorways(final)
    assert out["HUTTON"][0].whs is None


def test_color_case_insensitive_match():
    final = _final(
        _rec("HUTTON", ["Glasgow"], "S1 No.1 | B1 | g"),
        _rec("HUTTON", ["GLASGOW"], "S2 No.2 | B1 | g"),
    )
    out, skips = style_colorways(final)
    assert skips == []
    assert out["HUTTON"] == [Colorway(("Glasgow",), None, None)]


def test_style_key_is_normalized():
    out, _ = style_colorways(_final(_rec(" hutton  x ", ["A"], "S1 No.1 | B1 | g")))
    assert "HUTTON X" in out


# ---- rencana sisip & proyeksi ----

_COLS = ImColumns(description=1, colors=(2,), ws_tag=3, whs=4, last=4)


def _cws(n):
    return [Colorway((f"C{i}",), None, None) for i in range(n)]


def _layout(groups, values=None, formulas=None):
    return ImLayout(_COLS, groups, values or {}, formulas or {})


def test_no_insert_when_enough_rows():
    lay = _layout([ImGroup("S - X/S", (55, 56, 57), 58)])
    assert plan_inserts(lay, {"S": _cws(3)}) == []
    assert plan_inserts(lay, {"S": _cws(2)}) == []
    assert plan_inserts(lay, {"Z": _cws(9)}) == []


def test_insert_inside_group_two_or_more_rows():
    lay = _layout([ImGroup("S - X/S", (55, 56, 57), 58)])
    assert plan_inserts(lay, {"S": _cws(5)}) == [
        InsertOp("S - X/S", at_row=57, count=2, fill_from=56, fill_to=58,
                 subtotal_row_after=None, formula_fixes={})]


def test_single_row_group_inserts_below_and_widens_subtotal():
    lay = _layout([ImGroup("S - X/S", (60,), 61)],
                  formulas={61: {2: "=SUBTOTAL(9,B60)", 3: "=B600*2", 4: "=SUBTOTAL(9,$B$60)",
                                 5: "=SUBTOTAL(9,B60:B60)", 6: "=SUBTOTAL(9,B59:B60)"}})
    (op,) = plan_inserts(lay, {"S": _cws(3)})
    assert (op.at_row, op.count, op.fill_from, op.fill_to, op.subtotal_row_after) == (61, 2, 60, 62, 63)
    assert op.formula_fixes == {2: "=SUBTOTAL(9,B60:B62)", 4: "=SUBTOTAL(9,$B$60:$B$62)",
                                5: "=SUBTOTAL(9,B60:B62)"}


def test_single_row_group_without_subtotal():
    lay = _layout([ImGroup("S - X/S", (60,), None)])
    (op,) = plan_inserts(lay, {"S": _cws(2)})
    assert op.subtotal_row_after is None and op.formula_fixes == {}
    assert (op.at_row, op.count) == (61, 1)


def test_ops_sorted_bottom_up():
    lay = _layout([ImGroup("A - X", (55, 56), 57), ImGroup("B - X", (58, 59), 60)])
    ops = plan_inserts(lay, {"A": _cws(3), "B": _cws(4)})
    assert [o.at_row for o in ops] == [59, 56]


def test_match_style_longest_prefix():
    styles = {"AMELIA ROLLNECK", "AMELIA ROLLNECK CARDI"}
    assert match_style("AMELIA ROLLNECK CARDI - X/S", styles) == "AMELIA ROLLNECK CARDI"
    assert match_style("AMELIA ROLLNECK - X/S", styles) == "AMELIA ROLLNECK"
    assert match_style("OTHER - X/S", styles) is None


def test_project_layout_shifts_rows_and_copies_values():
    lay = _layout(
        [ImGroup("A - X", (55, 56), 57), ImGroup("B - X", (58, 59), 60)],
        values={(55, 2): "a", (56, 2): "b", (58, 2): "c", (59, 2): "d"},
        formulas={57: {2: "=SUBTOTAL(9,B55:B56)"}, 60: {2: "=SUBTOTAL(9,B58:B59)"}})
    ops = plan_inserts(lay, {"A": _cws(4), "B": _cws(3)})
    out = project_layout(lay, ops)
    assert out.groups == [ImGroup("A - X", (55, 56, 57, 58), 59), ImGroup("B - X", (60, 61, 62), 63)]
    assert out.values[(55, 2)] == "a"
    assert out.values[(56, 2)] == "a" and out.values[(57, 2)] == "a"  # sisipan = salinan fill_from
    assert out.values[(58, 2)] == "b"
    assert out.values[(60, 2)] == "c" and out.values[(61, 2)] == "c"  # sisipan = salinan baris c
    assert out.values[(62, 2)] == "d"
    assert set(out.subtotal_formulas) == {59, 63}

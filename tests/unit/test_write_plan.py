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
        _rec("HUTTON", ["G"], "S2 No.2 | B1 | g", whs="BOTH WHS"),
        _rec("HUTTON", ["N"], "S2 No.2 | B2 | g", whs="000"),
        _rec("HUTTON", ["P"], "S2 No.2 | B3 | g", whs="SY"),
    )
    out, skips = style_colorways(final)
    assert skips == []
    assert [c.colors[0] for c in out["HUTTON"]] == ["N", "P", "G"]
    assert [c.whs for c in out["HUTTON"]] == ["000", "SY", "BOTH WHS"]


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


# --- plan_values -----------------------------------------------------------

from matrix_to_im.write_plan import CellChange, plan_values  # noqa: E402

_VCOLS = ImColumns(description=1, colors=tuple(range(2, 20)), ws_tag=20, whs=21, last=21)


def _vlayout(groups, values=None):
    return ImLayout(_VCOLS, groups, values or {}, {})


def _grp(desc, rows, subtotal="auto"):
    """Grup dengan baris subtotal tepat di bawahnya (bentuk normal); subtotal=None untuk grup tanpa subtotal."""
    return ImGroup(desc, tuple(rows), rows[-1] + 1 if subtotal == "auto" else subtotal)


def _cw(colors, ws=None, whs=None):
    return Colorway(tuple(colors), ws, whs)


def _by_cell(plan):
    return {(c.row, c.col): c for c in plan.changes}


def test_rows_get_colorways_in_order():
    lay = _vlayout([_grp("S - X/S", [5, 6])])
    plan = plan_values(lay, {"S": [_cw(["A"]), _cw(["B"])]})
    got = _by_cell(plan)
    assert [got[(r, 2)].new for r in (5, 6)] == ["A", "B"]
    assert plan.short_groups == [] and plan.extra_rows == []


def test_extra_rows_clear_ws_tag_and_whs_keep_colors():
    """Baris grup melebihi jumlah colorway MATRIX: WS TAG COLOR dan WHS dikosongkan, COLOR tidak disentuh
    (keputusan user 11 Nov 2026; sebelumnya diisi colorway terakhir)."""
    vals = {(7, 2): "OLD - 1", (7, 20): "OLD TAG", (7, 21): "000", (8, 2): "OLD - 2", (8, 21): "SY"}
    lay = _vlayout([_grp("S - X/S", [5, 6, 7, 8])], vals)
    plan = plan_values(lay, {"S": [_cw(["A"], "TA", "000"), _cw(["B"], "TB", "SY")]})
    got = _by_cell(plan)
    assert [got[(r, 2)].new for r in (5, 6)] == ["A", "B"]
    assert not {(r, c) for r, c in got if r in (7, 8) and c < 20}
    assert got[(7, 20)] == CellChange(7, 20, "OLD TAG", None)
    assert got[(7, 21)].new is None and got[(8, 21)].new is None
    assert (8, 20) not in got
    assert plan.extra_rows == [7, 8]
    assert plan.changed_colorway_rows == [5, 6]


def test_extra_row_formula_not_cleared():
    lay = _vlayout([_grp("S - X/S", [5, 6])], {(6, 20): "=T5"})
    plan = plan_values(lay, {"S": [_cw(["A"], "TA", "000")]})
    assert (6, 20) not in _by_cell(plan)
    assert plan.warnings == ["sel T6 berisi rumus, tidak ditulis"]


def test_color_slots_overwrite_clear_old_yarn_keep_accessory():
    old = ["YCA BLACK B - 1708 B", "YCA CHOCOLATE ALE - 609", "XBP 729 2 CM-DARK ROAST", "YCA OLD - 1"]
    lay = _vlayout([_grp("S - X/S", [5])], {(5, 2 + i): v for i, v in enumerate(old)})
    got = _by_cell(plan_values(lay, {"S": [_cw(["YCA KHAKI - 749"])]}))
    assert got[(5, 2)].new == "YCA KHAKI - 749"
    assert got[(5, 3)].new is None and got[(5, 3)].old == old[1]
    assert (5, 4) not in got
    assert got[(5, 5)].new is None
    got3 = _by_cell(plan_values(lay, {"S": [_cw(["A - 1", "B - 2", "C - 3"])]}))
    assert got3[(5, 4)].new == "C - 3"


def test_more_than_18_yarns_truncated():
    lay = _vlayout([_grp("S - X/S", [5])])
    got = _by_cell(plan_values(lay, {"S": [_cw([f"Y{i}" for i in range(20)])]}))
    assert max(c for _, c in got) == 19 and got[(5, 19)].new == "Y17"


def test_need_review_never_replaces_filled_ws_tag():
    cws = {"S": [_cw(["A"], ws="Need Review")]}
    lay = _vlayout([_grp("S - X/S", [5])], {(5, 20): "BLACK MARL"})
    assert (5, 20) not in _by_cell(plan_values(lay, cws))
    got = _by_cell(plan_values(_vlayout([_grp("S - X/S", [5])]), cws))
    assert got[(5, 20)].new == "Need Review"


def test_null_values_not_written():
    lay = _vlayout([_grp("S - X/S", [5])], {(5, 20): "KEEP", (5, 21): "SY"})
    got = _by_cell(plan_values(lay, {"S": [_cw(["A"])]}))
    assert (5, 20) not in got and (5, 21) not in got


def test_unchanged_values_produce_no_change():
    lay = _vlayout([_grp("S - X/S", [5])], {(5, 2): "A", (5, 20): "T", (5, 21): "SY"})
    plan = plan_values(lay, {"S": [_cw(["A"], "T", "SY")]})
    assert plan.changes == [] and plan.changed_colorway_rows == []


def test_formula_target_cell_untouched_with_warning():
    lay = _vlayout([_grp("S - X/S", [5])], {(5, 2): "=B1"})
    plan = plan_values(lay, {"S": [_cw(["A"])]})
    assert plan.changes == []
    assert plan.warnings == ["sel B5 berisi rumus, tidak ditulis"]


def test_short_group_fills_existing_rows_and_reports():
    lay = _vlayout([_grp("S - X/S", [5, 6])])
    plan = plan_values(lay, {"S": [_cw(["A"]), _cw(["B"]), _cw(["C"]), _cw(["D"])]})
    assert plan.short_groups == ["S - X/S: kurang 2 baris"]
    got = _by_cell(plan)
    assert [got[(r, 2)].new for r in (5, 6)] == ["A", "B"]


def test_unmatched_and_missing_styles_reported():
    lay = _vlayout([_grp("ZED - X/S", [5]), _grp("ZED - S/M", [6]), _grp("S - X/S", [7])])
    plan = plan_values(lay, {"S": [_cw(["A"])], "GONE": [_cw(["B"])]})
    assert plan.unmatched_im_styles == ["ZED"]
    assert plan.missing_styles == ["GONE"]


def test_changed_colorway_rows_listed():
    lay = _vlayout([_grp("S - X/S", [5, 6])],
                   {(5, 2): "a  - 1", (5, 3): "XZIP", (6, 2): "OLD - 2"})
    plan = plan_values(lay, {"S": [_cw(["A - 1"]), _cw(["NEW - 3"])]})
    assert plan.changed_colorway_rows == [6]


def test_whs_000_is_string():
    lay = _vlayout([_grp("S - X/S", [5])])
    c = _by_cell(plan_values(lay, {"S": [_cw(["A"], whs="000")]}))[(5, 21)]
    assert c.new == "000" and isinstance(c.new, str)


def test_unmatched_excludes_one_segment_labels():
    """Label musim/judul (DESCRIPTION tanpa " - ") bukan style; tidak dihitung sebagai style IM tanpa MATRIX."""
    lay = _vlayout([_grp("FALL", [4]), _grp("ZED - X/S", [5]), _grp("S - X/S", [7])])
    plan = plan_values(lay, {"S": [_cw(["A"])]})
    assert plan.unmatched_im_styles == ["ZED"]


def test_warns_matched_group_without_subtotal():
    lay = _vlayout([_grp("S - X/S", [5, 6], subtotal=None), _grp("ZED - X/S", [9], subtotal=None)])
    plan = plan_values(lay, {"S": [_cw(["A"])]})
    assert plan.warnings == ["S - X/S (baris 5): grup tanpa baris subtotal"]


def test_warns_same_key_in_separate_parts_of_sheet():
    lay = _vlayout([_grp("S - X/S", [5]), _grp("S - S/M", [7]), _grp("S - X/S", [20, 21]),
                    _grp("ZED - X/S", [30]), _grp("ZED - S/M", [31]), _grp("ZED - X/S", [32])])
    plan = plan_values(lay, {"S": [_cw(["A"])]})
    assert plan.warnings == ["S - X/S: muncul di 2 bagian sheet; semua bagian diisi colorway yang sama"]
    assert {c.row for c in plan.changes if c.col == 2} == {5, 7, 20}
    assert plan.extra_rows == [21]


def test_adjacent_blocks_of_same_key_are_one_part():
    """Dua blok M/L berurutan (subtotal FAF lalu FAH) adalah bentuk normal F27: tidak diperingatkan."""
    pair = [_grp("S - S/M", [3]), _grp("S - M/L", [5, 6]), _grp("S - M/L", [8, 9]), _grp("S - X/L", [11])]
    assert plan_values(_vlayout(pair), {"S": [_cw(["A"])]}).warnings == []
    again = _vlayout([*pair, _grp("T - M/L", [13]), _grp("S - M/L", [20])])
    plan = plan_values(again, {"S": [_cw(["A"])], "T": [_cw(["B"])]})
    assert plan.warnings == ["S - M/L: muncul di 2 bagian sheet; semua bagian diisi colorway yang sama"]


def test_record_without_color_skips_style():
    """Record tanpa COLOR akan mengosongkan benang lama; style itu dilewati, bukan ditebak."""
    final = _final(
        _rec("AAA", ["A"], "S1 No.1 | B1 | g"),
        _rec("AAA", [], "S1 No.1 | B2 | g"),
        _rec("BBB", ["B"], "S1 No.2 | B3 | g"),
    )
    out, skips = style_colorways(final)
    assert "AAA" not in out and list(out) == ["BBB"]
    assert skips == [Skip("AAA", "record tanpa COLOR")]


def test_old_whs_label_rejected():
    """final.json lama berlabel "000 & SY" tidak boleh menulis label lama ke IM; minta buat ulang."""
    final = _final(_rec("HUTTON", ["A"], "S1 No.1 | B1 | g", whs="000 & SY"))
    with pytest.raises(Matrix2ImError, match=r"WHS '000 & SY'.*buat ulang final\.json"):
        style_colorways(final)

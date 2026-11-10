import pytest

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.write_plan import Colorway, Skip, parse_ref, style_colorways


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

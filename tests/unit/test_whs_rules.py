import pytest

from matrix_to_im.review import Code
from matrix_to_im.whs_rules import color_whs, final_whs


@pytest.mark.parametrize("group,font,want", [
    ("below_red", "NEW COLOR", "SY"), ("below_red", "000 ONLY COLOR", "SY"),
    ("above_black", "000 ONLY COLOR", "000"), ("between_black_red", "000 REPEAT COLOR", "000"),
    ("above_black", "NEW COLOR", "BOTH WHS"), ("between_black_red", "REPEAT COLOR", "BOTH WHS")])
def test_color_whs(group, font, want):
    assert color_whs(group, font) == (want, None)


@pytest.mark.parametrize("font", ["YR COLOR", "HOLD COLOR FOR SY", "DISCONTINUED COLOR FOR SY", "? (FF123456)"])
@pytest.mark.parametrize("group", ["above_black", "below_red"])
def test_color_whs_no_rule(group, font):
    v, p = color_whs(group, font)
    assert v is None and p.code is Code.FONT_COLOR_NO_RULE


def test_final_whs_intersection():
    assert final_whs("BOTH WHS", "BOTH WHS") == ("BOTH WHS", None)
    assert final_whs("SY", "SY") == ("SY", None)
    assert final_whs("000", "BOTH WHS") == ("000", None)      # dibatasi style: aturan biasa, tanpa review
    assert final_whs("SY", "BOTH WHS") == ("SY", None)
    v, p = final_whs("SY", "000")
    assert v is None and p.code is Code.WHS_CONFLICT
    v, p = final_whs(None, "000")
    assert v == "000" and p.code is Code.WHS_STYLE_UNKNOWN
    assert final_whs("000", None) == (None, None)

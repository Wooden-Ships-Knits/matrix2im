import pytest

from matrix_to_im.review import Code, Problem
from matrix_to_im.tag_rules import build_ws_tag_color, color_name, unique_yarns


def test_color_name():
    assert color_name("YCA KHAKI - 749") == "KHAKI"
    assert color_name("YCA BLACK B - 1708 B") == "BLACK"
    assert color_name("YCA BLACK B -1708 B") == "BLACK"
    assert color_name("bj red ginger") == "RED GINGER"
    assert color_name("POODLE OATMEAL - 12") == "OATMEAL"
    assert color_name("DARK DUSK") == "DARK DUSK"


def test_unique_yarns_case_insensitive_first_occurrence():
    assert unique_yarns(["HUSKY BEIGE", "PURE SNOW", "RED GINGER", "pure snow", "PURE SNOW"]) == ["HUSKY BEIGE", "PURE SNOW", "RED GINGER"]


def test_tag_substitution_marl():
    assert build_ws_tag_color("TAG: C1/C2 MARL", ["YCA KHAKI - 749", "YCA BLUSH CHAMPAGNE - 645"]) == ("KHAKI/BLUSH CHAMPAGNE MARL", None)


@pytest.mark.parametrize("tag", ["TAG: C1 (For Solid) | TAG: SNOW MARL", "TAG: C1 for marled",
                                 "TAG: C1/C2/C3 with abbreviated color names",
                                 "TAG: BLACK/LATTE/ RED | TAG: RED PAPRIKA/SHADLE/ KALE",
                                 "TAG: C1 (For Solid) | TAG: C1/C2 (For abbreviated)"])
def test_tag_conditional(tag):
    v, p = build_ws_tag_color(tag, ["A", "B", "C"])
    assert v is None and p.code is Code.TAG_CONDITIONAL


def test_tag_cn_out_of_range():
    v, p = build_ws_tag_color("TAG: C1/C2/C3", ["A", "B"])
    assert v is None and p.code is Code.TAG_CN_OUT_OF_RANGE and p.issue == "TAG minta C3 tapi colorway hanya punya 2 benang"


def test_tag_empty():
    assert build_ws_tag_color(None, ["A"])[1].code is Code.TAG_EMPTY


def test_tag_striped_and_double_suffix():
    assert build_ws_tag_color("TAG: C1 STRIPE", ["BLUSH CHAMPAGNE STRIPE"]) == ("BLUSH CHAMPAGNE STRIPE", None)
    assert build_ws_tag_color("TAG: C1 STRIPED", ["NAVY"]) == ("NAVY STRIPE", None)
    assert build_ws_tag_color("TAG: C1 MARL", ["OYSTER MARL"]) == ("OYSTER MARL", None)


def test_tag_without_cn_and_spacing():
    assert build_ws_tag_color("TAG: BLACK/LATTE/ RED", ["X"]) == ("BLACK/LATTE/RED", None)
    assert build_ws_tag_color("TAG C1", ["YCA SIENNA - 1"]) == ("SIENNA", None)


def test_c10_not_confused_with_c1():
    yarns = [f"Y{i}" for i in range(1, 11)]
    assert build_ws_tag_color("TAG: C10", yarns) == ("Y10", None)


@pytest.mark.parametrize("tag", ["TAG:", "TAG", "TAG: -", "TAG : / ."])
def test_tag_without_letters_or_digits_is_empty(tag):     # F1
    assert build_ws_tag_color(tag, ["A", "B"]) == (None, Problem(Code.TAG_EMPTY, "TAG kosong"))


def test_tag_c0_out_of_range():                            # F2
    v, p = build_ws_tag_color("TAG: C0", ["A", "B"])
    assert v is None and p == Problem(Code.TAG_CN_OUT_OF_RANGE, "TAG minta C0 tapi colorway hanya punya 2 benang")


@pytest.mark.parametrize("tag,yarns", [("TAG: C1/C(car)", ["FAWN"]), ("TAG: C1/GONDOLA color", ["SILVER MOON"]),
                                       ("TAG: C1 colour", ["A"])])
def test_tag_with_untranslated_instruction_is_unresolved(tag, yarns):
    value, problem = build_ws_tag_color(tag, yarns)
    assert value is None and problem.code is Code.TAG_UNRESOLVED
    assert problem.issue.startswith("TAG masih berisi instruksi yang belum diterjemahkan")


def test_yarn_name_containing_color_is_not_unresolved():
    # Pengecekan dilakukan pada teks TAG sebelum Cn diganti, jadi nama benang tidak ikut diperiksa.
    assert build_ws_tag_color("TAG: C1", ["MULTI COLOR"]) == ("MULTI COLOR", None)



# ---- TAG bersyarat (keputusan user 11 Nov 2026): solid = 1 benang, selain itu bukan solid ----

HUTTON = "TAG: C1 (For Solid) | TAG: SNOW MARL (For PURE SNOW/ALMOND BUTTER)"
SHEPHERD = "TAG: C1 (For Solid) TAG: C1 MARL (For Marled)"
ARIA = "TAG for Heathered: C1 HEATHERED TAG For solids: C1"


@pytest.mark.parametrize("tag, yarns, want", [
    (HUTTON, ["YCA NIGHT NAVY - 123"], "NIGHT NAVY"),
    (HUTTON, ["YCA PURE SNOW - 1", "YCA ALMOND BUTTER - 2"], "SNOW MARL"),
    (SHEPHERD, ["FAWN"], "FAWN"),
    (SHEPHERD, ["FAWN", "BLACK"], "FAWN MARL"),
    (ARIA, ["SILVER PELT"], "SILVER PELT"),
    (ARIA, ["YCA DARK DUSK - 899", "BJ BLUE FIRE"], "DARK DUSK HEATHERED"),
])
def test_conditional_tag_resolved_by_solid_or_colorway(tag, yarns, want):
    assert build_ws_tag_color(tag, yarns) == (want, None)


def test_conditional_no_matching_clause_needs_review():
    v, p = build_ws_tag_color(HUTTON, ["BLACK", "RED"])          # bukan solid, bukan PURE SNOW/ALMOND BUTTER
    assert v is None and p.code is Code.TAG_CONDITIONAL
    assert p.issue == "TAG bersyarat: tidak ada syarat yang cocok untuk colorway ini (bukan solid, 2 benang)"


def test_conditional_colorway_clause_beats_solid_clause():
    tag = "TAG: C1 (For Solid) | TAG: SNOW (For PURE SNOW)"
    assert build_ws_tag_color(tag, ["YCA PURE SNOW - 1"]) == ("SNOW", None)
    assert build_ws_tag_color(tag, ["FAWN"]) == ("FAWN", None)


def test_conditional_colorway_clause_needs_exact_yarn_order():
    v, p = build_ws_tag_color(HUTTON, ["ALMOND BUTTER", "PURE SNOW"])
    assert v is None and p.code is Code.TAG_CONDITIONAL


def test_conditional_two_clauses_same_condition_is_ambiguous():
    v, p = build_ws_tag_color("TAG: C1 (For Solid) | TAG: C1 MARL (For Solid)", ["A"])
    assert v is None and p.code is Code.TAG_CONDITIONAL


def test_conditional_chosen_instruction_follows_normal_rules():
    v, p = build_ws_tag_color(SHEPHERD.replace("C1 MARL", "C3 MARL"), ["A", "B"])
    assert v is None and p.code is Code.TAG_CN_OUT_OF_RANGE

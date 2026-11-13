import json

import pytest

from matrix_to_im.transform import transform

pytestmark = pytest.mark.regression


@pytest.fixture(scope="module")
def f27_final(f27_raw):
    return transform(f27_raw)


def test_records_identical_to_golden(f27_final, golden_final):
    assert json.dumps(f27_final["records"], ensure_ascii=False) == json.dumps(golden_final["records"],
                                                                              ensure_ascii=False)


def test_needs_review_matches_golden(f27_final, golden_final):
    got = [(x["record"], x["code"]) for x in f27_final["needs_review"]]
    want = [(x["record"], x["code"]) for x in golden_final["needs_review"]]
    assert got == want


def test_baseline(f27_final):
    s = f27_final["summary"]
    assert s["total_records"] == 724 and s["total_styles"] == 211
    assert s["WHS"] == {"BOTH WHS": 194, "000": 354, "SY": 176}
    # TAG bersyarat diputuskan per colorway sejak 11 Nov 2026: 45 record keluar dari Need Review.
    # Benang berulang dipertahankan sejak 13 Nov 2026 (Cn = benang ke-n): 21 TAG_CN_OUT_OF_RANGE terselesaikan.
    assert (s["WS_TAG_COLOR_filled"], s["WS_TAG_COLOR_need_review"], s["WS_TAG_COLOR_null"]) == (699, 25, 0)
    assert s["needs_review_items"] == 25
    assert s["needs_review_by_issue"] == {"TAG_CONDITIONAL": 12, "TAG_UNRESOLVED": 7, "COLOR_IS_DESCRIPTOR": 6}


def _items(final, code):
    return [x for x in final["needs_review"] if x["code"] == code]


def _review_records(final):
    return {x["record"] for x in final["needs_review"]}


def test_need_review_only_where_reviewed(f27_final):
    marked = {i for i, r in enumerate(f27_final["records"]) if r["WS TAG COLOR"] == "Need Review"}
    assert len(marked) == 25 and marked == _review_records(f27_final)


def test_hutton_in_two_sheets_are_separate_records_without_review(f27_final):
    refs = [r["_ref"] for r in f27_final["records"] if r["DESCRIPTION"] == "HUTTON MOCK T NECK CHUNKY"]
    assert len(refs) == 16
    assert sum(r.startswith("ESSENTIALS GRASSY No.20") for r in refs) == 8
    assert sum(r.startswith("LUXE No.11") for r in refs) == 8


def test_sloane_capped_to_000_without_review(f27_final):
    for i in range(130, 135):
        assert f27_final["records"][i]["WHS"] == "000"
    assert not _review_records(f27_final) & set(range(130, 135))


def test_rihanna_striped_descriptor_is_need_review(f27_final):
    rec = f27_final["records"][167]
    assert rec["COLOR1"] == "BLUSH CHAMPAGNE STRIPE" and rec["WS TAG COLOR"] == "Need Review"
    assert [x["record"] for x in _items(f27_final, "COLOR_IS_DESCRIPTOR")] == [167, 188, 189, 206, 207, 379]


def test_merry_striped_whs_without_review(f27_final):
    assert f27_final["records"][668]["WHS"] == "BOTH WHS" and f27_final["records"][669]["WHS"] == "000"
    assert _items(f27_final, "WHS_STYLE_UNKNOWN") == []


def test_unresolved_tags_tree_farm_and_telluride(f27_final):
    items = _items(f27_final, "TAG_UNRESOLVED")
    assert [x["record"] for x in items] == [588, 589, 706, 707, 708, 709, 710]
    assert {x["tag_instruction"] for x in items} == {"TAG: C1/C(car)", "TAG: C1/GONDOLA color"}


def test_colorway_count_equals_records(f27_raw, f27_final):
    n = sum(len(s[g]) for s in f27_raw["styles"] for g in
            ("color_above_black_line", "color_between_black_and_red_line", "color_below_red_line"))
    assert n == len(f27_final["records"])


def test_validate_f27_against_im_header(f27_raw, f27_final, im_path):
    from matrix_to_im.validate import read_im_header, validate
    assert validate(f27_raw, f27_final, read_im_header(im_path)) == []



def test_conditional_tags_resolved_by_solid_rule(f27_final):
    """Keputusan user 11 Nov 2026: (For Solid) = 1 benang, (For Marled/Heathered) = lebih dari 1 benang."""
    tag = {r["_ref"].split(" | ")[0] + " " + r["_ref"].split(" | ")[1]: r["WS TAG COLOR"] for r in f27_final["records"]}
    assert tag["ESSENTIALS GRASSY No.20 C164"] == "SNOW MARL"          # (For PURE SNOW/ALMOND BUTTER)
    assert tag["ESSENTIALS GRASSY No.20 C163"] == "NIGHT NAVY"         # solid
    assert tag["ESSENTIALS COTTON No.15 G105"] == "DARK DUSK HEATHERED"  # sama dengan comment sel di MATRIX
    assert tag["ESSENTIALS COTTON No.15 G106"] == "SILVER PELT"
    assert [tag[f"FALL No.26 E22{i}"] for i in (1, 2, 3)] == ["PURE SNOW", "WINTER COAL", "FAWN"]
    left = {x["DESCRIPTION"] for x in _items(f27_final, "TAG_CONDITIONAL")}
    assert left == {"CANDY CANES HEART CREW CHUNKY", "CHEERS COCKTAIL CREW CHUNKY", "MERRY COCKTAIL CREW CHUNKY",
                    "FA LA LA FAIR ISLE CREW CHUNKY"}


def test_repeated_yarns_kept_and_long_colorways(f27_final):
    """Keputusan user 13 Nov 2026: benang berulang dipertahankan (urutan = urutan warna)."""
    recs = f27_final["records"]
    assert recs[452]["WS TAG COLOR"] == "BLACK/BROWN BRICK MARL"           # TAG: C1/C3 MARL, C1 = C2
    lengths = [sum(k.startswith("COLOR") for k in r) for r in recs]
    assert max(lengths) == 32 and sum(n > 18 for n in lengths) == 8

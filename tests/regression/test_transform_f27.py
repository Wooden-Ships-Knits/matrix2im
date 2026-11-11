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
    assert (s["WS_TAG_COLOR_filled"], s["WS_TAG_COLOR_need_review"], s["WS_TAG_COLOR_null"]) == (633, 91, 0)
    assert s["needs_review_items"] == 91
    assert s["needs_review_by_issue"] == {"TAG_CONDITIONAL": 57, "TAG_CN_OUT_OF_RANGE": 21,
                                          "TAG_UNRESOLVED": 7, "COLOR_IS_DESCRIPTOR": 6}


def _items(final, code):
    return [x for x in final["needs_review"] if x["code"] == code]


def _review_records(final):
    return {x["record"] for x in final["needs_review"]}


def test_need_review_only_where_reviewed(f27_final):
    marked = {i for i, r in enumerate(f27_final["records"]) if r["WS TAG COLOR"] == "Need Review"}
    assert len(marked) == 91 and marked == _review_records(f27_final)


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

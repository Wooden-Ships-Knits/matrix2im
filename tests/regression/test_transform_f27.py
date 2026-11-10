import json

import pytest

from matrix_to_im.transform import transform

pytestmark = pytest.mark.regression

GOLDEN_ISSUE_TO_CODE = [("TAG kondisional", "TAG_CONDITIONAL"), ("TAG minta C", "TAG_CN_OUT_OF_RANGE"),
                        ("info:", "WHS_CAPPED_BY_STYLE"), ("COLOR berisi deskripsi", "COLOR_IS_DESCRIPTOR"),
                        ("style muncul di 2 sheet", "DUPLICATE_STYLE_ACROSS_SHEETS"),
                        ("WHS style tidak diketahui", "WHS_STYLE_UNKNOWN"),
                        ("warning ekstraksi", "EXTRACTION_WARNING")]


@pytest.fixture(scope="module")
def f27_final(f27_raw):
    return transform(f27_raw)


def golden_code(issue: str) -> str:
    for prefix, code in GOLDEN_ISSUE_TO_CODE:
        if issue.startswith(prefix):
            return code
    raise AssertionError(f"issue golden tak dikenal: {issue}")


def test_records_identical_to_golden(f27_final, golden_final):
    assert json.dumps(f27_final["records"], ensure_ascii=False) == json.dumps(golden_final["records"],
                                                                              ensure_ascii=False)


def test_needs_review_matches_golden(f27_final, golden_final):
    got = [(x["record"], x["code"]) for x in f27_final["needs_review"]]
    want = [(x["record"], golden_code(x["issue"])) for x in golden_final["needs_review"]]
    assert got == want


def test_baseline(f27_final):
    s = f27_final["summary"]
    assert s["total_records"] == 724 and s["total_styles"] == 211
    assert s["WHS"] == {"000 & SY": 194, "000": 354, "SY": 176}
    assert (s["WS_TAG_COLOR_filled"], s["WS_TAG_COLOR_null"]) == (646, 78)
    assert s["needs_review_items"] == 115
    assert s["needs_review_by_issue"] == {"TAG_CONDITIONAL": 57, "TAG_CN_OUT_OF_RANGE": 21,
                                          "WHS_CAPPED_BY_STYLE": 17, "DUPLICATE_STYLE_ACROSS_SHEETS": 8,
                                          "COLOR_IS_DESCRIPTOR": 6, "WHS_STYLE_UNKNOWN": 2,
                                          "EXTRACTION_WARNING": 4}


def _items(final, code):
    return [x for x in final["needs_review"] if x["code"] == code]


def test_duplicates_two_with_different_whs(f27_final):
    dups = _items(f27_final, "DUPLICATE_STYLE_ACROSS_SHEETS")
    assert len(dups) == 8 and all("HUTTON" in d["DESCRIPTION"] for d in dups)
    differing = [d for d in dups if len(d["whs_values"]) == 2]
    assert len(differing) == 2
    text = " ".join(d["colorway"].upper() for d in differing)
    assert "GLASGOW" in text and "FAWN" in text


def test_sloane_capped_to_000(f27_final):
    for i in range(130, 135):
        assert f27_final["records"][i]["WHS"] == "000"
    capped = [x for x in _items(f27_final, "WHS_CAPPED_BY_STYLE") if 130 <= x["record"] <= 134]
    assert len(capped) == 5


def test_rihanna_striped_tag(f27_final):
    assert f27_final["records"][167]["WS TAG COLOR"] == "BLUSH CHAMPAGNE STRIPE"


def test_merry_striped_uses_colour_whs(f27_final):
    assert f27_final["records"][668]["WHS"] == "000 & SY" and f27_final["records"][669]["WHS"] == "000"
    unknown = {x["record"] for x in _items(f27_final, "WHS_STYLE_UNKNOWN")}
    assert unknown == {668, 669}


def test_colorway_count_equals_records(f27_raw, f27_final):
    n = sum(len(s[g]) for s in f27_raw["styles"] for g in
            ("color_above_black_line", "color_between_black_and_red_line", "color_below_red_line"))
    assert n == len(f27_final["records"])


def test_validate_f27_against_im_header(f27_raw, f27_final, im_path):
    from matrix_to_im.validate import read_im_header, validate
    assert validate(f27_raw, f27_final, read_im_header(im_path)) == []

from matrix_to_im.review import Code


def test_review_codes_after_revision_2026_11_10():
    assert {c.value for c in Code} == {
        "TAG_CONDITIONAL", "TAG_CN_OUT_OF_RANGE", "TAG_EMPTY", "TAG_UNRESOLVED", "WHS_STYLE_UNKNOWN",
        "WHS_CONFLICT", "FONT_COLOR_NO_RULE", "COLOR_IS_DESCRIPTOR", "EXTRACTION_WARNING"}

import json

import pytest

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.io_utils import read_json, write_json
from matrix_to_im.transform import transform


def cw(yarns, cell="A1", font="NEW COLOR", comment=None):
    return {"colorway": "/".join(yarns), "yarns": yarns, "font_color": font, "cell": cell, "comment": comment}


def style(name="STYLE X", sheet="S1", section=None, no=1, whs="BOTH WHS", tag="TAG: C1/C2",
          above=(), between=(), below=()):
    return {"no": no, "style": name, "style_whs": whs, "tag": tag, "sheet": sheet, "section": section,
            "color_above_black_line": list(above), "color_between_black_and_red_line": list(between),
            "color_below_red_line": list(below)}


def raw_of(*styles, warnings=()):
    return {"source": "x", "styles": list(styles), "warnings": list(warnings)}


def raw_one(yarns, **kw):
    return raw_of(style(above=[cw(yarns, **kw)]))


def by_code(doc, code):
    return [x for x in doc["needs_review"] if x["code"] == code]


def test_record_keys_and_slots():
    rec = transform(raw_one(yarns=["A", "B", "C", "D"]))["records"][0]
    assert list(rec) == ["DESCRIPTION", "COLOR1", "COLOR2", "COLOR3", "COLOR 4", "WS TAG COLOR", "WHS", "_ref"]


def test_top_level_key_order_and_generated():
    doc = transform(raw_one(["A"]), generated="2026-01-02")
    assert list(doc) == ["target", "source", "generated", "rules", "summary", "records", "needs_review"]
    assert doc["generated"] == "2026-01-02"


def test_ref_with_and_without_section():
    raw = raw_of(style(sheet="LUXE", section="CHRISTMAS - COTTON", no=4, above=[cw(["A"], cell="D343")]),
                 style(name="S2", sheet="ESSENTIALS GRASSY", no=21, between=[cw(["A"], cell="D166")]))
    refs = [r["_ref"] for r in transform(raw)["records"]]
    assert refs == ["LUXE (CHRISTMAS - COTTON) No.4 | D343 | above_black",
                    "ESSENTIALS GRASSY No.21 | D166 | between_black_red"]


def test_stripe_yarns_written_once():
    rec = transform(raw_one(["A", "B", "a", "B"]))["records"][0]
    assert rec["COLOR1"] == "A" and rec["COLOR2"] == "B" and "COLOR3" not in rec


def test_too_many_yarns():
    yarns = [f"Y{i}" for i in range(1, 20)]
    doc = transform(raw_one(yarns))
    assert doc["records"][0]["COLOR18"] == "Y18" and "COLOR19" not in doc["records"][0]
    item = by_code(doc, "TOO_MANY_YARNS")[0]
    assert item["issue"] == "19 benang unik, slot COLOR hanya 18"
    assert item["colorway"].endswith("/Y18")


def test_descriptor_review_carries_comment():
    doc = transform(raw_one(["KHAKI STRIPE"], comment="Dev: X"))
    item = by_code(doc, "COLOR_IS_DESCRIPTOR")[0]
    assert item["comment_di_matrix"] == "Dev: X"
    assert item["issue"] == "COLOR berisi deskripsi, bukan nama benang: ['KHAKI STRIPE']"
    assert "tag_instruction" not in item
    assert list(item) == ["record", "code", "DESCRIPTION", "colorway", "issue", "comment_di_matrix", "_ref"]


def test_style_in_two_sheets_are_separate_records_without_review():      # keputusan 1
    raw = raw_of(style(sheet="S1", whs="000", between=[cw(["A", "B"], font="REPEAT COLOR")]),
                 style(sheet="S2", whs="BOTH WHS", between=[cw(["A", "B"], font="REPEAT COLOR")]))
    doc = transform(raw)
    assert [(r["WHS"], r["_ref"].split(" ")[0]) for r in doc["records"]] == [("000", "S1"), ("BOTH WHS", "S2")]
    assert doc["needs_review"] == []


def test_undecided_values_are_marked_and_reviewed():
    doc = transform(raw_of(style(tag="TAG for solids: C1 | for marl: C1", above=[cw(["A"], font="YR COLOR")])))
    rec = doc["records"][0]
    assert rec["WS TAG COLOR"] == "Need Review" and rec["WHS"] is None
    codes = [x["code"] for x in doc["needs_review"]]
    assert codes == ["TAG_CONDITIONAL", "FONT_COLOR_NO_RULE"]
    assert all("tag_instruction" in x for x in doc["needs_review"])
    assert list(doc["needs_review"][0]) == ["record", "code", "DESCRIPTION", "colorway", "issue",
                                            "tag_instruction", "_ref"]


def test_whs_capped_by_style_is_not_reviewed():                          # keputusan 2
    doc = transform(raw_of(style(whs="000", tag="TAG: C1", above=[cw(["A"])])))
    assert doc["records"][0]["WHS"] == "000" and doc["needs_review"] == []


def test_extraction_warnings_last_with_null_record():
    raw = raw_of(style(above=[cw(["A"], font="YR COLOR")]), warnings=["w1", "w2"])
    items = transform(raw)["needs_review"]
    assert [x["code"] for x in items][-2:] == ["EXTRACTION_WARNING"] * 2
    assert items[-1] == {"record": None, "code": "EXTRACTION_WARNING", "issue": "warning ekstraksi: w2"}


def test_review_sorted_by_record():
    raw = raw_of(style(name="S1", above=[cw(["A"], font="YR COLOR")]),
                 style(name="S2", above=[cw(["B"], font="YR COLOR")]),
                 style(name="S1", sheet="Z", above=[cw(["A"], font="YR COLOR")]), warnings=["w"])
    recs = [x["record"] for x in transform(raw)["needs_review"]]
    key = lambda r: r if isinstance(r, int) else (r[0] if isinstance(r, list) else 10**6)
    assert [key(r) for r in recs] == sorted(key(r) for r in recs)
    assert recs[-1] is None


def test_summary_counts_by_code():
    raw = raw_of(style(name="S1", tag="TAG: C1", above=[cw(["A"], font="YR COLOR"), cw(["B"])]),
                 style(name="S2", tag="TAG: C1", below=[cw(["C STRIPE"])]), warnings=["w"])
    s = transform(raw)["summary"]
    assert s["total_records"] == 3 and s["total_styles"] == 2
    assert s["WHS"] == {None: 1, "BOTH WHS": 1, "SY": 1}
    assert (s["WS_TAG_COLOR_filled"], s["WS_TAG_COLOR_need_review"], s["WS_TAG_COLOR_null"]) == (2, 1, 0)
    assert s["needs_review_by_issue"] == {"FONT_COLOR_NO_RULE": 1, "COLOR_IS_DESCRIPTOR": 1,
                                          "EXTRACTION_WARNING": 1}
    assert s["needs_review_items"] == 3


def test_bad_raw_raises():
    with pytest.raises(Matrix2ImError):
        transform({"source": "x"})
    with pytest.raises(Matrix2ImError):
        transform({"styles": "x"})


def test_write_json_creates_parent_and_utf8(tmp_path):
    path = tmp_path / "a" / "b" / "o.json"
    write_json(path, {"k": "é"})
    text = path.read_text(encoding="utf-8")
    assert "é" in text and text.startswith('{\n  "k"')
    assert read_json(path) == {"k": "é"}


def test_read_json_errors(tmp_path):
    with pytest.raises(Matrix2ImError):
        read_json(tmp_path / "missing.json")
    bad = tmp_path / "bad.json"
    bad.write_text("{oops", encoding="utf-8")
    with pytest.raises(Matrix2ImError):
        read_json(bad)


# --- pemeriksaan struktur raw JSON (Ruling 8) ---

@pytest.mark.parametrize("key", ["style", "tag", "sheet", "no", "section", "color_above_black_line",
                                 "color_between_black_and_red_line", "color_below_red_line"])
def test_style_missing_key_raises(key):
    st = style(above=[cw(["A"])])
    del st[key]
    with pytest.raises(Matrix2ImError, match=rf"style #1.*{key}"):
        transform(raw_of(st))


def test_error_names_style_index():
    bad = style(name="B", above=[cw(["A"])])
    del bad["tag"]
    with pytest.raises(Matrix2ImError, match=r"style #2"):
        transform(raw_of(style(above=[cw(["A"])]), bad))


@pytest.mark.parametrize("key", ["yarns", "font_color", "cell"])
def test_colorway_missing_key_raises(key):
    c = cw(["A"])
    del c[key]
    with pytest.raises(Matrix2ImError, match=rf"style #1.*colorway.*{key}"):
        transform(raw_of(style(above=[c])))


def test_group_not_a_list_raises():
    st = style(above=[cw(["A"])])
    st["color_below_red_line"] = "x"
    with pytest.raises(Matrix2ImError, match="color_below_red_line"):
        transform(raw_of(st))


def test_style_not_a_dict_raises():
    with pytest.raises(Matrix2ImError, match="style #1"):
        transform({"styles": ["x"], "warnings": []})


def test_warnings_not_a_list_raises():
    raw = raw_one(["A"])
    raw["warnings"] = "oops"
    with pytest.raises(Matrix2ImError, match="warnings"):
        transform(raw)


def test_tag_prefix_only_gives_need_review_and_tag_empty_with_instruction():   # F1 + revisi 10 Nov
    doc = transform(raw_of(style(tag="TAG:", above=[cw(["A"])])))
    assert doc["records"][0]["WS TAG COLOR"] == "Need Review"
    item = by_code(doc, "TAG_EMPTY")[0]
    assert item["record"] == 0 and item["tag_instruction"] == "TAG:" and item["issue"] == "TAG kosong"


@pytest.mark.parametrize("tag,yarns,code", [
    ("TAG: C1 (For Solid) | TAG: SNOW MARL", ["A"], "TAG_CONDITIONAL"),                    # keputusan 3
    ("TAG: C1/C2/C3", ["A", "B"], "TAG_CN_OUT_OF_RANGE"),                                   # keputusan 4
    ("TAG:", ["A"], "TAG_EMPTY"),
    ("TAG: C1/GONDOLA color", ["A"], "TAG_UNRESOLVED")])                                    # keputusan 7
def test_tag_problems_write_need_review(tag, yarns, code):
    doc = transform(raw_of(style(tag=tag, above=[cw(yarns)])))
    assert doc["records"][0]["WS TAG COLOR"] == "Need Review"
    assert [x["code"] for x in doc["needs_review"]] == [code]


def test_descriptor_writes_need_review_but_keeps_colors():                   # keputusan 6
    doc = transform(raw_of(style(tag="TAG: C1", above=[cw(["SILVERSTONE", "BREAKER WHITE MARL"])])))
    rec = doc["records"][0]
    assert (rec["COLOR1"], rec["COLOR2"], rec["WS TAG COLOR"]) == ("SILVERSTONE", "BREAKER WHITE MARL", "Need Review")
    assert [x["code"] for x in doc["needs_review"]] == ["COLOR_IS_DESCRIPTOR"]

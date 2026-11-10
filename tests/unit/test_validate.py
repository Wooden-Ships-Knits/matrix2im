import openpyxl
import pytest

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.validate import read_im_header, validate


def make_im(path, sheet="F27", row=54, cells=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for col, value in (cells or {}).items():
        ws.cell(row=row, column=col, value=value)
    wb.save(path)
    return path


def test_header_first_line_only(tmp_path):
    p = make_im(tmp_path / "im.xlsx", cells={
        1: "WS TAG COLOR\n(Note - Tag color untuk style striped)", 2: "COLOR 4", 3: 5, 4: None})
    assert read_im_header(p) == {"WS TAG COLOR", "COLOR 4", "5"}


def test_header_missing_sheet(tmp_path):
    p = make_im(tmp_path / "im.xlsx", sheet="Lain", cells={1: "X"})
    with pytest.raises(Matrix2ImError, match="Sheet 'F27' tidak ada"):
        read_im_header(p)


def test_header_empty_row(tmp_path):
    p = make_im(tmp_path / "im.xlsx", row=1, cells={1: "X"})
    with pytest.raises(Matrix2ImError, match="baris 54 kosong"):
        read_im_header(p)


def test_header_bad_file(tmp_path):
    with pytest.raises(Matrix2ImError):
        read_im_header(tmp_path / "tidak_ada.xlsx")


HEADER = {"DESCRIPTION", "COLOR1", "COLOR2", "COLOR 4", "WS TAG COLOR", "WHS"}


def raw_n(n=1):
    cw = {"colorway": "A", "yarns": ["A"], "font_color": "NEW COLOR", "cell": "A1", "comment": None}
    return {"styles": [{"color_above_black_line": [cw] * n, "color_between_black_and_red_line": [],
                        "color_below_red_line": []}]}


def rec(**kw):
    r = {"DESCRIPTION": "S", "COLOR1": "A", "WS TAG COLOR": "A", "WHS": "000", "_ref": "x"}
    r.update(kw)
    return r


def final_of(records, items=()):
    return {"records": records, "needs_review": list(items)}


def test_validate_ok():
    assert validate(raw_n(1), final_of([rec()]), HEADER) == []


def test_validate_count_mismatch():
    msgs = validate(raw_n(2), final_of([rec()]))
    assert any("jumlah colorway" in m for m in msgs)


def test_validate_unknown_key():
    msgs = validate(raw_n(1), final_of([rec(**{"COLOR 4": "Z"})]), HEADER - {"COLOR 4"})
    assert any("COLOR 4" in m for m in msgs)


def test_validate_bad_whs():
    msgs = validate(raw_n(1), final_of([rec(WHS="000/SY")]), HEADER)
    assert any("000/SY" in m for m in msgs)


def test_validate_null_without_review():
    msgs = validate(raw_n(1), final_of([rec(**{"WS TAG COLOR": None, "WHS": None})]), HEADER)
    assert len(msgs) == 2


def test_validate_null_with_review_ok():
    items = [{"record": 0, "code": "TAG_EMPTY"}, {"record": 0, "code": "WHS_CONFLICT"}]
    assert validate(raw_n(1), final_of([rec(**{"WS TAG COLOR": None, "WHS": None})], items), HEADER) == []


def test_duplicate_item_does_not_satisfy_null():
    items = [{"record": [0, 1], "code": "TAG_EMPTY"}]
    msgs = validate(raw_n(1), final_of([rec(**{"WS TAG COLOR": None})], items), HEADER)
    assert len(msgs) == 1


@pytest.mark.parametrize("key,value", [("COLOR1", ""), ("COLOR1", " A"), ("DESCRIPTION", "S "),
                                       ("WS TAG COLOR", "=C1"), ("COLOR2", 5)])
def test_validate_rejects_empty_whitespace_formula_values(key, value):     # F6
    msgs = validate(raw_n(1), final_of([rec(**{key: value})]), HEADER)
    assert msgs == [f"record 0: nilai '{key}' tidak valid ({value!r}); harus null atau teks tidak kosong "
                    "tanpa spasi di awal/akhir dan tidak diawali '='"]


def test_validate_accepts_valid_values_and_ignores_underscore_keys():
    ok = rec(COLOR2="PURE SNOW/ALMOND", _ref=" =bebas ")
    assert validate(raw_n(1), final_of([ok]), HEADER) == []


def test_need_review_without_item_is_rejected():
    msgs = validate(raw_n(1), final_of([rec(**{"WS TAG COLOR": "Need Review"})]), HEADER)
    assert msgs == ["record 0: WS TAG COLOR 'Need Review' tanpa item needs_review kode TAG_* atau COLOR_IS_DESCRIPTOR"]


@pytest.mark.parametrize("code", ["TAG_CONDITIONAL", "TAG_CN_OUT_OF_RANGE", "TAG_EMPTY", "TAG_UNRESOLVED",
                                  "COLOR_IS_DESCRIPTOR"])
def test_need_review_with_matching_item_ok(code):
    items = [{"record": 0, "code": code}]
    assert validate(raw_n(1), final_of([rec(**{"WS TAG COLOR": "Need Review"})], items), HEADER) == []


def test_need_review_with_unrelated_item_is_rejected():
    items = [{"record": 0, "code": "WHS_CONFLICT"}]
    assert len(validate(raw_n(1), final_of([rec(**{"WS TAG COLOR": "Need Review"})], items), HEADER)) == 1

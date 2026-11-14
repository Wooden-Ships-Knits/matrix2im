import copy
import io

import pytest
from helpers.matrix_factory import build_matrix
from openpyxl import load_workbook

from matrix_to_im.black_line_export import (
    ABOVE_SHEET, MISSING_SHEET, above_black_rows, above_black_xlsx_bytes, missing_above_rows,
)
from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.extract import extract
from matrix_to_im.transform import transform

HEADER_2 = ["COLLECTION", "SUB", "No", "DESCRIPTION", "COLOR1", "COLOR2", "WS TAG COLOR", "WHS", "SEL MATRIX",
            "NEEDS REVIEW"]
MISSING_HEADER = ["COLLECTION", "SUB", "No", "DESCRIPTION", "JUMLAH WARNA ANTARA HITAM & MERAH", "KETERANGAN"]
NEWPORT = ["TEST", "WINTER BEACH", 1, "NEWPORT CREW", "YCA KHAKI - 749", "YCA ALMOND BUTTER - 872",
           "KHAKI/ALMOND BUTTER MARL", "000", "A15", None]


@pytest.fixture(scope="module")
def data(tmp_path_factory):
    """raw + final dari MATRIX sintetis: 1 colorway di atas garis hitam (NEWPORT CREW A15), JUDE T NECK dan
    STRIPE CARDI tanpa garis hitam, SINGLE CARD seluruhnya di bawah garis merah."""
    raw = extract(build_matrix(tmp_path_factory.mktemp("m") / "m.xlsx"))
    return raw, transform(raw)


def sheets(body: bytes) -> dict[str, list[list]]:
    wb = load_workbook(io.BytesIO(body))
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb}


def test_rows_only_colorways_above_black_line(data):
    header, rows = above_black_rows(*data)
    assert header == HEADER_2
    assert rows == [NEWPORT]


def test_needs_review_issues_joined_per_record(data):
    raw, final = copy.deepcopy(data)
    final["needs_review"] += [{"record": 0, "code": "X", "issue": "satu"}, {"record": 0, "code": "Y", "issue": "dua"}]
    _, rows = above_black_rows(raw, final)
    assert rows[0][-1] == "satu; dua"


def test_color_columns_follow_widest_colorway(data):
    raw, final = copy.deepcopy(data)
    final["records"][0]["COLOR3"] = "EXTRA"
    header, rows = above_black_rows(raw, final)
    assert header[4:7] == ["COLOR1", "COLOR2", "COLOR3"]
    assert rows[0][4:7] == ["YCA KHAKI - 749", "YCA ALMOND BUTTER - 872", "EXTRA"]


@pytest.mark.parametrize("change", [
    lambda f: f["records"][0].update(_ref="TEST (WINTER BEACH) No.1 | A15 | between_black_red"),
    lambda f: f["records"][4].update(_ref="TEST (WINTER BEACH) No.2 | B99 | between_black_red"),
    lambda f: f["records"].pop(),
    lambda f: f["records"].append(dict(f["records"][0])),
])
def test_final_not_matching_raw_is_an_error(data, change):
    raw, final = copy.deepcopy(data)
    change(final)
    with pytest.raises(Matrix2ImError, match="Proses ulang"):
        above_black_rows(raw, final)


def test_missing_lists_cards_with_colors_above_red_but_none_above_black(data):
    assert missing_above_rows(data[0]) == [
        ["TEST", "WINTER BEACH", 2, "JUDE T NECK", 2, "garis hitam tidak ditemukan"],
        ["TEST", "WINTER BEACH", "S2 - A", "STRIPE CARDI", 1, "garis hitam tidak ditemukan"],
    ]


def test_missing_reason_when_black_line_exists(data):
    raw = copy.deepcopy(data[0])
    newport = raw["styles"][0]
    newport["color_between_black_and_red_line"][:0] = newport.pop("color_above_black_line")
    newport["color_above_black_line"] = []
    assert missing_above_rows(raw)[0] == ["TEST", "WINTER BEACH", 1, "NEWPORT CREW", 3,
                                          "semua warna di bawah garis hitam"]


def test_missing_reason_needs_exact_style_name(data):
    """Warning style 'JUDE T NECK' tidak boleh dianggap milik style bernama 'T NECK'."""
    raw = copy.deepcopy(data[0])
    raw["styles"][1]["style"] = "T NECK"
    assert missing_above_rows(raw)[0][-1] == "semua warna di bawah garis hitam"


def test_xlsx_has_both_sheets_with_frozen_filtered_header(data):
    body = above_black_xlsx_bytes(*data)
    got = sheets(body)
    assert list(got) == [ABOVE_SHEET, MISSING_SHEET] == ["Atas garis hitam", "Tanpa warna atas garis hitam"]
    assert got[ABOVE_SHEET] == [HEADER_2, NEWPORT]
    assert got[MISSING_SHEET][0] == MISSING_HEADER
    assert len(got[MISSING_SHEET]) == 3
    wb = load_workbook(io.BytesIO(body))
    for ws in wb:
        assert ws.freeze_panes == "A2"
        assert ws.auto_filter.ref == f"A1:{ws.cell(1, ws.max_column).column_letter}{ws.max_row}"
        assert all(c.font.b for c in ws[1])


def test_text_stays_text(data):
    raw, final = copy.deepcopy(data)
    final["records"][0]["COLOR1"] = "=1+1"
    final["records"][0]["COLOR2"] = "BEL\x07"
    ws = load_workbook(io.BytesIO(above_black_xlsx_bytes(raw, final)))[ABOVE_SHEET]
    assert (ws["E2"].value, ws["E2"].data_type) == ("=1+1", "s")
    assert ws["F2"].value == "BEL"
    assert (ws["H2"].value, ws["H2"].data_type) == ("000", "s")
    assert (ws["C2"].value, ws["C2"].data_type) == (1, "n")


def test_nothing_above_black_still_gives_both_sheets(data):
    raw, final = data
    raw = {**raw, "styles": raw["styles"][3:], "warnings": []}            # hanya SINGLE CARD (bawah merah)
    final = {**final, "records": final["records"][9:], "needs_review": []}
    got = sheets(above_black_xlsx_bytes(raw, final))
    assert got[ABOVE_SHEET] == [["COLLECTION", "SUB", "No", "DESCRIPTION", "COLOR1", "WS TAG COLOR", "WHS",
                                 "SEL MATRIX", "NEEDS REVIEW"]]
    assert got[MISSING_SHEET] == [MISSING_HEADER, ["(tidak ada)", None, None, None, None, None]]

import re

from matrix_to_im.extract import (font_meaning, is_new_repeat_label, is_punctuation_only,
                                  join_tag_cells, parse_no, split_yarns)
from matrix_to_im.legend_config import NUMBER_RE


def test_split_yarns_slash_and_newline():
    assert split_yarns("YCA KHAKI - 749/YCA ALMOND BUTTER - 872") == ["YCA KHAKI - 749", "YCA ALMOND BUTTER - 872"]
    assert split_yarns("SIENNA\nBOLD INK") == ["SIENNA", "BOLD INK"]
    assert split_yarns("BLACK/\nLATTE/ RED/") == ["BLACK", "LATTE", "RED"]
    assert split_yarns("  NIGHT  NAVY ") == ["NIGHT NAVY"]


def test_split_yarns_long_part_still_split():      # D2
    long = "A" * 40 + "/" + "B" * 30
    assert split_yarns(long) == ["A" * 40, "B" * 30]


def test_split_yarns_non_string():
    assert split_yarns(749) == ["749"]


def test_parse_no():
    assert parse_no(20) == 20 and parse_no(20.0) == 20 and parse_no("7") == 7
    assert parse_no("S2 - A") == "S2 - A" and parse_no("1A") == "1A" and parse_no(None) is None


def test_font_meaning_table():
    assert font_meaning("FF00B050") == ("NEW COLOR", True)
    assert font_meaning(None) == ("REPEAT COLOR", True)
    assert font_meaning("FF000000") == ("REPEAT COLOR", True)
    assert font_meaning("FF9900FF") == ("YR COLOR", True)
    assert font_meaning("FF0066FF") == ("DISCONTINUED COLOR FOR SY", True)
    assert font_meaning("FF123456") == ("? (FF123456)", False)


def test_font_meaning_ignores_alpha_byte():                # F8: openpyxl menulis "00B050" sebagai "0000B050"
    assert font_meaning("0000B050") == ("NEW COLOR", True)
    assert font_meaning("00FF00FF") == ("000 ONLY COLOR", True)
    assert font_meaning("00000000") == ("REPEAT COLOR", True)
    assert font_meaning("00123456") == ("? (00123456)", False)   # kode asli 8 digit tetap dilaporkan


def test_new_repeat_label_full_match_only():
    for ok in ("NEW", " exact repeat ", "REPEAT WITH UPDATE", "SIMILAR", "UPDATE", "REPEAT"):
        assert is_new_repeat_label(ok)
    for bad in ("NEWPORT CREW", "NEW COLOR", "REPEAT COLOR", None, 5):
        assert not is_new_repeat_label(bad)


def test_number_regex():
    for ok in ("20", "20.0", "1A", "S1", "S2 - A", "S10"):
        assert re.match(NUMBER_RE, ok)
    for bad in ("MAX Styles for 000 is 27", "S2-A", "A1", ""):
        assert not re.match(NUMBER_RE, bad)


def test_punctuation_only():
    assert is_punctuation_only(".") and is_punctuation_only("`") and is_punctuation_only("*")
    assert not is_punctuation_only("RED.")


def test_join_tag_cells():
    assert join_tag_cells(["TAG: BLACK/LATTE/\nRED", "TAG: RED PAPRIKA/SHADLE/\nKALE"]) == \
        "TAG: BLACK/LATTE/ RED | TAG: RED PAPRIKA/SHADLE/ KALE"
    assert join_tag_cells(["TAG: C1 (For Solid)\n\nTAG: SNOW MARL\n(For PURE SNOW/ALMOND BUTTER)"]) == \
        "TAG: C1 (For Solid) | TAG: SNOW MARL (For PURE SNOW/ALMOND BUTTER)"
    assert join_tag_cells([]) is None


def test_font_f19ef9_is_000_repeat():
    """Pink FFF19EF9 di MATRIX baru = 000 REPEAT COLOR (keputusan user 11 Nov 2026), jadi WHS warna 000."""
    assert font_meaning("FFF19EF9") == ("000 REPEAT COLOR", True)
    assert font_meaning("00F19EF9") == ("000 REPEAT COLOR", True)

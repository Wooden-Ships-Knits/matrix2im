"""Regression Tahap 1: ekstraksi file F27 nyata harus identik dengan golden."""
from __future__ import annotations

import hashlib
import json
import re
import time
from collections import Counter

import pytest

from matrix_to_im.extract import extract_workbook, to_raw_json

pytestmark = pytest.mark.regression


def _dump(obj) -> str:
    return json.dumps(obj, ensure_ascii=False)


def _names(colorways) -> list[str]:
    return [c.colorway for c in colorways]


def _find(f27_result, sheet, no, style=None):
    hits = [c for c in f27_result.cards if c.sheet == sheet and c.no == no
            and (style is None or c.style == style)]
    assert len(hits) == 1, hits
    return hits[0]


def test_styles_identical_to_golden(f27_raw, golden_raw):
    assert _dump(f27_raw["styles"]) == _dump(golden_raw["styles"])


def test_four_warnings_same_cells(f27_raw):
    refs = sorted(re.search(r"^(.+?) ([A-Z]+\d+)", w).groups() for w in f27_raw["warnings"])
    assert refs == [("ESSENTIALS GRASSY", "G169"), ("FALL", "D345"), ("FALL", "E227"), ("LUXE", "D324")]


def test_warnings_identical_to_golden(f27_raw, golden_raw):
    assert f27_raw["warnings"] == golden_raw["warnings"]


def test_baseline_counts(f27_raw):
    styles = f27_raw["styles"]
    assert len(styles) == 215 == f27_raw["total_styles"]
    assert len({s["style"] for s in styles}) == 211
    above = sum(len(s["color_above_black_line"]) for s in styles)
    between = sum(len(s["color_between_black_and_red_line"]) for s in styles)
    below = sum(len(s["color_below_red_line"]) for s in styles)
    assert (above, between, below) == (158, 390, 176)
    assert above + between + below == 724
    whs = Counter(s["style_whs"] for s in styles)
    assert whs == {"000 & SY": 85, "000": 70, "SY": 59, None: 1}


def test_hutton_no20(f27_result):
    card = _find(f27_result, "ESSENTIALS GRASSY", 20)
    assert card.style.startswith("HUTTON")
    assert _names(card.above_black) == ["NIGHT NAVY"]
    assert _names(card.between) == ["PURE SNOW/ALMOND BUTTER", "GLASGOW", "FAWN", "RED CHILE"]
    assert _names(card.below_red) == ["BLACK", "DARK ROAST", "SCANDIA BLUE"]
    assert card.style_whs == "000 & SY"


def test_jude_no21(f27_result):
    card = _find(f27_result, "ESSENTIALS GRASSY", 21)
    assert card.style.startswith("JUDE")
    assert _names(card.above_black) == ["DARK CEDAR"]
    assert _names(card.between) == ["ALMOND BUTTER", "SLEET", "SIENNA", "BOLD INK"]
    assert card.below_red == []
    assert card.tag == "TAG: C1"


def test_essentials_cotton_page1_has_8_cards(f27_result):
    cards = [c for c in f27_result.cards if c.sheet == "ESSENTIALS COTTON" and c.page == 1]
    assert len(cards) == 8 and cards[0].style == "PEBBLE STITCH RAGLAN SWEATSHIRT COTTON"


def test_fall_page5_all_below_red(f27_result):
    cards = [c for c in f27_result.cards if c.sheet == "FALL" and c.page == 5]
    assert len(cards) == 10 and sum(len(c.below_red) for c in cards) == 12
    assert all(not c.above_black and not c.between for c in cards)


def test_luxe_page7_single_card_s1(f27_result):
    cards = [c for c in f27_result.cards if c.sheet == "LUXE" and c.page == 7]
    assert len(cards) == 1
    assert cards[0].no == "S1" and cards[0].style == "CHRISTMAS TREE CREW COTTON"


def test_text_number_s2_a(f27_raw):
    hits = [s for s in f27_raw["styles"] if s["no"] == "S2 - A"]
    assert any(s["style"] == "CROP BOYFRIEND CREW COTTON" for s in hits)


def test_fa_la_la_s3_two_row_tag(f27_raw):
    hits = [s for s in f27_raw["styles"]
            if s["sheet"] == "LUXE" and s["section"] == "CHRISTMAS - GRASSY"
            and s["no"] == "S3" and s["style"] == "FA LA LA FAIR ISLE CREW CHUNKY"]
    assert len(hits) == 1
    assert hits[0]["tag"] == "TAG: BLACK/LATTE/ RED | TAG: RED PAPRIKA/SHADLE/ KALE"


def test_merry_striped_style_whs_null_with_warning(f27_raw):
    hits = [s for s in f27_raw["styles"]
            if s["sheet"] == "LUXE" and s["section"] == "CHRISTMAS - COTTON" and s["no"] == 4]
    assert len(hits) == 1 and hits[0]["style_whs"] is None
    assert any("D324" in w for w in f27_raw["warnings"])


def test_input_file_not_modified(matrix_path):
    def snapshot():
        return hashlib.sha256(matrix_path.read_bytes()).hexdigest(), matrix_path.stat().st_mtime_ns

    before = snapshot()
    extract_workbook(matrix_path)
    assert snapshot() == before


def test_deterministic_and_fast(matrix_path, f27_raw):
    start = time.perf_counter()
    second = to_raw_json(extract_workbook(matrix_path))
    elapsed = time.perf_counter() - start
    assert _dump(second) == _dump(f27_raw)
    assert elapsed < 15, f"extract_workbook {elapsed:.1f} detik"

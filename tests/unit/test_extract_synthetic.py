import hashlib

import pytest
from helpers.matrix_factory import build_matrix, build_no_pages
from openpyxl import Workbook, load_workbook

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.extract import extract, extract_workbook, to_raw_json
from matrix_to_im.legend_config import SOURCE_NAME


@pytest.fixture
def synthetic(tmp_path):
    return extract_workbook(build_matrix(tmp_path / "m.xlsx"))


def test_cards_and_order(synthetic):
    assert [(c.page, c.column, c.style) for c in synthetic.cards] == [
        (1, "A", "NEWPORT CREW"), (1, "B", "JUDE T NECK"), (1, "C", "STRIPE CARDI"), (2, "A", "SINGLE CARD")]


def test_numbers_section_style_whs(synthetic):
    a, b, c, d = synthetic.cards
    assert (a.no, a.style_whs, a.section) == (1, "000", "WINTER BEACH")
    assert (b.no, b.style_whs) == (2, "000 & SY")
    assert (c.no, c.style_whs, c.no_fill) == ("S2 - A", None, None)
    assert (d.no, d.style_whs, d.section) == (1, "SY", None)


def test_card_a_lines_and_groups(synthetic):
    a = synthetic.cards[0]
    assert (a.black_line_row, a.red_line_row) == (16, 18)
    assert [x.cell for x in a.above_black] == ["A15"]
    assert [x.cell for x in a.between] == ["A16", "A17"]
    assert [x.cell for x in a.below_red] == ["A18"]
    assert a.between[1].yarns == ["SIENNA", "BOLD INK"] and a.between[1].colorway == "SIENNA/BOLD INK"
    assert [x.font_color for x in a.above_black + a.between] == ["NEW COLOR", "REPEAT COLOR", "000 ONLY COLOR"]


def test_card_b_falls_back_to_page_red_line(synthetic):
    b = synthetic.cards[1]
    assert b.red_line_row == 18 and b.black_line_row is None
    assert [x.cell for x in b.between] == ["B15", "B16"] and [x.cell for x in b.below_red] == ["B18"]
    assert b.between[0].font_color == "000 REPEAT COLOR"


def test_card_c_multirow_tag_comment_and_tag_in_colour_area(synthetic):
    c = synthetic.cards[2]
    assert c.tag == "TAG: BLACK/LATTE/ RED | TAG: KALE"
    assert [x.cell for x in c.between] == ["C16"] and c.between[0].comment == "Dev: catatan"
    assert [x.cell for x in c.below_red] == ["C18"]            # C17 (TAG) bukan colorway


def test_red_line_on_first_colour_row_puts_all_below(synthetic):
    d = synthetic.cards[3]
    assert d.above_black == [] and d.between == [] and [x.cell for x in d.below_red] == ["A36", "A37"]
    assert d.below_red[1].font_color == "? (FF123456)"


def test_warnings(synthetic):
    w = "\n".join(synthetic.warnings)
    assert 'TEST A19: karakter "." di area warna diabaikan' in w
    assert "TEST C9 No.S2 - A STRIPE CARDI: background nomor tidak dikenali (None)" in w
    assert "C17" in w and "TAG" in w                          # D10
    assert "A37" in w and "FF123456" in w                     # D9
    assert "garis hitam tidak ditemukan" in w                 # kartu B dan C


def test_raw_json_schema(synthetic):
    raw = to_raw_json(synthetic)
    assert list(raw) == ["source", "total_styles", "styles", "warnings"] and raw["total_styles"] == 4
    assert list(raw["styles"][0]) == ["no", "style", "style_whs", "tag", "color_above_black_line",
        "color_between_black_and_red_line", "color_below_red_line", "sheet", "section"]
    assert list(raw["styles"][0]["color_above_black_line"][0]) == ["colorway", "yarns", "font_color", "cell", "comment"]


def test_extract_projects_colorways_to_raw(tmp_path):
    raw = extract(build_matrix(tmp_path / "m.xlsx"))
    assert raw["source"] == SOURCE_NAME and raw["styles"][3]["sheet"] == "TEST"
    assert raw["styles"][2]["color_between_black_and_red_line"] == [
        {"colorway": "RED CHILE", "yarns": ["RED CHILE"], "font_color": "REPEAT COLOR",
         "cell": "C16", "comment": "Dev: catatan"}]


def test_card_without_number_warns(tmp_path):                 # D8
    path = build_matrix(tmp_path / "m.xlsx")
    wb = load_workbook(path)
    wb["TEST"]["A30"] = None                                  # halaman 2 kehilangan baris nomor
    wb.save(tmp_path / "tanpa_nomor.xlsx")
    result = extract_workbook(tmp_path / "tanpa_nomor.xlsx")
    d = result.cards[3]
    assert d.no is None and d.no_cell is None and d.style_whs is None
    assert [w for w in result.warnings if "A33" in w] == [
        "TEST A33 SINGLE CARD: nomor style tidak ditemukan, WHS style tidak bisa ditentukan"]


def test_no_pages_is_fatal(tmp_path):
    with pytest.raises(Matrix2ImError, match="Tidak ada halaman"):
        extract_workbook(build_no_pages(tmp_path / "x.xlsx"))


def test_pages_without_cards_is_fatal(tmp_path):
    wb = Workbook()
    wb.active["A1"] = "FALL 27 - KOSONG - MATRIX MASTER (PAGE 1 OF 1)"
    wb.save(tmp_path / "kosong.xlsx")
    with pytest.raises(Matrix2ImError, match="Tidak ada kartu style"):
        extract_workbook(tmp_path / "kosong.xlsx")


def test_input_not_modified(tmp_path):
    path = build_matrix(tmp_path / "m.xlsx")
    before = (hashlib.sha256(path.read_bytes()).hexdigest(), path.stat().st_mtime_ns)
    extract_workbook(path)
    assert (hashlib.sha256(path.read_bytes()).hexdigest(), path.stat().st_mtime_ns) == before


def _variant(tmp_path, name, **cells):
    """Salinan workbook sintetis dengan beberapa sel diganti (disimpan openpyxl: formula tanpa cache)."""
    wb = load_workbook(build_matrix(tmp_path / "m.xlsx"))
    for ref, value in cells.items():
        wb["TEST"][ref] = value
    wb.save(tmp_path / name)
    return tmp_path / name


def test_uncached_formula_number_is_fatal(tmp_path):         # F4
    path = _variant(tmp_path, "tanpa_cache.xlsx", B9="=A9+1")
    with pytest.raises(Matrix2ImError) as exc:
        extract_workbook(path)
    assert str(exc.value) == (
        f"Nilai hasil formula nomor style tidak tersimpan di {path} (contoh: TEST B9). File kemungkinan "
        "disimpan tanpa kalkulasi; buka dan simpan ulang di Excel, lalu jalankan lagi.")


def test_number_row_with_only_uncached_formulas_is_fatal(tmp_path):
    # Halaman 2 hanya punya satu nomor; bila itu formula tanpa cache baris nomor tidak terbaca
    # sama sekali, dan itu harus error, bukan "nomor style tidak ditemukan".
    path = _variant(tmp_path, "tanpa_cache2.xlsx", A30="=A9")
    with pytest.raises(Matrix2ImError, match=r"\(contoh: TEST A30\)"):
        extract_workbook(path)


def test_missing_number_with_known_fill_says_whs_from_background(tmp_path):   # D8 teks
    result = extract_workbook(_variant(tmp_path, "b9_kosong.xlsx", B9=None))
    b = result.cards[1]
    assert (b.no, b.no_cell, b.style_whs) == (None, "B9", "000 & SY")
    assert [w for w in result.warnings if "B9" in w] == [
        "TEST B9 JUDE T NECK: nomor style tidak ditemukan, WHS style 000 & SY diambil dari warna background"]

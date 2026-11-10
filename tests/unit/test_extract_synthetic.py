import hashlib

import pytest
from helpers.matrix_factory import build_matrix, build_no_pages
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, PatternFill, Side

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.excel_reader import open_workbooks
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
    assert (c.no, c.style_whs, c.no_fill) == ("S2 - A", "000 & SY", None)    # tanpa fill = putih (keputusan 5)
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
    assert "A19" not in w and "karakter" not in w                 # tanda baca dibuang tanpa warning (keputusan 8)
    assert "C9" not in w and "background nomor" not in w         # tanpa fill = 000 & SY (keputusan 5)
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


def _with_cache(monkeypatch, path, cache):
    """Simulasikan file yang disimpan Excel: formula di workbook normal, nilai hasilnya di data_only."""
    normal, values = open_workbooks(path)
    for ref, value in cache.items():
        values["TEST"][ref] = value
    monkeypatch.setattr("matrix_to_im.extract.open_workbooks", lambda _path: (normal, values))


def test_formula_cells_read_cached_value_like_literals(tmp_path, monkeypatch):   # F4 + F5
    literal = to_raw_json(extract_workbook(build_matrix(tmp_path / "lit.xlsx")))
    path = _variant(tmp_path, "formula.xlsx", B9="=A9+1", C12="=A12", B14='="TAG: C1"', A16="=A15")
    _with_cache(monkeypatch, path, {"B9": 2, "C12": "STRIPE CARDI", "B14": "TAG: C1 (For Solid)",
                                    "A16": "DARK CEDAR"})
    assert to_raw_json(extract_workbook(path)) == literal


def test_uncached_formula_colour_cell_warns_and_is_skipped(tmp_path):    # F5
    result = extract_workbook(_variant(tmp_path, "warna.xlsx", A16="=A15"))
    a = result.cards[0]
    assert [x.cell for x in a.above_black + a.between + a.below_red] == ["A15", "A17", "A18"]
    assert "TEST A16: formula tanpa nilai tersimpan, diabaikan" in result.warnings
    assert not any(x.colorway.startswith("=") for c in result.cards
                   for x in c.above_black + c.between + c.below_red)


def test_uncached_formula_name_cell_skips_card_with_warning(tmp_path):
    result = extract_workbook(_variant(tmp_path, "nama.xlsx", C12="=A12"))
    assert [c.style for c in result.cards] == ["NEWPORT CREW", "JUDE T NECK", "SINGLE CARD"]
    assert "TEST C12: formula tanpa nilai tersimpan di sel nama style, kartu dilewati" in result.warnings


def test_uncached_formula_tag_cell_is_not_a_tag(tmp_path):
    result = extract_workbook(_variant(tmp_path, "tag.xlsx", B14='="TAG: C1"', B13="=A13"))
    b = result.cards[1]
    assert b.tag is None and [x.cell for x in b.between + b.below_red] == ["B15", "B16", "B18"]
    w = result.warnings
    assert "TEST B14: formula tanpa nilai tersimpan, diabaikan" in w
    assert "TEST B13: formula tanpa nilai tersimpan, diabaikan" in w   # sebelum blok TAG
    assert "TEST B12 JUDE T NECK: tidak ada baris TAG" in w


def test_six_digit_colours_written_by_openpyxl(tmp_path):   # F8: alpha "00" bukan warna lain
    wb = load_workbook(build_matrix(tmp_path / "m.xlsx"))
    ws = wb["TEST"]
    ws["A9"].fill = PatternFill("solid", fgColor="FF66FF")
    ws["A15"].font = Font(color="00B050")
    ws["A18"].border = Border(top=Side(style="medium", color="FF0000"))
    wb.save(tmp_path / "alpha00.xlsx")
    a = extract_workbook(tmp_path / "alpha00.xlsx").cards[0]
    assert (a.no_fill, a.style_whs) == ("00FF66FF", "000")
    assert a.above_black[0].font_color == "NEW COLOR"
    assert a.red_line_row == 18 and [x.cell for x in a.below_red] == ["A18"]



def _fill_variant(tmp_path, name, ref, rgb, value="keep"):
    """Salinan workbook sintetis dengan warna background satu sel diganti (rgb None = tanpa fill)."""
    wb = load_workbook(build_matrix(tmp_path / "m.xlsx"))
    cell = wb["TEST"][ref]
    cell.fill = PatternFill(fill_type="solid", fgColor=rgb) if rgb else PatternFill(fill_type=None)
    if value != "keep":
        cell.value = value
    wb.save(tmp_path / name)
    return tmp_path / name


def test_white_number_background_is_000_and_sy(tmp_path):                   # keputusan 5
    result = extract_workbook(_fill_variant(tmp_path, "putih.xlsx", "A9", "FFFFFFFF"))
    assert result.cards[0].style_whs == "000 & SY"
    assert not [w for w in result.warnings if "A9" in w]


def test_unknown_number_background_still_null_with_warning(tmp_path):
    result = extract_workbook(_fill_variant(tmp_path, "hijau.xlsx", "A9", "FF00B050"))
    assert result.cards[0].style_whs is None
    assert "TEST A9 No.1 NEWPORT CREW: background nomor tidak dikenali (FF00B050), WHS style tidak bisa ditentukan" \
        in result.warnings


def test_missing_number_without_fill_stays_unknown(tmp_path):
    # Aturan putih/tanpa fill hanya untuk sel nomor yang benar-benar ada isinya.
    result = extract_workbook(_fill_variant(tmp_path, "c9_kosong.xlsx", "C9", None, value=None))
    c = result.cards[2]
    assert (c.no, c.style_whs) == (None, None)
    assert "TEST C9 STRIPE CARDI: nomor style tidak ditemukan, WHS style tidak bisa ditentukan" in result.warnings

"""Tahap 1 (ekstraksi) MATRIX MASTER: halaman -> kartu style -> colorway, lalu raw JSON.

Semua posisi dicari dari isi sel (penanda halaman, label, garis), tidak ada nomor
baris/kolom tetap, karena tiap halaman MATRIX punya tinggi dan jumlah kartu berbeda.
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .errors import Matrix2ImError
from .excel_reader import (fill_key, font_rgb, is_black_line, is_red_line, line_above, norm,
                           open_workbooks)
from .legend_config import (BLACK_FONT_RGB, DEFAULT_FONT_MEANING, FONT_MEANING, LEGEND_LABELS,
                            NEW_REPEAT_LABELS, NUMBER_RE, PAGE_MARKER, SECTION_EXCLUDE_WORDS,
                            SOURCE_NAME, STYLE_WHS_BY_FILL, TAG_PREFIX, WIDTH_LABEL)

_NUMBER = re.compile(NUMBER_RE)


def split_yarns(text: object) -> list[str]:
    """Pecah sel benang pada '/' dan baris baru; selalu dipecah (D2), urutan dipertahankan."""
    parts = (norm(p) for p in re.split(r"[/\n]", str(text)))
    return [p for p in parts if p]


def parse_no(value: object) -> int | str | None:
    """Nomor style: angka jadi int agar 20.0 tidak tampil sebagai '20.0'; kode seperti '1A' tetap teks."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(value)
    text = norm(value)
    return int(text) if text.isdigit() else text


def font_meaning(rgb: str | None) -> tuple[str, bool]:
    """Arti warna font menurut legend; tanpa warna/hitam = REPEAT. Kode asing ditandai tidak dikenal."""
    if rgb is None or rgb in BLACK_FONT_RGB:
        return DEFAULT_FONT_MEANING, True
    if rgb in FONT_MEANING:
        return FONT_MEANING[rgb], True
    return f"? ({rgb})", False


def is_new_repeat_label(value: object) -> bool:
    """Full match saja, supaya nama style seperti 'NEWPORT CREW' tidak dikira label NEW."""
    return isinstance(value, str) and norm(value).upper() in NEW_REPEAT_LABELS


def is_punctuation_only(text: str) -> bool:
    """Sel yang hanya berisi tanda baca (mis. '.', '`') bukan data (D5)."""
    return re.fullmatch(r"[\W_]+", text) is not None


def join_tag_cells(values: list[str]) -> str | None:
    """Gabung sel TAG per kartu; baris kosong ganda jadi ' | ', baris baru tunggal jadi spasi (D6)."""
    cells = [norm(v).replace("\n\n", " | ").replace("\n", " ") for v in values]
    return " | ".join(cells) if cells else None


@dataclass
class Colorway:
    """Satu sel di area warna kartu. `row` hanya untuk pengelompokan garis, tidak diekspor."""

    colorway: str
    yarns: list[str]
    font_color: str
    cell: str
    comment: str | None
    row: int


@dataclass
class Card:
    """Satu kartu style (satu kolom di satu halaman) dengan warna per grup garis hitam/merah."""

    sheet: str
    page: int
    column: str
    section: str | None
    no: int | str | None
    no_cell: str | None
    no_fill: str | None
    style_whs: str | None
    style: str
    tag: str | None
    red_line_row: int | None
    black_line_row: int | None
    above_black: list[Colorway]
    between: list[Colorway]
    below_red: list[Colorway]


@dataclass
class ExtractResult:
    """Kartu berurutan sheet -> halaman -> kolom, plus warning yang perlu dicek manusia."""

    cards: list[Card]
    warnings: list[str]


@dataclass
class _Page:
    """Anchor satu halaman, dipakai bersama oleh semua kartunya."""

    ws: Worksheet
    values: Worksheet       # sheet yang sama dari workbook data_only: nomor style hasil formula
    number: int
    columns: range          # kolom kartu = kiri dari kolom label LEBAR/WIDTH
    nr_row: int
    width_row: int
    num_row: int | None
    section: str | None
    red_row: int | None


def _is_tag(value: object) -> bool:
    """Instruksi TAG dikenali dari prefix-nya, bukan dari posisi barisnya."""
    return isinstance(value, str) and norm(value).upper().startswith(TAG_PREFIX)


def _is_number(value: object) -> bool:
    return value is not None and _NUMBER.match(norm(value)) is not None


def _uncached_formula(ws: Worksheet, values: Worksheet, row: int, col: int) -> bool:
    """Sel formula yang hasilnya tidak tersimpan di file: workbook data_only hanya melihat None.

    Terjadi bila file disimpan oleh tool yang tidak menghitung formula (mis. openpyxl).
    """
    return ws.cell(row, col).data_type == "f" and values.cell(row, col).value is None


def _value(ws: Worksheet, values: Worksheet, row: int, col: int) -> object:
    """Isi sel sebagai data: untuk sel formula pakai nilai hasil dari workbook data_only.

    Workbook normal menyimpan teks formula (mis. '=C163'), yang bukan nama style/TAG/warna.
    """
    cell = ws.cell(row, col)
    return values.cell(row, col).value if cell.data_type == "f" else cell.value


def _page_bounds(ws: Worksheet) -> list[tuple[int, int]]:
    """(baris awal, baris akhir eksklusif) tiap halaman; satu sheet memuat banyak halaman cetak."""
    tops = [c.row for c in ws["A"] if isinstance(c.value, str) and c.value.startswith(PAGE_MARKER)]
    return list(zip(tops, tops[1:] + [ws.max_row + 1]))


def _width_label(ws: Worksheet, top: int, bottom: int) -> tuple[int, int] | None:
    """(baris, kolom) label LEBAR/WIDTH: barisnya menutup area warna, kolomnya membatasi kolom kartu."""
    for row in range(top, bottom):
        for cell in ws[row]:
            if isinstance(cell.value, str) and norm(cell.value) == WIDTH_LABEL:
                return row, cell.column
    return None


def _number_row(ws: Worksheet, values: Worksheet, top: int, nr_row: int, columns: range,
                source: Path) -> int | None:
    """Baris nomor = baris pertama di bawah legend yang punya nomor style di kolom kartu.

    Pencarian dimulai setelah baris legend terakhir (D3, pengganti offset tetap top+5
    prototipe) supaya isi legend tidak terbaca sebagai nomor, dan berhenti sebelum baris nama.
    Nomor style berupa formula (mis. =A9+1); formula tanpa nilai tersimpan di baris yang
    diperiksa (termasuk seluruh baris nomor) adalah error, karena kalau dilewati semua nomor
    hilang diam-diam dan WHS style ikut tidak bisa dicek.
    """
    legend_end = top
    for row in range(top, nr_row):
        if any(isinstance(c.value, str) and norm(c.value).upper() in LEGEND_LABELS for c in ws[row]):
            legend_end = row
    for row in range(legend_end + 1, nr_row - 1):
        found = False
        for col in columns:
            if _uncached_formula(ws, values, row, col):
                raise Matrix2ImError(
                    f"Nilai hasil formula nomor style tidak tersimpan di {source} "
                    f"(contoh: {ws.title} {ws.cell(row, col).coordinate}). File kemungkinan disimpan "
                    "tanpa kalkulasi; buka dan simpan ulang di Excel, lalu jalankan lagi.")
            found = found or _is_number(values.cell(row, col).value)
        if found:
            return row
    return None


def _section(ws: Worksheet, num_row: int | None) -> str | None:
    """Nama grup (mis. 'WINTER BEACH') di kolom A tepat di atas baris nomor.

    Teks legend/catatan di posisi itu (mengandung COLOR/STYLE/ORDER) bukan nama grup.
    """
    if num_row is None:
        return None
    value = ws.cell(num_row - 1, 1).value
    if not isinstance(value, str):
        return None
    text = norm(value)
    if text and not any(word in text.upper() for word in SECTION_EXCLUDE_WORDS):
        return text
    return None


def _page_red_line(ws: Worksheet, nr_row: int, width_row: int, columns: range) -> int | None:
    """Baris garis merah mayoritas di halaman, fallback untuk kartu tanpa garis merah sendiri.

    Seri dimenangkan baris paling atas (Counter menyimpan urutan pertama kali muncul).
    """
    hits = Counter(row for row in range(nr_row, width_row + 1) for col in columns
                   if line_above(ws, row, col, is_red_line))
    return hits.most_common(1)[0][0] if hits else None


def _read_page(ws: Worksheet, values: Worksheet, number: int, top: int, bottom: int,
               warnings: list[str], source: Path) -> _Page | None:
    """Anchor halaman; halaman yang anchor-nya tidak lengkap dilewati dengan warning, bukan error."""
    label = _width_label(ws, top, bottom)
    if label is None:
        warnings.append(f'{ws.title} hal.{number}: label "{WIDTH_LABEL}" tidak ditemukan')
        return None
    width_row, label_col = label
    columns = range(1, label_col)
    # Label NEW/REPEAT dicari dari awal halaman (D3); full match, jadi 'NEWPORT CREW' tidak ikut.
    nr_row = next((row for row in range(top, width_row)
                   if any(is_new_repeat_label(ws.cell(row, col).value) for col in columns)), None)
    if nr_row is None:
        warnings.append(f"{ws.title} hal.{number}: baris NEW/REPEAT tidak ditemukan")
        return None
    num_row = _number_row(ws, values, top, nr_row, columns, source)
    red_row = _page_red_line(ws, nr_row, width_row, columns)
    if red_row is None:
        warnings.append(
            f"{ws.title} hal.{number}: garis merah tidak ditemukan, semua warna dianggap box hijau")
    return _Page(ws, values, number, columns, nr_row, width_row, num_row, _section(ws, num_row), red_row)


def _tag_rows(ws: Worksheet, values: Worksheet, col: int, nr_row: int, width_row: int) -> list[int]:
    """Sel TAG pertama di bawah baris NEW/REPEAT plus sel TAG berurutan di bawahnya
    (instruksi TAG panjang kadang ditulis di beberapa sel)."""
    rows: list[int] = []
    for row in range(nr_row, width_row):
        if _is_tag(_value(ws, values, row, col)):
            rows.append(row)
        elif rows:
            break
    return rows


def _read_colorway(ws: Worksheet, values: Worksheet, row: int, col: int,
                   warnings: list[str]) -> Colorway | None:
    """Sel area warna -> Colorway.

    Sel TAG (D10), sel tanda baca (D5), formula tanpa nilai tersimpan, dan warna font di luar
    legend (D9) dicatat sebagai warning supaya tidak hilang diam-diam; tiga yang pertama bukan
    colorway.
    """
    cell = ws.cell(row, col)
    where = f"{ws.title} {cell.coordinate}"
    if _uncached_formula(ws, values, row, col):
        warnings.append(f"{where}: formula tanpa nilai tersimpan, diabaikan")
        return None
    value = _value(ws, values, row, col)
    if value is None:
        return None
    text = norm(value)
    if not text:
        return None
    if _is_tag(value):
        flat = text.replace("\n", " ")
        warnings.append(f'{where}: sel "{flat}" di area warna tidak dihitung sebagai colorway, '
                        "periksa manual")
        return None
    if is_punctuation_only(text):
        warnings.append(f'{where}: karakter "{text}" di area warna diabaikan')
        return None
    rgb = font_rgb(cell)
    meaning, known = font_meaning(rgb)
    if not known:
        warnings.append(f'{where}: warna font {rgb} tidak ada di legend, font_color diisi "{meaning}"')
    yarns = split_yarns(value)
    comment = norm(cell.comment.text) if cell.comment else None
    return Colorway("/".join(yarns), yarns, meaning, cell.coordinate, comment, row)


def _read_card(page: _Page, col: int, warnings: list[str]) -> Card | None:
    """Satu kolom kartu; None bila sel nama (tepat di atas baris NEW/REPEAT) kosong."""
    ws, values = page.ws, page.values
    name_row = page.nr_row - 1
    letter = get_column_letter(col)
    name_cell = f"{letter}{name_row}"
    if _uncached_formula(ws, values, name_row, col):
        warnings.append(f"{ws.title} {name_cell}: formula tanpa nilai tersimpan di sel nama style, "
                        "kartu dilewati")
        return None
    name = _value(ws, values, name_row, col)
    style = norm(name) if name is not None else ""
    if not style:
        return None

    tag_rows = _tag_rows(ws, values, col, page.nr_row, page.width_row)
    start = (tag_rows[-1] if tag_rows else page.nr_row) + 1
    # Sel di zona TAG (baris NEW/REPEAT sampai blok TAG) tidak dibaca area warna; formula tanpa
    # nilai di situ mungkin TAG yang hilang, jadi tetap diberi warning.
    for row in range(page.nr_row, start):
        if row not in tag_rows and _uncached_formula(ws, values, row, col):
            warnings.append(f"{ws.title} {ws.cell(row, col).coordinate}: formula tanpa nilai tersimpan, "
                            "diabaikan")
    # Garis dicari per kolom karena posisinya bisa beda antar kartu; garis merah halaman
    # hanya fallback. Garis hitam dicari mulai start+1 seperti prototipe (D15).
    red = next((row for row in range(start, page.width_row + 1)
                if line_above(ws, row, col, is_red_line)), page.red_row)
    black = next((row for row in range(start + 1, red or page.width_row)
                  if line_above(ws, row, col, is_black_line)), None)
    above: list[Colorway] = []
    between: list[Colorway] = []
    below: list[Colorway] = []
    for row in range(start, page.width_row):
        colorway = _read_colorway(ws, values, row, col, warnings)
        if colorway is None:
            continue
        if red is not None and row >= red:
            below.append(colorway)
        elif black is not None and row < black:
            above.append(colorway)
        else:
            between.append(colorway)

    no = no_cell = no_fill = None
    if page.num_row is not None:
        number_cell = ws.cell(page.num_row, col)
        no = parse_no(values.cell(page.num_row, col).value)
        no_cell, no_fill = number_cell.coordinate, fill_key(number_cell)
    style_whs = STYLE_WHS_BY_FILL.get(no_fill)   # WHS style = warna background sel nomor

    name_ref = f"{ws.title} {name_cell} {style}"
    if no is None:   # D8; sel nomor kosong yang ber-background tetap memberi WHS style
        whs_note = ("WHS style tidak bisa ditentukan" if style_whs is None
                    else f"WHS style {style_whs} diambil dari warna background")
        warnings.append(f"{ws.title} {no_cell or name_cell} {style}: nomor style tidak ditemukan, {whs_note}")
    elif style_whs is None:
        warnings.append(f"{ws.title} {no_cell} No.{no} {style}: "
                        f"background nomor tidak dikenali ({no_fill}), WHS style tidak bisa ditentukan")
    if (above or between) and black is None:
        warnings.append(f"{name_ref}: ada warna di atas garis merah tapi garis hitam tidak ditemukan")
    if not tag_rows:
        warnings.append(f"{name_ref}: tidak ada baris TAG")
    if not (above or between or below):
        warnings.append(f"{name_ref}: tidak ada warna")

    return Card(sheet=ws.title, page=page.number, column=letter, section=page.section,
                no=no, no_cell=no_cell, no_fill=no_fill, style_whs=style_whs, style=style,
                tag=join_tag_cells([_value(ws, values, row, col) for row in tag_rows]),
                red_line_row=red, black_line_row=black,
                above_black=above, between=between, below_red=below)


def extract_workbook(path: Path) -> ExtractResult:
    """Tahap 1: baca semua sheet dan halaman MATRIX menjadi kartu. Workbook tidak pernah di-save.

    Nol halaman atau nol kartu berarti file bukan MATRIX (atau formatnya berubah), jadi
    berhenti dengan error daripada menulis JSON kosong.
    """
    wb, wb_values = open_workbooks(path)
    cards: list[Card] = []
    warnings: list[str] = []
    page_count = 0
    for ws in wb.worksheets:
        values = wb_values[ws.title]
        for number, (top, bottom) in enumerate(_page_bounds(ws), 1):
            page_count += 1
            page = _read_page(ws, values, number, top, bottom, warnings, path)
            if page is None:
                continue
            for col in page.columns:
                card = _read_card(page, col, warnings)
                if card is not None:
                    cards.append(card)
    if page_count == 0:
        raise Matrix2ImError(
            f"Tidak ada halaman MATRIX (sel kolom A berawalan '{PAGE_MARKER}') di {path}")
    if not cards:
        detail = f": {'; '.join(warnings)}" if warnings else ""
        raise Matrix2ImError(
            f"Tidak ada kartu style yang terbaca dari {page_count} halaman MATRIX di {path}{detail}")
    return ExtractResult(cards, warnings)


def _colorway_json(colorway: Colorway) -> dict:
    return {"colorway": colorway.colorway, "yarns": list(colorway.yarns),
            "font_color": colorway.font_color, "cell": colorway.cell, "comment": colorway.comment}


def to_raw_json(result: ExtractResult) -> dict:
    """Proyeksi ke raw JSON; urutan key mengikuti golden F27 agar output bisa dibandingkan persis."""
    styles = [{
        "no": card.no,
        "style": card.style,
        "style_whs": card.style_whs,
        "tag": card.tag,
        "color_above_black_line": [_colorway_json(x) for x in card.above_black],
        "color_between_black_and_red_line": [_colorway_json(x) for x in card.between],
        "color_below_red_line": [_colorway_json(x) for x in card.below_red],
        "sheet": card.sheet,
        "section": card.section,
    } for card in result.cards]
    return {"source": SOURCE_NAME, "total_styles": len(styles), "styles": styles,
            "warnings": list(result.warnings)}


def extract(path: Path) -> dict:
    """Tahap 1 lengkap: file MATRIX -> raw JSON."""
    return to_raw_json(extract_workbook(path))

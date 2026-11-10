"""Workbook MATRIX sintetis kecil (sheet TEST, 2 halaman, 4 kartu) untuk test tanpa file F27.

Setiap sel mewakili satu aturan Tahap 1: legend di luar kolom kartu, nomor angka/teks
dengan fill berbeda, TAG dua baris, garis hitam/merah per kolom, fallback garis merah
halaman, dan sel yang harus jadi warning. Nomor ditulis sebagai nilai literal karena
openpyxl tidak menyimpan cache hasil formula.
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Border, Color, Font, PatternFill, Side

_VALUES = {
    # Halaman 1: legend (sebagian di kolom E/G, di kanan kolom label) dan section
    "A1": "FALL 27 - TEST - MATRIX MASTER (PAGE 1 OF 2)",
    "A2": "000 ONLY STYLE", "A3": "SY ONLY", "E4": "NEW COLOR", "E5": "REPEAT COLOR",
    "G5": "000 ONLY COLOR", "G6": "DISCONTINUED COLOR FOR SY", "C5": "MAX Styles for 000 is 27",
    "A8": "WINTER BEACH",
    "A9": 1, "B9": "2", "C9": "S2 - A",
    "A12": "NEWPORT CREW", "B12": "JUDE T NECK", "C12": "STRIPE CARDI",
    "A13": "NEW", "B13": "REPEAT WITH UPDATE", "C13": "EXACT REPEAT",
    "A14": "TAG: C1/C2 MARL", "B14": "TAG: C1 (For Solid)", "C14": "TAG: BLACK/LATTE/\nRED",
    "C15": "TAG: KALE",
    "A15": "YCA KHAKI - 749/YCA ALMOND BUTTER - 872", "A16": "DARK CEDAR", "A17": "SIENNA\nBOLD INK",
    "A18": "BLACK", "A19": ".",
    "B15": "NIGHT NAVY", "B16": "GLASGOW", "B18": "FAWN",
    "C16": "RED CHILE", "C17": "TAG: EXTRA", "C18": "BLUSH CHAMPAGNE STRIPE",
    "D20": "- LEBAR/WIDTH", "A20": "L: 60 cm",
    # Halaman 2: satu kartu, garis merah tepat di baris warna pertama
    "A23": "FALL 27 - TEST - MATRIX MASTER (PAGE 2 OF 2)",
    "A24": "000 ONLY STYLE", "A25": "SY ONLY", "E26": "NEW COLOR",
    "A30": 1,
    "A33": "SINGLE CARD", "A34": "EXACT REPEAT", "A35": "TAG: C1",
    "A36": "OLIVE", "A37": "SLEET",
    "B39": "- LEBAR/WIDTH",
}

_FILLS = {
    "A9": PatternFill("solid", fgColor="FFFF66FF"),
    "B9": PatternFill("solid", fgColor=Color(theme=0, tint=-0.1499984740745262)),
    "A30": PatternFill("solid", fgColor="FFFF0000"),
}

_FONTS = {
    "A15": Font(color="FF00B050"),
    "A16": Font(color=Color(theme=1)),
    "A17": Font(color="FFFF00FF"),
    "B15": Font(color="FFFF99FF"),
    "A37": Font(color="FF123456"),
}

_BORDERS = {
    "A16": Border(top=Side(style="medium")),                      # garis hitam (tanpa warna)
    "A18": Border(top=Side(style="medium", color="FFFF0000")),    # garis merah kolom A
    "C17": Border(bottom=Side(style="thick", color="FFE40500")),  # garis merah kolom C di baris 18
    "A36": Border(top=Side(style="medium", color="FFFF0000")),
}


def build_matrix(path: Path) -> Path:
    """Tulis workbook MATRIX sintetis ke `path` dan kembalikan path-nya."""
    wb = Workbook()
    ws = wb.active
    ws.title = "TEST"
    for ref, value in _VALUES.items():
        ws[ref] = value
    for ref, fill in _FILLS.items():
        ws[ref].fill = fill
    for ref, font in _FONTS.items():
        ws[ref].font = font
    for ref, border in _BORDERS.items():
        ws[ref].border = border
    ws["C16"].comment = Comment("Dev: catatan", "test")
    wb.save(path)
    return path


def build_no_pages(path: Path) -> Path:
    """Workbook tanpa sel penanda halaman, untuk menguji error fatal."""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Ini bukan MATRIX MASTER"
    ws["A2"] = "hanya teks biasa"
    wb.save(path)
    return path

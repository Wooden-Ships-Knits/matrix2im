"""Helper tingkat rendah untuk membaca warna, border, dan teks sel Excel."""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Callable

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook import Workbook

from . import legend_config as cfg
from .errors import Matrix2ImError


def open_workbooks(path: Path) -> tuple[Workbook, Workbook]:
    """Buka workbook dua kali: normal (gaya sel) dan data_only (nilai formula).

    Workbook tidak pernah di-save, jadi file input tidak berubah.
    """
    try:
        return (openpyxl.load_workbook(path),
                openpyxl.load_workbook(path, data_only=True))
    except (FileNotFoundError, zipfile.BadZipFile, InvalidFileException, KeyError, OSError) as exc:
        raise Matrix2ImError(f"File tidak bisa dibaca: {path} ({exc})") from exc


def norm(value: object) -> str:
    """Rapikan spasi/tab/NBSP tetapi pertahankan newline (dipakai sebagai pemisah di sel TAG)."""
    return re.sub(r"[ \t\xa0]+", " ", str(value)).strip()


def font_rgb(cell) -> str | None:
    """RGB font hanya bila bertipe rgb; theme/indexed dianggap hitam biasa oleh pemanggil."""
    color = cell.font.color if cell.font else None
    if color is not None and color.type == "rgb" and isinstance(color.rgb, str):
        return color.rgb
    return None


def fill_key(cell) -> str | None:
    """Kunci warna fill.

    Abu-abu angka disimpan sebagai theme 0 dengan tint sekitar -0.15 atau sebagai
    FFD9D9D9, maka theme dinormalisasi ke "theme<N><tint:+.2f>".
    """
    if cell.fill is None or cell.fill.fill_type is None:
        return None
    color = cell.fill.fgColor
    if color is None:
        return None
    if color.type == "rgb":
        return color.rgb
    if color.type == "theme":
        return f"theme{color.theme}{color.tint:+.2f}"
    if color.type == "indexed":
        return f"indexed{color.indexed}"
    return None


def is_red_line(side) -> bool:
    """Garis merah pemisah harus medium/thick/double (keputusan D1); thin diabaikan."""
    if side is None or side.style not in cfg.THICK_BORDER_STYLES:
        return False
    color = side.color
    return color is not None and color.type == "rgb" and color.rgb in cfg.RED_LINE_RGB


def is_black_line(side) -> bool:
    """Garis hitam tebal. Hitam tersimpan sebagai auto, indexed 8/64, theme 1, FF000000 atau 00000000."""
    if side is None or side.style not in cfg.THICK_BORDER_STYLES:
        return False
    color = side.color
    if color is None:
        return True
    if color.type == "auto":
        return True
    if color.type == "indexed":
        return color.indexed in cfg.BLACK_LINE_INDEXED
    if color.type == "theme":
        return color.theme in cfg.BLACK_LINE_THEME
    if color.type == "rgb":
        return color.rgb in cfg.BLACK_LINE_RGB
    return False


def line_above(ws, row: int, col: int, test: Callable) -> bool:
    """Garis di atas sel bisa tersimpan sebagai border top sel ini atau bottom sel di atasnya."""
    if test(ws.cell(row, col).border.top):
        return True
    return row > 1 and test(ws.cell(row - 1, col).border.bottom)

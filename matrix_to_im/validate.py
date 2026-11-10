"""Validasi otomatis hasil transform dan pencocokan key record dengan header IM MASTER."""
from __future__ import annotations

import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

from .errors import Matrix2ImError
from .legend_config import IM_HEADER_ROW, IM_SHEET, WHS_VALUES
from .review import Code
from .transform import GROUPS

_TAG_CODES = frozenset({Code.TAG_CONDITIONAL, Code.TAG_CN_OUT_OF_RANGE, Code.TAG_EMPTY})
_WHS_NULL_CODES = frozenset({Code.WHS_CONFLICT, Code.FONT_COLOR_NO_RULE})


def read_im_header(path: Path, sheet: str = IM_SHEET, row: int = IM_HEADER_ROW) -> set[str]:
    """Baca nama kolom IM MASTER dari satu baris header.

    Hanya baris pertama tiap sel dipakai: header "WS TAG COLOR" memuat catatan di baris
    kedua yang bukan bagian nama kolom.
    """
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except (FileNotFoundError, zipfile.BadZipFile, InvalidFileException, KeyError, OSError) as exc:
        raise Matrix2ImError(f"File tidak bisa dibaca: {path} ({exc})") from exc
    try:
        if sheet not in wb.sheetnames:
            raise Matrix2ImError(f"Sheet '{sheet}' tidak ada di IM MASTER")
        header: set[str] = set()
        for r in wb[sheet].iter_rows(min_row=row, max_row=row, values_only=True):
            for value in r:
                if value is None:
                    continue
                first = str(value).split("\n")[0].strip()
                if first:
                    header.add(first)
    finally:
        wb.close()
    if not header:
        raise Matrix2ImError(f"Header IM MASTER baris {row} kosong")
    return header


def _is_clean_value(value: object) -> bool:
    """Nilai record yang boleh ditulis ke IM: null, atau teks tidak kosong tanpa spasi di tepi.

    Teks berawalan '=' akan dibaca Excel sebagai formula, jadi juga ditolak.
    """
    return value is None or (isinstance(value, str) and value != "" and value == value.strip()
                             and not value.startswith("="))


def _count_colorways(raw: dict) -> int:
    return sum(len(style[key]) for style in raw["styles"] for key, _ in GROUPS)


def validate(raw: dict, final: dict, im_header: set[str] | None = None) -> list[str]:
    """Kembalikan semua pesan error (kosong = lolos); tidak berhenti di error pertama.

    Item DUPLICATE (record berupa list) sengaja tidak dihitung untuk pemeriksaan null:
    duplikat tidak menjelaskan kenapa TAG/WHS kosong.
    """
    errors: list[str] = []
    records = final["records"]
    expected = _count_colorways(raw)
    if expected != len(records):
        errors.append(f"jumlah colorway raw ({expected}) tidak sama dengan jumlah record ({len(records)})")

    codes_by_record: dict[int, set[str]] = {}
    for item in final.get("needs_review", []):
        if isinstance(item.get("record"), int):
            codes_by_record.setdefault(item["record"], set()).add(item["code"])

    for i, rec in enumerate(records):
        if im_header is not None:
            for key in rec:
                if not key.startswith("_") and key not in im_header:
                    errors.append(f"record {i}: key '{key}' tidak ada di header IM MASTER")
        for key, value in rec.items():
            if not key.startswith("_") and not _is_clean_value(value):
                errors.append(f"record {i}: nilai '{key}' tidak valid ({value!r}); harus null atau teks "
                              "tidak kosong tanpa spasi di awal/akhir dan tidak diawali '='")
        whs = rec.get("WHS")
        if whs is not None and whs not in WHS_VALUES:
            errors.append(f"record {i}: WHS '{whs}' bukan salah satu dari {list(WHS_VALUES)}")
        codes = codes_by_record.get(i, set())
        if rec.get("WS TAG COLOR") is None and not codes & _TAG_CODES:
            errors.append(f"record {i}: WS TAG COLOR null tanpa item needs_review kode TAG_*")
        if whs is None and not codes & _WHS_NULL_CODES:
            errors.append(f"record {i}: WHS null tanpa item needs_review "
                          f"kode WHS_CONFLICT atau FONT_COLOR_NO_RULE")
    return errors

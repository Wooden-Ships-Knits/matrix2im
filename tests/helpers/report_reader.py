"""Baca laporan .xlsx matrix2im kembali (read_only) untuk test: sel kosong menjadi "", baris tetap angka."""
from __future__ import annotations

import io

from openpyxl import load_workbook


def read_report(data: bytes) -> list[tuple]:
    """Semua baris sheet "Laporan" (header ikut), apa adanya kecuali None -> ""."""
    wb = load_workbook(io.BytesIO(data), read_only=True)
    try:
        assert wb.sheetnames == ["Laporan"], wb.sheetnames
        return [tuple("" if v is None else v for v in row) for row in wb["Laporan"].iter_rows(values_only=True)]
    finally:
        wb.close()


def report_kinds(rows: list[tuple]) -> dict[str, int]:
    """Jumlah baris per jenis (kolom terakhir), tanpa header."""
    counts: dict[str, int] = {}
    for row in rows[1:]:
        counts[row[-1]] = counts.get(row[-1], 0) + 1
    return counts

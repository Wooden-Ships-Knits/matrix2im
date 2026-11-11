"""Laporan perubahan IM MASTER (spec §6): semua perubahan terlihat sebelum konfirmasi dan tersimpan sesudahnya.

Satu format untuk pratinjau (UI/CLI) dan laporan run di folder cadangan, supaya TIM IM cukup memfilter
kolom `jenis` di Excel, tanpa batas 500 baris seperti tabel di UI. Formatnya workbook .xlsx baru, bukan CSV:
CSV berpemisah koma terbuka dalam satu kolom di Excel Windows dengan setelan regional Indonesia.

Larangan openpyxl save() berlaku untuk IM MASTER dan salinannya. Laporan ini workbook baru yang dibuat dari
nol (write-only), bukan turunan IM, jadi boleh disimpan dengan openpyxl.
"""
from __future__ import annotations

import io
from typing import TYPE_CHECKING, Iterable

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from .legend_config import IM_COLOR_KEYS, IM_WHS_KEY, IM_WS_TAG_KEY
from .write_plan import InsertOp

if TYPE_CHECKING:
    from .im_writer import WritePreview

REPORT_HEADER = ("DESCRIPTION", "sel", "baris", "kolom", "lama", "baru", "jenis")
REPORT_SHEET = "Laporan"
_BARIS = REPORT_HEADER.index("baris")


def inserted_rows(ops: list[InsertOp]) -> list[tuple[str, int]]:
    """(DESCRIPTION, baris) tiap baris sisipan, dalam koordinat file hasil (rumus sama dengan project_layout)."""
    rows: list[tuple[str, int]] = []
    for o in ops:
        start = o.at_row + sum(p.count for p in ops if p.at_row < o.at_row)
        rows.extend((o.description, r) for r in range(start, start + o.count))
    return sorted(rows, key=lambda item: item[1])


def _text(value: object) -> str:
    return "" if value is None else str(value)


def report_rows(preview: WritePreview, notes: Iterable[str] = ()) -> list[tuple[str, ...]]:
    """Baris laporan per jenis; nomor baris dalam koordinat file sesudah sisipan.

    Untuk jenis yang bukan perubahan nilai (dilewati, tidak ada di IM, IM tanpa MATRIX, peringatan),
    kolom `baru` berisi keterangan. `notes` = peringatan tambahan dari saat menulis (mis. alasan gagal).
    """
    cols = preview.projected.columns
    names = {**dict(zip(cols.colors, IM_COLOR_KEYS)), cols.ws_tag: IM_WS_TAG_KEY, cols.whs: IM_WHS_KEY}
    owner = {r: g.description for g in preview.projected.groups for r in g.rows}
    plan = preview.values
    rows: list[tuple[str, ...]] = []
    for c in plan.changes:
        letter = get_column_letter(c.col)
        rows.append((owner.get(c.row, ""), f"{letter}{c.row}", str(c.row), names.get(c.col, letter),
                     _text(c.old), _text(c.new), "ubah"))
    for description, row in inserted_rows(preview.inserts):
        rows.append((description, "", str(row), "", "", "baris sisipan (salinan baris di atasnya)", "sisip"))
    for row in plan.changed_colorway_rows:
        old, new = plan.colorway_text.get(row, ("", ""))
        rows.append((owner.get(row, ""), "", str(row), "COLOR", old, new, "baris colorway berubah"))
    rows += [(k.style, "", "", "", "", k.reason, "dilewati") for k in preview.skipped]
    rows += [(s, "", "", "", "", "tidak ditulis", "tidak ada di IM") for s in plan.missing_styles]
    rows += [(s, "", "", "", "", "tidak disentuh", "IM tanpa MATRIX") for s in plan.unmatched_im_styles]
    warnings = [*plan.warnings, *(f"grup kurang baris: {g}" for g in plan.short_groups), *notes]
    rows += [("", "", "", "", "", w, "peringatan") for w in warnings]
    return rows


def _cell(ws: object, index: int, value: str) -> object:
    """Nilai sel laporan: kosong = sel kosong, nomor baris = angka (bisa diurutkan), sisanya selalu teks."""
    if value == "":
        return None
    if index == _BARIS and value.isdigit():
        return int(value)
    cell = WriteOnlyCell(ws, value=ILLEGAL_CHARACTERS_RE.sub("", value))   # karakter kontrol ditolak xlsx
    cell.data_type = "s"     # teks berawalan "=" tidak boleh menjadi rumus; "000" tetap teks
    return cell


def report_xlsx_bytes(rows: Iterable[tuple[str, ...]]) -> bytes:
    """Workbook .xlsx baru (sheet "Laporan", header tebal) berisi semua baris laporan."""
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(REPORT_SHEET)
    header = []
    for name in REPORT_HEADER:
        cell = WriteOnlyCell(ws, value=name)
        cell.font = Font(bold=True)
        header.append(cell)
    ws.append(header)
    for row in rows:
        ws.append([_cell(ws, i, value) for i, value in enumerate(row)])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

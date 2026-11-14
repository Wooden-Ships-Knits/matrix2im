"""Excel untuk dicek tim: hanya colorway di atas garis hitam kartu MATRIX (permintaan user 14 Nov 2026).

Dibangun dari raw + final hasil Proses yang sama. Record final dipasangkan ke colorway raw lewat urutan
pembuatan di `transform` (style -> grup garis -> colorway), dan setiap pasangan dicek lewat `_ref`
(sel + posisi garis). Bila tidak cocok, export gagal; data tidak boleh tertukar diam-diam.

Sheet kedua mendaftar kartu yang punya warna di atas garis merah tapi tidak satu pun di atas garis
hitam. Warna kartu-kartu itu tidak muncul di sheet pertama, jadi tim tahu harus mengecek manual.
Workbook ini dibuat baru dari nol (write-only), bukan turunan IM, jadi boleh disimpan dengan openpyxl.
"""
from __future__ import annotations

import io
import re
from collections import defaultdict
from collections.abc import Iterator, Sequence

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from .errors import Matrix2ImError
from .extract import NO_BLACK_LINE
from .legend_config import IM_WHS_KEY, IM_WS_TAG_KEY, color_key
from .transform import GROUPS

ABOVE_SHEET = "Atas garis hitam"
MISSING_SHEET = "Tanpa warna atas garis hitam"
ABOVE_KEY = "color_above_black_line"
BETWEEN_KEY = "color_between_black_and_red_line"
CARD_HEADER = ("COLLECTION", "SUB", "No", "DESCRIPTION")
TAIL_HEADER = (IM_WS_TAG_KEY, IM_WHS_KEY, "SEL MATRIX", "NEEDS REVIEW")
MISSING_HEADER = (*CARD_HEADER, "JUMLAH WARNA ANTARA HITAM & MERAH", "KETERANGAN")
REASON_NO_LINE = "garis hitam tidak ditemukan"
REASON_ALL_BELOW = "semua warna di bawah garis hitam"
NONE_ROW = "(tidak ada)"
_CARD_WIDTHS = (22, 24, 8, 45)
_COLOR_WIDTH = 24
_TAIL_WIDTHS = (34, 10, 11, 60)
_MISSING_WIDTHS = (*_CARD_WIDTHS, 18, 34)


def _pairs(raw: dict, final: dict) -> Iterator[tuple[dict, dict, int, dict]]:
    """(style, colorway, indeks record, record) untuk setiap colorway di atas garis hitam."""
    records = final["records"]
    idx = 0
    for style in raw["styles"]:
        for key, group in GROUPS:
            for cw in style[key]:
                rec = records[idx] if idx < len(records) else None
                ref = None if rec is None else str(rec.get("_ref", ""))
                if ref is None or not ref.endswith(f" | {cw['cell']} | {group}"):
                    raise Matrix2ImError(
                        f"final.json tidak cocok dengan raw.json di record #{idx} ({ref!r}, seharusnya sel "
                        f"{cw['cell']} {group}); jalankan Proses ulang")
                if key == ABOVE_KEY:
                    yield style, cw, idx, rec
                idx += 1
    if idx != len(records):
        raise Matrix2ImError(f"final.json berisi {len(records)} record, raw.json {idx} colorway; "
                             "jalankan Proses ulang")


def _colors(rec: dict) -> list[str]:
    """COLOR1..COLORn berurutan (transform menulisnya tanpa celah)."""
    out, n = [], 1
    while color_key(n) in rec:
        out.append(rec[color_key(n)])
        n += 1
    return out


def above_black_rows(raw: dict, final: dict) -> tuple[list[str], list[list]]:
    """Header + 1 baris per colorway di atas garis hitam, urut seperti MATRIX."""
    issues: dict[int, list[str]] = defaultdict(list)
    for item in final.get("needs_review", []):
        if isinstance(item.get("record"), int):
            issues[item["record"]].append(str(item["issue"]))
    picked = [(style, cw, idx, rec, _colors(rec)) for style, cw, idx, rec in _pairs(raw, final)]
    width = max([1, *(len(colors) for *_, colors in picked)])
    header = [*CARD_HEADER, *(color_key(n) for n in range(1, width + 1)), *TAIL_HEADER]
    rows = [[style["sheet"], style["section"], style["no"], rec["DESCRIPTION"],
             *colors, *[None] * (width - len(colors)),
             rec.get(IM_WS_TAG_KEY), rec.get(IM_WHS_KEY), cw["cell"], "; ".join(issues[idx]) or None]
            for style, cw, idx, rec, colors in picked]
    return header, rows


def missing_above_rows(raw: dict) -> list[list]:
    """Kartu dengan warna di atas garis merah tapi tidak satu pun di atas garis hitam."""
    warnings = [w for w in raw.get("warnings", []) if w.endswith(f": {NO_BLACK_LINE}")]
    rows = []
    for style in raw["styles"]:
        if style[ABOVE_KEY] or not style[BETWEEN_KEY]:
            continue
        exact = re.compile(rf"{re.escape(style['sheet'])} [A-Z]{{1,3}}\d+ {re.escape(style['style'])}: "
                           rf"{re.escape(NO_BLACK_LINE)}")
        reason = REASON_NO_LINE if any(exact.fullmatch(w) for w in warnings) else REASON_ALL_BELOW
        rows.append([style["sheet"], style["section"], style["no"], style["style"],
                     len(style[BETWEEN_KEY]), reason])
    return rows


def _cell(ws: object, value: object) -> object:
    """Angka tetap angka; teks selalu teks ("000" tidak jadi 0, "=..." tidak jadi rumus)."""
    if value is None or value == "":
        return None
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    cell = WriteOnlyCell(ws, value=ILLEGAL_CHARACTERS_RE.sub("", str(value)))   # karakter kontrol ditolak xlsx
    cell.data_type = "s"
    return cell


def _sheet(wb: Workbook, title: str, header: Sequence[str], rows: list[list], widths: Sequence[float]) -> None:
    ws = wb.create_sheet(title)
    ws.freeze_panes = "A2"
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{len(rows) + 1}"
    bold = []
    for name in header:
        cell = WriteOnlyCell(ws, value=name)
        cell.font = Font(bold=True)
        bold.append(cell)
    ws.append(bold)
    for row in rows:
        ws.append([_cell(ws, v) for v in row])


def above_black_xlsx_bytes(raw: dict, final: dict) -> bytes:
    """Workbook .xlsx: sheet "Atas garis hitam" + sheet "Tanpa warna atas garis hitam"."""
    header, rows = above_black_rows(raw, final)
    width = len(header) - len(CARD_HEADER) - len(TAIL_HEADER)
    wb = Workbook(write_only=True)
    _sheet(wb, ABOVE_SHEET, header, rows, (*_CARD_WIDTHS, *[_COLOR_WIDTH] * width, *_TAIL_WIDTHS))
    missing = missing_above_rows(raw) or [[NONE_ROW]]
    _sheet(wb, MISSING_SHEET, MISSING_HEADER, missing, _MISSING_WIDTHS)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

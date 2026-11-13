"""Tata letak sheet IM MASTER: kolom menurut nama header, grup size, nilai sel target."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .errors import Matrix2ImError
from .xlsx_cells import read_fills
from .legend_config import (
    IM_COLOR_HEADER_RE, IM_DESCRIPTION_KEY, IM_GROUP_KEY_SEGMENTS, IM_HEADER_ROW, IM_SHEET, IM_SUBTOTAL_FUNC,
    IM_WHS_KEY, IM_WS_TAG_KEY, color_key,
)

_COLOR_HEADER = re.compile(IM_COLOR_HEADER_RE)


def normalize(text: object) -> str:
    """Rapatkan spasi, strip, huruf besar: kunci pembanding DESCRIPTION."""
    return re.sub(r"\s+", " ", "" if text is None else str(text)).strip().upper()


@dataclass(frozen=True)
class ImColumns:
    description: int
    colors: tuple[int, ...]
    ws_tag: int
    whs: int
    last: int


@dataclass(frozen=True)
class ImGroup:
    description: str
    rows: tuple[int, ...]
    subtotal_row: int | None


@dataclass
class ImLayout:
    columns: ImColumns
    groups: list[ImGroup]
    values: dict[tuple[int, int], object]
    subtotal_formulas: dict[int, dict[int, str]]
    # Warna latar COLOR1..n, WS TAG COLOR, dan WHS per baris grup (kunci xlsx_styles.fill_key: "FF99FF", None = tanpa warna).
    fills: dict[tuple[int, int], str | None] = field(default_factory=dict)


def _find(header: dict[str, int], key: str) -> int:
    if key not in header:
        raise Matrix2ImError(f"Kolom '{key}' tidak ditemukan di header IM MASTER")
    return header[key]


def _color_columns(header: dict[str, int]) -> tuple[int, ...]:
    """Kolom COLOR1..COLORn menurut nomor di header, berapa pun n (IM bisa ditambah kolom COLOR19..)."""
    found: dict[int, int] = {}
    for name, col in header.items():
        if m := _COLOR_HEADER.match(name):
            found.setdefault(int(m.group(1)), col)
    n = max(found, default=1)
    missing = [k for k in range(1, n + 1) if k not in found]
    if missing:
        raise Matrix2ImError(f"Kolom '{color_key(missing[0])}' tidak ditemukan di header IM MASTER "
                             f"(kolom COLOR harus berurutan dari COLOR1 sampai COLOR{n})")
    return tuple(found[k] for k in range(1, n + 1))


def read_layout(path: Path, sheet: str = IM_SHEET, header_row: int = IM_HEADER_ROW) -> ImLayout:
    """Satu kali lewat baris (read_only) agar sheet ~10 ribu baris x 244 kolom tetap ringan."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception as exc:
        raise Matrix2ImError(f"File tidak bisa dibaca: {path} ({exc})") from exc
    try:
        if sheet not in wb.sheetnames:
            raise Matrix2ImError(f"Sheet '{sheet}' tidak ada di IM MASTER")
        ws = wb[sheet]
        last = ws.max_column
        header: dict[str, int] = {}
        columns: ImColumns | None = None
        func_re = re.compile(rf"=\s*{re.escape(IM_SUBTOTAL_FUNC)}\(", re.IGNORECASE)
        groups: list[ImGroup] = []
        values: dict[tuple[int, int], object] = {}
        formulas: dict[int, dict[int, str]] = {}
        target_cols: tuple[int, ...] = ()
        cur_desc: str | None = None
        cur_rows: list[int] = []

        def close(subtotal_row: int | None) -> None:
            nonlocal cur_desc, cur_rows
            if cur_desc is not None:
                groups.append(ImGroup(cur_desc, tuple(cur_rows), subtotal_row))
            cur_desc, cur_rows = None, []

        for r, row in enumerate(ws.iter_rows(min_row=header_row, values_only=True), header_row):
            if r == header_row:
                for c, v in enumerate(row, 1):
                    if v is not None:
                        header.setdefault(str(v).split("\n")[0].strip(), c)
                ws_tag = _find(header, IM_WS_TAG_KEY)
                columns = ImColumns(
                    description=_find(header, IM_DESCRIPTION_KEY),
                    colors=_color_columns(header),
                    ws_tag=ws_tag, whs=ws_tag + 1, last=last,
                )
                nxt = row[ws_tag] if len(row) > ws_tag else None
                if nxt is None or str(nxt).split("\n")[0].strip() != IM_WHS_KEY:
                    raise Matrix2ImError("Kolom WHS tidak ditemukan tepat di kanan WS TAG COLOR")
                target_cols = (*columns.colors, columns.ws_tag, columns.whs)
                continue
            desc_raw = row[columns.description - 1] if len(row) >= columns.description else None
            desc = " - ".join(normalize(desc_raw).split(" - ")[:IM_GROUP_KEY_SEGMENTS])
            row_formulas = {c: v for c, v in enumerate(row, 1) if isinstance(v, str) and v.startswith("=")}
            if any(func_re.match(v) for v in row_formulas.values()):
                close(r)
                formulas[r] = row_formulas
                continue
            if not desc:
                close(None)
                continue
            if desc != cur_desc:
                close(None)
                cur_desc = desc
            cur_rows.append(r)
            for c in target_cols:
                values[(r, c)] = row[c - 1] if len(row) >= c else None
        close(None)
    finally:
        wb.close()
    if columns is None:
        raise Matrix2ImError(f"Header IM MASTER baris {header_row} kosong")
    try:
        fill_cols = (*columns.colors, columns.ws_tag, columns.whs)
        fills = read_fills(path, sheet, [(r, c) for g in groups for r in g.rows for c in fill_cols])
    except Matrix2ImError:
        raise
    except Exception as exc:
        raise Matrix2ImError(f"Warna sel tidak bisa dibaca: {path} ({exc})") from exc
    return ImLayout(columns, groups, values, formulas, fills)

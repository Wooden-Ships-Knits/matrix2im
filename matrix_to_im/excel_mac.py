"""Sisip baris di IM MASTER lewat Microsoft Excel for Mac (osascript).

Hanya Excel yang menggeser rumus, SUBTOTAL, format bersyarat, komentar, dan gambar dengan benar
saat baris disisipkan; modul ini hanya mengurus struktur pada salinan sementara, nilai ditulis modul lain.
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

from openpyxl.utils import get_column_letter

from .errors import Matrix2ImError
from .legend_config import EXCEL_APP, EXCEL_TIMEOUT_S
from .write_plan import InsertOp


class ExcelUnavailable(Matrix2ImError):
    """Microsoft Excel for Mac tidak ada di komputer ini."""


def excel_available() -> bool:
    return sys.platform == "darwin" and Path(EXCEL_APP).exists()


def _esc(text: str) -> str:
    return text.replace("\\", "\\\\").replace('"', '\\"')


def _runs(cols: list[int]) -> list[list[int]]:
    runs: list[list[int]] = []
    for c in sorted(cols):
        if runs and runs[-1][-1] == c - 1:
            runs[-1].append(c)
        else:
            runs.append([c])
    return runs


def _op_lines(op: InsertOp, last_col: str) -> list[str]:
    lines = [
        f'      insert into range (range "{op.at_row}:{op.at_row + op.count - 1}" of ws) shift shift down',
        f'      fill down range "A{op.fill_from}:{last_col}{op.fill_to}" of ws',
    ]
    # Rumus diperbaiki di sini juga: subtotal_row_after hanya berlaku sebelum op yang lebih atas dijalankan.
    if op.formula_fixes and op.subtotal_row_after is not None:
        row = op.subtotal_row_after
        for run in _runs(list(op.formula_fixes)):
            vals = ", ".join(f'"{_esc(op.formula_fixes[c])}"' for c in run)
            rng = f"{get_column_letter(run[0])}{row}:{get_column_letter(run[-1])}{row}"
            lines.append(f'      set formula of range "{rng}" of ws to {{{{{vals}}}}}')
    return lines


def build_insert_script(ops: list[InsertOp], sheet: str, last_col: str,
                        timeout_s: int = EXCEL_TIMEOUT_S) -> str:
    body = [line for op in ops for line in _op_lines(op, last_col)]
    return "\n".join([
        "on run argv",
        "  set f to (POSIX file (item 1 of argv)) as text",
        '  set pw to system attribute "MATRIX2IM_PW"',
        "  set wb to missing value",
        f"  with timeout of {timeout_s} seconds",
        '  tell application "Microsoft Excel"',
        "    try",
        "      open workbook workbook file name f write reserved password pw "
        "update links do not update links ignore read only recommended true",
        "      set wb to active workbook",
        f'      set ws to worksheet "{_esc(sheet)}" of wb',
        "      set screen updating to false",
        "      set calculation to calculation manual",
        *body,
        "      set calculation to calculation automatic",
        "      set screen updating to true",
        "      save wb",
        "      close wb saving no",
        '      return "OK"',
        "    on error errMsg number errNum",
        "      try",
        "        set calculation to calculation automatic",
        "        set screen updating to true",
        "      end try",
        "      try",
        "        if wb is not missing value then close wb saving no",
        "      end try",
        '      return "ERROR " & errNum & ": " & errMsg',
        "    end try",
        "  end tell",
        "  end timeout",
        "end run",
        "",
    ])


def run_excel_inserts(path: Path, password: str, ops: list[InsertOp], sheet: str, last_col: str,
                      timeout_s: int = EXCEL_TIMEOUT_S) -> float:
    """Jalankan sisipan di path; password hanya lewat environment, bukan argumen atau skrip."""
    if not excel_available():
        raise ExcelUnavailable("Microsoft Excel for Mac tidak ditemukan di komputer ini.")
    script = build_insert_script(ops, sheet, last_col, timeout_s)
    start = time.monotonic()
    try:
        proc = subprocess.run(["osascript", "-", str(path)], input=script,
                              env={**os.environ, "MATRIX2IM_PW": password},
                              capture_output=True, text=True, timeout=timeout_s + 15)
    except subprocess.TimeoutExpired as exc:
        raise Matrix2ImError(
            f"Excel tidak merespons dalam {timeout_s} detik. "
            "Periksa apakah ada dialog terbuka di Excel PC pusat.") from exc
    out = (proc.stdout or "").strip()
    if proc.returncode != 0 or out.startswith("ERROR"):
        detail = out if out.startswith("ERROR") else (proc.stderr or out).strip()
        msg = f"Excel gagal: {detail}"
        if "-1743" in detail:
            msg += (" Izinkan Terminal mengontrol Microsoft Excel di "
                    "System Settings → Privacy & Security → Automation.")
        raise Matrix2ImError(msg)
    return time.monotonic() - start

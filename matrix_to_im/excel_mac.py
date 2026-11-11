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
    """Skrip osascript; argv = (path POSIX ter-resolve, nama file). Hanya workbook yang dibuka skrip ini disentuh."""
    body = [line for op in ops for line in _op_lines(op, last_col)]
    return "\n".join([
        "on run argv",
        "  set p to item 1 of argv",
        "  set nm to item 2 of argv",
        "  set f to (POSIX file p) as text",
        '  set pw to system attribute "MATRIX2IM_PW"',
        "  set wb to missing value",
        f"  with timeout of {timeout_s} seconds",
        '  tell application "Microsoft Excel"',
        # Excel tidak bisa membuka dua workbook bernama sama: open akan gagal atau menyerahkan workbook lain.
        "    repeat with w in (get every workbook)",
        '      if (name of w) is nm then return "ERROR 0: Excel sudah membuka workbook bernama " & nm & "; tutup dulu"',
        "    end repeat",
        # Alert Excel di PC pusat tanpa pengawas memblokir Excel dan server. Yang diperiksa commit_write sesudah
        # langkah ini: tata letak grup, rencana nilai, fileSharing, daftar media/drawing/comment/link/embedding;
        # jumlah conditional formatting hanya diperingatkan.
        "    set display alerts to false",
        "    try",
        "      open workbook workbook file name f write reserved password pw "
        "update links do not update links ignore read only recommended true",
        # Bukan active workbook: hanya workbook ber-full name path kita yang boleh disimpan atau ditutup.
        "      repeat with w in (get every workbook)",
        "        if (full name of w) is p then set wb to contents of w",
        "      end repeat",
        '      if wb is missing value then error "workbook yang dibuka tidak ditemukan di Excel: " & p',
        f'      set ws to worksheet "{_esc(sheet)}" of wb',
        "      set screen updating to false",
        "      set calculation to calculation manual",
        *body,
        "      set calculation to calculation automatic",
        "      set screen updating to true",
        "      save wb",
        "      close wb saving no",
        "      set wb to missing value",
        "      try",
        "        set display alerts to true",
        "      end try",
        '      return "OK"',
        "    on error errMsg number errNum",
        # Tiap pemulihan di try sendiri (alert dulu): satu yang gagal tak boleh membuat alert tetap mati.
        "      try",
        "        set display alerts to true",
        "      end try",
        "      try",
        "        set calculation to calculation automatic",
        "      end try",
        "      try",
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


def _close_script(timeout_s: int) -> str:
    return "\n".join([
        "on run argv",
        "  set p to item 1 of argv",
        '  if application "Microsoft Excel" is not running then return "NONE"',
        f"  with timeout of {max(timeout_s - 5, 1)} seconds",
        '  tell application "Microsoft Excel"',
        "    set display alerts to false",     # alasan sama dengan skrip sisip
        "    try",                             # daftar workbook gagal dibaca pun, alert tetap dinyalakan lagi
        "      repeat with wb in (get every workbook)",
        "        try",
        "          if (full name of wb) is p then close wb saving no",
        "        end try",
        "      end repeat",
        "    end try",
        "    try",
        "      set display alerts to true",
        "    end try",
        "  end tell",
        "  end timeout",
        '  return "OK"',
        "end run",
        "",
    ])


def close_workbook_if_open(path: Path, timeout_s: int = 30) -> None:
    """Tutup (tanpa simpan) hanya workbook sementara kita bila Excel masih menahannya.

    Bila osascript mati karena timeout, Excel tetap bekerja dan workbook yang terbuka memblokir
    open workbook berikutnya tanpa pesan. Best effort: tidak pernah melempar, tidak menyentuh
    workbook lain, tidak menutup Excel. Path lewat argv, bukan disisipkan ke skrip.
    """
    try:
        subprocess.run(["osascript", "-", str(path.resolve())], input=_close_script(timeout_s),
                       capture_output=True, text=True, timeout=timeout_s)
    except Exception:
        pass


def run_excel_inserts(path: Path, password: str, ops: list[InsertOp], sheet: str, last_col: str,
                      timeout_s: int = EXCEL_TIMEOUT_S) -> float:
    """Jalankan sisipan di path; password hanya lewat environment, bukan argumen atau skrip.

    Nama file ikut dikirim agar skrip bisa menolak bila Excel sudah membuka workbook lain bernama sama.
    """
    if not excel_available():
        raise ExcelUnavailable("Microsoft Excel for Mac tidak ditemukan di komputer ini.")
    script = build_insert_script(ops, sheet, last_col, timeout_s)
    start = time.monotonic()
    try:
        proc = subprocess.run(["osascript", "-", str(path.resolve()), path.name], input=script,
                              env={**os.environ, "MATRIX2IM_PW": password},
                              capture_output=True, text=True, timeout=timeout_s + 15)
    except subprocess.TimeoutExpired as exc:
        close_workbook_if_open(path)
        raise Matrix2ImError(
            f"Excel tidak merespons dalam {timeout_s} detik. "
            "Periksa apakah ada dialog terbuka di Excel PC pusat.") from exc
    out = (proc.stdout or "").strip()
    if proc.returncode != 0 or out.startswith("ERROR"):
        close_workbook_if_open(path)    # juga saat osascript mati/crash: workbook bisa tertinggal terbuka
        detail = out if out.startswith("ERROR") else (proc.stderr or out).strip()
        msg = f"Excel gagal: {detail}" if detail else f"Excel berhenti tanpa pesan (kode {proc.returncode})"
        if "-1743" in detail:
            msg += (" Izinkan Terminal mengontrol Microsoft Excel di "
                    "System Settings → Privacy & Security → Automation.")
        raise Matrix2ImError(msg)
    return time.monotonic() - start

"""Orkestrasi penulisan final.json ke IM MASTER: rencana (dry-run) dan eksekusi berlapis verifikasi.

File asli hanya disentuh pada langkah terakhir, ditimpa di tempat; semua kerja lain di salinan sementara.
"""
from __future__ import annotations

import binascii
import hashlib
import json
import logging
import os
import shutil
import tempfile
import time
import traceback
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable
from xml.etree import ElementTree as ET

from openpyxl.utils import get_column_letter

from .errors import Matrix2ImError
from .excel_mac import excel_available, run_excel_inserts
from .im_layout import ImLayout, read_layout
from .im_lock import check_modify_password, owner_file, owner_name
from .legend_config import IM_SHEET
from .settings import Settings
from .write_plan import (
    InsertOp, Skip, ValuePlan, match_style, plan_inserts, plan_values, project_layout, style_colorways,
)
from .xlsx_cells import apply_cell_values, sheet_xml_path

LOG_NAME = "matrix2im-log.jsonl"
_CHANGED_MSG = "IM MASTER berubah atau sedang dibuka sejak rencana dibuat; buat rencana lagi"
_UNREADABLE = (zipfile.BadZipFile, KeyError, ValueError, binascii.Error, ET.ParseError)


class ImMasterBusy(Matrix2ImError):
    """IM MASTER sedang dibuka orang lain; UI memetakannya ke 409 tanpa mencocokkan teks pesan."""


@dataclass(frozen=True)
class Fingerprint:
    """Sidik file IM MASTER saat rencana dibuat; beda berarti ada orang lain yang mengubahnya."""
    size: int
    mtime_ns: int
    sha256: str


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def fingerprint(path: Path) -> Fingerprint:
    """Ukuran, mtime, dan sha256: pembanding murah-tapi-pasti bahwa file tidak berubah sejak rencana."""
    st = path.stat()
    return Fingerprint(st.st_size, st.st_mtime_ns, _sha256(path))


@dataclass
class WritePreview:
    """Hasil dry-run: apa yang akan disisipkan dan diubah, dan sidik file yang dijadikan dasar rencana."""
    im_master: Path
    fingerprint: Fingerprint
    engine: str
    inserts: list[InsertOp]
    projected: ImLayout
    values: ValuePlan
    skipped: list[Skip]
    summary: dict


@dataclass
class WriteReport:
    """Hasil commit untuk UI/CLI; semua teks Indonesia dan bebas password."""
    ok: bool
    message: str
    summary: dict
    backup: Path | None
    log_file: Path
    seconds: float
    inserted_rows: list[int]


def _check_password(path: Path, password: str) -> None:
    """Password cocok; file rusak/bukan xlsx dijadikan pesan Indonesia, bukan error mentah zip/XML."""
    try:
        check_modify_password(path, password)
    except _UNREADABLE as exc:
        raise Matrix2ImError(f"IM MASTER tidak bisa dibaca sebagai file Excel: {exc}") from exc


def _no_owner(path: Path) -> None:
    owner = owner_file(path)
    if owner is not None:
        raise ImMasterBusy(f"IM MASTER sedang dibuka oleh {owner_name(owner)}. "
                           "Minta yang bersangkutan menutup file, lalu coba lagi.")


def prepare_write(final: dict, settings: Settings, password: str) -> WritePreview:
    """Dry-run: hitung sisipan dan perubahan sel tanpa mengubah apa pun."""
    if settings.im_master is None:
        raise Matrix2ImError("Setting im_master belum diisi")
    if settings.backup_dir is None:
        raise Matrix2ImError("Setting backup_dir belum diisi")
    path = settings.im_master
    if not path.is_file():
        raise Matrix2ImError(f"IM MASTER tidak ditemukan: {path}")
    _check_password(path, password)
    _no_owner(path)
    fp = fingerprint(path)
    layout = read_layout(path)
    colorways, skipped = style_colorways(final)
    warnings: list[str] = []
    engine = "none"
    if settings.insert_engine == "excel":
        if excel_available():
            engine = "excel"
        else:
            warnings.append("Excel tidak tersedia: hanya mengisi baris yang sudah ada")
    ops = plan_inserts(layout, colorways) if engine == "excel" else []
    projected = project_layout(layout, ops)
    values = plan_values(projected, colorways)
    group_size = {r: len(g.rows) for g in layout.groups for r in g.rows}
    for op in ops:
        if group_size.get(op.fill_from) == 1:
            warnings.append(f"{op.description}: grup hanya 1 baris, baris disisipkan di bawahnya; rumus di luar "
                            "baris subtotal yang merujuk baris itu tidak ikut melebar")
    values.warnings.extend(warnings)
    matched = [match_style(g.description, colorways) for g in layout.groups]
    summary = {
        "styles": len({m for m in matched if m is not None}),
        "groups": sum(m is not None for m in matched),
        "cells_changed": len(values.changes),
        "rows_inserted": sum(o.count for o in ops),
        "groups_inserted": len(ops),
        "changed_colorway_rows": len(values.changed_colorway_rows),
        "short_groups": len(values.short_groups),
        "skipped": len(skipped),
        "unmatched_im_styles": len(values.unmatched_im_styles),
        "missing_styles": len(values.missing_styles),
        "warnings": len(values.warnings),
        "engine": engine,
    }
    return WritePreview(path, fp, engine, ops, projected, values, skipped, summary)


def _inserted_rows(ops: list[InsertOp]) -> list[int]:
    """Baris baru hasil sisipan, dalam koordinat akhir (rumus yang sama dengan project_layout)."""
    rows: list[int] = []
    for o in ops:
        start = o.at_row + sum(p.count for p in ops if p.at_row < o.at_row)
        rows.extend(range(start, start + o.count))
    return sorted(rows)


def _structure(layout: ImLayout) -> list[tuple]:
    return [(g.description, g.rows, g.subtotal_row) for g in layout.groups]


def _verify_written(tmp: Path, tmp2: Path, password: str, changes: list) -> None:
    layout = read_layout(tmp2)
    for ch in changes:
        got = layout.values.get((ch.row, ch.col))
        blank = got is None or str(got).strip() == ""
        if (ch.new is None and not blank) or (ch.new is not None and (blank or str(got) != ch.new)):
            raise Matrix2ImError(f"Verifikasi gagal: sel {get_column_letter(ch.col)}{ch.row} tidak sesuai rencana")
    _check_password(tmp2, password)
    with zipfile.ZipFile(tmp) as a, zipfile.ZipFile(tmp2) as b:
        skip = {sheet_xml_path(a, IM_SHEET), "xl/workbook.xml"}
        crc_a = {i.filename: i.CRC for i in a.infolist() if i.filename not in skip}
        crc_b = {i.filename: i.CRC for i in b.infolist() if i.filename not in skip}
    if crc_a != crc_b:
        raise Matrix2ImError("Verifikasi gagal: bagian lain dari file berubah selain sheet dan workbook.xml")


def _write_bytes_in_place(target: Path, data: bytes) -> None:
    with open(target, "r+b") as f:
        f.write(data)
        f.truncate()
        f.flush()
        os.fsync(f.fileno())


def _overwrite_in_place(target: Path, source: Path, backup: Path, password: str = "") -> None:
    """Timpa isi file asli di tempat (identitas file di Google Drive tetap); pulihkan dari cadangan bila gagal.

    Pesan jujur: "tidak diubah" hanya bila belum ada byte tertulis, "dipulihkan" hanya bila sha256 cocok cadangan.
    """
    data = source.read_bytes()
    touched = False
    try:
        with open(target, "r+b") as f:
            touched = True              # mulai dari sini file asli bisa sudah berubah
            f.write(data)
            f.truncate()
            f.flush()
            os.fsync(f.fileno())
        if _sha256(target) != hashlib.sha256(data).hexdigest():
            raise Matrix2ImError("hasil tulis tidak cocok dengan data yang dikirim")
    except Exception as exc:
        detail = str(exc).replace(password, "***") if password else str(exc)
        if not touched:
            note = "File asli tidak diubah."
        else:
            try:
                _write_bytes_in_place(target, backup.read_bytes())
                ok = _sha256(target) == _sha256(backup)
            except Exception:  # noqa: BLE001 - pesan pemulihan manual harus selalu sampai
                ok = False
            note = ("File asli sudah dipulihkan dari cadangan." if ok else
                    f"File asli mungkin rusak; pulihkan manual dari cadangan {backup}.")
        raise Matrix2ImError(f"Gagal menyimpan ke IM MASTER ({detail}). {note}") from exc


def _log(log_file: Path, entry: dict) -> bool:
    try:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        return True
    except OSError:
        return False


def _execute(preview: WritePreview, final: dict, settings: Settings, password: str,
             progress: Callable[[str], None], inserter: Callable[..., float], state: dict) -> None:
    """Langkah 1-9; state["backup"] diisi segera setelah cadangan dibuat agar tercatat walau gagal sesudahnya."""
    path, backup_dir = preview.im_master, settings.backup_dir
    if not path.is_file():
        raise Matrix2ImError(f"IM MASTER tidak ditemukan: {path}")
    progress("cek password")
    _check_password(path, password)
    if fingerprint(path) != preview.fingerprint or owner_file(path) is not None:
        raise Matrix2ImError(_CHANGED_MSG)
    progress("cadangan")
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = backup_dir / f"{path.stem} {datetime.now():%Y-%m-%d %H%M%S} sebelum matrix2im{path.suffix}"
    shutil.copyfile(path, backup)
    if _sha256(backup) != preview.fingerprint.sha256:
        backup.unlink(missing_ok=True)
        raise Matrix2ImError(_CHANGED_MSG)
    state["backup"] = backup
    parent = backup_dir / "kerja"
    parent.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="kerja-", dir=parent))
    tmp, tmp2 = work / f"im-kerja{path.suffix}", work / f"im-hasil{path.suffix}"
    try:
        shutil.copyfile(backup, tmp)     # backup sudah lolos cek sha256: tmp pasti isi yang direncanakan
        if preview.inserts:
            progress("Excel: sisip baris")
            inserter(tmp, password, preview.inserts, IM_SHEET, get_column_letter(preview.projected.columns.last))
        real = read_layout(tmp)
        if _structure(real) != _structure(preview.projected):
            raise Matrix2ImError("Tata letak IM MASTER setelah sisip baris tidak sama dengan rencana; "
                                 "penulisan dibatalkan, file asli tidak diubah")
        progress("tulis nilai")
        colorways, _ = style_colorways(final)
        planned = {(c.row, c.col, c.new) for c in preview.values.changes}
        if {(c.row, c.col, c.new) for c in plan_values(real, colorways).changes} != planned:
            raise Matrix2ImError("Perubahan nilai pada file hasil sisip tidak sama dengan rencana; "
                                 "penulisan dibatalkan, file asli tidak diubah")
        apply_cell_values(tmp, tmp2, IM_SHEET, preview.values.changes)
        progress("verifikasi")
        _verify_written(tmp, tmp2, password, preview.values.changes)
        progress("simpan ke IM MASTER")
        if fingerprint(path) != preview.fingerprint or owner_file(path) is not None:
            raise Matrix2ImError(_CHANGED_MSG)
        _overwrite_in_place(path, tmp2, backup, password)
    finally:
        shutil.rmtree(work, ignore_errors=True)   # best effort; folder milik run ini, tak boleh mengubah hasil


def commit_write(preview: WritePreview, final: dict, settings: Settings, password: str, user: str,
                 matrix_name: str, progress: Callable[[str], None] = lambda s: None,
                 inserter: Callable[..., float] = run_excel_inserts) -> WriteReport:
    """Jalankan rencana. Kegagalan sebelum langkah simpan tidak menyentuh file asli; semua hasil dicatat di log."""
    start = time.monotonic()
    if settings.backup_dir is None:
        raise Matrix2ImError("Setting backup_dir belum diisi")
    log_file = settings.backup_dir / LOG_NAME
    state: dict = {"backup": None}
    try:
        _execute(preview, final, settings, password, progress, inserter, state)
        ok = True
        message = (f"IM MASTER diperbarui: {preview.summary['cells_changed']} sel diubah, "
                   f"{preview.summary['rows_inserted']} baris disisipkan.")
    except Matrix2ImError as exc:
        ok, message = False, str(exc)
    except Exception as exc:  # noqa: BLE001 - apa pun harus berakhir sebagai laporan, bukan crash
        ok, message = False, f"Gagal tak terduga: {type(exc).__name__}: {exc}"
        trace = "".join(traceback.format_exception(exc))
        logging.getLogger(__name__).error("commit_write gagal: %s",
                                          trace.replace(password, "***") if password else trace)
    if password:
        message = message.replace(password, "***")
    backup = state["backup"]
    entry = {"waktu": datetime.now().isoformat(timespec="seconds"), "user": user, "matrix": matrix_name,
             "im_master": str(preview.im_master), "engine": preview.engine, "summary": preview.summary,
             "backup": str(backup) if backup else None, "hasil": "ok" if ok else "gagal", "pesan": message}
    if not _log(log_file, entry):
        message += " (catatan log gagal ditulis)"
    if ok:
        progress("selesai")
    return WriteReport(ok, message, preview.summary, backup, log_file, time.monotonic() - start,
                       _inserted_rows(preview.inserts) if ok else [])

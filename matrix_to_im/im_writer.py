"""Orkestrasi penulisan final.json ke IM MASTER: rencana (dry-run) dan eksekusi berlapis verifikasi.

File asli hanya disentuh pada langkah terakhir, ditimpa di tempat; semua kerja lain di salinan sementara.
"""
from __future__ import annotations

import binascii
import hashlib
import json
import logging
import os
import re
import secrets
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
from .im_lock import FileSharing, check_modify_password, owner_file, owner_name, read_file_sharing
from .legend_config import IM_KEEP_PARTS, IM_SHEET
from .settings import Settings
from .write_plan import (
    InsertOp, Skip, ValuePlan, match_style, plan_inserts, plan_values, project_layout, style_colorways,
)
from .write_report import inserted_rows, report_rows, report_xlsx_bytes
from .xlsx_cells import apply_cell_values, sheet_xml_path, styles_xml_path
from .xlsx_styles import check_appended

try:
    import fcntl
except ImportError:   # bukan Unix: penulisan IM hanya dari PC pusat (Mac), kunci antarproses dilewati
    fcntl = None  # type: ignore[assignment]

LOG_NAME = "matrix2im-log.jsonl"
LOCK_NAME = "matrix2im.lock"
_LOCKED_MSG = "Ada penulisan lain yang sedang berjalan (proses lain di PC ini)"
_CHANGED_MSG = "IM MASTER berubah atau sedang dibuka sejak rencana dibuat; buat rencana lagi"
_UNREADABLE = (zipfile.BadZipFile, KeyError, ValueError, binascii.Error, ET.ParseError)
_ABORT_NOTE = "penulisan dibatalkan, file asli tidak diubah"
# Atribut FileSharing -> nama atribut XML, supaya pesan bisa dicocokkan langsung dengan workbook.xml.
_SHARING_XML = {"user_name": "userName", "read_only_recommended": "readOnlyRecommended",
                "algorithm": "algorithmName", "spin_count": "spinCount", "hash_value": "hashValue",
                "salt_value": "saltValue", "legacy": "reservationPassword"}
_CF_TOKEN = b"<conditionalFormatting"
_WORKSHEET_RE = re.compile(r"xl/worksheets/[^/]+\.xml")


class ImMasterBusy(Matrix2ImError):
    """IM MASTER sedang dibuka orang lain; UI memetakannya ke 409 tanpa mencocokkan teks pesan."""


DRIVE_PERMISSION_HINT = ("Beri izin akses folder Google Drive ke aplikasi yang menjalankan matrix2im "
                         '(README, bagian "Menyiapkan PC pusat").')


def drive_access_error(path: Path, exc: OSError) -> Matrix2ImError:
    """PermissionError dari folder Google Drive biasanya berarti izin macOS belum diberikan, bukan file rusak."""
    return Matrix2ImError(f"IM MASTER tidak bisa dibuka: {path} ({exc.strerror or exc}). {DRIVE_PERMISSION_HINT}")


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
    laporan: Path | None = None      # laporan run (.xlsx) di folder cadangan


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
    try:
        if not path.is_file():
            raise Matrix2ImError(f"IM MASTER tidak ditemukan: {path}")
        _check_password(path, password)
        _no_owner(path)
        fp = fingerprint(path)
    except OSError as exc:
        raise drive_access_error(path, exc) from exc
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
    skipped_styles = {k.style for k in skipped}     # ada di MATRIX tapi dilewati: bukan "IM tanpa MATRIX"
    values.unmatched_im_styles = [s for s in values.unmatched_im_styles if s not in skipped_styles]
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
        "fills_changed": len(values.fills),
        "rows_inserted": sum(o.count for o in ops),
        "groups_inserted": len(ops),
        "changed_colorway_rows": len(values.changed_colorway_rows),
        "short_groups": len(values.short_groups),
        "extra_rows": len(values.extra_rows),
        "skipped": len(skipped),
        "unmatched_im_styles": len(values.unmatched_im_styles),
        "missing_styles": len(values.missing_styles),
        "warnings": len(values.warnings),
        "engine": engine,
    }
    return WritePreview(path, fp, engine, ops, projected, values, skipped, summary)


def _inserted_rows(ops: list[InsertOp]) -> list[int]:
    """Baris baru hasil sisipan, dalam koordinat akhir (rumus yang sama dengan project_layout)."""
    return [row for _, row in inserted_rows(ops)]


def _structure(layout: ImLayout) -> list[tuple]:
    return [(g.description, g.rows, g.subtotal_row) for g in layout.groups]


def _sharing_diff(a: FileSharing | None, b: FileSharing | None) -> list[str]:
    """Nama atribut XML fileSharing yang berbeda antara dua file."""
    if a is None or b is None:
        return [] if a is b else ["fileSharing"]
    return [xml for attr, xml in _SHARING_XML.items() if getattr(a, attr) != getattr(b, attr)]


def _names(zf: zipfile.ZipFile, prefix: str) -> set[str]:
    return {n for n in zf.namelist() if n.startswith(prefix)}


def _short(names: list[str], limit: int = 5) -> str:
    return ", ".join(names[:limit]) + (f" (+{len(names) - limit} lagi)" if len(names) > limit else "")


def _count_token(zf: zipfile.ZipFile, token: bytes) -> int:
    """Jumlah token di semua XML worksheet, dibaca bertahap: sheet F27 ~120 MB setelah diekstrak."""
    total = 0
    for name in zf.namelist():
        if not _WORKSHEET_RE.fullmatch(name):
            continue
        tail = b""
        with zf.open(name) as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                buf = tail + chunk
                total += buf.count(token)
                tail = buf[-(len(token) - 1):]   # token yang terpotong di batas potongan tetap terhitung
    return total


def _check_excel_kept(backup: Path, tmp: Path) -> list[str]:
    """Bandingkan hasil Excel dengan cadangan (sha256 sudah terverifikasi); kembalikan peringatan.

    Yang membatalkan: atribut fileSharing atau daftar nama media/drawing/comment/link/embedding berubah.
    Jumlah blok conditional formatting yang berubah hanya peringatan. Isi bagian-bagian itu tidak dibandingkan.
    """
    try:
        problems = []
        if diff := _sharing_diff(read_file_sharing(backup), read_file_sharing(tmp)):
            problems.append(f"fileSharing {', '.join(diff)}")
        with zipfile.ZipFile(backup) as za, zipfile.ZipFile(tmp) as zb:
            for prefix in IM_KEEP_PARTS:
                old, new = _names(za, prefix), _names(zb, prefix)
                if old != new:
                    what = [f"hilang {_short(sorted(old - new))}"] if old - new else []
                    what += [f"baru {_short(sorted(new - old))}"] if new - old else []
                    problems.append(f"{prefix.rstrip('/')}: {', '.join(what)}")
            cf_old, cf_new = _count_token(za, _CF_TOKEN), _count_token(zb, _CF_TOKEN)
    except _UNREADABLE as exc:
        raise Matrix2ImError(f"File hasil Excel tidak bisa dibaca ({exc}); {_ABORT_NOTE}") from exc
    if problems:
        raise Matrix2ImError(f"Excel mengubah bagian IM MASTER yang harus utuh ({'; '.join(problems)}); {_ABORT_NOTE}")
    if cf_old != cf_new:
        return [f"jumlah blok conditional formatting berubah setelah Excel ({cf_old} menjadi {cf_new}); "
                "periksa warna kondisional di IM MASTER"]
    return []


def _verify_written(tmp: Path, tmp2: Path, password: str, changes: list, fills: list = (),
                    before_fills: dict | None = None) -> None:
    """before_fills = warna WS TAG COLOR/WHS semua baris grup di tmp; sesudah ditulis harus sama kecuali yang direncanakan."""
    layout = read_layout(tmp2)
    for ch in changes:
        got = layout.values.get((ch.row, ch.col))
        blank = got is None or str(got).strip() == ""
        if (ch.new is None and not blank) or (ch.new is not None and (blank or str(got) != ch.new)):
            raise Matrix2ImError(f"Verifikasi gagal: sel {get_column_letter(ch.col)}{ch.row} tidak sesuai rencana")
    expected = {**(before_fills or {}), **{(f.row, f.col): f.new for f in fills}}
    for (row, col), want in sorted(expected.items()):
        if (row, col) not in layout.fills or layout.fills[(row, col)] != want:
            raise Matrix2ImError(f"Verifikasi gagal: warna sel {get_column_letter(col)}{row} tidak sesuai rencana")
    _check_password(tmp2, password)
    # workbook.xml dikecualikan dari cek CRC di bawah, jadi atribut kunci (mis. readOnlyRecommended) dicek di sini.
    if diff := _sharing_diff(read_file_sharing(tmp), read_file_sharing(tmp2)):
        raise Matrix2ImError(f"Verifikasi gagal: fileSharing berubah ({', '.join(diff)})")
    with zipfile.ZipFile(tmp) as a, zipfile.ZipFile(tmp2) as b:
        styles = styles_xml_path(a)
        skip = {sheet_xml_path(a, IM_SHEET), "xl/workbook.xml", styles}
        crc_a = {i.filename: i.CRC for i in a.infolist() if i.filename not in skip}
        crc_b = {i.filename: i.CRC for i in b.infolist() if i.filename not in skip}
        styles_a, styles_b = a.read(styles), b.read(styles)
    if crc_a != crc_b or (styles_a != styles_b and not fills):
        raise Matrix2ImError("Verifikasi gagal: bagian lain dari file berubah selain sheet dan workbook.xml")
    if styles_a != styles_b:
        check_appended(styles_a, styles_b)    # warna: hanya boleh menambah format baru di akhir styles.xml


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
    except BaseException as exc:        # termasuk Ctrl+C/SIGTERM: file setengah tertulis lebih buruk dari gagal
        detail = str(exc).replace(password, "***") if password else str(exc)
        if not touched:
            note = "File asli tidak diubah."
        else:
            try:
                _write_bytes_in_place(target, backup.read_bytes())
                ok = _sha256(target) == _sha256(backup)
            except BaseException:  # noqa: BLE001 - juga Ctrl+C kedua: pesan pemulihan manual harus selalu sampai
                ok = False
            note = ("File asli sudah dipulihkan dari cadangan." if ok else
                    f"File asli mungkin rusak; pulihkan manual dari cadangan {backup}.")
        if not isinstance(exc, Exception):
            exc.add_note(note)          # commit_write mencatatnya ke log, lalu interupsi diteruskan
            raise
        raise Matrix2ImError(f"Gagal menyimpan ke IM MASTER ({detail}). {note}") from exc


def _log(log_file: Path, entry: dict) -> bool:
    try:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        return True
    except OSError:
        return False


def _cells(items: list) -> set[tuple]:
    """Pembanding rencana: sel dan nilai/warna barunya (CellChange atau FillChange)."""
    return {(c.row, c.col, c.new) for c in items}


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
    backup = backup_dir / f"{path.stem} {state['stamp']} sebelum matrix2im{path.suffix}"
    shutil.copyfile(path, backup)
    if _sha256(backup) != preview.fingerprint.sha256:
        backup.unlink(missing_ok=True)
        raise Matrix2ImError(_CHANGED_MSG)
    state["backup"] = backup
    parent = backup_dir / "kerja"
    parent.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="kerja-", dir=parent))
    # Nama unik per run: skrip Excel menolak bila workbook bernama sama sudah terbuka.
    tmp = work / f"im-kerja-{secrets.token_hex(4)}{path.suffix}"
    tmp2 = work / f"im-hasil{path.suffix}"
    try:
        shutil.copyfile(backup, tmp)     # backup sudah lolos cek sha256: tmp pasti isi yang direncanakan
        if preview.inserts:
            progress("Excel: sisip baris")
            inserter(tmp, password, preview.inserts, IM_SHEET, get_column_letter(preview.projected.columns.last))
            state["warnings"].extend(_check_excel_kept(backup, tmp))
        real = read_layout(tmp)
        if _structure(real) != _structure(preview.projected):
            raise Matrix2ImError("Tata letak IM MASTER setelah sisip baris tidak sama dengan rencana; "
                                 "penulisan dibatalkan, file asli tidak diubah")
        progress("tulis nilai")
        colorways, _ = style_colorways(final)
        again = plan_values(real, colorways)
        if (_cells(again.changes) != _cells(preview.values.changes)
                or _cells(again.fills) != _cells(preview.values.fills)):
            raise Matrix2ImError("Perubahan nilai atau warna pada file hasil sisip tidak sama dengan rencana; "
                                 "penulisan dibatalkan, file asli tidak diubah")
        apply_cell_values(tmp, tmp2, IM_SHEET, preview.values.changes, preview.values.fills)
        progress("verifikasi")
        _verify_written(tmp, tmp2, password, preview.values.changes, preview.values.fills, real.fills)
        progress("simpan ke IM MASTER")
        if fingerprint(path) != preview.fingerprint or owner_file(path) is not None:
            raise Matrix2ImError(_CHANGED_MSG)
        _overwrite_in_place(path, tmp2, backup, password)
        state["written"] = True
    finally:
        shutil.rmtree(work, ignore_errors=True)   # best effort; folder milik run ini, tak boleh mengubah hasil


def _acquire_lock(backup_dir: Path):
    """Kunci file di folder cadangan untuk seluruh run: CLI dan server di PC yang sama tidak boleh menulis
    bersamaan (kunci server hanya berlaku di dalam satu proses). None bila proses lain memegangnya."""
    backup_dir.mkdir(parents=True, exist_ok=True)
    handle = open(backup_dir / LOCK_NAME, "a+b")
    if fcntl is not None:
        try:
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            handle.close()
            return None
    return handle


def _write_run_report(preview: WritePreview, backup_dir: Path, stamp: str, notes: list[str]) -> Path | None:
    """Laporan run (.xlsx), sepasang dengan nama cadangan; gagal menulis tidak boleh mengubah hasil penulisan."""
    path = backup_dir / f"{preview.im_master.stem} {stamp} laporan matrix2im.xlsx"
    try:
        path.write_bytes(report_xlsx_bytes(report_rows(preview, notes)))
    except Exception:  # noqa: BLE001 - laporan pelengkap; IM MASTER mungkin sudah tertulis dengan benar
        return None
    return path


def _interrupted(exc: BaseException, state: dict) -> str:
    """Pesan log untuk Ctrl+C/SIGTERM yang jujur tentang keadaan file asli."""
    notes = " ".join(getattr(exc, "__notes__", []))
    if notes:
        tail = notes
    elif state["written"]:
        tail = "IM MASTER sudah diperbarui sebelum dihentikan."
    else:
        tail = "File asli tidak diubah."
    return f"Penulisan dihentikan paksa ({type(exc).__name__}). {tail}"


def commit_write(preview: WritePreview, final: dict, settings: Settings, password: str, user: str,
                 matrix_name: str, progress: Callable[[str], None] = lambda s: None,
                 inserter: Callable[..., float] = run_excel_inserts) -> WriteReport:
    """Jalankan rencana. Kegagalan sebelum langkah simpan tidak menyentuh file asli; semua hasil dicatat di log."""
    start = time.monotonic()
    if settings.backup_dir is None:
        raise Matrix2ImError("Setting backup_dir belum diisi")
    log_file = settings.backup_dir / LOG_NAME
    stamp = f"{datetime.now():%Y-%m-%d %H%M%S}"   # sama untuk cadangan dan laporan run
    state: dict = {"backup": None, "warnings": [], "written": False, "stamp": stamp}
    lock = None

    def record(ok: bool, message: str) -> WriteReport:
        if password:
            message = message.replace(password, "***")
        laporan = None
        if lock is not None:     # hanya pemegang kunci yang menulis ke folder cadangan (nama bisa bentrok)
            notes = state["warnings"] + ([] if ok else [f"Penulisan gagal: {message}"])
            laporan = _write_run_report(preview, settings.backup_dir, stamp, notes)
            if laporan is None:
                message += " (laporan gagal ditulis)"
        backup = state["backup"]
        entry = {"waktu": datetime.now().isoformat(timespec="seconds"), "user": user, "matrix": matrix_name,
                 "im_master": str(preview.im_master), "engine": preview.engine, "summary": preview.summary,
                 "backup": str(backup) if backup else None, "hasil": "ok" if ok else "gagal", "pesan": message,
                 "laporan": str(laporan) if laporan else None}
        if not _log(log_file, entry):
            message += " (catatan log gagal ditulis)"
        return WriteReport(ok, message, preview.summary, backup, log_file, time.monotonic() - start,
                           _inserted_rows(preview.inserts) if ok else [], laporan)

    try:
        try:
            lock = _acquire_lock(settings.backup_dir)
            if lock is None:
                raise Matrix2ImError(_LOCKED_MSG)     # sebelum menyentuh apa pun
            _execute(preview, final, settings, password, progress, inserter, state)
            ok = True
            message = (f"IM MASTER diperbarui: {preview.summary['cells_changed']} sel diubah, "
                       f"{preview.summary['fills_changed']} warna latar diubah (WS TAG COLOR dan WHS), "
                       f"{preview.summary['rows_inserted']} baris disisipkan.")
            if missing := len(preview.values.missing_styles):
                message += f" {missing} style MATRIX tidak ada di IM (tidak ditulis)."
            message += "".join(f" PERINGATAN: {w}." for w in state["warnings"])
        except Matrix2ImError as exc:
            ok, message = False, str(exc)
        except Exception as exc:  # noqa: BLE001 - apa pun harus berakhir sebagai laporan, bukan crash
            ok, message = False, f"Gagal tak terduga: {type(exc).__name__}: {exc}"
            trace = "".join(traceback.format_exception(exc))
            logging.getLogger(__name__).error("commit_write gagal: %s",
                                              trace.replace(password, "***") if password else trace)
        except BaseException as exc:   # Ctrl+C/SIGTERM: catat dulu, lalu teruskan
            record(False, _interrupted(exc, state))
            raise
        report = record(ok, message)
    finally:
        if lock is not None:
            lock.close()                # menutup file melepas flock
    if report.ok:
        progress("selesai")
    return report

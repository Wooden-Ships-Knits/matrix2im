"""Antarmuka command line `matrix2im`: extract, transform, run, write, ui."""
from __future__ import annotations

import argparse
import dataclasses
import getpass
import logging
import sys
from pathlib import Path
from typing import NoReturn

from .errors import Matrix2ImError
from .extract import extract
from .im_lock import read_file_sharing
from .im_writer import WritePreview, commit_write, prepare_write
from .io_utils import read_json, write_json
from .settings import Settings, load_settings
from .transform import transform
from .ui_server import serve
from .validate import read_im_header, validate
from .write_report import report_rows, report_xlsx_bytes

log = logging.getLogger("matrix2im")
_HANDLER_FLAG = "_matrix2im_handler"


def _setup_logging() -> None:
    """Log ke stderr; handler lama dari main() sebelumnya dibuang agar tidak dobel di test.

    Handler lain (mis. milik pytest caplog) sengaja tidak disentuh.
    """
    root = logging.getLogger()
    for handler in [h for h in root.handlers if getattr(h, _HANDLER_FLAG, False)]:
        root.removeHandler(handler)
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    setattr(handler, _HANDLER_FLAG, True)
    root.addHandler(handler)
    root.setLevel(logging.INFO)


def format_summary(raw: dict, final: dict) -> str:
    """Ringkasan hasil (produk perintah, bukan log) dalam Bahasa Indonesia."""
    s = final["summary"]
    whs = ", ".join(f"{'null' if k is None else k}: {v}" for k, v in s["WHS"].items()) or "-"
    lines = [
        "Ringkasan:",
        f"  Jumlah style (kartu MATRIX): {raw.get('total_styles', s['total_styles'])}",
        f"  Jumlah record (colorway): {s['total_records']}",
        f"  WHS per nilai: {whs}",
        f"  WS TAG COLOR: {s['WS_TAG_COLOR_filled']} terisi, {s.get('WS_TAG_COLOR_need_review', 0)} Need Review, "
        f"{s['WS_TAG_COLOR_null']} null",
        f"  needs_review: {s['needs_review_items']} item",
    ]
    for code, count in s["needs_review_by_issue"].items():
        lines.append(f"    {code}: {count}")
    return "\n".join(lines)


def _extract_to(matrix: Path, raw_path: Path) -> dict:
    raw = extract(matrix)
    write_json(raw_path, raw)
    log.info("Tahap 1 selesai: %d style, %d warning -> %s",
             len(raw["styles"]), len(raw.get("warnings", [])), raw_path)
    return raw


def _remove_stale(path: Path) -> None:
    """Hapus final lama di path tujuan agar tidak dikira hasil run yang baru saja gagal."""
    if not path.is_file():
        return
    try:
        path.unlink()
    except OSError as exc:
        log.warning("File lama %s tidak bisa dihapus: %s", path, exc)
    else:
        log.warning("File lama %s dihapus karena Tahap 2 gagal", path)


def _transform_to(raw: dict, final_path: Path, im_master: Path | None) -> dict:
    """Tahap 2 + validasi; final hanya ditulis bila lolos.

    Bila gagal (validasi, header IM, transformasi), final lama dari run sebelumnya dihapus:
    kalau tidak, final itu tertinggal di samping raw.json baru seolah keduanya satu pasang.
    """
    try:
        final = transform(raw)
        header = read_im_header(im_master) if im_master else None
        errors = validate(raw, final, header)
        if errors:
            for message in errors:
                log.error(message)
            raise Matrix2ImError(f"Validasi gagal ({len(errors)} error); {final_path} tidak ditulis")
    except Exception:
        _remove_stale(final_path)
        raise
    write_json(final_path, final)
    log.info("Tahap 2 selesai: %d record -> %s", len(final["records"]), final_path)
    return final


def _finish(raw: dict, final: dict, strict: bool) -> int:
    print(format_summary(raw, final))
    if strict and final["summary"]["needs_review_items"]:
        log.error("--strict: needs_review tidak kosong (%d item)", final["summary"]["needs_review_items"])
        return 2
    return 0


def _cmd_extract(args: argparse.Namespace) -> int:
    _extract_to(args.matrix, args.output)
    return 0


def _cmd_transform(args: argparse.Namespace) -> int:
    raw = read_json(args.raw)
    final = _transform_to(raw, args.output, args.im_master)
    print(format_summary(raw, final))
    return 0


def _cmd_run(args: argparse.Namespace) -> int:
    raw = _extract_to(args.matrix, args.output / "raw.json")
    final = _transform_to(raw, args.output / "final.json", args.im_master)
    return _finish(raw, final, args.strict)


def _cmd_ui(args: argparse.Namespace) -> int:
    settings = load_settings(args.settings)
    serve(host=args.host or settings.host, port=args.port or settings.port,
          open_browser=not args.no_browser, settings=settings)
    return 0


_SUMMARY_LABELS = (("styles", "Style cocok"), ("groups", "Grup IM cocok"), ("cells_changed", "Sel diubah"),
                   ("rows_inserted", "Baris disisipkan"), ("groups_inserted", "Grup yang disisipi"),
                   ("changed_colorway_rows", "Baris colorway berubah"), ("short_groups", "Grup kekurangan baris"),
                   ("extra_rows", "Baris lebih (WS TAG COLOR & WHS dikosongkan)"),
                   ("skipped", "Style dilewati"), ("unmatched_im_styles", "Style IM tanpa data MATRIX"),
                   ("missing_styles", "Style MATRIX tidak ada di IM (tidak ditulis)"), ("warnings", "Peringatan"))


def format_write_preview(preview: WritePreview) -> str:
    """Ringkasan rencana tulis (sebelum konfirmasi dan pada --dry-run) dalam Bahasa Indonesia."""
    lines = ["Rencana tulis ke IM MASTER:", f"  File IM MASTER: {preview.im_master}",
             f"  Mesin sisip: {preview.engine}"]
    lines += [f"  {label}: {preview.summary[key]}" for key, label in _SUMMARY_LABELS if key in preview.summary]
    for warning in preview.values.warnings:
        lines.append(f"  PERINGATAN: {warning}")
    for skip in preview.skipped:
        lines.append(f"  DILEWATI: {skip.style}: {skip.reason}")
    for style in preview.values.missing_styles:
        lines.append(f"  TIDAK ADA DI IM (tidak ditulis): {style}")
    return "\n".join(lines)


def _read_password() -> str:
    """Password hanya dari getpass (tty) atau baris pertama stdin; tidak pernah dari argv."""
    if sys.stdin.isatty():
        return getpass.getpass("Password IM MASTER: ")
    return sys.stdin.readline().rstrip("\n").rstrip("\r")


def _report_target(path: Path, settings: Settings) -> Path:
    """Path akhir laporan --report; menolak IM MASTER atau workbook berkunci (mis. cadangan IM).

    Salah ketik path laporan tidak boleh menimpa IM MASTER atau cadangannya: penulisan laporan terjadi
    sebelum ada cadangan, bahkan saat --dry-run.
    """
    if path.suffix.lower() != ".xlsx":
        path = path.with_name(path.name + ".xlsx")
    target = path.expanduser().resolve()
    refused = settings.im_master is not None and target == Path(settings.im_master).expanduser().resolve()
    if not refused and target.is_file():
        try:
            refused = read_file_sharing(target) is not None
        except Exception:   # bukan xlsx yang terbaca: bukan workbook IM, boleh ditimpa sebagai laporan
            refused = False
    if refused:
        raise Matrix2ImError(f"Path --report harus file laporan baru, bukan IM MASTER atau cadangannya: {path}")
    return path


def _write_preview_report(path: Path, preview: WritePreview) -> Path:
    """Laporan .xlsx semua perubahan rencana (tanpa batas baris), sama dengan unduhan di UI; path akhirnya."""
    try:
        path.write_bytes(report_xlsx_bytes(report_rows(preview)))
    except OSError as exc:
        raise Matrix2ImError(f"Laporan tidak bisa ditulis: {path} ({exc})") from exc
    return path


def _cmd_write(args: argparse.Namespace) -> int:
    settings = load_settings(args.settings)
    overrides = {k: v for k, v in (("im_master", args.im_master), ("backup_dir", args.backup_dir),
                                   ("insert_engine", args.engine)) if v is not None}
    settings = dataclasses.replace(settings, **overrides)
    report_path = _report_target(args.report, settings) if args.report else None
    final = read_json(args.final_json)
    interactive = sys.stdin.isatty()
    if not interactive and not args.dry_run and not args.yes:
        raise Matrix2ImError("Input bukan terminal; tambahkan --yes untuk menulis tanpa konfirmasi")
    password = _read_password()
    preview = prepare_write(final, settings, password)
    print(format_write_preview(preview))
    if report_path is not None:
        print(f"Laporan rencana (Excel): {_write_preview_report(report_path, preview)}")
    if args.dry_run:
        print("Dry-run: IM MASTER tidak diubah.")
        return 0
    if not args.yes and input("Ketik 'ya' untuk menulis ke IM MASTER: ").strip() != "ya":
        raise Matrix2ImError("Dibatalkan, IM MASTER tidak diubah")
    report = commit_write(preview, final, settings, password, args.user or getpass.getuser(),
                          args.final_json.name, progress=lambda phase: log.info("%s", phase))
    if not report.ok:
        log.error("%s", report.message)
        if report.laporan:
            log.error("Laporan: %s", report.laporan)
        log.error("Log: %s", report.log_file)
        return 1
    print(report.message)
    if report.backup:
        print(f"Cadangan: {report.backup}")
    if report.laporan:
        print(f"Laporan: {report.laporan}")
    print(f"Log: {report.log_file}")
    return 0


# Pesan argparse yang paling sering muncul, diterjemahkan agar pesan salah pakai tetap Indonesia.
_ARGPARSE_ID = (("the following arguments are required", "argumen wajib belum diisi"),
                ("unrecognized arguments", "argumen tidak dikenal"),
                ("expected one argument", "butuh satu nilai"),
                ("invalid choice", "pilihan tidak valid"),
                ("choose from", "pilih dari"),
                ("argument ", "argumen "))


class _Parser(argparse.ArgumentParser):
    """argparse bawaan keluar dengan kode 2 saat argumen salah, padahal kode 2 khusus untuk
    --strict (PROMPT §5); salah pakai diperlakukan sebagai error fatal biasa (exit 1)."""

    def error(self, message: str) -> NoReturn:
        for english, indonesian in _ARGPARSE_ID:
            message = message.replace(english, indonesian)
        self.exit(1, f"ERROR: Argumen tidak valid: {message}. "
                     f"Jalankan '{self.prog} -h' untuk bantuan.\n")


def _parser() -> argparse.ArgumentParser:
    parser = _Parser(prog="matrix2im",
                     description="Ubah F27 MATRIX MASTER (Excel) menjadi JSON untuk IM MASTER.")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("extract", help="Tahap 1: MATRIX .xlsx -> raw JSON")
    p.add_argument("matrix", type=Path)
    p.add_argument("-o", "--output", type=Path, required=True, help="file raw JSON")
    p.set_defaults(func=_cmd_extract)

    p = sub.add_parser("transform", help="Tahap 2: raw JSON -> JSON final")
    p.add_argument("raw", type=Path)
    p.add_argument("-o", "--output", type=Path, required=True, help="file JSON final")
    p.add_argument("--im-master", type=Path, help="cocokkan key record dengan header IM MASTER")
    p.set_defaults(func=_cmd_transform)

    p = sub.add_parser("run", help="Tahap 1 + 2, tulis raw.json dan final.json ke folder")
    p.add_argument("matrix", type=Path)
    p.add_argument("--im-master", type=Path, required=True)
    p.add_argument("-o", "--output", type=Path, required=True, help="folder output")
    p.add_argument("--strict", action="store_true", help="exit 2 bila needs_review tidak kosong")
    p.set_defaults(func=_cmd_run)

    p = sub.add_parser("write", help="tulis final.json ke IM MASTER (password dibaca dari prompt/stdin)")
    p.add_argument("final_json", type=Path)
    p.add_argument("--settings", type=Path, help="file setting (default matrix2im.local.toml)")
    p.add_argument("--im-master", type=Path, help="timpa lokasi IM MASTER dari setting")
    p.add_argument("--backup-dir", type=Path, help="timpa folder cadangan dari setting")
    p.add_argument("--engine", choices=("excel", "none"), help="timpa mesin sisip baris")
    p.add_argument("--user", help="nama pengguna untuk log (default: user OS)")
    p.add_argument("--dry-run", action="store_true", help="tampilkan ringkasan saja, file tidak diubah")
    p.add_argument("--yes", action="store_true", help="tulis tanpa konfirmasi")
    p.add_argument("--report", type=Path,
                   help="tulis laporan .xlsx semua perubahan rencana ke file ini (juga saat --dry-run)")
    p.set_defaults(func=_cmd_write)

    p = sub.add_parser("ui", help="buka UI web untuk upload Excel dan melihat hasil")
    p.add_argument("--port", type=int, default=None, help="port (default dari setting, lalu 8765)")
    p.add_argument("--host", help="alamat bind (default dari setting, lalu 127.0.0.1)")
    p.add_argument("--settings", type=Path, help="file setting (default matrix2im.local.toml)")
    p.add_argument("--no-browser", action="store_true", help="jangan buka browser otomatis")
    p.set_defaults(func=_cmd_ui)
    return parser


def main(argv: list[str] | None = None) -> int:
    """Titik masuk CLI; tidak pernah menampilkan traceback (exit 0/1/2).

    Salah pakai argumen keluar lewat SystemExit(1) dari _Parser; -h/--help keluar dengan 0.
    """
    args = _parser().parse_args(argv)
    _setup_logging()
    try:
        return args.func(args)
    except Matrix2ImError as exc:
        log.error(str(exc))
        return 1
    except Exception as exc:  # noqa: BLE001 - openpyxl bisa melempar ValueError/TypeError pada file rusak
        log.error(f"Error tidak terduga: {type(exc).__name__}: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

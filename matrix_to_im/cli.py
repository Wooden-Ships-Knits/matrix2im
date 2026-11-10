"""Antarmuka command line `matrix2im`: extract, transform, run."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import NoReturn

from .errors import Matrix2ImError
from .extract import extract
from .io_utils import read_json, write_json
from .transform import transform
from .validate import read_im_header, validate

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
        f"  WS TAG COLOR: {s['WS_TAG_COLOR_filled']} terisi, {s['WS_TAG_COLOR_null']} null",
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


def _transform_to(raw: dict, final_path: Path, im_master: Path | None) -> dict:
    final = transform(raw)
    header = read_im_header(im_master) if im_master else None
    errors = validate(raw, final, header)
    if errors:
        for message in errors:
            log.error(message)
        raise Matrix2ImError(f"Validasi gagal ({len(errors)} error); {final_path} tidak ditulis")
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

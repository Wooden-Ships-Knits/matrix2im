"""Tahap ekstraksi MATRIX MASTER. Bagian ini: helper murni (tanpa I/O Excel)."""
from __future__ import annotations

import re

from .excel_reader import norm
from .legend_config import (BLACK_FONT_RGB, DEFAULT_FONT_MEANING, FONT_MEANING,
                            NEW_REPEAT_LABELS)


def split_yarns(text: object) -> list[str]:
    """Pecah sel benang pada '/' dan baris baru; selalu dipecah (D2), urutan dipertahankan."""
    parts = (norm(p) for p in re.split(r"[/\n]", str(text)))
    return [p for p in parts if p]


def parse_no(value: object) -> int | str | None:
    """Nomor style: angka jadi int agar 20.0 tidak tampil sebagai '20.0'; kode seperti '1A' tetap teks."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(value)
    text = norm(value)
    return int(text) if text.isdigit() else text


def font_meaning(rgb: str | None) -> tuple[str, bool]:
    """Arti warna font menurut legend; tanpa warna/hitam = REPEAT. Kode asing ditandai tidak dikenal."""
    if rgb is None or rgb in BLACK_FONT_RGB:
        return DEFAULT_FONT_MEANING, True
    if rgb in FONT_MEANING:
        return FONT_MEANING[rgb], True
    return f"? ({rgb})", False


def is_new_repeat_label(value: object) -> bool:
    """Full match saja, supaya nama style seperti 'NEWPORT CREW' tidak dikira label NEW."""
    return isinstance(value, str) and norm(value).upper() in NEW_REPEAT_LABELS


def is_punctuation_only(text: str) -> bool:
    """Sel yang hanya berisi tanda baca (mis. '.', '`') bukan data (D5)."""
    return re.fullmatch(r"[\W_]+", text) is not None


def join_tag_cells(values: list[str]) -> str | None:
    """Gabung sel TAG per kartu; baris kosong ganda jadi ' | ', baris baru tunggal jadi spasi (D6)."""
    cells = [norm(v).replace("\n\n", " | ").replace("\n", " ") for v in values]
    return " | ".join(cells) if cells else None

"""Kode review yang terbaca mesin, supaya needs_review bisa difilter tanpa mem-parse teks."""
from dataclasses import dataclass
from enum import StrEnum


class Code(StrEnum):
    """Nilai sama dengan nama member agar JSON stabil dan mudah dicari."""
    TAG_CONDITIONAL = "TAG_CONDITIONAL"
    TAG_CN_OUT_OF_RANGE = "TAG_CN_OUT_OF_RANGE"
    TAG_EMPTY = "TAG_EMPTY"
    TAG_UNRESOLVED = "TAG_UNRESOLVED"
    WHS_STYLE_UNKNOWN = "WHS_STYLE_UNKNOWN"
    WHS_CONFLICT = "WHS_CONFLICT"
    FONT_COLOR_NO_RULE = "FONT_COLOR_NO_RULE"
    COLOR_IS_DESCRIPTOR = "COLOR_IS_DESCRIPTOR"
    EXTRACTION_WARNING = "EXTRACTION_WARNING"


@dataclass(frozen=True)
class Problem:
    """Satu temuan review: kode untuk mesin, issue untuk manusia (Bahasa Indonesia)."""
    code: Code
    issue: str

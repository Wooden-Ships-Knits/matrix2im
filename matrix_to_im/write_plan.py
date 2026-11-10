"""Rencana penulisan ke IM MASTER: dari final.json menjadi daftar colorway per style."""
from __future__ import annotations

import re
from dataclasses import dataclass

from .errors import Matrix2ImError
from .im_layout import normalize
from .legend_config import IM_COLOR_KEYS, WHS_SETS

_REF_RE = re.compile(r"^(?P<sheet>.+?)(?: \((?P<section>[^)]*)\))? No\.(?P<no>.+?) \| ")
_WHS_BY_SET = {parts: name for name, parts in WHS_SETS.items()}


@dataclass(frozen=True)
class Colorway:
    colors: tuple[str, ...]
    ws_tag: str | None
    whs: str | None


@dataclass(frozen=True)
class Skip:
    style: str
    reason: str


def parse_ref(ref: str) -> tuple[str, str]:
    """Kembalikan (sheet, kunci kartu); kunci kartu = teks sebelum ' | ' pertama."""
    m = _REF_RE.match(ref)
    if not m:
        raise Matrix2ImError(f"Format _ref tidak dikenali: {ref!r}")
    return m.group("sheet"), ref.split(" | ", 1)[0]


def _colorway(rec: dict) -> Colorway:
    colors = tuple(str(rec[k]) for k in IM_COLOR_KEYS if rec.get(k))
    return Colorway(colors, rec.get("WS TAG COLOR"), rec.get("WHS"))


def _key(cw: Colorway) -> tuple[str, ...]:
    return tuple(normalize(c) for c in cw.colors)


def _merge_whs(values: list[str | None]) -> str | None:
    """Gabungan bagian WHS; None di salah satu sheet berarti tidak ditulis."""
    if any(v is None for v in values):
        return None
    parts: set[str] = set()
    for v in values:
        if v not in WHS_SETS:
            return None
        parts |= WHS_SETS[v]
    return _WHS_BY_SET.get(frozenset(parts))


def _merge_sheets(style: str, by_sheet: dict[str, list[Colorway]]) -> list[Colorway] | Skip:
    sheets = list(by_sheet)
    first = sheets[0]
    base_keys = [_key(c) for c in by_sheet[first]]
    for other in sheets[1:]:
        keys = [_key(c) for c in by_sheet[other]]
        if len(keys) != len(base_keys) or set(keys) != set(base_keys):
            return Skip(style, f"colorway berbeda antar-sheet: {first} {len(base_keys)} colorway, "
                               f"{other} {len(keys)} colorway")
    lookup = [{_key(c): c for c in by_sheet[s]} for s in sheets]
    merged: list[Colorway] = []
    for cw in by_sheet[first]:
        k = _key(cw)
        variants = [m[k] for m in lookup]
        tags = {normalize(v.ws_tag) if v.ws_tag is not None else None for v in variants}
        if len(tags) > 1:
            return Skip(style, f"WS TAG COLOR berbeda antar-sheet untuk colorway {' / '.join(cw.colors)}")
        merged.append(Colorway(cw.colors, cw.ws_tag, _merge_whs([v.whs for v in variants])))
    return merged


def style_colorways(final: dict) -> tuple[dict[str, list[Colorway]], list[Skip]]:
    """Satu daftar colorway terurut per style (key = DESCRIPTION ternormalisasi)."""
    grouped: dict[str, dict[str, list[Colorway]]] = {}
    for rec in final["records"]:
        sheet, _card = parse_ref(rec["_ref"])
        style = normalize(rec["DESCRIPTION"])
        grouped.setdefault(style, {}).setdefault(sheet, []).append(_colorway(rec))
    result: dict[str, list[Colorway]] = {}
    skips: list[Skip] = []
    for style, by_sheet in grouped.items():
        if len(by_sheet) == 1:
            result[style] = next(iter(by_sheet.values()))
            continue
        merged = _merge_sheets(style, by_sheet)
        if isinstance(merged, Skip):
            skips.append(merged)
        else:
            result[style] = merged
    return result, skips

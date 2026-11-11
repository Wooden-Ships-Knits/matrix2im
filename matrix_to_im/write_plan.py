"""Rencana penulisan ke IM MASTER: dari final.json menjadi daftar colorway per style."""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Iterable

from openpyxl.utils import get_column_letter

from .errors import Matrix2ImError
from .im_layout import ImGroup, ImLayout, normalize
from .legend_config import IM_ACCESSORY_RE, IM_COLOR_KEYS, NEED_REVIEW, WHS_SETS

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
        # Colorway tanpa benang akan mengosongkan semua benang lama di baris IM: lewati, jangan tebak.
        if any(not cw.colors for cws in by_sheet.values() for cw in cws):
            skips.append(Skip(style, "record tanpa COLOR"))
            continue
        if len(by_sheet) == 1:
            result[style] = next(iter(by_sheet.values()))
            continue
        merged = _merge_sheets(style, by_sheet)
        if isinstance(merged, Skip):
            skips.append(merged)
        else:
            result[style] = merged
    return result, skips


@dataclass(frozen=True)
class InsertOp:
    """Satu sisipan baris di dalam grup; at_row/fill_* dalam koordinat sebelum sisipan op ini."""
    description: str
    at_row: int
    count: int
    fill_from: int
    fill_to: int
    subtotal_row_after: int | None
    formula_fixes: dict[int, str]


def match_style(description: str, styles: Iterable[str]) -> str | None:
    """Style terpanjang yang menjadi awalan 'STYLE - ' dari DESCRIPTION grup."""
    best: str | None = None
    for s in styles:
        if description.startswith(s + " - ") and (best is None or len(s) > len(best)):
            best = s
    return best


_REF = r"(\$?[A-Z]{1,3}\$?)"


def _widen(formula: str, r: int, k: int) -> str:
    """Perlebar rujukan sel tunggal / range satu-baris di baris r menjadi r..r+k."""
    rng = re.compile(rf"(?<![A-Za-z$:]){_REF}{r}:{_REF}{r}(?!\d)")
    one = re.compile(rf"(?<![A-Za-z$:]){_REF}{r}(?![\d:])")
    formula = rng.sub(lambda m: f"{m.group(1)}{r}:{m.group(2)}{r + k}", formula)
    return one.sub(lambda m: f"{m.group(1)}{r}:{m.group(1)}{r + k}", formula)


def plan_inserts(layout: ImLayout, colorways: dict[str, list[Colorway]]) -> list[InsertOp]:
    """Sisipan agar tiap grup punya baris sebanyak colorway; urut at_row menurun (eksekusi bawah ke atas)."""
    ops: list[InsertOp] = []
    for g in layout.groups:
        style = match_style(g.description, colorways)
        if style is None:
            continue
        r, n = len(g.rows), len(colorways[style])
        if n <= r:
            continue
        k = n - r
        if r >= 2:
            last = g.rows[-1]
            ops.append(InsertOp(g.description, last, k, last - 1, last + k - 1, None, {}))
            continue
        row = g.rows[0]
        fixes: dict[int, str] = {}
        after = None
        if g.subtotal_row is not None:
            after = g.subtotal_row + k
            for col, f in layout.subtotal_formulas.get(g.subtotal_row, {}).items():
                if isinstance(f, str):
                    new = _widen(f, row, k)
                    if new != f:
                        fixes[col] = new
        ops.append(InsertOp(g.description, row + 1, k, row, row + k, after, fixes))
    return sorted(ops, key=lambda o: o.at_row, reverse=True)


def project_layout(layout: ImLayout, ops: list[InsertOp]) -> ImLayout:
    """Tata letak setelah semua op dijalankan, dihitung murni di Python (pembanding hasil Excel)."""
    def shift(x: int) -> int:
        return x + sum(o.count for o in ops if o.at_row <= x)

    inserted: dict[int, int] = {}  # baris sisipan baru -> baris sumber (terproyeksi)
    group_new: dict[int, list[int]] = {}
    for idx, g in enumerate(layout.groups):
        for o in ops:
            if o.description == g.description and g.rows[0] <= o.at_row <= g.rows[-1] + 1:
                start = o.at_row + sum(p.count for p in ops if p.at_row < o.at_row)
                new = list(range(start, start + o.count))
                group_new.setdefault(idx, []).extend(new)
                for row in new:
                    inserted[row] = shift(o.fill_from)
    groups = [
        ImGroup(g.description, tuple(sorted([shift(x) for x in g.rows] + group_new.get(i, []))),
                None if g.subtotal_row is None else shift(g.subtotal_row))
        for i, g in enumerate(layout.groups)
    ]
    values = {(shift(r), c): v for (r, c), v in layout.values.items()}
    for new_row, src in inserted.items():
        for (r, c), v in list(values.items()):
            if r == src:
                values[(new_row, c)] = v
    formulas = {shift(r): dict(f) for r, f in layout.subtotal_formulas.items()}
    return ImLayout(layout.columns, groups, values, formulas)


_ACCESSORY = re.compile(IM_ACCESSORY_RE)


@dataclass(frozen=True)
class CellChange:
    """Satu sel yang berubah; new=None berarti sel dikosongkan."""
    row: int
    col: int
    old: object
    new: str | None


@dataclass
class ValuePlan:
    changes: list[CellChange] = field(default_factory=list)
    changed_colorway_rows: list[int] = field(default_factory=list)
    short_groups: list[str] = field(default_factory=list)
    unmatched_im_styles: list[str] = field(default_factory=list)
    missing_styles: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    colorway_text: dict[int, tuple[str, str]] = field(default_factory=dict)   # baris -> (benang lama, baru)


def _blank(v: object) -> bool:
    return v is None or str(v).strip() == ""


def _is_formula(v: object) -> bool:
    return isinstance(v, str) and v.startswith("=")


def plan_values(layout: ImLayout, colorways: dict[str, list[Colorway]]) -> ValuePlan:
    """Sel COLOR/WS TAG/WHS yang berubah per baris grup; rumus tidak pernah disentuh."""
    plan = ValuePlan()
    cols = layout.columns
    matched: set[str] = set()
    unmatched: set[str] = set()
    # Bagian = deret grup berurutan dengan kunci sama. Blok M/L berurutan (subtotal FAF lalu FAH) adalah satu
    # bagian dan bentuk normal; kunci yang muncul lagi di bagian lain sheet diperingatkan (tetap diisi).
    parts: Counter[str] = Counter()
    prev: str | None = None
    for g in layout.groups:
        key = g.description if match_style(g.description, colorways) else None
        if key is not None and key != prev:
            parts[key] += 1
        prev = key
    for key, count in parts.items():
        if count > 1:
            plan.warnings.append(f"{key}: muncul di {count} bagian sheet; semua bagian diisi colorway yang sama")
    for g in layout.groups:
        style = match_style(g.description, colorways)
        if style is None:
            if " - " in g.description:   # tanpa " - " = label musim/judul, bukan style
                unmatched.add(g.description.split(" - ", 1)[0])
            continue
        matched.add(style)
        if g.subtotal_row is None:
            plan.warnings.append(f"{g.description} (baris {g.rows[0]}): grup tanpa baris subtotal")
        cws = colorways[style]
        n, r = len(cws), len(g.rows)
        if r < n:
            plan.short_groups.append(f"{g.description}: kurang {n - r} baris")
        for i, row in enumerate(g.rows):
            cw = cws[min(i, n - 1)]
            yarns = cw.colors[:len(cols.colors)]

            def put(col: int, new: str | None) -> None:
                old = layout.values.get((row, col))
                if _is_formula(old):
                    plan.warnings.append(f"sel {get_column_letter(col)}{row} berisi rumus, tidak ditulis")
                elif (None if _blank(old) else str(old)) != new and not (_blank(old) and new is None):
                    plan.changes.append(CellChange(row, col, old, new))

            old_yarns: list[str] = []
            old_text: list[str] = []
            for j, col in enumerate(cols.colors):
                old = layout.values.get((row, col))
                if not _blank(old) and not _is_formula(old) and not _ACCESSORY.match(str(old)):
                    old_yarns.append(normalize(old))
                    old_text.append(str(old))
                if j < len(yarns):
                    put(col, yarns[j])
                elif not _blank(old) and not _ACCESSORY.match(str(old)):
                    put(col, None)
            if tuple(old_yarns) != tuple(normalize(y) for y in yarns):
                plan.changed_colorway_rows.append(row)
                plan.colorway_text[row] = (" / ".join(old_text), " / ".join(yarns))
            if cw.ws_tag is not None and not (
                    cw.ws_tag == NEED_REVIEW and not _blank(layout.values.get((row, cols.ws_tag)))):
                put(cols.ws_tag, cw.ws_tag)
            if cw.whs is not None:
                put(cols.whs, cw.whs)
    plan.unmatched_im_styles = sorted(unmatched)
    plan.missing_styles = [s for s in colorways if s not in matched]
    return plan

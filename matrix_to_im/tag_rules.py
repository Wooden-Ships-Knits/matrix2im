"""Aturan WS TAG COLOR: instruksi TAG + daftar benang -> nilai kolom IM."""
import re

from .legend_config import (PUNCTUATION_ONLY_RE, TAG_ALIAS, TAG_CN_RE, TAG_CONDITIONAL_MARKERS,
                            TAG_DOUBLE_SUFFIX_RE, TAG_PREFIX_RE, TAG_REPLACEMENTS, TAG_UNRESOLVED_RE,
                            YARN_CODE_SUFFIX_RE, YARN_PREFIX_RE)
from .review import Code, Problem

_PREFIX = re.compile(TAG_PREFIX_RE)
_CN = re.compile(TAG_CN_RE)
_REPLACEMENTS = tuple((re.compile(pattern), repl) for pattern, repl in TAG_REPLACEMENTS)
_PUNCTUATION_ONLY = re.compile(PUNCTUATION_ONLY_RE)
_UNRESOLVED = re.compile(TAG_UNRESOLVED_RE, re.IGNORECASE)


def color_name(yarn: str) -> str:
    """Nama benang -> nama warna: buang prefix pemasok dan kode angka, lalu terapkan alias.

    Tag di IM memakai nama warna saja, bukan kode benang pabrik.
    """
    n = re.sub(YARN_PREFIX_RE, "", yarn.strip().upper())
    n = re.sub(YARN_CODE_SUFFIX_RE, "", n).strip()
    return TAG_ALIAS.get(n, n)


def unique_yarns(yarns: list[str]) -> list[str]:
    """Hapus benang kembar tanpa peduli huruf besar/kecil; urutan dan ejaan kemunculan pertama dipertahankan."""
    seen: set[str] = set()
    out: list[str] = []
    for y in yarns:
        if y.upper() not in seen:
            seen.add(y.upper())
            out.append(y)
    return out


def build_ws_tag_color(tag: str | None, yarns: list[str]) -> tuple[str | None, Problem | None]:
    """Ganti C1, C2, ... di instruksi TAG dengan nama warna benang (yarns = hasil unique_yarns).

    Seluruh hasil di-uppercase (keputusan D7, sama dengan prototipe).
    """
    if not tag:
        return None, Problem(Code.TAG_EMPTY, "TAG kosong")
    t = tag.upper()
    if any(m in t for m in TAG_CONDITIONAL_MARKERS):
        return None, Problem(Code.TAG_CONDITIONAL, "TAG kondisional / singkatan: perlu aturan khusus")
    body = _PREFIX.sub("", t).strip()
    # "TAG:" saja atau "TAG: -" tidak memuat nama warna; string kosong/tanda baca bukan nilai IM.
    if not body or _PUNCTUATION_ONLY.fullmatch(body):
        return None, Problem(Code.TAG_EMPTY, "TAG kosong")
    # C0 juga di luar jangkauan: tanpa cek ini yarns[-1] diam-diam memilih benang terakhir.
    missing = [m for m in _CN.findall(body) if not 1 <= int(m) <= len(yarns)]
    if missing:
        return None, Problem(Code.TAG_CN_OUT_OF_RANGE,
                             f"TAG minta C{missing[0]} tapi colorway hanya punya {len(yarns)} benang")
    # Dicek pada teks TAG sebelum Cn diganti, supaya nama benang (mis. "MULTI COLOR") tidak ikut terdeteksi.
    if _UNRESOLVED.search(body):
        return None, Problem(Code.TAG_UNRESOLVED, f"TAG masih berisi instruksi yang belum diterjemahkan: {tag}")
    val = _CN.sub(lambda m: color_name(yarns[int(m.group(1)) - 1]), body)
    for pattern, repl in _REPLACEMENTS:   # mis. STRIPED -> STRIPE (catatan header IM)
        val = pattern.sub(repl, val)
    # Nama benang sudah berakhir dengan sufiks itu ("OYSTER MARL" + " MARL"), jadi lipat jadi satu.
    val = re.sub(TAG_DOUBLE_SUFFIX_RE, r"\1", val)
    return re.sub(r"\s*/\s*", "/", val).strip(), None

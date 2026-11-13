"""Aturan WS TAG COLOR: instruksi TAG + daftar benang -> nilai kolom IM."""
import re

from .legend_config import (PUNCTUATION_ONLY_RE, TAG_ALIAS, TAG_CN_RE, TAG_COND_NOT_SOLID, TAG_COND_SOLID,
                            TAG_COND_UNSUPPORTED, TAG_CONDITIONAL_MARKERS, TAG_DOUBLE_SUFFIX_RE, TAG_PREFIX_RE,
                            TAG_REPLACEMENTS, TAG_UNRESOLVED_RE, YARN_CODE_SUFFIX_RE, YARN_PREFIX_RE)
from .review import Code, Problem

_PREFIX = re.compile(TAG_PREFIX_RE)
_CN = re.compile(TAG_CN_RE)
# Deret Cn bergaris miring (C1/C2, C1/C2/C3): nama warna yang sama cukup ditulis sekali.
_CN_RUN = re.compile(TAG_CN_RE + r"(?:\s*/\s*" + TAG_CN_RE.replace("(", "(?:") + r")+")
_REPLACEMENTS = tuple((re.compile(pattern), repl) for pattern, repl in TAG_REPLACEMENTS)
_PUNCTUATION_ONLY = re.compile(PUNCTUATION_ONLY_RE)
_UNRESOLVED = re.compile(TAG_UNRESOLVED_RE, re.IGNORECASE)
# Dua bentuk bagian TAG bersyarat: "TAG: <instruksi> (FOR <syarat>)" dan "TAG FOR <syarat>: <instruksi>".
_CLAUSE_PAREN = re.compile(r"^TAG\s*:?\s*(?P<instr>.+?)\s*\(\s*FOR\s+(?P<cond>[^)]+?)\s*\)$")
_CLAUSE_PREFIX = re.compile(r"^TAG\s+FOR\s+(?P<cond>[^:]+?)\s*:\s*(?P<instr>.+)$")


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


def _clauses(tag: str) -> list[tuple[str, str]] | None:
    """(syarat, instruksi) per bagian TAG bersyarat; None bila ada bagian yang bentuknya tidak dikenali."""
    parts = [c.strip() for piece in tag.split("|") for c in re.split(r"(?=\bTAG\b)", piece) if c.strip()]
    out = []
    for part in parts:
        m = _CLAUSE_PAREN.match(part) or _CLAUSE_PREFIX.match(part)
        if m is None:
            return None
        out.append((re.sub(r"\s+", " ", m.group("cond")).strip(), m.group("instr").strip()))
    return out


def _conditional_instruction(tag: str, yarns: list[str]) -> tuple[str | None, str]:
    """Instruksi TAG yang berlaku untuk colorway ini, atau (None, alasan) bila tidak bisa diputuskan.

    Urutan: syarat nama colorway (persis, urutan benang sama), lalu solid (1 benang) / bukan solid.
    Keduanya dinilai dari benang unik: benang berulang (stripe) tidak membuat colorway bukan solid.
    Lebih dari satu bagian yang cocok dianggap ambigu.
    """
    if any(m in tag for m in TAG_COND_UNSUPPORTED):
        return None, "TAG kondisional / singkatan: perlu aturan khusus"
    clauses = _clauses(tag)
    if not clauses:
        return None, "TAG kondisional / singkatan: perlu aturan khusus"
    distinct = unique_yarns(yarns)
    names = [color_name(y) for y in distinct]
    solid = len(distinct) == 1
    by_colorway = [i for c, i in clauses if c not in TAG_COND_SOLID | TAG_COND_NOT_SOLID
                   and [n.strip() for n in c.split("/")] == names]
    by_kind = [i for c, i in clauses if c in (TAG_COND_SOLID if solid else TAG_COND_NOT_SOLID)]
    for found in (by_colorway, by_kind):
        if len(found) == 1:
            return found[0], ""
        if len(found) > 1:
            return None, "TAG bersyarat: lebih dari satu syarat cocok untuk colorway ini"
    kind = "solid, 1 benang" if solid else f"bukan solid, {len(distinct)} benang"
    return None, f"TAG bersyarat: tidak ada syarat yang cocok untuk colorway ini ({kind})"


def build_ws_tag_color(tag: str | None, yarns: list[str]) -> tuple[str | None, Problem | None]:
    """Ganti C1, C2, ... di instruksi TAG dengan nama warna benang.

    yarns = benang sesuai urutan MATRIX termasuk yang berulang, jadi Cn = benang ke-n (keputusan user 13 Nov 2026).

    Seluruh hasil di-uppercase (keputusan D7, sama dengan prototipe).
    """
    if not tag:
        return None, Problem(Code.TAG_EMPTY, "TAG kosong")
    t = tag.upper()
    if any(m in t for m in TAG_CONDITIONAL_MARKERS):
        # TAG bersyarat dipilih per colorway (solid / bukan solid / nama colorway), lalu diproses seperti TAG biasa.
        chosen, why = _conditional_instruction(t, yarns)
        if chosen is None:
            return None, Problem(Code.TAG_CONDITIONAL, why)
        return build_ws_tag_color("TAG: " + chosen, yarns)
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
    # Benang berbeda bernama warna sama (mis. TWILIGHT SKY dua kali, C1/C2 MARL): TAG cukup "TWILIGHT SKY MARL";
    # kolom COLOR tetap terpisah (permintaan user 13 Nov 2026).
    val = _CN_RUN.sub(lambda m: "/".join(dict.fromkeys(
        color_name(yarns[int(n) - 1]) for n in _CN.findall(m.group(0)))), body)
    val = _CN.sub(lambda m: color_name(yarns[int(m.group(1)) - 1]), val)
    for pattern, repl in _REPLACEMENTS:   # mis. STRIPED -> STRIPE (catatan header IM)
        val = pattern.sub(repl, val)
    # Nama benang sudah berakhir dengan sufiks itu ("OYSTER MARL" + " MARL"), jadi lipat jadi satu.
    val = re.sub(TAG_DOUBLE_SUFFIX_RE, r"\1", val)
    return re.sub(r"\s*/\s*", "/", val).strip(), None

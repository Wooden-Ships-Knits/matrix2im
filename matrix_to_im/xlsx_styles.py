"""Warna latar sel lewat styles.xml: baca warna sel dan tambahkan format baru tanpa mengubah format lama.

Sel menunjuk format lewat atribut s (indeks cellXfs); format menunjuk warna lewat fillId (indeks fills).
Format dan warna lama tidak pernah diubah atau dihapus: yang dibutuhkan dicari dulu, bila tidak ada
ditambahkan di akhir daftar, jadi sel lain yang memakai format lama tidak ikut berubah.
"""
from __future__ import annotations

import re

from .errors import Matrix2ImError

_FILLS = re.compile(rb"(<fills\b[^>]*>)(.*?)</fills>", re.S)
_CELL_XFS = re.compile(rb"(<cellXfs\b[^>]*>)(.*?)</cellXfs>", re.S)
_FILL_EL = re.compile(rb"<fill\b[^>]*/>|<fill\b[^>]*>.*?</fill>", re.S)
_XF_EL = re.compile(rb"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", re.S)
ATTR_RE = re.compile(rb'([\w:]+)="([^"]*)"')
_PATTERN = re.compile(rb"<patternFill\b([^>]*?)/?>")
_FG = re.compile(rb"<fgColor\b([^>]*?)/?>")
_OPEN = re.compile(rb"<xf\b[^>]*?(/?)>")
_COUNT = re.compile(rb'\scount="\d+"')


def fill_key(fill: bytes) -> str | None:
    """Kunci pembanding warna latar: None = tanpa warna, "FF99FF" = solid rgb (6 hex, tanpa alpha).

    Bentuk lain (theme, pola, gradien) diberi kunci teks yang stabil; hanya dipakai untuk membedakan.
    """
    m = _PATTERN.search(fill)
    if m is None:
        return "gradien" if b"<gradientFill" in fill else None
    ptype = dict(ATTR_RE.findall(m.group(1))).get(b"patternType", b"none").decode()
    if ptype == "none":
        return None
    fg = _FG.search(fill)
    attrs = dict(ATTR_RE.findall(fg.group(1))) if fg else {}
    if ptype == "solid" and set(attrs) == {b"rgb"}:
        return attrs[b"rgb"].decode().upper()[-6:]
    return " ".join([ptype, *(f"{k.decode()}={v.decode()}" for k, v in sorted(attrs.items()))])


def _new_fill(key: str | None) -> bytes:
    if key is None:
        return b'<fill><patternFill patternType="none"/></fill>'
    if not re.fullmatch(r"[0-9A-F]{6}", key):
        raise Matrix2ImError(f"Warna latar '{key}' bukan kode 6 hex")
    return (b'<fill><patternFill patternType="solid"><fgColor rgb="FF' + key.encode()
            + b'"/><bgColor indexed="64"/></patternFill></fill>')


def set_attr(tag: bytes, name: bytes, value: bytes) -> bytes:
    """Ganti (atau tambahkan) satu atribut di tag pembuka XML."""
    pat = re.compile(rb'(?<![\w:])' + name + rb'="[^"]*"')
    new = name + b'="' + value + b'"'
    if pat.search(tag):
        return pat.sub(lambda _m: new, tag, count=1)
    end = re.match(rb"<[\w:]+", tag).end()
    return tag[:end] + b" " + new + tag[end:]


def _canon(xf: bytes) -> tuple:
    """Bentuk pembanding format: atribut tanpa urutan + isi elemen (alignment/protection) apa adanya.

    applyFill="1" diabaikan: format tanpa atribut itu (mis. buatan openpyxl) tetap menampilkan warnanya.
    """
    head = _OPEN.match(xf)
    attrs = [a for a in ATTR_RE.findall(head.group(0)) if a != (b"applyFill", b"1")]
    return tuple(sorted(attrs)), xf[head.end():]


def _sections(xml: bytes) -> tuple[re.Match, re.Match]:
    fills, xfs = _FILLS.search(xml), _CELL_XFS.search(xml)
    if fills is None or xfs is None:
        raise Matrix2ImError("styles.xml tanpa daftar fills/cellXfs; warna sel tidak bisa ditulis")
    if xfs.start() < fills.end():       # skema: fills selalu sebelum cellXfs
        raise Matrix2ImError("styles.xml: urutan fills/cellXfs tidak dikenali; warna sel tidak bisa ditulis")
    return fills, xfs


class StyleSheet:
    """styles.xml yang hanya boleh ditambah: warna dan format baru masuk di akhir daftar."""

    def __init__(self, xml: bytes):
        self.xml = xml
        fills, xfs = _sections(xml)
        self.fills: list[bytes] = _FILL_EL.findall(fills.group(2))
        self.xfs: list[bytes] = _XF_EL.findall(xfs.group(2))
        self._old = (len(self.fills), len(self.xfs))
        self._fill_keys = [fill_key(f) for f in self.fills]
        self._canon: dict[tuple, int] = {}
        for i, xf in enumerate(self.xfs):
            self._canon.setdefault(_canon(xf), i)
        self._cache: dict[tuple[int, str | None], int] = {}

    def _xf(self, s: int) -> bytes:
        if not 0 <= s < len(self.xfs):
            raise Matrix2ImError(f"Sel menunjuk format sel {s} yang tidak ada di styles.xml")
        return self.xfs[s]

    def _fill_id(self, xf: bytes) -> int:
        fid = int(dict(ATTR_RE.findall(_OPEN.match(xf).group(0))).get(b"fillId", b"0"))
        if not 0 <= fid < len(self.fills):
            raise Matrix2ImError(f"Format sel menunjuk warna {fid} yang tidak ada di styles.xml")
        return fid

    def xf_fill_key(self, s: int) -> str | None:
        """Kunci warna latar format sel ke-s."""
        return self._fill_keys[self._fill_id(self._xf(s))]

    def restyle(self, s: int, key: str | None) -> int:
        """Indeks format yang sama dengan format s kecuali warna latarnya = key (dipakai ulang bila sudah ada)."""
        if (s, key) in self._cache:
            return self._cache[(s, key)]
        xf = self._xf(s)
        if self._fill_keys[self._fill_id(xf)] == key:
            idx = s
        else:
            if key in self._fill_keys:
                fid = self._fill_keys.index(key)
            else:
                self.fills.append(_new_fill(key))
                self._fill_keys.append(key)
                fid = len(self.fills) - 1
            head = _OPEN.match(xf).group(0)
            new_head = set_attr(set_attr(head, b"fillId", str(fid).encode()), b"applyFill", b"1")
            new = new_head + xf[len(head):]
            idx = self._canon.get(_canon(new), -1)
            if idx < 0:
                self.xfs.append(new)
                idx = len(self.xfs) - 1
                self._canon[_canon(new)] = idx
        self._cache[(s, key)] = idx
        return idx

    def to_bytes(self) -> bytes:
        """styles.xml dengan warna/format baru di akhir daftar; tanpa tambahan = byte asli."""
        n_fills, n_xfs = self._old
        if (len(self.fills), len(self.xfs)) == self._old:
            return self.xml
        out = self.xml
        for pattern, items, old in ((_CELL_XFS, self.xfs, n_xfs), (_FILLS, self.fills, n_fills)):
            m = pattern.search(out)      # dicari ulang di hasil sebelumnya, jadi offset selalu benar
            head = m.group(1)
            if _COUNT.search(head):
                head = _COUNT.sub(b' count="%d"' % len(items), head, count=1)
            close = m.group(0)[m.group(0).rindex(b"</"):]
            out = out[:m.start()] + head + m.group(2) + b"".join(items[old:]) + close + out[m.end():]
        return out


def check_appended(old: bytes, new: bytes) -> None:
    """Pastikan new = old + warna/format baru di akhir daftar saja; selain itu batalkan."""
    try:
        a, b = StyleSheet(old), StyleSheet(new)
        fa, xa = _sections(old)
        fb, xb = _sections(new)
    except Matrix2ImError as exc:
        raise Matrix2ImError(f"Verifikasi gagal: styles.xml tidak terbaca ({exc})") from exc
    problems = []
    if b.fills[:len(a.fills)] != a.fills or b.xfs[:len(a.xfs)] != a.xfs:
        problems.append("warna/format lama berubah")
    for m, items in ((fb, b.fills), (xb, b.xfs)):
        count = re.search(rb'\scount="(\d+)"', m.group(1))
        if count and int(count.group(1)) != len(items):
            problems.append("atribut count tidak cocok")

    def rest(xml: bytes, f: re.Match, x: re.Match) -> bytes:
        return xml[:f.start()] + _COUNT.sub(b"", f.group(1)) + xml[f.end():x.start()] + \
            _COUNT.sub(b"", x.group(1)) + xml[x.end():]

    if rest(old, fa, xa) != rest(new, fb, xb):
        problems.append("bagian lain berubah")
    if problems:
        raise Matrix2ImError(f"Verifikasi gagal: styles.xml ({', '.join(problems)})")

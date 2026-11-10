"""Verifikasi Password to modify IM MASTER (fileSharing di xl/workbook.xml) dan deteksi pemilik Excel."""
from __future__ import annotations

import base64
import hashlib
import hmac
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from matrix_to_im import legend_config as cfg
from matrix_to_im.errors import Matrix2ImError

_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


class WrongPassword(Matrix2ImError):
    """Password to modify yang dimasukkan tidak cocok."""


class NotLocked(Matrix2ImError):
    """File tidak memakai Password to modify."""


@dataclass(frozen=True)
class FileSharing:
    user_name: str | None
    algorithm: str | None
    hash_value: str | None
    salt_value: str | None
    spin_count: int | None
    read_only_recommended: bool
    legacy: bool


def read_file_sharing(path: Path) -> FileSharing | None:
    """Baca elemen fileSharing langsung dari zip; None bila tidak ada (tanpa openpyxl agar file tak tersentuh)."""
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read("xl/workbook.xml"))
    el = root.find(f"{_MAIN_NS}fileSharing")
    if el is None:
        return None
    spin = el.get("spinCount")
    return FileSharing(
        user_name=el.get("userName"),
        algorithm=el.get("algorithmName"),
        hash_value=el.get("hashValue"),
        salt_value=el.get("saltValue"),
        spin_count=int(spin) if spin else None,
        read_only_recommended=el.get("readOnlyRecommended") in ("1", "true"),
        legacy=el.get("reservationPassword") is not None,
    )


def hash_modify_password(password: str, salt_b64: str, spin_count: int, algorithm: str = "SHA-512") -> str:
    """Hash ECMA-376: H(salt+pw utf-16le), lalu spin_count kali H(h + i little-endian 4 byte)."""
    name = cfg.IM_LOCK_ALGORITHMS.get(algorithm)
    if name is None:
        raise Matrix2ImError(f"Algoritma kunci IM MASTER tidak didukung: {algorithm}")
    h = hashlib.new(name, base64.b64decode(salt_b64) + password.encode("utf-16-le")).digest()
    for i in range(spin_count):
        h = hashlib.new(name, h + i.to_bytes(4, "little")).digest()
    return base64.b64encode(h).decode("ascii")


def check_modify_password(path: Path, password: str) -> None:
    """Lolos diam-diam bila password benar; selain itu lempar error yang menjelaskan sebabnya."""
    fs = read_file_sharing(path)
    if fs is not None and fs.legacy:
        raise Matrix2ImError("Kunci IM MASTER memakai format lama (reservationPassword) yang belum didukung")
    if fs is None or not fs.hash_value:
        raise NotLocked("IM MASTER tidak dikunci Password to modify; penulisan dibatalkan. "
                        "Periksa apakah file yang dipilih benar.")
    if not fs.salt_value or fs.spin_count is None:
        raise Matrix2ImError("Data kunci IM MASTER tidak lengkap (saltValue/spinCount)")
    actual = hash_modify_password(password, fs.salt_value, fs.spin_count, fs.algorithm or "SHA-512")
    if not hmac.compare_digest(actual, fs.hash_value):
        raise WrongPassword("Password IM MASTER salah")


def owner_file(path: Path) -> Path | None:
    """File pemilik Excel (~$nama) di folder yang sama; nama bisa terpotong 2 huruf pertama."""
    for name in (cfg.OWNER_FILE_PREFIX + path.name,
                 cfg.OWNER_FILE_PREFIX + path.name[cfg.OWNER_FILE_TRUNCATE:]):
        candidate = path.with_name(name)
        if candidate.is_file():
            return candidate
    return None


def owner_name(owner: Path) -> str:
    """Nama pemilik: byte pertama = panjang n, lalu n byte UTF-8."""
    data = owner.read_bytes()
    if data:
        n = data[0]
        if 0 < n <= cfg.OWNER_NAME_MAX_LEN:
            name = data[1:1 + n].decode("utf-8", errors="replace").strip()
            if name:
                return name
    return cfg.OWNER_NAME_FALLBACK

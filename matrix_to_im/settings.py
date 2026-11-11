"""Setting lokal (matrix2im.local.toml): lokasi IM MASTER, folder cadangan, alamat UI, mesin sisip."""
from __future__ import annotations

import tomllib
from dataclasses import dataclass
from pathlib import Path

from .errors import Matrix2ImError

DEFAULT_SETTINGS_FILE = "matrix2im.local.toml"
_ENGINES = ("excel", "none")


@dataclass(frozen=True)
class Settings:
    im_master: Path | None = None
    backup_dir: Path | None = None
    host: str = "127.0.0.1"
    port: int = 8765
    insert_engine: str = "excel"


def _path(value: object, key: str) -> Path | None:
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        raise Matrix2ImError(f"Setting {key} harus berupa teks path")
    return Path(value).expanduser()


def load_settings(path: Path | None = None) -> Settings:
    """File default yang tidak ada berarti setting kosong; file yang disebut eksplisit wajib ada."""
    explicit = path is not None
    path = path if explicit else Path(DEFAULT_SETTINGS_FILE)
    if not path.is_file():
        if explicit:
            raise Matrix2ImError(f"File setting tidak ditemukan: {path}")
        return Settings()
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
    except (tomllib.TOMLDecodeError, UnicodeDecodeError) as exc:
        raise Matrix2ImError(f"File setting tidak valid: {path} ({exc})") from exc
    ui, im = data.get("ui", {}), data.get("im", {})
    base = Settings()
    engine = im.get("insert_engine", base.insert_engine)
    if engine not in _ENGINES:
        raise Matrix2ImError(f"Setting insert_engine harus 'excel' atau 'none', bukan {engine!r}")
    port = ui.get("port", base.port)
    if isinstance(port, bool) or not isinstance(port, int):
        raise Matrix2ImError("Setting port harus berupa angka")
    return Settings(im_master=_path(im.get("master"), "im.master"),
                    backup_dir=_path(im.get("backup_dir"), "im.backup_dir"),
                    host=str(ui.get("host", base.host)), port=port, insert_engine=engine)

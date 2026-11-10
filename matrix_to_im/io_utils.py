"""Baca/tulis JSON dengan format tetap (UTF-8, tanpa escape ASCII, indent 2)."""
import json
from pathlib import Path

from .errors import Matrix2ImError


def read_json(path: Path) -> dict:
    """Baca JSON; file hilang atau rusak jadi Matrix2ImError agar CLI bisa keluar bersih."""
    path = Path(path)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise Matrix2ImError(f"File tidak ditemukan: {path}") from None
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise Matrix2ImError(f"File JSON rusak: {path} ({exc})") from exc
    except OSError as exc:
        raise Matrix2ImError(f"Tidak bisa membaca file: {path} ({exc})") from exc


def write_json(path: Path, data: dict) -> None:
    """Tulis JSON dan buat folder parent bila belum ada."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

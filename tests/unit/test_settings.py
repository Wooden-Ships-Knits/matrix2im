from pathlib import Path

import pytest

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.settings import DEFAULT_SETTINGS_FILE, Settings, load_settings


def test_load_settings_default_missing_file_gives_defaults(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    assert DEFAULT_SETTINGS_FILE == "matrix2im.local.toml"
    s = load_settings()
    assert s == Settings()
    assert (s.host, s.port, s.insert_engine) == ("127.0.0.1", 8765, "excel")
    assert s.im_master is None and s.backup_dir is None


def test_load_settings_explicit_missing_file_raises(tmp_path):
    with pytest.raises(Matrix2ImError, match="tidak ditemukan"):
        load_settings(tmp_path / "tidak-ada.toml")


def test_load_settings_reads_toml_and_expands_user(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    cfg = tmp_path / "s.toml"
    cfg.write_text('[ui]\nhost = "0.0.0.0"\nport = 9000\n'
                   '[im]\nmaster = "~/im.xlsx"\nbackup_dir = "~/cadangan"\ninsert_engine = "none"\n',
                   encoding="utf-8")
    s = load_settings(cfg)
    assert s.im_master == Path(str(tmp_path / "im.xlsx"))
    assert s.backup_dir == tmp_path / "cadangan"
    assert (s.host, s.port, s.insert_engine) == ("0.0.0.0", 9000, "none")


def test_invalid_engine_raises(tmp_path):
    cfg = tmp_path / "s.toml"
    cfg.write_text('[im]\ninsert_engine = "libreoffice"\n', encoding="utf-8")
    with pytest.raises(Matrix2ImError, match="insert_engine"):
        load_settings(cfg)


def test_invalid_toml_raises(tmp_path):
    cfg = tmp_path / "s.toml"
    cfg.write_text("[im\n", encoding="utf-8")
    with pytest.raises(Matrix2ImError, match="tidak valid"):
        load_settings(cfg)


@pytest.mark.parametrize("body", ["ui = 1\n", 'im = "x"\n'])
def test_non_table_section_raises_indonesian_error(tmp_path, body):
    cfg = tmp_path / "s.toml"
    cfg.write_text(body, encoding="utf-8")
    with pytest.raises(Matrix2ImError, match="harus berupa tabel"):
        load_settings(cfg)

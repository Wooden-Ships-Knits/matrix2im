import hashlib
import io
import json

import pytest
from helpers.im_factory import TEST_PASSWORD, GroupSpec, add_file_sharing, build_im_sheet
from helpers.matrix_factory import build_im, build_matrix, build_no_pages

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.cli import format_summary, main


@pytest.fixture
def matrix(tmp_path):
    return build_matrix(tmp_path / "m.xlsx")


@pytest.fixture
def im(tmp_path):
    return build_im(tmp_path / "im.xlsx")


def load(path):
    return json.loads(path.read_text(encoding="utf-8"))


def test_extract_ok(tmp_path, matrix):
    out = tmp_path / "raw.json"
    assert main(["extract", str(matrix), "-o", str(out)]) == 0
    assert len(load(out)["styles"]) == 4


def test_extract_logs_warning_count(tmp_path, matrix, caplog):
    caplog.set_level("INFO")
    assert main(["extract", str(matrix), "-o", str(tmp_path / "raw.json")]) == 0
    assert "warning" in caplog.text.lower()


def test_transform_ok(tmp_path, matrix, im):
    raw, final = tmp_path / "raw.json", tmp_path / "final.json"
    assert main(["extract", str(matrix), "-o", str(raw)]) == 0
    assert main(["transform", str(raw), "-o", str(final), "--im-master", str(im)]) == 0
    assert load(final)["records"]
    assert main(["transform", str(raw), "-o", str(tmp_path / "f2.json")]) == 0


def test_run_writes_both_and_prints_summary(tmp_path, matrix, im, capsys):
    out = tmp_path / "out"
    assert main(["run", str(matrix), "--im-master", str(im), "-o", str(out)]) == 0
    assert (out / "raw.json").exists() and (out / "final.json").exists()
    captured = capsys.readouterr()
    assert "needs_review" in captured.out
    assert "WHS" in captured.out
    assert "Traceback" not in captured.err


def test_run_requires_im_master(tmp_path, matrix):
    with pytest.raises(SystemExit) as exc:
        main(["run", str(matrix), "-o", str(tmp_path / "out")])
    assert exc.value.code == 1 and not (tmp_path / "out").exists()


@pytest.mark.parametrize("argv", [[], ["run", "m.xlsx", "-o", "out"], ["extract", "m.xlsx"],
                                  ["extract", "m.xlsx", "-o", "r.json", "--tidak-ada"], ["hapus"]])
def test_usage_error_exits_1_with_indonesian_message(argv, capsys):    # F3: exit 2 hanya untuk --strict
    with pytest.raises(SystemExit) as exc:
        main(argv)
    err = capsys.readouterr().err
    assert exc.value.code == 1
    assert "Argumen tidak valid:" in err and "-h' untuk bantuan." in err
    assert "error:" not in err and "Traceback" not in err


def test_required_argument_detail_in_indonesian(capsys):
    with pytest.raises(SystemExit):
        main(["run", "m.xlsx", "-o", "out"])
    assert "argumen wajib belum diisi: --im-master" in capsys.readouterr().err


@pytest.mark.parametrize("argv", [["-h"], ["run", "--help"]])
def test_help_exits_0(argv, capsys):
    with pytest.raises(SystemExit) as exc:
        main(argv)
    assert exc.value.code == 0 and "matrix2im" in capsys.readouterr().out


def test_run_strict_with_review_returns_2(tmp_path, matrix, im):
    out = tmp_path / "out"
    assert main(["run", str(matrix), "--im-master", str(im), "-o", str(out), "--strict"]) == 2
    assert (out / "final.json").exists()
    assert load(out / "final.json")["needs_review"]


def test_missing_input_returns_1_without_traceback(tmp_path, capsys):
    assert main(["extract", str(tmp_path / "x.xlsx"), "-o", str(tmp_path / "r.json")]) == 1
    err = capsys.readouterr().err
    assert "Traceback" not in err and "ERROR" in err


def test_workbook_without_pages_returns_1(tmp_path, capsys):
    path = build_no_pages(tmp_path / "np.xlsx")
    assert main(["extract", str(path), "-o", str(tmp_path / "r.json")]) == 1
    assert "Traceback" not in capsys.readouterr().err
    assert not (tmp_path / "r.json").exists()


@pytest.mark.parametrize("content", ["{bukan json", '{"a": 1}'])
def test_transform_bad_json_returns_1(tmp_path, capsys, content):
    bad = tmp_path / "bad.json"
    bad.write_text(content, encoding="utf-8")
    assert main(["transform", str(bad), "-o", str(tmp_path / "f.json")]) == 1
    assert "Traceback" not in capsys.readouterr().err
    assert not (tmp_path / "f.json").exists()


def test_transform_malformed_style_returns_1(tmp_path, capsys):
    bad = tmp_path / "bad.json"
    bad.write_text('{"styles": [{"style": "X"}], "warnings": []}', encoding="utf-8")
    assert main(["transform", str(bad), "-o", str(tmp_path / "f.json")]) == 1
    err = capsys.readouterr().err
    assert "Traceback" not in err and "style #1" in err


def test_unexpected_exception_returns_1(tmp_path, matrix, capsys, monkeypatch):
    def boom(path):
        raise ValueError("rusak")
    monkeypatch.setattr("matrix_to_im.cli.extract", boom)
    assert main(["extract", str(matrix), "-o", str(tmp_path / "r.json")]) == 1
    err = capsys.readouterr().err
    assert "Error tidak terduga: ValueError: rusak" in err and "Traceback" not in err


def test_im_header_missing_key_returns_1(tmp_path, matrix, capsys):
    im = build_im(tmp_path / "im.xlsx", drop="COLOR2")
    out = tmp_path / "out"
    assert main(["run", str(matrix), "--im-master", str(im), "-o", str(out)]) == 1
    assert (out / "raw.json").exists()
    assert not (out / "final.json").exists()
    assert "COLOR2" in capsys.readouterr().err


def test_transform_validation_failure_writes_nothing(tmp_path, matrix):
    im = build_im(tmp_path / "im.xlsx", drop="COLOR2")
    raw, final = tmp_path / "raw.json", tmp_path / "final.json"
    main(["extract", str(matrix), "-o", str(raw)])
    assert main(["transform", str(raw), "-o", str(final), "--im-master", str(im)]) == 1
    assert not final.exists()


def test_failed_run_removes_stale_final(tmp_path, matrix, caplog):     # F7
    out = tmp_path / "out"
    out.mkdir()
    (out / "final.json").write_text('{"lama": true}', encoding="utf-8")
    im = build_im(tmp_path / "im.xlsx", drop="COLOR2")
    assert main(["run", str(matrix), "--im-master", str(im), "-o", str(out)]) == 1
    assert (out / "raw.json").exists() and not (out / "final.json").exists()
    assert any(r.levelname == "WARNING" and "dihapus" in r.getMessage() and "final.json" in r.getMessage()
               for r in caplog.records)


def test_failed_transform_removes_stale_final(tmp_path, matrix):
    raw, final = tmp_path / "raw.json", tmp_path / "final.json"
    assert main(["extract", str(matrix), "-o", str(raw)]) == 0
    final.write_text('{"lama": true}', encoding="utf-8")
    im = build_im(tmp_path / "im.xlsx", drop="COLOR2")
    assert main(["transform", str(raw), "-o", str(final), "--im-master", str(im)]) == 1
    assert not final.exists()


def test_run_with_unreadable_im_removes_stale_final(tmp_path, matrix):
    out = tmp_path / "out"
    out.mkdir()
    (out / "final.json").write_text('{"lama": true}', encoding="utf-8")
    assert main(["run", str(matrix), "--im-master", str(tmp_path / "nope.xlsx"), "-o", str(out)]) == 1
    assert (out / "raw.json").exists() and not (out / "final.json").exists()


def test_im_master_missing_returns_1(tmp_path, matrix):
    assert main(["run", str(matrix), "--im-master", str(tmp_path / "nope.xlsx"),
                 "-o", str(tmp_path / "out")]) == 1


def test_output_parent_created(tmp_path, matrix, im):
    assert main(["extract", str(matrix), "-o", str(tmp_path / "a" / "b" / "raw.json")]) == 0
    assert (tmp_path / "a" / "b" / "raw.json").exists()
    assert main(["run", str(matrix), "--im-master", str(im), "-o", str(tmp_path / "baru")]) == 0
    assert (tmp_path / "baru" / "final.json").exists()


def test_format_summary_contents():
    raw = {"total_styles": 4}
    final = {"summary": {"total_records": 7, "total_styles": 4, "WHS": {"SY": 3, None: 1},
                         "WS_TAG_COLOR_filled": 4, "WS_TAG_COLOR_need_review": 2, "WS_TAG_COLOR_null": 1,
                         "needs_review_items": 3,
                         "needs_review_by_issue": {"TAG_CONDITIONAL": 2, "TAG_EMPTY": 1}}}
    text = format_summary(raw, final)
    assert "WS TAG COLOR: 4 terisi, 2 Need Review, 1 null" in text
    for needle in ("4", "7", "SY", "null", "terisi", "TAG_CONDITIONAL", "TAG_EMPTY", "needs_review"):
        assert needle in text


# ---------- write ----------
class _Tty(io.StringIO):
    def isatty(self):
        return True


def _rec(style, colors, ws=None, whs=None, n=1):
    rec = {"DESCRIPTION": style}
    for i, c in enumerate(colors, 1):
        rec[f"COLOR{i}"] = c
    rec["WS TAG COLOR"], rec["WHS"] = ws, whs
    rec["_ref"] = f"SHEET No.{n} | B1 | x"
    return rec


@pytest.fixture
def write_env(tmp_path, monkeypatch):
    im = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("AAA - K", [["a1"], ["a2"]])])
    add_file_sharing(im)
    final = tmp_path / "final.json"
    final.write_text(json.dumps({"records": [_rec("AAA", ["x1"], "T1", "000", 1),
                                             _rec("AAA", ["x2"], "T2", "SY", 2)]}), encoding="utf-8")
    argv = ["write", str(final), "--im-master", str(im), "--backup-dir", str(tmp_path / "cadangan"),
            "--engine", "none", "--user", "andi"]
    monkeypatch.setattr("sys.stdin", io.StringIO(TEST_PASSWORD + "\n"))
    return im, argv


def _sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_write_dry_run_prints_summary_and_leaves_file(write_env, capsys):
    im, argv = write_env
    before = _sha(im)
    assert main(argv + ["--dry-run"]) == 0
    out = capsys.readouterr()
    assert _sha(im) == before
    assert str(im) in out.out and "none" in out.out
    assert "Traceback" not in out.err


def test_write_wrong_password_exit_1(write_env, monkeypatch, capsys):
    im, argv = write_env
    before = _sha(im)
    monkeypatch.setattr("sys.stdin", io.StringIO("salah\n"))
    assert main(argv + ["--yes"]) == 1
    assert _sha(im) == before
    assert "Traceback" not in capsys.readouterr().err


def test_write_yes_writes_engine_none(write_env, tmp_path, capsys):
    im, argv = write_env
    before = _sha(im)
    assert main(argv + ["--yes"]) == 0
    assert _sha(im) != before
    assert "IM MASTER diperbarui" in capsys.readouterr().out
    assert (tmp_path / "cadangan").is_dir()


def test_write_non_tty_without_yes_exit_1(write_env, capsys):
    im, argv = write_env
    before = _sha(im)
    assert main(argv) == 1
    assert _sha(im) == before
    assert "--yes" in capsys.readouterr().err


def test_write_tty_confirmation(write_env, monkeypatch, capsys):
    im, argv = write_env
    before = _sha(im)
    monkeypatch.setattr("sys.stdin", _Tty())
    monkeypatch.setattr("getpass.getpass", lambda prompt="": TEST_PASSWORD)
    monkeypatch.setattr("builtins.input", lambda prompt="": "tidak")
    assert main(argv) == 1
    assert _sha(im) == before
    assert "Dibatalkan, IM MASTER tidak diubah" in capsys.readouterr().err
    monkeypatch.setattr("builtins.input", lambda prompt="": "ya")
    assert main(argv) == 0
    assert _sha(im) != before


def test_write_report_not_ok_exit_1(write_env, monkeypatch, capsys):
    from pathlib import Path

    from matrix_to_im.im_writer import WriteReport
    im, argv = write_env
    monkeypatch.setattr("matrix_to_im.cli.commit_write", lambda *a, **k: WriteReport(
        False, "Gagal contoh", {}, None, Path("/x/log.jsonl"), 0.1, []))
    assert main(argv + ["--yes"]) == 1
    err = capsys.readouterr()
    assert "Gagal contoh" in err.err + err.out and "log.jsonl" in err.err + err.out


def test_write_options_override_settings_file(write_env, tmp_path, monkeypatch):
    im, argv = write_env
    toml = tmp_path / "s.toml"
    toml.write_text('[im]\nmaster = "/tidak/ada.xlsx"\ninsert_engine = "excel"\n', encoding="utf-8")
    seen = {}

    def fake_prepare(final, settings, password):
        seen["settings"] = settings
        raise Matrix2ImError("stop")

    monkeypatch.setattr("matrix_to_im.cli.prepare_write", fake_prepare)
    assert main(argv + ["--settings", str(toml), "--dry-run"]) == 1
    assert seen["settings"].im_master == im and seen["settings"].insert_engine == "none"


def test_write_missing_settings_exit_1(tmp_path, monkeypatch, capsys):
    final = tmp_path / "final.json"
    final.write_text('{"records": []}', encoding="utf-8")
    monkeypatch.setattr("sys.stdin", io.StringIO("x\n"))
    monkeypatch.chdir(tmp_path)
    assert main(["write", str(final), "--dry-run"]) == 1
    err = capsys.readouterr().err
    assert "belum diisi" in err and "Traceback" not in err


def test_password_never_in_output(write_env, monkeypatch, capsys):
    im, argv = write_env
    assert main(argv + ["--yes"]) == 0
    out = capsys.readouterr()
    assert TEST_PASSWORD not in out.out + out.err
    monkeypatch.setattr("sys.stdin", io.StringIO(TEST_PASSWORD + "-salah\n"))
    assert main(argv + ["--yes"]) == 1
    out = capsys.readouterr()
    assert TEST_PASSWORD not in out.out + out.err


def test_no_password_option():
    with pytest.raises(SystemExit):
        main(["write", "f.json", "--password", "x"])


# ---------- ui ----------
def test_ui_passes_host_and_settings(tmp_path, monkeypatch):
    calls = []
    monkeypatch.setattr("matrix_to_im.cli.serve", lambda **kw: calls.append(kw))
    toml = tmp_path / "s.toml"
    toml.write_text('[ui]\nhost = "0.0.0.0"\nport = 9000\n', encoding="utf-8")
    assert main(["ui", "--settings", str(toml), "--no-browser"]) == 0
    assert calls[0]["host"] == "0.0.0.0" and calls[0]["port"] == 9000
    assert calls[0]["settings"].port == 9000
    assert main(["ui", "--settings", str(toml), "--host", "127.0.0.1", "--port", "9100", "--no-browser"]) == 0
    assert calls[1]["host"] == "127.0.0.1" and calls[1]["port"] == 9100


def test_write_summary_lists_missing_styles_and_final_count(write_env, tmp_path, monkeypatch, capsys):
    im, argv = write_env
    final = tmp_path / "final.json"
    data = json.loads(final.read_text(encoding="utf-8"))
    data["records"].append(_rec("ZZZ", ["z"], "T", "000", 9))
    final.write_text(json.dumps(data), encoding="utf-8")
    assert main(argv + ["--dry-run"]) == 0
    out = capsys.readouterr().out
    assert "Style MATRIX tidak ada di IM (tidak ditulis): 1" in out
    assert "  TIDAK ADA DI IM (tidak ditulis): ZZZ" in out
    monkeypatch.setattr("sys.stdin", io.StringIO(TEST_PASSWORD + "\n"))
    assert main(argv + ["--yes"]) == 0
    assert "1 style MATRIX tidak ada di IM (tidak ditulis)" in capsys.readouterr().out

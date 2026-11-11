import os
import secrets
import subprocess

import pytest
from openpyxl import load_workbook

from matrix_to_im import excel_mac
from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.excel_mac import (
    ExcelUnavailable, build_insert_script, excel_available, run_excel_inserts,
)
from matrix_to_im.im_layout import read_layout
from matrix_to_im.im_lock import check_modify_password
from matrix_to_im.write_plan import Colorway, InsertOp, plan_inserts
from tests.helpers.im_factory import TEST_PASSWORD, GroupSpec, add_file_sharing, build_im_sheet


def _op(at, count, fill_from, fill_to, fixes=None, after=None):
    return InsertOp("X - BOTH WHS", at, count, fill_from, fill_to, after, fixes or {})


def test_script_ops_in_given_order_and_fill_limited_to_last_col():
    script = build_insert_script([_op(100, 1, 99, 100), _op(57, 2, 56, 58)], "F27", "IJ")
    assert script.index('insert into range (range "100:100"') < script.index('insert into range (range "57:58"')
    assert 'insert into range (range "57:58" of ws) shift shift down' in script
    assert 'fill down range "A56:IJ58" of ws' in script
    assert 'system attribute "MATRIX2IM_PW"' in script
    assert "calculation manual" in script and "calculation automatic" in script


def test_script_formula_fixes_grouped_in_runs_and_escaped():
    fixes = {2: '=IF(A1="x",1,2)', 3: "=C1", 5: "=E1"}
    script = build_insert_script([_op(60, 2, 59, 61, fixes, 62)], "F27", "IJ")
    assert 'set formula of range "B62:C62" of ws to {{"=IF(A1=\\"x\\",1,2)", "=C1"}}' in script
    assert 'set formula of range "E62:E62" of ws to {{"=E1"}}' in script
    # fixes mengikuti insert + fill down op yang sama, sebelum op berikutnya
    two = build_insert_script([_op(60, 2, 59, 61, fixes, 62), _op(10, 1, 9, 10)], "F27", "IJ")
    assert two.index('range "B62:C62"') < two.index('insert into range (range "10:10"')


_OPEN = "open workbook workbook file name"


def test_script_never_uses_active_workbook_and_has_timeout():
    script = build_insert_script([_op(57, 2, 56, 58)], "F27", "IJ", 180)
    assert "active workbook" not in script
    assert "set wb to open workbook" not in script
    assert "with timeout of 180 seconds" in script and "end timeout" in script
    assert script.index("set wb to missing value") < script.index(_OPEN)


def test_script_name_collision_guard_before_open_touches_nothing():
    """Workbook lain bernama sama: Excel tak bisa membuka file kita; berhenti sebelum apa pun diubah."""
    script = build_insert_script([_op(57, 2, 56, 58)], "F27", "IJ")
    guard = ('if (name of w) is nm then return "ERROR 0: Excel sudah membuka workbook bernama " '
             '& nm & "; tutup dulu"')
    assert "set nm to item 2 of argv" in script
    assert guard in script
    assert script.index(guard) < script.index(_OPEN)
    assert script.index(guard) < script.index("set display alerts to false")   # bahkan alert belum disentuh


def test_script_verifies_full_name_after_open_before_setting_wb():
    script = build_insert_script([_op(57, 2, 56, 58)], "F27", "IJ")
    check = "if (full name of w) is p then set wb to contents of w"
    assert "set p to item 1 of argv" in script and check in script
    assert script.index(_OPEN) < script.index(check) < script.index('set ws to worksheet "F27" of wb')
    sets = [line.strip() for line in script.splitlines() if line.strip().startswith("set wb to")]
    assert sets == ["set wb to missing value", "set wb to missing value"]   # awal, dan sesudah ditutup
    assert script.count("set wb to") == 3 and script.count(check) == 1
    missing = 'if wb is missing value then error "workbook yang dibuka tidak ditemukan di Excel: " & p'
    assert script.index(check) < script.index(missing) < script.index("set ws to worksheet")


def test_script_closes_only_the_verified_workbook():
    script = build_insert_script([_op(57, 2, 56, 58)], "F27", "IJ")
    closes = [line.strip() for line in script.splitlines() if "close " in line]
    assert closes == ["close wb saving no", "if wb is not missing value then close wb saving no"]
    error_path = script[script.index("on error"):]
    assert "if wb is not missing value then close wb saving no" in error_path
    assert "save " not in error_path


def test_script_error_path_restores_each_setting_in_own_try_alerts_first():
    script = build_insert_script([_op(57, 2, 56, 58)], "F27", "IJ")
    lines = [line.strip() for line in script[script.index("on error"):].splitlines()]
    restores = ["set display alerts to true", "set calculation to calculation automatic",
                "set screen updating to true", "if wb is not missing value then close wb saving no"]
    pos = [lines.index(r) for r in restores]
    assert pos == sorted(pos)
    for i in pos:
        assert lines[i - 1] == "try" and lines[i + 1] == "end try", lines[i - 1:i + 2]


def _capture(monkeypatch, result=None, exc=None):
    seen = {}

    def fake_run(args, **kw):
        seen["args"], seen["kw"] = args, kw
        if exc:
            raise exc
        return result or subprocess.CompletedProcess(args, 0, stdout="OK\n", stderr="")

    monkeypatch.setattr(excel_mac, "excel_available", lambda: True)
    monkeypatch.setattr(subprocess, "run", fake_run)
    return seen


def test_password_only_in_env(monkeypatch, tmp_path):
    seen = _capture(monkeypatch)
    secs = run_excel_inserts(tmp_path / "a.xlsx", "sandi-rahasia", [_op(57, 2, 56, 58)], "F27", "IJ")
    assert isinstance(secs, float)
    assert "sandi-rahasia" not in " ".join(seen["args"])
    assert "sandi-rahasia" not in seen["kw"]["input"]
    assert seen["kw"]["env"]["MATRIX2IM_PW"] == "sandi-rahasia"


def test_timeout_message(monkeypatch, tmp_path):
    _capture(monkeypatch, exc=subprocess.TimeoutExpired("osascript", 5))
    with pytest.raises(Matrix2ImError, match="Excel tidak merespons dalam 5 detik"):
        run_excel_inserts(tmp_path / "a.xlsx", "pw", [], "F27", "IJ", timeout_s=5)


def test_error_output_message_and_automation_hint(monkeypatch, tmp_path):
    _capture(monkeypatch, subprocess.CompletedProcess([], 0, stdout="ERROR -1743: tidak diizinkan\n", stderr=""))
    with pytest.raises(Matrix2ImError) as ei:
        run_excel_inserts(tmp_path / "a.xlsx", "pw", [], "F27", "IJ")
    assert "Excel gagal: ERROR -1743" in str(ei.value)
    assert "Automation" in str(ei.value)


def test_nonzero_returncode_fails(monkeypatch, tmp_path):
    _capture(monkeypatch, subprocess.CompletedProcess([], 1, stdout="", stderr="boom"))
    with pytest.raises(Matrix2ImError, match="Excel gagal: boom"):
        run_excel_inserts(tmp_path / "a.xlsx", "pw", [], "F27", "IJ")


def test_unavailable_raises(monkeypatch, tmp_path):
    monkeypatch.setattr(excel_mac, "excel_available", lambda: False)
    with pytest.raises(ExcelUnavailable):
        run_excel_inserts(tmp_path / "a.xlsx", "pw", [], "F27", "IJ")


def _record_runs(monkeypatch, first_result=None, first_exc=None, cleanup_exc=None):
    """Panggilan subprocess.run: pertama skrip sisip, kedua (bila ada) skrip pembersih."""
    calls = []

    def fake_run(args, **kw):
        calls.append((args, kw))
        if len(calls) == 1:
            if first_exc:
                raise first_exc
            return first_result
        if cleanup_exc:
            raise cleanup_exc
        return subprocess.CompletedProcess(args, 0, stdout="OK\n", stderr="")

    monkeypatch.setattr(excel_mac, "excel_available", lambda: True)
    monkeypatch.setattr(subprocess, "run", fake_run)
    return calls


def _assert_cleanup(calls, path, pw="sandi-rahasia"):
    assert len(calls) == 2
    args, kw = calls[1]
    assert str(path.resolve()) in args
    assert pw not in " ".join(args) and pw not in kw["input"] and "MATRIX2IM_PW" not in kw["input"]
    assert str(path.resolve()) not in kw["input"]
    assert "saving no" in kw["input"] and "full name" in kw["input"]
    assert "quit" not in kw["input"] and "with timeout of 25 seconds" in kw["input"]
    assert args[0] == "osascript" and calls[0][0][2] == str(path.resolve())


def test_cleanup_after_timeout(monkeypatch, tmp_path):
    calls = _record_runs(monkeypatch, first_exc=subprocess.TimeoutExpired("osascript", 5))
    with pytest.raises(Matrix2ImError, match="Excel tidak merespons"):
        run_excel_inserts(tmp_path / "a.xlsx", "sandi-rahasia", [], "F27", "IJ", timeout_s=5)
    _assert_cleanup(calls, tmp_path / "a.xlsx")


def test_cleanup_after_error_output(monkeypatch, tmp_path):
    calls = _record_runs(monkeypatch, subprocess.CompletedProcess([], 0, stdout="ERROR -50: x\n", stderr=""))
    with pytest.raises(Matrix2ImError, match="Excel gagal: ERROR -50"):
        run_excel_inserts(tmp_path / "a.xlsx", "sandi-rahasia", [], "F27", "IJ")
    _assert_cleanup(calls, tmp_path / "a.xlsx")


def test_failing_cleanup_does_not_mask_original_error(monkeypatch, tmp_path):
    _record_runs(monkeypatch, first_exc=subprocess.TimeoutExpired("osascript", 5),
                 cleanup_exc=subprocess.TimeoutExpired("osascript", 30))
    with pytest.raises(Matrix2ImError, match="Excel tidak merespons"):
        run_excel_inserts(tmp_path / "a.xlsx", "pw", [], "F27", "IJ", timeout_s=5)
    _record_runs(monkeypatch, subprocess.CompletedProcess([], 0, stdout="ERROR -1: x", stderr=""),
                 cleanup_exc=OSError("osascript hilang"))
    with pytest.raises(Matrix2ImError, match="Excel gagal: ERROR -1"):
        run_excel_inserts(tmp_path / "a.xlsx", "pw", [], "F27", "IJ")


def test_close_workbook_if_open_never_raises(monkeypatch, tmp_path):
    def boom(*a, **k):
        raise OSError("tidak ada osascript")

    monkeypatch.setattr(subprocess, "run", boom)
    assert excel_mac.close_workbook_if_open(tmp_path / "a.xlsx") is None


def test_success_does_not_run_cleanup(monkeypatch, tmp_path):
    calls = _record_runs(monkeypatch, subprocess.CompletedProcess([], 0, stdout="OK\n", stderr=""))
    run_excel_inserts(tmp_path / "a.xlsx", "pw", [], "F27", "IJ")
    assert len(calls) == 1


_NEEDS_EXCEL = pytest.mark.skipif(os.environ.get("MATRIX2IM_EXCEL_TESTS") != "1" or not excel_available(),
                                  reason="butuh MATRIX2IM_EXCEL_TESTS=1 dan Microsoft Excel")


def _excel_full_names() -> list[str]:
    """Full name semua workbook yang sedang terbuka di Excel (hanya untuk test bermarker excel)."""
    script = "\n".join(['tell application "Microsoft Excel"', '  set out to ""',
                        "  repeat with w in (get every workbook)",
                        "    set out to out & (full name of w) & linefeed", "  end repeat",
                        "  return out", "end tell", ""])
    proc = subprocess.run(["osascript", "-"], input=script, capture_output=True, text=True, timeout=60)
    assert proc.returncode == 0, proc.stderr
    return [line for line in proc.stdout.splitlines() if line]


@pytest.mark.excel
@_NEEDS_EXCEL
def test_real_excel_error_after_open_leaves_no_workbook_open(tmp_path):
    """Error sesudah open berhasil (sheet tidak ada): workbook kita tertutup tanpa disimpan (I2 + R7)."""
    path = tmp_path / f"im-kerja-{secrets.token_hex(4)}.xlsx"
    build_im_sheet(path, [GroupSpec("AAA - K", [["a"], ["b"]])])
    add_file_sharing(path)
    before = path.read_bytes()
    ops = [InsertOp("AAA - K", 56, 1, 55, 56, None, {})]
    with pytest.raises(Matrix2ImError, match="Excel gagal: ERROR"):
        run_excel_inserts(path, TEST_PASSWORD, ops, "SHEET TIDAK ADA", "Z", timeout_s=180)
    assert str(path.resolve()) not in _excel_full_names()
    assert path.read_bytes() == before


@pytest.mark.excel
@_NEEDS_EXCEL
def test_real_excel_inserts_and_widens_subtotals(tmp_path):
    path = tmp_path / "im.xlsx"
    build_im_sheet(path, [GroupSpec("AAA - K", [["a"], ["b"], ["c"]]), GroupSpec("BBB - K", [["x"]])])
    add_file_sharing(path)
    layout = read_layout(path)
    cws = {"AAA": [Colorway((c,), None, None) for c in "abcde"],
           "BBB": [Colorway((c,), None, None) for c in "xyz"]}
    ops = plan_inserts(layout, cws)
    assert len(ops) == 2
    run_excel_inserts(path, TEST_PASSWORD, ops, "F27", "Z", timeout_s=180)
    ws = load_workbook(path)["F27"]
    # AAA: baris 55-59, subtotal 60; BBB: baris 61-63, subtotal 64
    assert ws["B60"].value == "=SUBTOTAL(9,B55:B59)"
    assert ws["B64"].value == "=SUBTOTAL(9,B61:B63)"
    for row in (58, 59, 62, 63):
        assert ws.cell(row, 26).value == "catatan"
    check_modify_password(path, TEST_PASSWORD)


def test_script_alerts_off_before_open_and_restored_on_both_paths():
    script = build_insert_script([_op(57, 2, 56, 58)], "F27", "IJ")
    assert script.index("set display alerts to false") < script.index("open workbook workbook file name")
    assert script.index('tell application "Microsoft Excel"') < script.index("set display alerts to false")
    restore = "set display alerts to true"
    assert script.count(restore) == 2
    assert script.index(restore) < script.index("save wb") or script.index(restore) < script.index('return "OK"')
    assert script.rindex(restore) > script.index("on error")


def test_close_script_alerts_off_and_restored():
    script = excel_mac._close_script(30)
    assert script.index("set display alerts to false") < script.index("close wb saving no")
    assert script.index("set display alerts to true") > script.index("close wb saving no")


def test_close_script_loop_inside_try_so_alerts_always_restored():
    lines = [line.strip() for line in excel_mac._close_script(30).splitlines()]
    loop, end = lines.index("repeat with wb in (get every workbook)"), lines.index("end repeat")
    assert lines[loop - 1] == "try" and lines[end + 1] == "end try"
    restore = lines.index("set display alerts to true")
    assert restore > end + 1 and lines[restore - 1] == "try"


def test_insert_args_carry_resolved_path_and_name(monkeypatch, tmp_path):
    seen = _capture(monkeypatch)
    run_excel_inserts(tmp_path / "im-kerja-0a1b2c3d.xlsx", "pw", [_op(57, 2, 56, 58)], "F27", "IJ")
    path = tmp_path / "im-kerja-0a1b2c3d.xlsx"
    assert seen["args"] == ["osascript", "-", str(path.resolve()), "im-kerja-0a1b2c3d.xlsx"]


def test_cleanup_after_nonzero_returncode_without_error_output(monkeypatch, tmp_path):
    calls = _record_runs(monkeypatch, subprocess.CompletedProcess([], -9, stdout="", stderr=""))
    with pytest.raises(Matrix2ImError):
        run_excel_inserts(tmp_path / "a.xlsx", "sandi-rahasia", [], "F27", "IJ")
    _assert_cleanup(calls, tmp_path / "a.xlsx")


def test_empty_detail_message_has_return_code(monkeypatch, tmp_path):
    _capture(monkeypatch, subprocess.CompletedProcess([], -9, stdout="", stderr=""))
    with pytest.raises(Matrix2ImError) as ei:
        run_excel_inserts(tmp_path / "a.xlsx", "pw", [], "F27", "IJ")
    assert str(ei.value) == "Excel berhenti tanpa pesan (kode -9)"

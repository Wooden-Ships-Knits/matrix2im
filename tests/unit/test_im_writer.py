import dataclasses
from copy import copy
import json
import logging
import os
import re
import shutil
import zipfile
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from matrix_to_im import im_writer
from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.im_layout import read_layout
from matrix_to_im.write_plan import CellChange, FillChange
from matrix_to_im.xlsx_cells import read_fills
from matrix_to_im.xlsx_styles import check_appended
from matrix_to_im.im_lock import check_modify_password
from matrix_to_im.im_writer import commit_write, fingerprint, prepare_write
from matrix_to_im.settings import Settings
from tests.helpers.im_factory import TEST_PASSWORD, GroupSpec, add_file_sharing, build_im_sheet
from tests.helpers.report_reader import read_report, report_kinds

PW = TEST_PASSWORD


def _rec(style, colors, ws=None, whs=None, n=1):
    rec = {"DESCRIPTION": style}
    for i, c in enumerate(colors, 1):
        rec[f"COLOR{i}"] = c
    rec["WS TAG COLOR"], rec["WHS"] = ws, whs
    rec["_ref"] = f"SHEET No.{n} | B1 | x"
    return rec


def _make(tmp_path, groups, engine="none"):
    im = build_im_sheet(tmp_path / "im.xlsx", groups)
    add_file_sharing(im)
    backup = tmp_path / "cadangan"
    return im, Settings(im_master=im, backup_dir=backup, insert_engine=engine)


def _simple(tmp_path, engine="none"):
    groups = [GroupSpec("AAA - K", [["a1"], ["a2"]])]
    im, st = _make(tmp_path, groups, engine)
    final = {"records": [_rec("AAA", ["x1"], "T1", "000", 1), _rec("AAA", ["x2"], "T2", "SY", 2)]}
    return im, st, final


def fake_inserter(path, password, ops, sheet, last_col, drop=0):
    """Sisip via openpyxl pada file sintetis (tanpa gambar) lalu pulihkan fileSharing.

    Seperti fill down Excel, baris sisipan menyalin nilai dan format (termasuk warna) baris sumber.
    """
    wb = load_workbook(path)
    ws = wb[sheet]
    for op in ops:
        ws.insert_rows(op.at_row, op.count - drop)
        for r in range(op.at_row, op.at_row + op.count - drop):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = ws.cell(op.fill_from, c).value
                ws.cell(r, c)._style = copy(ws.cell(op.fill_from, c)._style)
    wb.save(path)
    add_file_sharing(path)
    return 0.1


def _values(path, row, cols=(5, 24, 25)):
    wb = load_workbook(path, read_only=True)
    try:
        ws = wb["F27"]
        return [ws.cell(row, c).value for c in cols]
    finally:
        wb.close()


@pytest.fixture
def excel_on(monkeypatch):
    monkeypatch.setattr(im_writer, "excel_available", lambda: True)


def _kerja_empty(st):
    k = st.backup_dir / "kerja"
    return not k.exists() or list(k.iterdir()) == []


# ---------- prepare ----------

def test_prepare_requires_settings(tmp_path):
    _, _, final = _simple(tmp_path)
    with pytest.raises(Matrix2ImError, match="im_master"):
        prepare_write(final, Settings(backup_dir=tmp_path / "c"), PW)
    with pytest.raises(Matrix2ImError, match="backup_dir"):
        prepare_write(final, Settings(im_master=tmp_path / "im.xlsx"), PW)
    with pytest.raises(Matrix2ImError, match="tidak ditemukan"):
        prepare_write(final, Settings(im_master=tmp_path / "nope.xlsx", backup_dir=tmp_path / "c"), PW)


def test_prepare_wrong_password(tmp_path):
    _, st, final = _simple(tmp_path)
    with pytest.raises(Matrix2ImError, match="Password IM MASTER salah"):
        prepare_write(final, st, "salah")


def test_prepare_non_zip_file_gives_indonesian_error(tmp_path):
    bad = tmp_path / "im.xlsx"
    bad.write_text("bukan excel")
    with pytest.raises(Matrix2ImError, match="IM MASTER tidak bisa dibaca sebagai file Excel"):
        prepare_write({"records": []}, Settings(im_master=bad, backup_dir=tmp_path / "c"), PW)


def test_prepare_blocks_when_owner_file_present(tmp_path):
    im, st, final = _simple(tmp_path)
    owner = im.with_name("~$" + im.name)
    owner.write_bytes(bytes([5]) + "Budi!".encode())
    with pytest.raises(Matrix2ImError, match="IM MASTER sedang dibuka oleh Budi!"):
        prepare_write(final, st, PW)


def test_prepare_summary_counts(tmp_path, excel_on):
    groups = [GroupSpec("AAA - K", [["a"]]),                       # kurang: 1 baris, butuh 2
              GroupSpec("BBB - K", [["b1"], ["b2"], ["b3"]]),      # lebih: 3 baris, butuh 2
              GroupSpec("CCC - K", [["c"]])]                       # tak cocok
    im, st = _make(tmp_path, groups, "excel")
    final = {"records": [_rec("AAA", ["x1"], "T", "000", 1), _rec("AAA", ["x2"], "T", "000", 2),
                         _rec("BBB", ["y1"], "T", "000", 3), _rec("BBB", ["y2"], "T", "000", 4),
                         _rec("ZZZ", ["z"], "T", "000", 5)]}
    p = prepare_write(final, st, PW)
    s = p.summary
    assert s["engine"] == "excel"
    assert (s["styles"], s["groups"]) == (2, 2)
    assert (s["rows_inserted"], s["groups_inserted"], s["short_groups"]) == (1, 1, 0)
    assert s["unmatched_im_styles"] == 1 and s["missing_styles"] == 1 and s["skipped"] == 0
    assert s["cells_changed"] == len(p.values.changes) == 12      # 4 baris x (COLOR1, WS TAG, WHS)
    assert s["changed_colorway_rows"] == 4
    assert s["extra_rows"] == 1 and p.values.extra_rows == [60]   # baris ke-3 BBB (sesudah sisip)
    assert s["warnings"] == len(p.values.warnings) == 1
    assert "grup hanya 1 baris" in p.values.warnings[0] and "AAA - K" in p.values.warnings[0]
    assert p.fingerprint == fingerprint(im) and p.im_master == im
    assert len(p.inserts) == 1 and p.projected.groups[0].rows == (55, 56)


def test_excel_unavailable_falls_back_to_none(tmp_path, monkeypatch):
    monkeypatch.setattr(im_writer, "excel_available", lambda: False)
    groups = [GroupSpec("AAA - K", [["a"]])]
    im, st = _make(tmp_path, groups, "excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    p = prepare_write(final, st, PW)
    assert p.engine == "none" and p.inserts == [] and p.summary["engine"] == "none"
    assert "Excel tidak tersedia: hanya mengisi baris yang sudah ada" in p.values.warnings
    assert p.summary["short_groups"] == 1


def test_engine_none_requested_has_no_unavailable_warning(tmp_path):
    _, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    assert p.engine == "none" and p.values.warnings == []


# ---------- commit ----------

def test_commit_engine_none_in_place_backup_log(tmp_path):
    im, st, final = _simple(tmp_path)
    before = im.read_bytes()
    ino = os.stat(im).st_ino
    p = prepare_write(final, st, PW)
    phases = []
    rep = commit_write(p, final, st, PW, "andi", "matrix.xlsx", progress=phases.append)
    assert rep.ok, rep.message
    assert os.stat(im).st_ino == ino
    assert rep.backup is not None and rep.backup.read_bytes() == before
    assert rep.backup.parent == st.backup_dir and "sebelum matrix2im" in rep.backup.name
    assert _values(im, 55) == ["x1", "T1", "000"] and _values(im, 56) == ["x2", "T2", "SY"]
    check_modify_password(im, PW)
    assert phases == ["cek password", "cadangan", "tulis nilai", "verifikasi",
                      "simpan ke IM MASTER", "selesai"]
    lines = (st.backup_dir / "matrix2im-log.jsonl").read_text(encoding="utf-8").splitlines()
    assert len(lines) == 1
    entry = json.loads(lines[0])
    assert set(entry) == {"waktu", "user", "matrix", "im_master", "engine", "summary", "backup", "hasil", "pesan",
                          "laporan"}
    assert entry["hasil"] == "ok" and entry["user"] == "andi" and PW not in lines[0]
    assert rep.log_file == st.backup_dir / "matrix2im-log.jsonl" and rep.inserted_rows == []
    assert _kerja_empty(st)


def test_commit_with_fake_inserter_extends_groups(tmp_path, excel_on):
    groups = [GroupSpec("AAA - K", [["a"]]), GroupSpec("BBB - K", [["b1"], ["b2"]])]
    im, st = _make(tmp_path, groups, "excel")
    final = {"records": [_rec("AAA", [f"x{i}"], "T", "000", i) for i in range(1, 4)]
             + [_rec("BBB", ["y1"], "T", "000", 9), _rec("BBB", ["y2"], "T", "000", 10),
                _rec("BBB", ["y3"], "T", "000", 11)]}
    p = prepare_write(final, st, PW)
    assert p.summary["rows_inserted"] == 3
    phases = []
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", progress=phases.append, inserter=fake_inserter)
    assert rep.ok, rep.message
    assert "Excel: sisip baris" in phases and phases.index("Excel: sisip baris") < phases.index("tulis nilai")
    lay = read_layout(im)
    assert [len(g.rows) for g in lay.groups] == [3, 3]
    assert [_values(im, r)[0] for r in (55, 56, 57)] == ["x1", "x2", "x3"]
    assert sorted(rep.inserted_rows) == [56, 57, 60]
    check_modify_password(im, PW)
    assert _kerja_empty(st)


def test_commit_aborts_if_file_changed_after_prepare(tmp_path):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    im.write_bytes(im.read_bytes() + b"x")
    changed = im.read_bytes()
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert not rep.ok
    assert "IM MASTER berubah atau sedang dibuka sejak rencana dibuat; buat rencana lagi" in rep.message
    assert im.read_bytes() == changed
    entry = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text().splitlines()[0])
    assert entry["hasil"] == "gagal" and entry["pesan"] == rep.message
    assert _kerja_empty(st)


def test_commit_aborts_if_layout_differs_from_projection(tmp_path, excel_on):
    groups = [GroupSpec("AAA - K", [["a"]])]
    im, st = _make(tmp_path, groups, "excel")
    final = {"records": [_rec("AAA", [f"x{i}"], n=i) for i in range(1, 4)]}
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino

    def short_inserter(path, password, ops, sheet, last_col):
        return fake_inserter(path, password, ops, sheet, last_col, drop=1)

    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=short_inserter)
    assert not rep.ok and "ata letak" in rep.message
    assert im.read_bytes() == before and os.stat(im).st_ino == ino
    assert rep.backup is not None and rep.backup.read_bytes() == before
    assert _kerja_empty(st)


def test_commit_aborts_when_inserter_fails(tmp_path, excel_on):
    groups = [GroupSpec("AAA - K", [["a"]])]
    im, st = _make(tmp_path, groups, "excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    p = prepare_write(final, st, PW)
    before = im.read_bytes()

    def boom(*a, **k):
        raise Matrix2ImError("Excel gagal: macet")

    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=boom)
    assert not rep.ok and "Excel gagal: macet" in rep.message
    assert im.read_bytes() == before and _kerja_empty(st)


def test_commit_aborts_when_owner_file_appears(tmp_path):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before = im.read_bytes()
    im.with_name("~$" + im.name).write_bytes(bytes([3]) + b"Ani")
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert not rep.ok and "berubah atau sedang dibuka" in rep.message
    assert im.read_bytes() == before


def test_commit_wrong_password_reports_and_logs(tmp_path):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before = im.read_bytes()
    rep = commit_write(p, final, st, "salah-sandi", "andi", "m.xlsx")
    assert not rep.ok and "Password IM MASTER salah" in rep.message
    assert im.read_bytes() == before
    log = (st.backup_dir / "matrix2im-log.jsonl").read_text()
    assert "salah-sandi" not in log and json.loads(log)["hasil"] == "gagal"


def test_password_not_in_log_report_or_messages(tmp_path, excel_on):
    groups = [GroupSpec("AAA - K", [["a"]])]
    im, st = _make(tmp_path, groups, "excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    p = prepare_write(final, st, PW)

    def leaky(*a, **k):
        raise RuntimeError(f"gagal dengan sandi {PW}")

    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=leaky)
    assert not rep.ok
    blob = repr(rep) + (st.backup_dir / "matrix2im-log.jsonl").read_text() + repr(p.summary)
    assert PW not in blob
    for f in tmp_path.rglob("*"):
        if f.is_file() and f.suffix in (".jsonl", ".txt", ".log"):
            assert PW not in f.read_text()
    ok_rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=fake_inserter)
    assert ok_rep.ok, ok_rep.message
    assert PW not in repr(ok_rep) + (st.backup_dir / "matrix2im-log.jsonl").read_text()


# ---------- jalur batal dan pemulihan ----------

def _aborted(st, rep, text, im=None, expect=None, ino=None):
    """Kontrak semua jalur batal: ok False, pesan Indonesia, folder kerja kosong, log 'gagal'."""
    assert not rep.ok and text in rep.message, rep.message
    assert _kerja_empty(st)
    last = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text().splitlines()[-1])
    assert last["hasil"] == "gagal" and text in last["pesan"]
    if im is not None:
        assert im.read_bytes() == expect
        if ino is not None:
            assert os.stat(im).st_ino == ino


def _rewrite_entry(path, name, fn):
    """Ganti isi satu entri zip (fixture sintetis saja)."""
    tmp = path.with_suffix(".rw")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w") as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            dst.writestr(item.filename, fn(data) if item.filename == name else data)
    tmp.replace(path)


def test_commit_aborts_at_step9_recheck_when_owner_file_appears(tmp_path):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino

    def progress(phase):
        if phase == "simpan ke IM MASTER":
            im.with_name("~$" + im.name).write_bytes(bytes([3]) + b"Ani")

    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", progress=progress)
    _aborted(st, rep, "berubah atau sedang dibuka", im, before, ino)


def test_commit_aborts_at_step9_recheck_when_file_changed(tmp_path):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    ino = os.stat(im).st_ino

    def progress(phase):
        if phase == "simpan ke IM MASTER":
            im.write_bytes(im.read_bytes() + b"x")

    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", progress=progress)
    _aborted(st, rep, "berubah atau sedang dibuka", im, p_bytes := im.read_bytes(), ino)
    assert p_bytes.endswith(b"x") and rep.backup is not None and rep.backup.read_bytes() != p_bytes


def test_commit_aborts_when_values_differ_from_plan(tmp_path, monkeypatch):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino
    real = im_writer.plan_values

    def other(layout, colorways):
        plan = real(layout, colorways)
        return dataclasses.replace(plan, changes=[*plan.changes, CellChange(55, 5, None, "liar")])

    monkeypatch.setattr(im_writer, "plan_values", other)
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, "Perubahan nilai", im, before, ino)


def _patch_apply(monkeypatch, mutate_changes=None, mutate_out=None, mutate_fills=None):
    real = im_writer.apply_cell_values

    def wrapped(src, dst, sheet, changes, fills=()):
        changes, fills = list(changes), list(fills)
        real(src, dst, sheet, mutate_changes(changes) if mutate_changes else changes,
             mutate_fills(fills) if mutate_fills else fills)
        if mutate_out:
            mutate_out(dst)

    monkeypatch.setattr(im_writer, "apply_cell_values", wrapped)


@pytest.mark.parametrize("case,text", [
    ("value", "Verifikasi gagal: sel"),
    ("lock", "tidak dikunci"),
    ("other_entry", "bagian lain"),
    ("read_only", "readOnlyRecommended"),
])
def test_commit_aborts_on_step8_verification(tmp_path, monkeypatch, case, text):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino
    if case == "value":
        _patch_apply(monkeypatch, mutate_changes=lambda ch: [dataclasses.replace(ch[0], new="lain"), *ch[1:]])
    elif case == "lock":
        _patch_apply(monkeypatch, mutate_out=lambda dst: _rewrite_entry(
            dst, "xl/workbook.xml", lambda d: re.sub(rb"<fileSharing[^>]*/>", b"", d)))
    elif case == "read_only":
        _patch_apply(monkeypatch, mutate_out=lambda dst: _rewrite_entry(
            dst, "xl/workbook.xml", lambda d: d.replace(b' readOnlyRecommended="1"', b"")))
    else:
        _patch_apply(monkeypatch, mutate_out=lambda dst: _rewrite_entry(
            dst, "docProps/core.xml", lambda d: d + b"<!--x-->"))
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, text, im, before, ino)


def test_step9_write_failure_restores_original(tmp_path, monkeypatch):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino
    real, calls = os.fsync, []

    def flaky(fd):
        calls.append(fd)
        if len(calls) == 1:
            raise OSError("disk penuh")
        return real(fd)

    monkeypatch.setattr(im_writer.os, "fsync", flaky)
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, "File asli sudah dipulihkan dari cadangan", im, before, ino)


def test_step9_restore_failure_says_may_be_damaged_with_backup_path(tmp_path, monkeypatch):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)

    def always(fd):
        raise OSError("disk rusak")

    monkeypatch.setattr(im_writer.os, "fsync", always)
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert not rep.ok and "mungkin rusak" in rep.message and str(rep.backup) in rep.message
    assert "dipulihkan dari" not in rep.message and _kerja_empty(st)


def test_step9_open_failure_says_original_not_changed(tmp_path, monkeypatch):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino
    import builtins

    def fake_open(file, mode="r", *a, **k):
        if mode == "r+b" and Path(file) == im:
            raise PermissionError("ditolak")
        return builtins.open(file, mode, *a, **k)

    monkeypatch.setattr(im_writer, "open", fake_open, raising=False)
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, "File asli tidak diubah", im, before, ino)
    assert "rusak" not in rep.message


def test_commit_aborts_when_backup_is_corrupt(tmp_path, monkeypatch):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before = im.read_bytes()
    real = shutil.copyfile

    def corrupt(src, dst, *a, **k):
        out = real(src, dst, *a, **k)
        if Path(dst).parent == st.backup_dir:
            with open(dst, "ab") as f:
                f.write(b"x")
        return out

    monkeypatch.setattr(im_writer.shutil, "copyfile", corrupt)
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, "berubah atau sedang dibuka", im, before)
    assert rep.backup is None and not list(st.backup_dir.glob("*sebelum matrix2im*"))


def _littering(inner=fake_inserter, then_raise=None):
    def ins(path, password, ops, sheet, last_col):
        (path.parent / "~$im-kerja.xlsx").write_bytes(b"sisa")
        sub = path.parent / "sub"
        sub.mkdir()
        (sub / "x").write_bytes(b"1")
        if then_raise:
            raise Matrix2ImError(then_raise)
        return inner(path, password, ops, sheet, last_col)
    return ins


def test_cleanup_leftovers_do_not_change_success(tmp_path, excel_on):
    im, st = _make(tmp_path, [GroupSpec("AAA - K", [["a"]])], "excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    p = prepare_write(final, st, PW)
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=_littering())
    assert rep.ok, rep.message
    assert json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text())["hasil"] == "ok"
    assert _kerja_empty(st)


def test_cleanup_leftovers_keep_real_abort_message(tmp_path, excel_on):
    im, st = _make(tmp_path, [GroupSpec("AAA - K", [["a"]])], "excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    p = prepare_write(final, st, PW)
    before = im.read_bytes()
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=_littering(then_raise="Excel gagal: macet"))
    _aborted(st, rep, "Excel gagal: macet", im, before)


def test_commit_missing_im_master_gives_indonesian_message(tmp_path):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    im.unlink()
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert not rep.ok and rep.message == f"IM MASTER tidak ditemukan: {im}"


def test_unexpected_error_logs_scrubbed_traceback(tmp_path, excel_on, caplog):
    im, st = _make(tmp_path, [GroupSpec("AAA - K", [["a"]])], "excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    p = prepare_write(final, st, PW)

    def leaky(*a, **k):
        raise RuntimeError(f"sandi {PW}")

    with caplog.at_level(logging.ERROR):
        rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=leaky)
    assert not rep.ok and "Traceback" in caplog.text and "RuntimeError" in caplog.text
    assert PW not in caplog.text


def test_single_row_warning_only_for_single_row_groups(tmp_path, excel_on):
    groups = [GroupSpec("AAA - K", [["a"]]),                          # 1 baris, butuh 3
              GroupSpec("BBB - K", [["b1"], ["b2"]])]                 # 2 baris, butuh 3
    im, st = _make(tmp_path, groups, "excel")
    final = {"records": [_rec("AAA", [f"x{i}"], "T", "000", i) for i in range(3)]
             + [_rec("BBB", [f"y{i}"], "T", "000", 10 + i) for i in range(3)]}
    p = prepare_write(final, st, PW)
    assert len(p.inserts) == 2
    warns = [w for w in p.values.warnings if "grup hanya 1 baris" in w]
    assert len(warns) == 1 and "AAA - K" in warns[0]


def test_temp_copy_has_unique_name_per_run(tmp_path, excel_on):
    """Nama unik: workbook lain bernama sama di Excel tak mungkin tertukar dengan salinan kerja run ini."""
    im, st = _make(tmp_path, [GroupSpec("AAA - K", [["a"]])], "excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    seen = []

    def spy(path, password, ops, sheet, last_col):
        seen.append(path)
        raise Matrix2ImError("Excel gagal: berhenti di sini")

    for _ in range(2):
        commit_write(prepare_write(final, st, PW), final, st, PW, "andi", "m.xlsx", inserter=spy)
    assert len(seen) == 2 and seen[0].name != seen[1].name
    for path in seen:
        assert re.fullmatch(r"im-kerja-[0-9a-f]{8}\.xlsx", path.name), path.name
        assert path.parent.parent == st.backup_dir / "kerja"


# ---------- I3: Excel tidak boleh menghilangkan bagian file ----------

_PARTS = {"xl/media/image1.png": b"\x89PNG-uji", "xl/drawings/drawing1.xml": b"<xdr:wsDr/>",
          "xl/comments1.xml": b"<comments/>", "xl/externalLinks/externalLink1.xml": b"<externalLink/>"}
_KEEP = ("xl/media/", "xl/drawings/", "xl/comments", "xl/threadedComments/", "xl/externalLinks/",
         "xl/embeddings/")


def _add_parts(path, parts):
    """Tambahkan entri zip ke file sintetis (meniru gambar/drawing/comment di IM asli)."""
    tmp = path.with_suffix(".parts")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            dst.writestr(item, src.read(item.filename))
        for name, data in parts.items():
            dst.writestr(name, data)
    tmp.replace(path)


def _drop_entry(path, name):
    tmp = path.with_suffix(".drop")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename != name:
                dst.writestr(item, src.read(item.filename))
    tmp.replace(path)


def faithful_inserter(mutate=None):
    """fake_inserter yang, seperti Excel, mempertahankan entri yang dibuang openpyxl; lalu mutasi opsional."""
    def ins(path, password, ops, sheet, last_col):
        with zipfile.ZipFile(path) as z:
            keep = {n: z.read(n) for n in z.namelist() if n.startswith(_KEEP)}
        fake_inserter(path, password, ops, sheet, last_col)
        with zipfile.ZipFile(path) as z:
            missing = {n: d for n, d in keep.items() if n not in z.namelist()}
        if missing:
            _add_parts(path, missing)
        if mutate:
            mutate(path)
        return 0.1
    return ins


def _excel_env(tmp_path, cf=False):
    im = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("AAA - K", [["a"]])])
    if cf:
        from openpyxl.formatting.rule import CellIsRule
        from openpyxl.styles import PatternFill
        wb = load_workbook(im)
        wb["F27"].conditional_formatting.add(
            "E55:E60", CellIsRule(operator="equal", formula=['"x1"'], fill=PatternFill("solid", fgColor="FFFF0000")))
        wb.save(im)
    add_file_sharing(im)
    _add_parts(im, _PARTS)
    st = Settings(im_master=im, backup_dir=tmp_path / "cadangan", insert_engine="excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    return im, st, final


@pytest.mark.parametrize("mutate,names", [
    (lambda p: _drop_entry(p, "xl/media/image1.png"), ("xl/media", "image1.png")),
    (lambda p: _drop_entry(p, "xl/comments1.xml"), ("xl/comments", "comments1.xml")),
    (lambda p: _rewrite_entry(p, "xl/workbook.xml", lambda d: d.replace(b' readOnlyRecommended="1"', b"")),
     ("readOnlyRecommended",)),
    (lambda p: _rewrite_entry(p, "xl/workbook.xml", lambda d: d.replace(b'userName="Tester"', b'userName="Excel"')),
     ("userName",)),
])
def test_commit_aborts_when_excel_loses_parts_or_lock_attributes(tmp_path, excel_on, mutate, names):
    im, st, final = _excel_env(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=faithful_inserter(mutate))
    _aborted(st, rep, "Excel mengubah bagian IM MASTER yang harus utuh", im, before, ino)
    for name in names:
        assert name in rep.message, rep.message
    assert "file asli tidak diubah" in rep.message


def test_commit_keeps_parts_with_faithful_excel(tmp_path, excel_on):
    im, st, final = _excel_env(tmp_path)
    p = prepare_write(final, st, PW)
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=faithful_inserter())
    assert rep.ok, rep.message
    with zipfile.ZipFile(im) as z:
        assert set(_PARTS) <= set(z.namelist())
    assert "PERINGATAN" not in rep.message


def test_cf_count_change_after_excel_is_warning_only(tmp_path, excel_on):
    im, st, final = _excel_env(tmp_path, cf=True)
    p = prepare_write(final, st, PW)

    def drop_cf(path):
        _rewrite_entry(path, "xl/worksheets/sheet1.xml",
                       lambda d: re.sub(rb"<conditionalFormatting.*?</conditionalFormatting>", b"", d, flags=re.S))

    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=faithful_inserter(drop_cf))
    assert rep.ok, rep.message
    assert "PERINGATAN" in rep.message and "conditional formatting" in rep.message and "1 menjadi 0" in rep.message
    entry = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text().splitlines()[-1])
    assert entry["hasil"] == "ok" and "conditional formatting" in entry["pesan"]


def test_cf_count_streams_and_counts_token_across_chunk_boundary(tmp_path):
    path = tmp_path / "x.zip"
    token = b"<conditionalFormatting"
    body = b"a" * ((1 << 20) - 5) + token + b"b" * 10 + token
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/worksheets/sheet1.xml", body)
        z.writestr("xl/worksheets/sheet2.xml", token)
        z.writestr("xl/styles.xml", token)          # bukan worksheet: tidak dihitung
    with zipfile.ZipFile(path) as z:
        assert im_writer._count_token(z, token) == 3


# ---------- M1: dihentikan paksa saat menimpa; M4: kunci antarproses ----------

def _lock_free(st):
    import fcntl
    with open(st.backup_dir / "matrix2im.lock", "a+b") as fd:
        try:
            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            return False
        return True


def test_keyboard_interrupt_during_in_place_write_restores_logs_and_reraises(tmp_path, monkeypatch):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino
    real, calls = os.fsync, []

    def interrupt(fd):
        calls.append(fd)
        if len(calls) == 1:
            assert im.read_bytes() != before     # byte baru sudah masuk ke file asli
            raise KeyboardInterrupt
        return real(fd)

    monkeypatch.setattr(im_writer.os, "fsync", interrupt)
    with pytest.raises(KeyboardInterrupt):
        commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert im.read_bytes() == before and os.stat(im).st_ino == ino
    last = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text().splitlines()[-1])
    assert last["hasil"] == "gagal" and "KeyboardInterrupt" in last["pesan"]
    assert "File asli sudah dipulihkan dari cadangan" in last["pesan"]
    assert _kerja_empty(st) and _lock_free(st)


def test_keyboard_interrupt_before_step9_says_original_untouched(tmp_path, excel_on):
    im, st = _make(tmp_path, [GroupSpec("AAA - K", [["a"]])], "excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    p = prepare_write(final, st, PW)
    before = im.read_bytes()

    def stop(*a, **k):
        raise KeyboardInterrupt

    with pytest.raises(KeyboardInterrupt):
        commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=stop)
    last = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text().splitlines()[-1])
    assert last["hasil"] == "gagal" and "File asli tidak diubah" in last["pesan"]
    assert im.read_bytes() == before and _kerja_empty(st)


def test_commit_refuses_while_other_process_holds_lock(tmp_path):
    import fcntl
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    before = im.read_bytes()
    st.backup_dir.mkdir(parents=True)
    with open(st.backup_dir / "matrix2im.lock", "a+b") as other:      # proses lain di PC ini
        fcntl.flock(other, fcntl.LOCK_EX | fcntl.LOCK_NB)
        rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert not rep.ok and rep.message == "Ada penulisan lain yang sedang berjalan (proses lain di PC ini)"
    assert im.read_bytes() == before and rep.backup is None
    assert not list(st.backup_dir.glob("*sebelum matrix2im*")) and not (st.backup_dir / "kerja").exists()
    last = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text().splitlines()[-1])
    assert last["hasil"] == "gagal" and last["pesan"] == rep.message
    rep2 = commit_write(p, final, st, PW, "andi", "m.xlsx")          # kunci dilepas: boleh lagi
    assert rep2.ok, rep2.message


def test_lock_held_during_run_and_released_after(tmp_path):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    seen = []
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", progress=lambda ph: seen.append(_lock_free(st)))
    assert rep.ok and seen and not any(seen[:-1])
    assert _lock_free(st)
    bad = commit_write(p, final, st, "salah-sandi", "andi", "m.xlsx")
    assert not bad.ok and _lock_free(st)


# ---------- I1: style MATRIX yang tidak ada di IM ----------

def test_final_message_mentions_missing_styles_count(tmp_path):
    im, st, final = _simple(tmp_path)
    final["records"] += [_rec("ZZZ", ["z"], n=7), _rec("YYY", ["y"], n=8)]
    p = prepare_write(final, st, PW)
    assert p.values.missing_styles == ["ZZZ", "YYY"]
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert rep.ok and "2 style MATRIX tidak ada di IM (tidak ditulis)" in rep.message


def test_final_message_without_missing_styles_has_no_count(tmp_path):
    im, st, final = _simple(tmp_path)
    rep = commit_write(prepare_write(final, st, PW), final, st, PW, "andi", "m.xlsx")
    assert rep.ok and "tidak ada di IM" not in rep.message


# ---------- I4/R18: laporan .xlsx tiap run ----------

def _report_path(st, im):
    found = sorted(st.backup_dir.glob(f"{im.stem} * laporan matrix2im.xlsx"))
    assert len(found) == 1, found
    return found[0]


def test_run_report_written_on_success_and_logged(tmp_path):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert rep.ok, rep.message
    path = _report_path(st, im)
    assert rep.laporan == path
    stamp = rep.backup.name.removeprefix(f"{im.stem} ").removesuffix(" sebelum matrix2im.xlsx")
    assert path.name == f"{im.stem} {stamp} laporan matrix2im.xlsx"          # pasangan cadangan
    entry = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text().splitlines()[-1])
    assert entry["laporan"] == str(path)
    rows = read_report(path.read_bytes())
    assert rows[0] == ("DESCRIPTION", "sel", "baris", "kolom", "lama", "baru", "jenis")
    assert report_kinds(rows)["ubah"] == len(p.values.changes)
    assert ("AAA - K", "E55", 55, "COLOR1", "a1", "x1", "ubah") in rows
    assert all(PW not in str(v) for row in rows for v in row)


def test_run_report_written_on_abort_with_reason(tmp_path, excel_on):
    im, st = _make(tmp_path, [GroupSpec("AAA - K", [["a"]])], "excel")
    final = {"records": [_rec("AAA", ["x1"]), _rec("AAA", ["x2"], n=2)]}
    p = prepare_write(final, st, PW)

    def boom(*a, **k):
        raise Matrix2ImError("Excel gagal: macet")

    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=boom)
    assert not rep.ok and rep.laporan is not None and rep.laporan == _report_path(st, im)
    rows = read_report(rep.laporan.read_bytes())
    assert ("", "", "", "", "", "Penulisan gagal: Excel gagal: macet", "peringatan") in rows
    assert report_kinds(rows)["sisip"] == 1
    entry = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text().splitlines()[-1])
    assert entry["laporan"] == str(rep.laporan)


def test_wrong_password_report_has_no_password(tmp_path):
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    rep = commit_write(p, final, st, "salah-sandi-77", "andi", "m.xlsx")
    assert not rep.ok and rep.laporan is not None
    cells = [str(v) for row in read_report(rep.laporan.read_bytes()) for v in row]
    assert not [c for c in cells if "salah-sandi-77" in c or PW in c]


def test_no_run_report_when_other_process_holds_lock(tmp_path):
    import fcntl
    im, st, final = _simple(tmp_path)
    p = prepare_write(final, st, PW)
    st.backup_dir.mkdir(parents=True)
    with open(st.backup_dir / "matrix2im.lock", "a+b") as other:
        fcntl.flock(other, fcntl.LOCK_EX | fcntl.LOCK_NB)
        rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert not rep.ok and rep.laporan is None
    assert not list(st.backup_dir.glob("*laporan matrix2im*"))
    entry = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text().splitlines()[-1])
    assert entry["laporan"] is None


# ---------- M5: izin folder Google Drive ----------

@pytest.mark.parametrize("where", ["is_file", "read"])
def test_prepare_permission_error_points_to_drive_permission(tmp_path, monkeypatch, where):
    im, st, final = _simple(tmp_path)
    if where == "is_file":
        real = Path.is_file

        def denied(self):
            if self == im:
                raise PermissionError(1, "Operation not permitted", str(self))
            return real(self)

        monkeypatch.setattr(Path, "is_file", denied)
    else:
        def denied(path, password):
            raise PermissionError(13, "Permission denied", str(path))

        monkeypatch.setattr(im_writer, "check_modify_password", denied)
    with pytest.raises(Matrix2ImError) as ei:
        prepare_write(final, st, PW)
    assert "izin akses folder Google Drive" in str(ei.value) and str(im) in str(ei.value)


# ---------- R19: style dilewati bukan "IM tanpa MATRIX" ----------

def test_skipped_style_not_counted_as_unmatched_im_style(tmp_path):
    groups = [GroupSpec("AAA - K", [["a"]]), GroupSpec("BBB - K", [["b"]]), GroupSpec("CCC - K", [["c"]])]
    im, st = _make(tmp_path, groups)
    b1, b2 = _rec("BBB", ["b1"], n=3), _rec("BBB", ["b2"], n=4)
    b2["_ref"] = "SHEET2 No.4 | B1 | x"                     # colorway beda antar-sheet -> dilewati
    final = {"records": [_rec("AAA", ["x1"]), b1, b2]}
    p = prepare_write(final, st, PW)
    assert [k.style for k in p.skipped] == ["BBB"]
    assert p.values.unmatched_im_styles == ["CCC"] and p.summary["unmatched_im_styles"] == 1



# ---------- warna latar WS TAG COLOR ----------

def _painted(tmp_path, engine="none", groups=None):
    groups = groups or [GroupSpec("AAA - K", [["a1"], ["a2"], ["a3"]], ws_tags=["T1", "T2", "OLD"],
                                  whs=["000", "SY", "000"], ws_tag_fills=["FF99FF", "FFFF00", "FF99FF"],
                                  whs_fills=["FF99FF", None, None])]
    im, st = _make(tmp_path, groups, engine)
    final = {"records": [_rec("AAA", ["a1"], "T1", "000", 1), _rec("AAA", ["a2"], "T2", "BOTH WHS", 2)]}
    return im, st, final


def test_commit_paints_ws_tag_by_whs_and_clears_extra_row(tmp_path):
    im, st, final = _painted(tmp_path)
    p = prepare_write(final, st, PW)
    # 56 dan 57 dua kolom (WS TAG/WHS); COLOR1 baris 55 dan 56 menjadi abu (benang final)
    assert p.summary["fills_changed"] == 6 and p.summary["extra_rows"] == 1
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert rep.ok, rep.message
    assert "6 warna latar diubah" in rep.message
    assert read_fills(im, "F27", [(55, 5), (56, 5), (57, 5)]) == {
        (55, 5): "7F7F7F", (56, 5): "7F7F7F", (57, 5): "FFFF00"}      # baris lebih: warna COLOR dibiarkan
    cells = [(r, c) for r in (55, 56, 57) for c in (24, 25)]
    assert read_fills(im, "F27", cells) == {
        (55, 24): "FF99FF", (56, 24): "969696", (57, 24): "FFFF00",
        (55, 25): "FF99FF", (56, 25): "969696", (57, 25): "FFFF00"}
    assert _values(im, 56) == ["a2", "T2", "BOTH WHS"]
    assert _values(im, 57) == ["a3", None, None]                       # COLOR baris lebih dibiarkan
    with zipfile.ZipFile(rep.backup) as a, zipfile.ZipFile(im) as b:
        check_appended(a.read("xl/styles.xml"), b.read("xl/styles.xml"))
    check_modify_password(im, PW)


def test_commit_inserted_rows_get_whs_fill(tmp_path, excel_on):
    groups = [GroupSpec("AAA - K", [["a1"]], ws_tags=["T1"], whs=["000"], ws_tag_fills=["FFFF00"])]
    im, st, final = _painted(tmp_path, "excel", groups)
    p = prepare_write(final, st, PW)
    assert [(f.row, f.col, f.old, f.new) for f in p.values.fills if f.col in (24, 25)] == [
        (55, 24, "FFFF00", "FF99FF"), (55, 25, None, "FF99FF"), (56, 24, "FFFF00", "969696"), (56, 25, None, "969696")]
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx", inserter=fake_inserter)
    assert rep.ok, rep.message
    assert read_fills(im, "F27", [(55, 24), (56, 24), (55, 25), (56, 25)]) == {
        (55, 24): "FF99FF", (56, 24): "969696", (55, 25): "FF99FF", (56, 25): "969696"}
    assert read_fills(im, "F27", [(55, 5), (56, 5)]) == {(55, 5): "7F7F7F", (56, 5): "7F7F7F"}   # sisipan juga


def test_commit_aborts_when_fill_not_written_as_planned(tmp_path, monkeypatch):
    im, st, final = _painted(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino
    _patch_apply(monkeypatch, mutate_fills=lambda fs: [dataclasses.replace(fs[0], new="FFFF00"), *fs[1:]])
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, "Verifikasi gagal: warna sel E55", im, before, ino)    # fs[0] = COLOR1 baris 55


def test_commit_aborts_when_styles_change_beyond_append(tmp_path, monkeypatch):
    im, st, final = _painted(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino
    _patch_apply(monkeypatch, mutate_out=lambda dst: _rewrite_entry(
        dst, "xl/styles.xml", lambda b: b.replace(b'<sz val="11"', b'<sz val="12"', 1)))
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, "styles.xml", im, before, ino)


def test_commit_aborts_when_styles_change_without_planned_fills(tmp_path, monkeypatch):
    im, st = _make(tmp_path, [GroupSpec("AAA - K", [["a1"]], color1_fill="7F7F7F")])
    final = {"records": [_rec("AAA", ["x1"], "T1", None, 1)]}       # WHS tak diketahui, COLOR1 sudah abu: tanpa warna
    p = prepare_write(final, st, PW)
    assert p.values.fills == []
    before, ino = im.read_bytes(), os.stat(im).st_ino
    _patch_apply(monkeypatch, mutate_out=lambda dst: _rewrite_entry(
        dst, "xl/styles.xml", lambda b: b.replace(b'<sz val="11"', b'<sz val="12"', 1)))
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, "bagian lain", im, before, ino)


def test_commit_aborts_when_fills_differ_from_plan(tmp_path, monkeypatch):
    im, st, final = _painted(tmp_path)
    p = prepare_write(final, st, PW)
    before, ino = im.read_bytes(), os.stat(im).st_ino
    real = im_writer.plan_values

    def other(layout, colorways):
        plan = real(layout, colorways)
        return dataclasses.replace(plan, fills=[*plan.fills, FillChange(55, 24, "FF99FF", "FFCC99")])

    monkeypatch.setattr(im_writer, "plan_values", other)
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, "warna", im, before, ino)



def _row_filled(tmp_path, rgb="FF99FF"):
    """Grup 1 baris tanpa sel WS TAG di XML, tapi baris 55 berwarna lewat gaya baris (customFormat)."""
    im, st = _make(tmp_path, [GroupSpec("AAA - K", [["a1"]])])
    wb = load_workbook(im)
    wb["F27"].row_dimensions[55].fill = PatternFill("solid", fgColor="FF" + rgb)
    wb.save(im)
    add_file_sharing(im)
    return im, st


def test_commit_keeps_inherited_fill_when_text_written_into_absent_cell(tmp_path):
    im, st = _row_filled(tmp_path)
    final = {"records": [_rec("AAA", ["a1"], "T1", "000", 1)]}
    p = prepare_write(final, st, PW)
    assert [f for f in p.values.fills if f.col in (24, 25)] == []
    assert {(c.row, c.col) for c in p.values.changes} >= {(55, 24)}
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    assert rep.ok, rep.message
    assert read_fills(im, "F27", [(55, 24)]) == {(55, 24): "FF99FF"}
    assert _values(im, 55) == ["a1", "T1", "000"]


def test_commit_aborts_on_unplanned_fill_change(tmp_path, monkeypatch):
    """Verifikasi warna mencakup semua baris grup, bukan hanya sel di rencana."""
    im, st, final = _painted(tmp_path)
    p = prepare_write(final, st, PW)
    assert (55, 24) not in {(f.row, f.col) for f in p.values.fills}
    before, ino = im.read_bytes(), os.stat(im).st_ino
    _patch_apply(monkeypatch, mutate_fills=lambda fs: [*fs, FillChange(55, 24, "FF99FF", "FFCC99")])
    rep = commit_write(p, final, st, PW, "andi", "m.xlsx")
    _aborted(st, rep, "Verifikasi gagal: warna sel X55", im, before, ino)


# ---------- pilih collection (permintaan user 12 Nov 2026) ----------

def _two_collections(tmp_path):
    """AAA di collection FALL, BBB di collection LUXE; IM berisi keduanya (baris 55 dan 57)."""
    groups = [GroupSpec("AAA - K", [["a"]]), GroupSpec("BBB - K", [["b"]])]
    im, st = _make(tmp_path, groups)
    a, b = _rec("AAA", ["x1"], "T1", "000"), _rec("BBB", ["y1"], "T2", "SY")
    a["_ref"], b["_ref"] = "FALL No.1 | B1 | x", "LUXE No.1 | B1 | x"
    return im, st, {"records": [a, b]}


def test_prepare_with_collection_plans_only_chosen_styles(tmp_path):
    im, st, final = _two_collections(tmp_path)
    p = prepare_write(final, st, PW, collections=["FALL"])
    assert {c.row for c in p.values.changes} == {55}
    assert p.values.unmatched_im_styles == []          # BBB tidak dipilih, bukan "IM tanpa data MATRIX"
    assert p.summary["collections"] == ["FALL"] and p.summary["not_selected_styles"] == 1


def test_prepare_without_collection_plans_all_styles(tmp_path):
    im, st, final = _two_collections(tmp_path)
    p = prepare_write(final, st, PW)
    assert {c.row for c in p.values.changes} == {55, 57}
    assert p.summary["collections"] is None and p.summary["not_selected_styles"] == 0


def test_commit_writes_only_chosen_collection_even_given_full_final(tmp_path):
    im, st, final = _two_collections(tmp_path)
    p = prepare_write(final, st, PW, collections=["FALL"])
    rep = commit_write(p, final, st, PW, "andi", "matrix.xlsx")
    assert rep.ok, rep.message
    assert _values(im, 55) == ["x1", "T1", "000"] and _values(im, 57) == ["b", None, None]
    entry = json.loads((st.backup_dir / "matrix2im-log.jsonl").read_text(encoding="utf-8").splitlines()[0])
    assert entry["summary"]["collections"] == ["FALL"]

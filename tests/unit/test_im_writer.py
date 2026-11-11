import json
import os
from pathlib import Path

import pytest
from openpyxl import load_workbook

from matrix_to_im import im_writer
from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.im_layout import read_layout
from matrix_to_im.im_lock import check_modify_password
from matrix_to_im.im_writer import commit_write, fingerprint, prepare_write
from matrix_to_im.settings import Settings
from tests.helpers.im_factory import TEST_PASSWORD, GroupSpec, add_file_sharing, build_im_sheet

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
    """Sisip via openpyxl pada file sintetis (tanpa gambar) lalu pulihkan fileSharing."""
    wb = load_workbook(path)
    ws = wb[sheet]
    for op in ops:
        ws.insert_rows(op.at_row, op.count - drop)
        for r in range(op.at_row, op.at_row + op.count - drop):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = ws.cell(op.fill_from, c).value
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
    assert s["cells_changed"] == len(p.values.changes) > 0
    assert s["changed_colorway_rows"] == len(p.values.changes and p.values.changed_colorway_rows)
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
    assert set(entry) == {"waktu", "user", "matrix", "im_master", "engine", "summary", "backup", "hasil", "pesan"}
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

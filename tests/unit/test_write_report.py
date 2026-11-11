"""Laporan perubahan IM MASTER (spec §6): isi, jenis baris, dan workbook .xlsx."""
import pytest

from matrix_to_im import im_writer
from matrix_to_im.im_writer import prepare_write
from matrix_to_im.settings import Settings
from matrix_to_im.write_report import REPORT_HEADER, report_rows, report_xlsx_bytes
from tests.helpers.im_factory import TEST_PASSWORD, GroupSpec, add_file_sharing, build_im_sheet
from tests.helpers.report_reader import read_report, report_kinds


def _rec(style, colors, n, sheet="SHEET", ws=None, whs=None):
    rec = {"DESCRIPTION": style, "WS TAG COLOR": ws, "WHS": whs, "_ref": f"{sheet} No.{n} | B1 | x"}
    for i, c in enumerate(colors, 1):
        rec[f"COLOR{i}"] = c
    return rec


@pytest.fixture
def preview(tmp_path, monkeypatch):
    """Rencana dengan sisipan, style dilewati, style MATRIX tak ada di IM, dan style IM tanpa MATRIX."""
    monkeypatch.setattr(im_writer, "excel_available", lambda: True)
    groups = [GroupSpec("AAA - K", [["old-a"]]),                     # butuh 2 baris: 1 sisipan
              GroupSpec("BBB - K", [["b"]]),                         # style dilewati (beda antar-sheet)
              GroupSpec("CCC - K", [["c"]])]                         # IM tanpa MATRIX
    im = build_im_sheet(tmp_path / "im.xlsx", groups)
    add_file_sharing(im)
    final = {"records": [_rec("AAA", ["x1"], 1, ws="T1", whs="000"), _rec("AAA", ["x2", "y2"], 2),
                         _rec("BBB", ["b1"], 3, "S1"), _rec("BBB", ["b2"], 4, "S2"),
                         _rec("ZZZ", ["z"], 5)]}
    st = Settings(im_master=im, backup_dir=tmp_path / "cadangan", insert_engine="excel")
    return prepare_write(final, st, TEST_PASSWORD)


def _by_kind(rows):
    out = {}
    for row in rows:
        out.setdefault(row[-1], []).append(row)
    return out


def test_report_rows_cover_every_kind(preview):
    rows = report_rows(preview)
    kinds = _by_kind(rows)
    assert {"ubah", "sisip", "baris colorway berubah", "dilewati", "tidak ada di IM", "IM tanpa MATRIX",
            "peringatan"} <= set(kinds)
    assert all(len(r) == len(REPORT_HEADER) for r in rows)
    assert len(kinds["ubah"]) == len(preview.values.changes)
    assert kinds["dilewati"] == [("BBB", "", "", "", "", preview.skipped[0].reason, "dilewati")]
    assert [r[0] for r in kinds["tidak ada di IM"]] == ["ZZZ"]
    assert [r[0] for r in kinds["IM tanpa MATRIX"]] == ["CCC"]          # BBB dilewati, bukan "IM tanpa MATRIX"
    assert any("grup hanya 1 baris" in r[5] for r in kinds["peringatan"])


def test_change_rows_name_cell_row_column_old_new(preview):
    ubah = {r[1]: r for r in _by_kind(report_rows(preview))["ubah"]}
    assert ubah["E55"] == ("AAA - K", "E55", "55", "COLOR1", "old-a", "x1", "ubah")
    assert ubah["X55"][3] == "WS TAG COLOR" and ubah["X55"][5] == "T1"
    assert ubah["Y55"][3] == "WHS" and ubah["Y55"][5] == "000"
    assert ubah["G56"][3] == "COLOR2" and ubah["G56"][5] == "y2"


def test_insert_rows_carry_row_numbers_in_written_file(preview):
    sisip = _by_kind(report_rows(preview))["sisip"]
    assert [(r[0], r[2]) for r in sisip] == [("AAA - K", "56")]
    assert [int(r[2]) for r in sisip] == im_writer._inserted_rows(preview.inserts)


def test_changed_colorway_rows_show_old_and_new_yarns(preview):
    rows = {r[2]: r for r in _by_kind(report_rows(preview))["baris colorway berubah"]}
    assert rows["55"][4:6] == ("old-a", "x1")
    assert rows["56"][4:6] == ("old-a", "x2 / y2")      # baris sisipan menyalin baris di atasnya


def test_notes_become_warning_rows(preview):
    rows = report_rows(preview, notes=["Penulisan gagal: contoh"])
    assert ("", "", "", "", "", "Penulisan gagal: contoh", "peringatan") in rows


def test_xlsx_workbook_header_rows_and_known_change(preview):
    rows = report_rows(preview)
    data = report_xlsx_bytes(rows)
    assert data[:2] == b"PK"
    got = read_report(data)
    assert got[0] == REPORT_HEADER == ("DESCRIPTION", "sel", "baris", "kolom", "lama", "baru", "jenis")
    assert len(got) == len(rows) + 1
    assert ("AAA - K", "E55", 55, "COLOR1", "old-a", "x1", "ubah") in got      # baris sebagai angka
    expected = {}
    for r in rows:
        expected[r[-1]] = expected.get(r[-1], 0) + 1
    assert report_kinds(got) == expected
    assert all(TEST_PASSWORD not in str(v) for row in got for v in row)


def test_xlsx_header_bold_text_kept_as_text_and_illegal_chars_dropped():
    from io import BytesIO

    from openpyxl import load_workbook
    rows = [("=SUM(A1)", "", "", "WHS", "000", "a\x0bb", "peringatan")]
    wb = load_workbook(BytesIO(report_xlsx_bytes(rows)))          # bukan read_only: perlu font
    ws = wb["Laporan"]
    assert ws["A1"].font.b and ws["A2"].data_type == "s" and ws["A2"].value == "=SUM(A1)"
    assert ws["E2"].value == "000" and ws["F2"].value == "ab" and ws["B2"].value is None

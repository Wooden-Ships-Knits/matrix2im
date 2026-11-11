"""Regression tulis IM MASTER: file `data/` hanya dibaca; hasil tulis selalu ke salinan di tmp_path."""
import os
import re
import shutil
import sys
import time
import zipfile

import pytest

from matrix_to_im.excel_mac import excel_available
from matrix_to_im.im_layout import read_layout
from matrix_to_im.im_lock import check_modify_password
from matrix_to_im.im_writer import commit_write, prepare_write
from matrix_to_im.legend_config import IM_SHEET
from matrix_to_im.settings import Settings
from matrix_to_im.write_plan import plan_inserts, plan_values, project_layout, style_colorways
from matrix_to_im.xlsx_cells import apply_cell_values, sheet_xml_path

pytestmark = pytest.mark.regression

# Angka rencana sisip F27 (salinan uji "AI TEST V2") setelah R14 (baris subtotal = baris ber-rumus
# SUBTOTAL) dan R15 (kunci grup = "STYLE - SIZE"; teks sesudahnya adalah catatan).
# Rekonsiliasi dengan perkiraan spike awal (yang mengelompokkan menurut label):
#  - 27 op grup 1 baris di spike = 20 + 7 baris subtotal yang salah dibaca sebagai grup 1 baris oleh
#    aturan label lama (ANSEL x4 di 9305/9314/9323/9332, OLIVIA 2976, WRAP 3298, CHIN-CHIN 5632);
#  - 184 op / 321 baris (hasil aturan label) menjadi 177 / 286, sama dengan spike;
#  - 179 style IM tak cocok di spike = 195 - 16 label satu segmen (GRASSY, FALL, COTTON, dst.).
EXPECTED_OPS = 177
EXPECTED_ROWS = 286
EXPECTED_SINGLE_ROW_OPS = 20
EXPECTED_UNMATCHED_IM_STYLES = 195


@pytest.fixture(scope="module")
def colorways(golden_final):
    result, skips = style_colorways(golden_final)
    assert skips == []
    return result


@pytest.fixture(scope="module")
def layout(im_write_path):
    return read_layout(im_write_path)


def test_subtotal_rows_by_formula_f27(im_write_path, layout):
    """R14: baris subtotal = baris ber-rumus SUBTOTAL; label bervariasi dan tidak menentukan."""
    import openpyxl
    wb = openpyxl.load_workbook(im_write_path, read_only=True)   # hanya baca, tidak pernah save
    try:
        sub_rows = {r for r, row in enumerate(wb[IM_SHEET].iter_rows(values_only=True), 1)
                    if any(isinstance(v, str) and re.match(r"=\s*subtotal\(", v, re.I) for v in row)}
    finally:
        wb.close()
    assert sub_rows == set(layout.subtotal_formulas)
    group_rows = {r for g in layout.groups for r in g.rows}
    assert not sub_rows & group_rows
    assigned = {g.subtotal_row for g in layout.groups if g.subtotal_row is not None}
    assert assigned == sub_rows
    assert {2976, 9305} <= sub_rows        # label varian "-BOTH WHS" dan "- SY ONLY" yang dulu terlewat
    assert 8207 in group_rows and 8207 not in sub_rows   # WHITNEY ... BOTH WHS - ALT 1: baris data
    assert {2976, 9305} <= assigned


def test_annotated_whitney_rows_merge_into_size_blocks_f27(layout):
    """R15: baris berkatalog "- ALT n" / "BOTH WHS" adalah baris colorway di blok sizenya."""
    by_first = {g.rows[0]: g for g in layout.groups}
    sm, ml = by_first[8207], by_first[8217]
    assert (sm.description, sm.rows, sm.subtotal_row) == (
        "WHITNEY POCKET TUNIC CHUNKY - S/M", tuple(range(8207, 8216)), 8216)
    assert (ml.description, ml.rows, ml.subtotal_row) == (
        "WHITNEY POCKET TUNIC CHUNKY - M/L", tuple(range(8217, 8226)), 8226)
    assert all(len(g.description.split(" - ")) <= 2 for g in layout.groups)


def test_style_colorways_f27(golden_final):
    colorways, skips = style_colorways(golden_final)
    assert len(colorways) == 211
    assert skips == []
    hutton = next(v for k, v in colorways.items() if k.startswith("HUTTON"))
    assert len(hutton) == 8
    by_color = {cw.colors: cw for cw in hutton}
    # "000 & SY" adalah nilai WHS (bukan WS TAG COLOR, yang masih "Need Review" di golden)
    assert by_color[("GLASGOW",)].whs == "000 & SY"
    assert by_color[("FAWN",)].whs == "000 & SY"
    crop = colorways["CROP BOYFRIEND CREW COTTON"]
    assert len(crop) == 5


def test_layout_and_insert_plan_f27(layout, colorways):
    ops = plan_inserts(layout, colorways)
    group_len = {(g.description, g.rows[0]): len(g.rows) for g in layout.groups}
    single = [o for o in ops if group_len.get((o.description, o.fill_from)) == 1]
    assert len(ops) == EXPECTED_OPS
    assert sum(o.count for o in ops) == EXPECTED_ROWS
    assert len(single) == EXPECTED_SINGLE_ROW_OPS
    # urutan eksekusi bawah ke atas dan tidak ada op yang saling menimpa
    assert [o.at_row for o in ops] == sorted((o.at_row for o in ops), reverse=True)


def test_value_plan_f27_runs_fast(im_write_path, colorways):
    start = time.perf_counter()
    layout = read_layout(im_write_path)
    ops = plan_inserts(layout, colorways)
    plan = plan_values(project_layout(layout, ops), colorways)
    elapsed = time.perf_counter() - start
    assert elapsed < 60
    assert plan.missing_styles == []
    assert len(plan.unmatched_im_styles) == EXPECTED_UNMATCHED_IM_STYLES


def test_apply_values_on_copy_keeps_parts(tmp_path, im_write_path, colorways, layout):
    """Tanpa sisip (engine none): hanya sheet XML dan workbook.xml yang boleh berbeda."""
    plan = plan_values(layout, colorways)
    assert plan.changes
    dst = tmp_path / "hasil.xlsx"
    start = time.perf_counter()
    apply_cell_values(im_write_path, dst, IM_SHEET, plan.changes)
    assert time.perf_counter() - start < 30
    with zipfile.ZipFile(im_write_path) as a, zipfile.ZipFile(dst) as b:
        assert a.namelist() == b.namelist()
        sheet = _sheet_part(a)
        changed = {n for n in a.namelist() if a.read(n) != b.read(n)}
        assert changed <= {sheet, "xl/workbook.xml"}
        assert sheet in changed
    password = os.environ.get("MATRIX2IM_IM_PASSWORD")
    if password:
        check_modify_password(dst, password)


def _sheet_part(zf: zipfile.ZipFile) -> str:
    return sheet_xml_path(zf, IM_SHEET)


def _count(zf: zipfile.ZipFile, pattern: str) -> int:
    return sum(1 for n in zf.namelist() if re.fullmatch(pattern, n))


def _xml_count(zf: zipfile.ZipFile, token: bytes) -> int:
    """Jumlah token di semua sheet XML (bukan hanya F27)."""
    return sum(zf.read(n).count(token) for n in zf.namelist()
               if re.fullmatch(r"xl/worksheets/[^/]+\.xml", n))


@pytest.mark.excel
@pytest.mark.regression
@pytest.mark.skipif(os.environ.get("MATRIX2IM_EXCEL_TESTS") != "1" or sys.platform != "darwin"
                    or not excel_available(),
                    reason="butuh MATRIX2IM_EXCEL_TESTS=1, macOS, dan Microsoft Excel")
def test_full_write_on_copy(tmp_path, im_write_path, golden_final):
    password = os.environ.get("MATRIX2IM_IM_PASSWORD")
    if not password:
        pytest.skip("butuh MATRIX2IM_IM_PASSWORD (password ubah salinan uji)")
    target = tmp_path / im_write_path.name
    shutil.copyfile(im_write_path, target)
    settings = Settings(im_master=target, backup_dir=tmp_path, insert_engine="excel")
    start = time.perf_counter()
    preview = prepare_write(golden_final, settings, password)
    report = commit_write(preview, golden_final, settings, password, "uji regression", "F27 MATRIX MASTER")
    elapsed = time.perf_counter() - start
    assert report.ok, report.message
    assert elapsed < 300
    with zipfile.ZipFile(im_write_path) as src, zipfile.ZipFile(target) as out:
        assert _count(out, r"xl/media/.*") == _count(src, r"xl/media/.*")
        assert _count(out, r"xl/drawings/drawing[^/]*\.xml") == _count(src, r"xl/drawings/drawing[^/]*\.xml")
        for token in (b"<conditionalFormatting", b"<cfRule"):
            assert _xml_count(out, token) == _xml_count(src, token), token
    # baris 229 di file sumber (grup 1 baris) setelah semua sisipan di atasnya
    src_group = next(g for g in read_layout(im_write_path).groups if g.rows == (229,))
    assert src_group.description == "BOYFRIEND RAGLAN CREW CHUNKY COTTON - X/L"
    shifted = 229 + sum(o.count for o in preview.inserts if o.at_row <= 229)
    planned = next(g for g in preview.projected.groups if shifted in g.rows)
    assert planned.description == src_group.description and len(planned.rows) == 5
    after = read_layout(target)
    grp = next(g for g in after.groups if shifted in g.rows)
    assert grp.rows == planned.rows and grp.subtotal_row == planned.subtotal_row is not None
    first, last = grp.rows[0], grp.rows[-1]
    formulas = after.subtotal_formulas[grp.subtotal_row].values()
    assert any(re.search(rf"[A-Z]{first}:[A-Z]+\$?{last}\b", f) for f in formulas), (first, last)
    check_modify_password(target, password)

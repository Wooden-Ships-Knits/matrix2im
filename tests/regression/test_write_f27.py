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

# Angka rencana sisip F27 (salinan uji "AI TEST V2"). Spike awal memperkirakan 177 op / 286 baris /
# 179 style IM tak cocok karena mengelompokkan baris hanya di baris subtotal. Aturan pengelompokan
# Task 2 juga memecah grup saat DESCRIPTION berubah (varian "- SY ONLY", "-BOTH WHS", label musim
# seperti GRASSY/FALL, dll.), sehingga hasil nyata: 184 op, 321 baris, 27 op grup 1 baris, 195 style.
# Penjelasan rinci ada di .superpowers/sdd/2026-11-10-tulis-im-master/task-12-report.md (bagian R3).
EXPECTED_OPS = 184
EXPECTED_ROWS = 321
EXPECTED_SINGLE_ROW_OPS = 27
EXPECTED_UNMATCHED_IM_STYLES = 195


@pytest.fixture(scope="module")
def colorways(golden_final):
    result, skips = style_colorways(golden_final)
    assert skips == []
    return result


@pytest.fixture(scope="module")
def layout(im_write_path):
    return read_layout(im_write_path)


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
    part = _sheet_part(zf)
    return zf.read(part).count(token)


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
    desc = "BOYFRIEND RAGLAN CREW CHUNKY COTTON - X/L"
    before = [g for g in read_layout(im_write_path).groups if g.description == desc and g.rows[0] == 229]
    assert before and len(before[0].rows) == 1
    after = read_layout(target)
    grp = next(g for g in after.groups if g.description == desc and g.subtotal_row is not None
               and len(g.rows) > 1)
    first, last = grp.rows[0], grp.rows[-1]
    formulas = after.subtotal_formulas[grp.subtotal_row].values()
    assert any(re.search(rf"[A-Z]{first}:[A-Z]+\$?{last}\b", f) for f in formulas), (first, last)

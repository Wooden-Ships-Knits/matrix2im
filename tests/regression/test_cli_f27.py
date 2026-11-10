import hashlib
import json
import time

import pytest

from matrix_to_im.cli import main

pytestmark = pytest.mark.regression


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load(path):
    return json.loads(path.read_text(encoding="utf-8"))


def test_run_f27_matches_golden(tmp_path, matrix_path, im_path, golden_raw, golden_final, capsys):
    before = (sha(matrix_path), sha(im_path))
    start = time.perf_counter()
    code = main(["run", str(matrix_path), "--im-master", str(im_path), "-o", str(tmp_path)])
    elapsed = time.perf_counter() - start
    assert code == 0
    assert elapsed < 15
    raw, final = load(tmp_path / "raw.json"), load(tmp_path / "final.json")
    assert raw["styles"] == golden_raw["styles"]
    assert final["records"] == golden_final["records"]
    # golden dari prototipe belum punya field tambahan `code` dan `whs_values` (D12); selebihnya harus sama.
    without_code = [{k: v for k, v in item.items() if k not in ("code", "whs_values")} for item in final["needs_review"]]
    assert without_code == golden_final["needs_review"]
    out = capsys.readouterr().out
    assert "needs_review" in out and str(raw["total_styles"]) in out
    assert (sha(matrix_path), sha(im_path)) == before


def test_run_f27_strict_returns_2(tmp_path, matrix_path, im_path):
    assert main(["run", str(matrix_path), "--im-master", str(im_path), "-o", str(tmp_path), "--strict"]) == 2
    assert (tmp_path / "final.json").exists()

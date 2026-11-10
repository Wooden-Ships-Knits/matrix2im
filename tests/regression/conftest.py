"""Fixture session untuk regression F27: ekstraksi nyata dijalankan sekali per sesi."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from matrix_to_im.extract import ExtractResult, extract_workbook, to_raw_json

SKIP_MESSAGE = "File data F27 tidak ditemukan (data/F27_MATRIX_MASTER.xlsx); regression test di-skip"
GOLDEN_DIR = Path(__file__).resolve().parents[2] / "reference" / "golden"


@pytest.fixture(scope="session")
def matrix_path(data_finder) -> Path:
    path = data_finder("matrix")
    if path is None:
        pytest.skip(SKIP_MESSAGE)
    return path


@pytest.fixture(scope="session")
def im_path(data_finder) -> Path:
    path = data_finder("im")
    if path is None:
        pytest.skip(SKIP_MESSAGE)
    return path


@pytest.fixture(scope="session")
def f27_result(matrix_path: Path) -> ExtractResult:
    return extract_workbook(matrix_path)


@pytest.fixture(scope="session")
def f27_raw(f27_result: ExtractResult) -> dict:
    return to_raw_json(f27_result)


@pytest.fixture(scope="session")
def golden_raw() -> dict:
    return json.loads((GOLDEN_DIR / "F27_MATRIX_MASTER_extract.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="session")
def golden_final() -> dict:
    return json.loads((GOLDEN_DIR / "F27_IM_MASTER_final.json").read_text(encoding="utf-8"))

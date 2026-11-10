"""Fixture bersama; file data nyata hanya dipakai test regression."""
from __future__ import annotations

import os
from glob import glob
from pathlib import Path
from typing import Literal

ROOT = Path(__file__).resolve().parent.parent

_KINDS = {
    "matrix": ("MATRIX2IM_MATRIX", "F27_MATRIX_MASTER.xlsx", "*F27*MATRIX*MASTER*.xlsx"),
    "im": ("MATRIX2IM_IM", "F27_IM_MASTER.xlsx", "*F27*IM*MASTER*.xlsx"),
}


def find_data_file(kind: Literal["matrix", "im"]) -> Path | None:
    """Cari file data: env var (tanpa fallback), nama spec, lalu glob (nama asli 'Copy of ...')."""
    env, spec_name, pattern = _KINDS[kind]
    override = os.environ.get(env)
    if override:
        path = Path(override)
        return path if path.exists() else None
    exact = ROOT / "data" / spec_name
    if exact.exists():
        return exact
    matches = sorted(glob(str(ROOT / "data" / pattern)))
    return Path(matches[0]) if matches else None


import pytest  # noqa: E402


@pytest.fixture(scope="session")
def data_finder():
    return find_data_file

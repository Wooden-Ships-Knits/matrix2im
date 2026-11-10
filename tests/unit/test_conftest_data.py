def test_env_override_missing_returns_none(data_finder, monkeypatch, tmp_path):
    monkeypatch.setenv("MATRIX2IM_MATRIX", str(tmp_path / "tidak_ada.xlsx"))
    assert data_finder("matrix") is None


def test_env_override_existing(data_finder, monkeypatch, tmp_path):
    f = tmp_path / "x.xlsx"
    f.touch()
    monkeypatch.setenv("MATRIX2IM_IM", str(f))
    assert data_finder("im") == f


def test_glob_finds_distinct_files(data_finder, monkeypatch):
    monkeypatch.delenv("MATRIX2IM_MATRIX", raising=False)
    monkeypatch.delenv("MATRIX2IM_IM", raising=False)
    m, i = data_finder("matrix"), data_finder("im")
    if m is None or i is None:
        return  # data/ tidak ada di mesin ini
    assert m != i and "MATRIX" in m.name and "IM MASTER" in i.name

import json
import threading
import urllib.error
import urllib.parse
import urllib.request

import pytest
from helpers.matrix_factory import build_im, build_matrix

from matrix_to_im.cli import main
from matrix_to_im.ui_server import UiState, make_server


@pytest.fixture
def state(tmp_path):
    st = UiState(tmp_path / "work")
    yield st
    st.cleanup()


@pytest.fixture
def base(state):
    srv = make_server("127.0.0.1", 0, state)
    thread = threading.Thread(target=srv.serve_forever, kwargs={"poll_interval": 0.05}, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{srv.server_address[1]}"
    srv.shutdown()
    srv.server_close()


def call(url, method="GET", data=None, headers=None):
    req = urllib.request.Request(url, data=data, method=method, headers=headers or {})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status, resp.headers, resp.read()
    except urllib.error.HTTPError as err:
        return err.code, err.headers, err.read()


def upload(base, kind, path, name=None):
    name = urllib.parse.quote(name or path.name)
    return call(f"{base}/api/upload/{kind}", "POST", path.read_bytes(), {"X-Filename": name})


def run(base):
    status, _, body = call(f"{base}/api/run", "POST", b"")
    return status, json.loads(body)


def test_index_page_is_html(base):
    status, headers, body = call(f"{base}/")
    assert status == 200
    assert headers["Content-Type"].startswith("text/html")
    assert b"matrix2im" in body


def test_run_without_matrix_is_400(base):
    status, payload = run(base)
    assert status == 400
    assert "MATRIX" in payload["error"]


def test_upload_reports_name_and_size(base, tmp_path):
    matrix = build_matrix(tmp_path / "Copy of F27 MATRIX.xlsx")
    status, _, body = upload(base, "matrix", matrix)
    assert status == 200
    assert json.loads(body) == {"name": "Copy of F27 MATRIX.xlsx", "size": matrix.stat().st_size}


def test_upload_and_run_without_im(base, tmp_path):
    upload(base, "matrix", build_matrix(tmp_path / "m.xlsx"))
    status, payload = run(base)
    assert status == 200
    assert list(payload) == ["files", "seconds", "raw", "final", "validation_errors", "im_checked"]
    assert payload["files"] == {"matrix": "m.xlsx", "im": None}
    assert payload["raw"]["total_styles"] == 4
    colorways = sum(len(s[g]) for s in payload["raw"]["styles"]
                    for g in ("color_above_black_line", "color_between_black_and_red_line", "color_below_red_line"))
    assert payload["final"]["summary"]["total_records"] == colorways
    assert payload["validation_errors"] == [] and payload["im_checked"] is False


def test_run_with_im_checks_header(base, tmp_path):
    upload(base, "matrix", build_matrix(tmp_path / "m.xlsx"))
    upload(base, "im", build_im(tmp_path / "im.xlsx"))
    status, payload = run(base)
    assert status == 200
    assert payload["files"]["im"] == "im.xlsx"
    assert payload["im_checked"] is True and payload["validation_errors"] == []


def test_im_missing_key_is_reported_not_fatal(base, tmp_path):
    upload(base, "matrix", build_matrix(tmp_path / "m.xlsx"))
    upload(base, "im", build_im(tmp_path / "im.xlsx", drop="COLOR2"))
    status, payload = run(base)
    assert status == 200
    assert any("COLOR2" in e for e in payload["validation_errors"])


def test_delete_im_upload(base, tmp_path):
    upload(base, "matrix", build_matrix(tmp_path / "m.xlsx"))
    upload(base, "im", build_im(tmp_path / "im.xlsx"))
    status, _, _ = call(f"{base}/api/upload/im", "DELETE")
    assert status == 200
    _, payload = run(base)
    assert payload["im_checked"] is False and payload["files"]["im"] is None


def test_bad_matrix_is_422_with_indonesian_message(base, tmp_path):
    bad = tmp_path / "rusak.xlsx"
    bad.write_bytes(b"bukan file excel")
    upload(base, "matrix", bad)
    status, payload = run(base)
    assert status == 422
    assert "tidak bisa dibaca" in payload["error"]


def test_download_before_run_is_404(base):
    status, _, _ = call(f"{base}/api/download/final.json")
    assert status == 404


def test_download_after_run(base, tmp_path):
    upload(base, "matrix", build_matrix(tmp_path / "m.xlsx"))
    run(base)
    for name, key in (("raw.json", "styles"), ("final.json", "records")):
        status, headers, body = call(f"{base}/api/download/{name}")
        assert status == 200
        assert headers["Content-Disposition"] == f'attachment; filename="{name}"'
        assert key in json.loads(body)


@pytest.mark.parametrize("method,path", [("GET", "/tidak-ada"), ("POST", "/api/upload/lain"),
                                         ("GET", "/api/download/lain.json"), ("POST", "/api/lain")])
def test_unknown_routes_are_404(base, method, path):
    status, _, _ = call(f"{base}{path}", method, b"" if method == "POST" else None)
    assert status == 404


def test_cleanup_removes_workdir(tmp_path):
    st = UiState(tmp_path / "work")
    (st.workdir / "x").write_text("x")
    st.cleanup()
    assert not st.workdir.exists()


def test_cli_ui_subcommand_calls_serve(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr("matrix_to_im.cli.serve", lambda **kw: calls.append(kw))
    monkeypatch.chdir(tmp_path)
    assert main(["ui", "--port", "9999", "--no-browser"]) == 0
    assert len(calls) == 1
    assert calls[0]["port"] == 9999 and calls[0]["open_browser"] is False
    assert calls[0]["host"] == "127.0.0.1"


def test_sigterm_handler_stops_like_ctrl_c():
    from matrix_to_im.ui_server import _stop_on_sigterm

    with pytest.raises(KeyboardInterrupt):
        _stop_on_sigterm(15, None)


def _fake_server_factory(record, fail_first=True):
    class _Srv:
        server_address = ("127.0.0.1", 5555)

        def serve_forever(self):
            raise KeyboardInterrupt

        def server_close(self):
            pass

    def fake(host, port, state):
        record.append((host, port))
        if fail_first and len(record) == 1:
            raise OSError("in use")
        return _Srv()

    return fake


def test_ui_non_loopback_port_in_use_raises(monkeypatch):
    from matrix_to_im.errors import Matrix2ImError
    from matrix_to_im.ui_server import serve

    record = []
    monkeypatch.setattr("matrix_to_im.ui_server.make_server", _fake_server_factory(record))
    with pytest.raises(Matrix2ImError, match="Port 8765 sudah dipakai di 192.0.2.1"):
        serve(host="192.0.2.1", port=8765, open_browser=False)
    assert record == [("192.0.2.1", 8765)]


def test_ui_loopback_port_in_use_falls_back(monkeypatch):
    from matrix_to_im.ui_server import serve

    record = []
    monkeypatch.setattr("matrix_to_im.ui_server.make_server", _fake_server_factory(record))
    serve(host="localhost", port=8765, open_browser=False)
    assert record == [("localhost", 8765), ("localhost", 0)]


def test_ui_wildcard_host_opens_loopback_url(monkeypatch, caplog):
    from matrix_to_im.settings import Settings
    from matrix_to_im.ui_server import serve

    record, opened = [], []
    monkeypatch.setattr("matrix_to_im.ui_server.make_server", _fake_server_factory(record, fail_first=False))
    monkeypatch.setattr("matrix_to_im.ui_server.webbrowser.open", opened.append)
    caplog.set_level("INFO")
    serve(host="0.0.0.0", port=8765, open_browser=True, settings=Settings())
    assert record == [("0.0.0.0", 8765)]
    assert opened == ["http://127.0.0.1:5555/"]
    assert "PC lain" in caplog.text and ":5555/" in caplog.text

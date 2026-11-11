import json
import logging
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import pytest
from helpers.matrix_factory import build_im, build_matrix

from matrix_to_im.cli import main
from matrix_to_im.settings import Settings
from matrix_to_im.ui_server import SessionStore, UiState, WriteCoordinator, make_server
from tests.helpers.im_factory import TEST_PASSWORD, GroupSpec, add_file_sharing, build_im_sheet


@pytest.fixture
def state(tmp_path):
    st = UiState(tmp_path / "work")
    yield st
    st.cleanup()


@pytest.fixture
def sessions(tmp_path):
    store = SessionStore(tmp_path / "sessions")
    yield store
    store.cleanup_all()


@pytest.fixture
def serve_with(sessions):
    """Pabrik server port 0 (loopback) dengan Settings tertentu; dimatikan saat test selesai."""
    started = []

    def start(settings=None):
        srv = make_server("127.0.0.1", 0, sessions, settings or Settings())
        thread = threading.Thread(target=srv.serve_forever, kwargs={"poll_interval": 0.05}, daemon=True)
        thread.start()
        started.append((srv, thread))
        return f"http://127.0.0.1:{srv.server_address[1]}"

    yield start
    for srv, thread in started:
        srv.shutdown()
        thread.join(timeout=10)
        srv.server_close()


@pytest.fixture
def base(serve_with):
    return serve_with()


def call(url, method="GET", data=None, headers=None):
    req = urllib.request.Request(url, data=data, method=method, headers=headers or {})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status, resp.headers, resp.read()
    except urllib.error.HTTPError as err:
        return err.code, err.headers, err.read()


@pytest.fixture(autouse=True)
def _fresh_cookie_jar():
    """Jar per test: nomor port 0 bisa dipakai ulang server lain dan membawa cookie basi."""
    _COOKIES.clear()
    yield
    _COOKIES.clear()


_COOKIES: dict[str, str] = {}


def own(base):
    """Header Cookie sesi bawaan test lama (satu browser per server)."""
    if base not in _COOKIES:
        _, headers, _ = call(f"{base}/api/write/config")
        _COOKIES[base] = headers["Set-Cookie"].split(";", 1)[0]
    return {"Cookie": _COOKIES[base]}


def upload(base, kind, path, name=None):
    name = urllib.parse.quote(name or path.name)
    return call(f"{base}/api/upload/{kind}", "POST", path.read_bytes(), {"X-Filename": name, **own(base)})


def run(base):
    status, _, body = call(f"{base}/api/run", "POST", b"", own(base))
    return status, json.loads(body)


def test_index_page_is_html(base):
    status, headers, body = call(f"{base}/")
    assert status == 200
    assert headers["Content-Type"].startswith("text/html")
    assert b"matrix2im" in body


def test_index_has_write_panel(base):
    """Panel tulis ada di halaman, memakai endpoint tulis, dan tidak memuat apa pun dari jaringan."""
    _, _, body = call(f"{base}/")
    html = body.decode("utf-8")
    for needle in ("Tulis semua ke IM", 'id="write-password"', 'type="password"', 'autocomplete="off"',
                   "/api/write/config", "/api/write/prepare", "/api/write/commit", "/api/write/status",
                   "Saya sudah memastikan tidak ada yang sedang membuka IM MASTER", "Buat rencana",
                   "Sedang ada penulisan oleh"):
        assert needle in html, needle
    assert "http://" not in html and "https://" not in html
    assert "localStorage" not in html and "sessionStorage" not in html


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
    status, _, _ = call(f"{base}/api/upload/im", "DELETE", headers=own(base))
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
        status, headers, body = call(f"{base}/api/download/{name}", headers=own(base))
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
        writer = WriteCoordinator()

        def serve_forever(self):
            raise KeyboardInterrupt

        def server_close(self):
            pass

    def fake(host, port, sessions, settings):
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


# ---------- sesi per browser ----------

def _cookie_of(headers):
    raw = headers["Set-Cookie"]
    return raw.split(";", 1)[0]


def _new_session(base):
    status, headers, _ = call(f"{base}/api/write/config")
    assert status == 200
    return _cookie_of(headers)


def _state_of(sessions, cookie):
    return sessions.get(cookie)[1]


def test_cookie_set_httponly_samesite(base):
    _, headers, _ = call(f"{base}/api/write/config")
    raw = headers["Set-Cookie"]
    assert raw.startswith("m2i_session=") and len(raw.split(";", 1)[0]) > len("m2i_session=") + 15
    assert "HttpOnly" in raw and "SameSite=Strict" in raw and "Path=/" in raw
    cookie = _cookie_of(headers)
    _, headers2, _ = call(f"{base}/api/write/config", headers={"Cookie": cookie})
    assert headers2["Set-Cookie"] is None   # sesi yang sama tidak diberi cookie lagi


def test_sessions_isolated_uploads(base, tmp_path):
    a = _new_session(base)
    b = _new_session(base)
    assert a != b
    matrix = build_matrix(tmp_path / "m.xlsx")
    call(f"{base}/api/upload/matrix", "POST", matrix.read_bytes(), {"X-Filename": "m.xlsx", "Cookie": a})
    status, _, _ = call(f"{base}/api/run", "POST", b"", {"Cookie": a})
    assert status == 200
    status, _, body = call(f"{base}/api/run", "POST", b"", {"Cookie": b})
    assert status == 400 and "MATRIX" in json.loads(body)["error"]
    status, _, _ = call(f"{base}/api/download/final.json", headers={"Cookie": b})
    assert status == 404
    status, _, _ = call(f"{base}/api/download/final.json", headers={"Cookie": a})
    assert status == 200


def test_session_store_get_new_and_existing(sessions):
    sid, st, is_new = sessions.get(None)
    assert is_new and st.workdir == sessions.root / sid and st.workdir.is_dir()
    sid2, st2, is_new2 = sessions.get(f"x=1; m2i_session={sid}")
    assert (sid2, st2, is_new2) == (sid, st, False)
    _, st3, new3 = sessions.get("m2i_session=tidak-ada")
    assert new3 and st3 is not st


def test_session_cleanup_expired_keeps_running_write(tmp_path):
    now = [0.0]
    store = SessionStore(tmp_path / "s", ttl_s=100, clock=lambda: now[0])
    sid_old, st_old, _ = store.get(None)
    sid_busy, st_busy, _ = store.get(None)
    st_busy.write_running = True
    now[0] = 50
    sid_fresh, st_fresh, _ = store.get(None)
    now[0] = 120
    store.cleanup_expired()
    assert not st_old.workdir.exists()
    assert st_busy.workdir.exists() and st_fresh.workdir.exists()
    assert store.get(f"m2i_session={sid_old}")[2] is True   # sesi lama hilang -> sesi baru
    st_busy.write_running = False
    now[0] = 300
    store.cleanup_expired()
    assert not st_busy.workdir.exists()
    store.cleanup_all()
    assert not st_fresh.workdir.exists()


def test_expired_sessions_cleaned_on_request(tmp_path):
    now = [0.0]
    store = SessionStore(tmp_path / "s", ttl_s=10, clock=lambda: now[0])
    srv = make_server("127.0.0.1", 0, store, Settings())
    thread = threading.Thread(target=srv.serve_forever, kwargs={"poll_interval": 0.05}, daemon=True)
    thread.start()
    try:
        base = f"http://127.0.0.1:{srv.server_address[1]}"
        cookie = _new_session(base)
        old = _state_of(store, cookie).workdir
        now[0] = 60
        _new_session(base)
        assert not old.exists()
    finally:
        srv.shutdown()
        thread.join(timeout=10)
        srv.server_close()
        store.cleanup_all()


def test_run_uses_setting_im_master_when_no_im_upload(serve_with, tmp_path):
    im = build_im(tmp_path / "im-setting.xlsx")
    base = serve_with(Settings(im_master=im, backup_dir=tmp_path / "bak"))
    cookie = _new_session(base)
    call(f"{base}/api/upload/matrix", "POST", build_matrix(tmp_path / "m.xlsx").read_bytes(),
         {"X-Filename": "m.xlsx", "Cookie": cookie})
    status, _, body = call(f"{base}/api/run", "POST", b"", {"Cookie": cookie})
    payload = json.loads(body)
    assert status == 200 and payload["im_checked"] is True and payload["files"]["im"] == "im-setting.xlsx"


# ---------- endpoint tulis ----------

PW = TEST_PASSWORD


def _write_env(tmp_path, engine="none"):
    groups = [GroupSpec("AAA - K", [["a1"], ["a2"]])]
    im = build_im_sheet(tmp_path / "im-tulis.xlsx", groups)
    add_file_sharing(im)
    settings = Settings(im_master=im, backup_dir=tmp_path / "cadangan", insert_engine=engine)

    def rec(color, ws, whs, n):
        return {"DESCRIPTION": "AAA", "COLOR1": color, "WS TAG COLOR": ws, "WHS": whs,
                "_ref": f"SHEET No.{n} | B1 | x"}

    final = {"records": [rec("x1", "T1", "000", 1), rec("x2", "T2", "SY", 2)]}
    return im, settings, final


def _inject(sessions, cookie, final, matrix_name="matrix.xlsx"):
    st = _state_of(sessions, cookie)
    st.last = {"final": final, "raw": {}, "matrix_name": matrix_name}
    st.uploads["matrix"] = (st.workdir / "matrix.xlsx", matrix_name)
    return st


def post_json(base, path, payload, cookie=None):
    headers = {"Content-Type": "application/json"}
    if cookie:
        headers["Cookie"] = cookie
    status, _, body = call(f"{base}{path}", "POST", json.dumps(payload).encode(), headers)
    return status, json.loads(body)


def get_json(base, path, cookie=None):
    status, _, body = call(f"{base}{path}", headers={"Cookie": cookie} if cookie else {})
    return status, json.loads(body)


def wait_state(base, cookie, want, timeout=20):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        _, st = get_json(base, "/api/write/status", cookie)
        if st["state"] in want:
            return st
        time.sleep(0.05)
    raise AssertionError(f"status tidak mencapai {want}: {st}")


def test_write_config_disabled_without_settings(base):
    status, payload = get_json(base, "/api/write/config")
    assert status == 200
    assert payload["enabled"] is False and payload["reason"]
    assert set(payload) == {"enabled", "reason", "im_name", "engine"}


def test_write_config_enabled(serve_with, tmp_path):
    im, settings, _ = _write_env(tmp_path)
    _, payload = get_json(serve_with(settings), "/api/write/config")
    assert payload["enabled"] is True and payload["im_name"] == "im-tulis.xlsx" and payload["engine"] == "none"


def test_prepare_without_result_400(serve_with, tmp_path):
    im, settings, _ = _write_env(tmp_path)
    base = serve_with(settings)
    status, payload = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, _new_session(base))
    assert status == 400 and "Jalankan proses dulu" in payload["error"]


def test_prepare_requires_name_and_password(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    status, payload = post_json(base, "/api/write/prepare", {"name": "  ", "password": PW}, cookie)
    assert status == 400 and payload["error"] == "Isi nama Anda dulu"
    status, _ = post_json(base, "/api/write/prepare", {"name": "andi"}, cookie)
    assert status == 400


def test_prepare_bad_json_and_oversize(serve_with, tmp_path):
    base = serve_with(Settings())
    cookie = _new_session(base)
    status, _, _ = call(f"{base}/api/write/prepare", "POST", b"{bukan json", {"Cookie": cookie})
    assert status == 400
    status, _, _ = call(f"{base}/api/write/prepare", "POST", b"[1]", {"Cookie": cookie})
    assert status == 400
    big = json.dumps({"name": "a", "password": "x" * (65 * 1024)}).encode()
    status, _, _ = call(f"{base}/api/write/prepare", "POST", big, {"Cookie": cookie})
    assert status == 413


def test_prepare_wrong_password_401(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    status, payload = post_json(base, "/api/write/prepare", {"name": "andi", "password": "pw-Keliru-31"}, cookie)
    assert status == 401 and "pw-Keliru-31" not in json.dumps(payload)


def test_prepare_owner_file_is_409(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    im.with_name("~$" + im.name).write_bytes(bytes([5]) + "Budi!".encode())
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    status, payload = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    assert status == 409 and "Budi!" in payload["error"]


def test_prepare_other_error_is_422(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    im.unlink()
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    status, _ = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    assert status == 422


def test_prepare_commit_status_end_to_end_engine_none(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    before = im.read_bytes()
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    status, plan = post_json(base, "/api/write/prepare", {"name": " andi ", "password": PW}, cookie)
    assert status == 200, plan
    assert {"summary", "changes", "changes_total", "inserts", "skipped", "short_groups",
            "unmatched_im_styles", "warnings", "engine", "token"} <= set(plan)
    assert plan["changes_total"] == plan["summary"]["cells_changed"] == len(plan["changes"]) > 0
    assert plan["engine"] == "none" and plan["inserts"] == []
    assert im.read_bytes() == before   # prepare tidak mengubah apa pun
    status, _ = post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, cookie)
    assert status == 202
    st = wait_state(base, cookie, ("done", "error"))
    assert st["state"] == "done", st
    assert st["user"] == "andi" and st["started"]
    assert {"backup", "seconds", "summary", "inserted_rows", "message"} <= set(st["report"])
    assert isinstance(st["report"]["backup"], str) and Path(st["report"]["backup"]).is_file()
    assert im.read_bytes() != before
    assert PW not in json.dumps(st) and plan["token"] not in json.dumps(st)
    # token sekali pakai
    status, payload = post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, cookie)
    assert status == 409 and payload["error"] == "rencana sudah dipakai; buat rencana lagi"
    assert _state_of(sessions, cookie).write_running is False


def test_prepare_caps_changes_at_500(serve_with, sessions, tmp_path, monkeypatch):
    from matrix_to_im import ui_server
    from matrix_to_im.write_plan import CellChange

    im, settings, final = _write_env(tmp_path)
    real = ui_server.prepare_write

    def many(final, settings, password):
        p = real(final, settings, password)
        p.values.changes = [CellChange(i, 5, None, "x") for i in range(1, 701)]
        return p

    monkeypatch.setattr(ui_server, "prepare_write", many)
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    assert len(plan["changes"]) == 500 and plan["changes_total"] == 700


def test_commit_needs_matching_token_of_own_session(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    base = serve_with(settings)
    a, b = _new_session(base), _new_session(base)
    _inject(sessions, a, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, a)
    status, _ = post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, b)
    assert status == 409   # sesi B tidak punya rencana
    status, _ = post_json(base, "/api/write/commit", {"token": "salah", "password": PW}, a)
    assert status == 403
    status, _ = post_json(base, "/api/write/commit", {"password": PW}, a)
    assert status == 400
    # token yang benar masih berlaku setelah percobaan gagal
    status, _ = post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, a)
    assert status == 202
    assert wait_state(base, a, ("done", "error"))["state"] == "done"


def test_second_prepare_replaces_token(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    _, p1 = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    _, p2 = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    assert p1["token"] != p2["token"]
    status, _ = post_json(base, "/api/write/commit", {"token": p1["token"], "password": PW}, cookie)
    assert status == 403


def test_second_commit_while_running_409(serve_with, sessions, tmp_path, monkeypatch):
    from matrix_to_im import ui_server
    from matrix_to_im.im_writer import WriteReport

    im, settings, final = _write_env(tmp_path)
    gate, entered = threading.Event(), threading.Event()
    calls = []

    def slow_commit(preview, final, settings, password, user, matrix_name, progress=lambda s: None, **kw):
        calls.append((user, matrix_name))
        progress("tulis nilai")
        entered.set()
        assert gate.wait(timeout=20)
        return WriteReport(True, "ok", preview.summary, None, tmp_path / "log.jsonl", 0.5, [])

    monkeypatch.setattr(ui_server, "commit_write", slow_commit)
    base = serve_with(settings)
    a, b = _new_session(base), _new_session(base)
    _inject(sessions, a, final, "m-a.xlsx")
    _inject(sessions, b, final, "m-b.xlsx")
    _, plan_a = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, a)
    _, plan_b = post_json(base, "/api/write/prepare", {"name": "budi", "password": PW}, b)
    try:
        assert post_json(base, "/api/write/commit", {"token": plan_a["token"], "password": PW}, a)[0] == 202
        assert entered.wait(timeout=20)
        status, payload = post_json(base, "/api/write/commit", {"token": plan_b["token"], "password": PW}, b)
        assert status == 409 and payload["error"] == "sedang ada penulisan oleh andi"
        status, payload = post_json(base, "/api/write/prepare", {"name": "budi", "password": PW}, b)
        assert status == 409 and payload["error"] == "sedang ada penulisan oleh andi"
        _, st = get_json(base, "/api/write/status", b)   # sesi lain ikut melihat status
        assert st["state"] == "running" and st["user"] == "andi" and st["phase"] == "tulis nilai"
        assert _state_of(sessions, a).write_running is True
    finally:
        gate.set()
    st = wait_state(base, b, ("done", "error"))
    assert st["state"] == "done" and st["report"]["backup"] is None
    assert calls == [("andi", "m-a.xlsx")]
    # token B masih utuh dan bisa dipakai setelah job pertama selesai
    assert post_json(base, "/api/write/commit", {"token": plan_b["token"], "password": PW}, b)[0] == 202
    wait_state(base, b, ("done", "error"))


def test_commit_failure_report_is_error_state(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    post_json(base, "/api/write/commit", {"token": plan["token"], "password": "salah-pw"}, cookie)
    st = wait_state(base, cookie, ("done", "error"))
    assert st["state"] == "error" and st["error"] and "report" not in st
    assert "salah-pw" not in json.dumps(st)


def test_idle_status(base):
    _, st = get_json(base, "/api/write/status")
    assert st == {"state": "idle", "phase": None, "user": None, "started": None}


def test_password_not_logged(serve_with, sessions, tmp_path, caplog):
    secret = "Rahasia-Unik-98765"
    im, settings, final = _write_env(tmp_path)
    caplog.set_level(logging.DEBUG)
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    post_json(base, "/api/write/prepare", {"name": "andi", "password": secret}, cookie)   # salah
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    post_json(base, "/api/write/commit", {"token": plan["token"], "password": secret}, cookie)  # salah
    wait_state(base, cookie, ("done", "error"))
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, cookie)
    wait_state(base, cookie, ("done", "error"))
    assert "/api/write/commit" in caplog.text   # ada log request, jadi assert di bawah tidak kosong
    assert secret not in caplog.text and PW not in caplog.text and plan["token"] not in caplog.text


def test_make_server_binds_given_host(sessions):
    srv = make_server("127.0.0.1", 0, sessions, Settings())
    try:
        assert srv.server_address[0] == "127.0.0.1"
    finally:
        srv.server_close()


def test_write_coordinator_runs_one_job():
    from matrix_to_im.im_writer import WriteReport

    coord = WriteCoordinator()
    gate = threading.Event()

    def job(progress):
        progress("fase")
        assert gate.wait(timeout=20)
        return WriteReport(True, "ok", {"x": 1}, None, Path("log"), 1.234, [3])

    assert coord.start("andi", job) is True
    assert coord.start("budi", job) is False
    assert coord.status()["state"] == "running" and coord.status()["user"] == "andi"
    gate.set()
    deadline = time.monotonic() + 20
    while coord.status()["state"] == "running" and time.monotonic() < deadline:
        time.sleep(0.02)
    st = coord.status()
    assert st["state"] == "done" and st["report"]["inserted_rows"] == [3] and st["report"]["seconds"] == 1.23

    def boom(progress):
        raise RuntimeError("meledak")

    assert coord.start("citra", boom) is True
    deadline = time.monotonic() + 20
    while coord.status()["state"] == "running" and time.monotonic() < deadline:
        time.sleep(0.02)
    assert coord.status()["state"] == "error" and "RuntimeError" in coord.status()["error"]


def test_lan_address_never_raises(monkeypatch):
    import socket

    from matrix_to_im.ui_server import _lan_address

    assert _lan_address() is None or isinstance(_lan_address(), str)

    class Boom:
        def __init__(self, *a, **k):
            raise OSError("tidak ada jaringan")

    monkeypatch.setattr(socket, "socket", Boom)
    assert _lan_address() is None


def test_serve_logs_lan_address_for_wildcard(monkeypatch, caplog):
    from matrix_to_im.ui_server import serve

    record = []
    monkeypatch.setattr("matrix_to_im.ui_server.make_server", _fake_server_factory(record, fail_first=False))
    monkeypatch.setattr("matrix_to_im.ui_server._lan_address", lambda: "192.168.7.9")
    caplog.set_level("INFO")
    serve(host="0.0.0.0", port=8765, open_browser=False)
    assert "http://192.168.7.9:5555/" in caplog.text


def test_writer_released_even_if_session_lock_held(serve_with, sessions, tmp_path, monkeypatch):
    """Upload/run memegang state.lock lama; selesainya penulisan tak boleh menunggunya."""
    from matrix_to_im import ui_server
    from matrix_to_im.im_writer import WriteReport

    im, settings, final = _write_env(tmp_path)
    gate = threading.Event()

    def commit(preview, final, settings, password, user, matrix_name, progress=lambda s: None, **kw):
        assert gate.wait(timeout=20)
        return WriteReport(True, "ok", preview.summary, None, tmp_path / "log.jsonl", 0.1, [])

    monkeypatch.setattr(ui_server, "commit_write", commit)
    base = serve_with(settings)
    cookie = _new_session(base)
    st = _inject(sessions, cookie, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    assert post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, cookie)[0] == 202
    assert st.lock.acquire(timeout=10)   # simulasi upload yang sedang berjalan
    try:
        gate.set()
        assert wait_state(base, cookie, ("done", "error"), timeout=10)["state"] == "done"
        deadline = time.monotonic() + 10
        while st.write_running and time.monotonic() < deadline:
            time.sleep(0.02)
        assert st.write_running is False
    finally:
        st.lock.release()


@pytest.mark.parametrize("header", ['prefs={"x":1}; m2i_session=XYZ', "foo bar; m2i_session=XYZ",
                                    "x=a b; m2i_session=XYZ", "m2i_session=XYZ"])
def test_session_cookie_survives_malformed_other_cookies(sessions, header):
    sid, st, _ = sessions.get(None)
    sessions._sessions["XYZ"] = sessions._sessions.pop(sid)
    got, st2, is_new = sessions.get(header)
    assert (got, st2, is_new) == ("XYZ", st, False)


def test_matrix_name_comes_from_run_and_new_upload_drops_plan(serve_with, tmp_path, monkeypatch):
    from matrix_to_im import ui_server
    from matrix_to_im.im_writer import WriteReport

    im, settings, _ = _write_env(tmp_path)
    seen = []

    def commit(preview, final, settings, password, user, matrix_name, progress=lambda s: None, **kw):
        seen.append(matrix_name)
        return WriteReport(True, "ok", preview.summary, None, tmp_path / "log.jsonl", 0.1, [])

    monkeypatch.setattr(ui_server, "commit_write", commit)
    base = serve_with(settings)
    own_cookie = own(base)["Cookie"]
    m1, m2 = build_matrix(tmp_path / "m1.xlsx"), build_matrix(tmp_path / "m2.xlsx")
    upload(base, "matrix", m1)
    assert run(base)[0] == 200
    upload(base, "matrix", m2)   # belum di-run: hasil run masih milik M1
    # IM sintetis tulis tidak cocok dengan matrix sintetis, tapi prepare tetap jalan (nama gaya tak cocok saja)
    status, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, own_cookie)
    assert status == 200, plan
    # upload baru membuat rencana basi
    upload(base, "matrix", m2)
    status, payload = post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, own_cookie)
    assert status == 409 and payload["error"] == "rencana sudah dipakai; buat rencana lagi"
    # prepare lagi tanpa run ulang: nama tetap M1
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, own_cookie)
    assert post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, own_cookie)[0] == 202
    wait_state(base, own_cookie, ("done", "error"))
    assert seen == ["m1.xlsx"]
    # run baru juga membuang rencana
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, own_cookie)
    assert run(base)[0] == 200
    status, payload = post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, own_cookie)
    assert status == 409


def test_coordinator_leaves_running_when_report_building_fails():
    coord = WriteCoordinator()

    class Odd:
        ok = True
        backup = None

        @property
        def seconds(self):
            raise ValueError("rusak")

    assert coord.start("andi", lambda progress: Odd()) is True
    coord.wait(20)
    st = coord.status()
    assert st["state"] == "error" and "ValueError" in st["error"] and "rusak" not in st["error"]


def test_serve_waits_for_running_write_before_cleanup(monkeypatch, caplog):
    from matrix_to_im.im_writer import WriteReport
    from matrix_to_im.ui_server import serve

    record = []
    factory = _fake_server_factory(record, fail_first=False)
    writer = WriteCoordinator()
    gate = threading.Event()
    finished = []

    def job(progress):
        assert gate.wait(timeout=20)
        finished.append(True)
        return WriteReport(True, "ok", {}, None, Path("log"), 0.1, [])

    assert writer.start("andi", job)

    def fake(host, port, sessions, settings):
        srv = factory(host, port, sessions, settings)
        srv.writer = writer
        return srv

    monkeypatch.setattr("matrix_to_im.ui_server.make_server", fake)
    cleaned = []
    monkeypatch.setattr("matrix_to_im.ui_server.SessionStore.cleanup_all", lambda self: cleaned.append(finished[:]))
    threading.Timer(0.3, gate.set).start()
    caplog.set_level("INFO")
    serve(host="127.0.0.1", port=8765, open_browser=False)
    assert "Menunggu penulisan IM MASTER selesai sebelum berhenti" in caplog.text
    assert cleaned == [[True]]   # cleanup baru jalan setelah job selesai


def test_prepare_shows_inserts_and_skipped(serve_with, sessions, tmp_path, monkeypatch):
    from matrix_to_im import ui_server
    from matrix_to_im.write_plan import InsertOp, Skip

    im, settings, final = _write_env(tmp_path)
    real = ui_server.prepare_write

    def with_ops(final, settings, password):
        p = real(final, settings, password)
        p.inserts = [InsertOp("AAA - K", 57, 2, 56, 56, None, {})]
        p.skipped = [Skip("ZZZ", "tidak ada di IM")]
        return p

    monkeypatch.setattr(ui_server, "prepare_write", with_ops)
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    assert plan["inserts"] == [{"description": "AAA - K", "count": 2}]
    assert plan["skipped"] == [{"style": "ZZZ", "reason": "tidak ada di IM"}]


def test_prepare_changes_have_cell_and_description(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    ch = plan["changes"][0]
    assert ch["description"] == "AAA - K"
    assert ch["cell"] == f"{'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[ch['col'] - 1]}{ch['row']}"
    assert {"row", "col", "old", "new"} <= set(ch)


def test_failed_commit_status_carries_backup(serve_with, sessions, tmp_path, monkeypatch):
    from matrix_to_im import ui_server
    from matrix_to_im.im_writer import WriteReport

    im, settings, final = _write_env(tmp_path)
    bak = tmp_path / "cadangan" / "x.xlsx"

    def failing(preview, final, settings, password, user, matrix, progress=None):
        return WriteReport(False, "gagal simpan", {}, bak, tmp_path / "log", 0.1, [])

    monkeypatch.setattr(ui_server, "commit_write", failing)
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, cookie)
    st = wait_state(base, cookie, ("done", "error"))
    assert st["state"] == "error" and st["error"] == "gagal simpan" and st["backup"] == str(bak)


def test_writer_thread_is_not_daemon():
    """Thread tulis bukan daemon: proses tidak boleh mati di tengah menimpa IM MASTER."""
    from matrix_to_im.im_writer import WriteReport

    coord = WriteCoordinator()
    gate = threading.Event()

    def job(progress):
        assert gate.wait(timeout=20)
        return WriteReport(True, "ok", {}, None, Path("log"), 0.1, [])

    assert coord.start("andi", job)
    try:
        assert coord._thread is not None and coord._thread.daemon is False
    finally:
        gate.set()
        coord.wait(20)


def test_prepare_returns_missing_styles_and_panel_shows_them(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    final = {"records": [*final["records"], {"DESCRIPTION": "ZZZ", "COLOR1": "z", "WS TAG COLOR": None,
                                             "WHS": None, "_ref": "SHEET No.9 | B1 | x"}]}
    base = serve_with(settings)
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    assert plan["missing_styles"] == ["ZZZ"] and plan["summary"]["missing_styles"] == 1
    _, _, body = call(f"{base}/")
    html = body.decode("utf-8")
    assert html.count("Style MATRIX tidak ada di IM (tidak ditulis)") >= 2   # tile dan judul daftar
    assert "p.missing_styles" in html and "s.missing_styles" in html


# ---------- I4: CSV rencana per sesi ----------

def test_preview_csv_for_own_plan_only(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    base = serve_with(settings)
    a, b = _new_session(base), _new_session(base)
    status, headers, body = call(f"{base}/api/write/preview.csv", headers={"Cookie": a})
    assert status == 404 and "rencana" in json.loads(body)["error"]      # belum ada rencana
    _inject(sessions, a, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, a)
    status, headers, body = call(f"{base}/api/write/preview.csv", headers={"Cookie": a})
    assert status == 200
    assert headers["Content-Type"].startswith("text/csv") and "attachment" in headers["Content-Disposition"]
    assert body.startswith(b"\xef\xbb\xbf")
    lines = body.decode("utf-8-sig").splitlines()
    assert lines[0] == "DESCRIPTION,sel,baris,kolom,lama,baru,jenis"
    assert sum(line.endswith(",ubah") for line in lines) == plan["changes_total"]
    assert PW.encode() not in body and plan["token"].encode() not in body
    status, _, _ = call(f"{base}/api/write/preview.csv", headers={"Cookie": b})   # sesi lain
    assert status == 404


def test_panel_has_csv_link_and_run_report_path(serve_with, sessions, tmp_path):
    im, settings, final = _write_env(tmp_path)
    base = serve_with(settings)
    _, _, body = call(f"{base}/")
    html = body.decode("utf-8")
    assert "Unduh semua perubahan (CSV)" in html and "/api/write/preview.csv" in html
    assert "r.laporan" in html and "st.laporan" in html
    cookie = _new_session(base)
    _inject(sessions, cookie, final)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    post_json(base, "/api/write/commit", {"token": plan["token"], "password": PW}, cookie)
    st = wait_state(base, cookie, ("done", "error"))
    assert st["state"] == "done" and st["report"]["laporan"].endswith("laporan matrix2im.csv")
    assert Path(st["report"]["laporan"]).is_file()
    post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    _, plan = post_json(base, "/api/write/prepare", {"name": "andi", "password": PW}, cookie)
    post_json(base, "/api/write/commit", {"token": plan["token"], "password": "salah-pw"}, cookie)
    st = wait_state(base, cookie, ("error",))
    assert st["laporan"] and st["laporan"].endswith("laporan matrix2im.csv")

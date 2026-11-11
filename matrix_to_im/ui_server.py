"""UI web lokal untuk mencoba matrix2im: upload Excel di browser, lihat hasil kedua tahap.

Hanya memakai standard library (http.server) supaya tidak menambah dependensi. Secara
default server hanya mendengarkan 127.0.0.1: file MATRIX berisi data produk internal. PC
pusat boleh memilih host LAN lewat matrix2im.local.toml atau `ui --host`. File yang
diupload disimpan sebagai salinan di folder kerja sementara, jadi file asli pengguna tidak
pernah disentuh, dan folder itu dihapus saat server berhenti.
"""
from __future__ import annotations

import hmac
import json
import logging
import secrets
import shutil
import socket
import signal
import tempfile
import threading
import time
import urllib.parse
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from importlib import resources
from pathlib import Path
from typing import Callable

from openpyxl.utils import get_column_letter

from .errors import Matrix2ImError
from .legend_config import EXCEL_TIMEOUT_S
from .im_lock import WrongPassword
from .im_writer import ImMasterBusy, WritePreview, WriteReport, commit_write, prepare_write
from .settings import Settings
from .extract import extract
from .transform import transform
from .validate import read_im_header, validate
from .write_report import report_csv_bytes, report_rows

log = logging.getLogger("matrix2im.ui")

UPLOAD_KINDS = ("matrix", "im")
MAX_UPLOAD_BYTES = 600 * 1024 * 1024   # F27 = 125 MB; batas ini hanya mencegah upload tak wajar
_CHUNK = 1024 * 1024
MAX_JSON_BYTES = 64 * 1024
MAX_CHANGES_SHOWN = 500
COOKIE_NAME = "m2i_session"


class UiState:
    """State satu sesi UI: salinan file upload dan hasil run terakhir (untuk download)."""

    def __init__(self, workdir: Path) -> None:
        self.workdir = workdir
        self.workdir.mkdir(parents=True, exist_ok=True)
        self.uploads: dict[str, tuple[Path, str]] = {}   # kind -> (salinan, nama asli)
        self.last: dict | None = None
        self.lock = threading.Lock()   # satu run pada satu waktu; openpyxl memakan memori besar
        # Rencana tulis (preview, final, user, matrix, token). Password sengaja tidak pernah disimpan.
        self.write: dict | None = None
        self.write_running = False     # sesi dengan job tulis berjalan tidak boleh dibersihkan
        # Kunci kecil untuk write/write_running; tidak pernah dipegang saat I/O, beda dengan `lock`
        # (yang dipegang upload/run berjam-jam) supaya penulis global bisa dilepas kapan saja.
        self.wlock = threading.Lock()

    def drop_plan(self) -> None:
        """Upload atau run baru membuat rencana tulis lama basi; jangan sentuh bila job sedang jalan."""
        with self.wlock:
            if not self.write_running:
                self.write = None

    def cleanup(self) -> None:
        shutil.rmtree(self.workdir, ignore_errors=True)


class SessionStore:
    """Satu UiState per browser (cookie), masing-masing di subfolder sendiri; kedaluwarsa menurut TTL."""

    def __init__(self, root: Path, ttl_s: int = 12 * 3600, clock: Callable[[], float] = time.monotonic) -> None:
        self.root = root
        self.ttl_s = ttl_s
        self._clock = clock
        self._lock = threading.Lock()
        self._sessions: dict[str, tuple[UiState, float]] = {}   # id -> (state, terakhir dipakai)

    def get(self, cookie_header: str | None) -> tuple[str, UiState, bool]:
        """(id, state, baru?). Cookie tak dikenal atau kedaluwarsa menghasilkan sesi baru."""
        sid = None
        for part in (cookie_header or "").split(";"):   # manual: cookie lain yang rusak tak boleh menjatuhkan sesi
            name, sep, value = part.strip().partition("=")
            if sep and name == COOKIE_NAME:
                sid = value.strip()
                break
        with self._lock:
            entry = self._sessions.get(sid) if sid else None
            if entry is not None:
                self._sessions[sid] = (entry[0], self._clock())
                return sid, entry[0], False
            sid = secrets.token_urlsafe(16)
            state = UiState(self.root / sid)
            self._sessions[sid] = (state, self._clock())
            return sid, state, True

    def cleanup_expired(self) -> None:
        """Hapus sesi lewat TTL, kecuali yang sedang menulis IM MASTER atau memproses upload/run."""
        now = self._clock()
        with self._lock:
            for sid, (state, seen) in list(self._sessions.items()):
                if now - seen > self.ttl_s and not state.write_running and not state.lock.locked():
                    del self._sessions[sid]
                    state.cleanup()

    def cleanup_all(self) -> None:
        """Hapus semua sesi dan folder root; dipanggil saat server berhenti."""
        with self._lock:
            for state, _ in self._sessions.values():
                state.cleanup()
            self._sessions.clear()
        shutil.rmtree(self.root, ignore_errors=True)


class WriteCoordinator:
    """Satu job tulis IM MASTER untuk seluruh sesi: file itu bersama, jadi tidak boleh ada dua penulisan."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._status: dict = {"state": "idle", "phase": None, "user": None, "started": None}
        self._thread: threading.Thread | None = None

    def status(self) -> dict:
        with self._lock:
            return dict(self._status)

    def _set_phase(self, phase: str) -> None:
        with self._lock:
            self._status["phase"] = phase

    def _finish(self, **fields: object) -> None:
        with self._lock:
            self._status.update(fields)

    def start(self, user: str, job: Callable[[Callable[[str], None]], WriteReport]) -> bool:
        """Jalankan job di thread; False bila sudah ada job berjalan."""
        with self._lock:
            if self._status["state"] == "running":
                return False
            self._status = {"state": "running", "phase": "mulai", "user": user,
                            "started": time.strftime("%Y-%m-%dT%H:%M:%S")}
        # Bukan daemon: proses tidak boleh berakhir di tengah menimpa IM MASTER (serve juga menunggunya).
        self._thread = threading.Thread(target=self._run, args=(job,), daemon=False, name="tulis-im")
        self._thread.start()
        return True

    def wait(self, timeout: float) -> None:
        """Tunggu job yang sedang berjalan (dipakai saat server berhenti)."""
        thread = self._thread
        if thread is not None:
            thread.join(timeout)

    def _run(self, job: Callable[[Callable[[str], None]], WriteReport]) -> None:
        fields: dict = {"state": "error", "error": "Gagal tak terduga: penulisan berhenti tanpa laporan", "backup": None}
        try:   # state harus selalu keluar dari "running", apa pun yang terjadi
            try:
                rep = job(self._set_phase)
                if rep.ok:
                    fields = {"state": "done", "phase": "selesai", "report": {
                        "backup": str(rep.backup) if rep.backup else None, "seconds": round(rep.seconds, 2),
                        "summary": rep.summary, "inserted_rows": rep.inserted_rows, "message": rep.message,
                        "laporan": str(rep.laporan) if rep.laporan else None}}
                else:
                    fields = {"state": "error", "error": rep.message,
                              "backup": str(rep.backup) if rep.backup else None,
                              "laporan": str(rep.laporan) if rep.laporan else None}
            except Matrix2ImError as exc:
                fields = {"state": "error", "error": str(exc), "backup": None}
            except Exception as exc:  # noqa: BLE001 - hanya nama tipe; pesan asli bisa memuat data sensitif
                log.error("Job tulis IM MASTER gagal: %s", type(exc).__name__)
                fields = {"state": "error", "error": f"Gagal tak terduga: {type(exc).__name__}", "backup": None}
        finally:
            self._finish(**fields)


def _lan_address() -> str | None:
    """IP keluar PC ini (best effort): UDP connect tidak mengirim paket apa pun."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("10.255.255.255", 1))
            return sock.getsockname()[0]
    except OSError:
        return None


def run_pipeline(matrix: Path, im: Path | None) -> dict:
    """Tahap 1 + Tahap 2 + validasi, sama seperti `matrix2im run`, tapi hasil validasi
    dikembalikan (bukan menggagalkan proses) agar pengguna tetap bisa melihat datanya."""
    started = time.perf_counter()
    raw = extract(matrix)
    final = transform(raw)
    header = read_im_header(im) if im else None
    errors = validate(raw, final, header)
    return {"seconds": round(time.perf_counter() - started, 2), "raw": raw, "final": final,
            "validation_errors": errors, "im_checked": header is not None}


class _Handler(BaseHTTPRequestHandler):
    server: "_UiServer"
    state: UiState
    _new_cookie: str | None = None

    def _begin(self) -> None:
        """Pilih sesi dari cookie dan bersihkan sesi kedaluwarsa; dipanggil di awal tiap request."""
        sessions = self.server.sessions
        sessions.cleanup_expired()
        sid, self.state, is_new = sessions.get(self.headers.get("Cookie"))
        self._new_cookie = f"{COOKIE_NAME}={sid}; HttpOnly; SameSite=Strict; Path=/" if is_new else None

    def log_message(self, format: str, *args: object) -> None:  # noqa: A002 - signature bawaan
        log.debug("%s - %s", self.address_string(), format % args)

    # ---------- respons ----------
    def _send(self, status: int, body: bytes, content_type: str, extra: dict[str, str] | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        if self._new_cookie:
            self.send_header("Set-Cookie", self._new_cookie)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, status: int, payload: object, extra: dict[str, str] | None = None) -> None:
        body = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
        self._send(status, body, "application/json; charset=utf-8", extra)

    def _error(self, status: int, message: str) -> None:
        self._json(status, {"error": message})

    # ---------- routing ----------
    def do_GET(self) -> None:  # noqa: N802 - nama method bawaan http.server
        self._begin()
        path = urllib.parse.urlparse(self.path).path
        if path in ("/", "/index.html"):
            page = resources.files("matrix_to_im").joinpath("ui/index.html").read_bytes()
            self._send(HTTPStatus.OK, page, "text/html; charset=utf-8")
        elif path in ("/api/download/raw.json", "/api/download/final.json"):
            self._download(path.rsplit("/", 1)[1])
        elif path == "/api/write/config":
            self._write_config()
        elif path == "/api/write/status":
            self._json(HTTPStatus.OK, self.server.writer.status())
        elif path == "/api/write/preview.csv":
            self._write_preview_csv()
        else:
            self._error(HTTPStatus.NOT_FOUND, "Halaman tidak ditemukan")

    def do_POST(self) -> None:  # noqa: N802
        self._begin()
        path = urllib.parse.urlparse(self.path).path
        kind = path.removeprefix("/api/upload/") if path.startswith("/api/upload/") else None
        if kind in UPLOAD_KINDS:
            self._upload(kind)
        elif path == "/api/run":
            self._drain()
            self._run()
        elif path == "/api/write/prepare":
            self._write_prepare()
        elif path == "/api/write/commit":
            self._write_commit()
        else:
            self._drain()
            self._error(HTTPStatus.NOT_FOUND, "Halaman tidak ditemukan")

    def do_DELETE(self) -> None:  # noqa: N802
        self._begin()
        path = urllib.parse.urlparse(self.path).path
        kind = path.removeprefix("/api/upload/") if path.startswith("/api/upload/") else None
        if kind not in UPLOAD_KINDS:
            self._error(HTTPStatus.NOT_FOUND, "Halaman tidak ditemukan")
            return
        state = self.state
        with state.lock:
            entry = state.uploads.pop(kind, None)
        if entry:
            entry[0].unlink(missing_ok=True)
        self._json(HTTPStatus.OK, {"ok": True})

    # ---------- aksi ----------
    def _content_length(self) -> int:
        try:
            return max(0, int(self.headers.get("Content-Length", "0")))
        except ValueError:
            return 0

    def _drain(self) -> None:
        """Buang body yang tidak dipakai agar koneksi tetap bersih untuk browser."""
        remaining = self._content_length()
        while remaining > 0:
            chunk = self.rfile.read(min(_CHUNK, remaining))
            if not chunk:
                break
            remaining -= len(chunk)

    def _upload(self, kind: str) -> None:
        """Body request = isi file mentah (bukan multipart), ditulis bertahap ke disk
        supaya file 125 MB tidak perlu disalin berkali-kali di memori."""
        size = self._content_length()
        if size == 0:
            self._error(HTTPStatus.BAD_REQUEST, "File kosong atau tidak terkirim")
            return
        if size > MAX_UPLOAD_BYTES:
            self._drain()
            self._error(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, "File terlalu besar (maks 600 MB)")
            return
        name = urllib.parse.unquote(self.headers.get("X-Filename", "")) or f"{kind}.xlsx"
        state = self.state
        target = state.workdir / f"{kind}.xlsx"
        remaining = size
        with state.lock, target.open("wb") as out:
            while remaining > 0:
                chunk = self.rfile.read(min(_CHUNK, remaining))
                if not chunk:
                    break
                out.write(chunk)
                remaining -= len(chunk)
            if remaining:
                target.unlink(missing_ok=True)
                state.uploads.pop(kind, None)
            else:
                state.uploads[kind] = (target, name)
        state.drop_plan()
        if remaining:
            self._error(HTTPStatus.BAD_REQUEST, "Upload terputus sebelum selesai")
            return
        self._json(HTTPStatus.OK, {"name": name, "size": size})

    def _run(self) -> None:
        state = self.state
        with state.lock:
            matrix = state.uploads.get("matrix")
            im = state.uploads.get("im")
            if im is None:   # tanpa upload IM, header diperiksa terhadap IM MASTER di setting
                master = self.server.settings.im_master
                if master is not None and master.is_file():
                    im = (master, master.name)
            if matrix is None:
                self._error(HTTPStatus.BAD_REQUEST, "Upload file MATRIX MASTER dulu")
                return
            try:
                result = run_pipeline(matrix[0], im[0] if im else None)
            except Matrix2ImError as exc:
                state.last = None
                state.drop_plan()
                self._error(HTTPStatus.UNPROCESSABLE_ENTITY, str(exc))
                return
            except Exception as exc:  # noqa: BLE001 - tampilkan di UI, bukan traceback
                state.last = None
                state.drop_plan()
                log.exception("Error tidak terduga saat memproses")
                self._error(HTTPStatus.INTERNAL_SERVER_ERROR, f"Error tidak terduga: {type(exc).__name__}: {exc}")
                return
            state.last = {**result, "matrix_name": matrix[1]}
        state.drop_plan()
        files = {"matrix": matrix[1], "im": im[1] if im else None}
        self._json(HTTPStatus.OK, {"files": files, **result})

    # ---------- tulis IM MASTER ----------
    def _read_json(self) -> dict | None:
        """Body JSON (maks 64 KB) sebagai dict; bila gagal, respons error sudah dikirim dan hasilnya None.
        Isi body tidak pernah dicatat ke log karena memuat password."""
        size = self._content_length()
        if size > MAX_JSON_BYTES:
            self._drain()
            self._error(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, "Data terlalu besar (maks 64 KB)")
            return None
        try:
            data = json.loads(self.rfile.read(size).decode("utf-8")) if size else None
        except (ValueError, UnicodeDecodeError):
            data = None
        if not isinstance(data, dict):
            self._error(HTTPStatus.BAD_REQUEST, "Data tidak valid; kirim JSON berisi isian formulir")
            return None
        return data

    def _busy_message(self) -> str | None:
        st = self.server.writer.status()
        return f"sedang ada penulisan oleh {st['user']}" if st["state"] == "running" else None

    def _write_config(self) -> None:
        cfg = self.server.settings
        if cfg.im_master is None:
            reason = "Setting im_master belum diisi (matrix2im.local.toml)"
        elif cfg.backup_dir is None:
            reason = "Setting backup_dir belum diisi (matrix2im.local.toml)"
        else:
            reason = ""
        self._json(HTTPStatus.OK, {"enabled": not reason, "reason": reason,
                                   "im_name": cfg.im_master.name if cfg.im_master else None,
                                   "engine": cfg.insert_engine})

    def _write_prepare(self) -> None:
        body = self._read_json()
        if body is None:
            return
        state = self.state
        last = state.last   # atribut biasa: jangan menunggu `lock` yang dipegang upload/run
        if last is None:
            self._error(HTTPStatus.BAD_REQUEST, "Jalankan proses dulu sebelum menulis ke IM MASTER")
            return
        name, password = body.get("name"), body.get("password")
        if not isinstance(name, str) or not name.strip():
            self._error(HTTPStatus.BAD_REQUEST, "Isi nama Anda dulu")
            return
        if not isinstance(password, str) or not password:
            self._error(HTTPStatus.BAD_REQUEST, "Isi password IM MASTER")
            return
        if (busy := self._busy_message()) is not None:
            self._error(HTTPStatus.CONFLICT, busy)
            return
        final = last["final"]
        try:
            preview = prepare_write(final, self.server.settings, password)
        except WrongPassword as exc:
            self._error(HTTPStatus.UNAUTHORIZED, str(exc))
            return
        except ImMasterBusy as exc:
            self._error(HTTPStatus.CONFLICT, str(exc))
            return
        except Matrix2ImError as exc:
            self._error(HTTPStatus.UNPROCESSABLE_ENTITY, str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - jangan bocorkan traceback/password ke browser
            log.error("Error tidak terduga saat membuat rencana tulis: %s", type(exc).__name__)
            self._error(HTTPStatus.INTERNAL_SERVER_ERROR, f"Error tidak terduga: {type(exc).__name__}")
            return
        token = secrets.token_urlsafe(16)
        with state.wlock:
            state.write = {"preview": preview, "final": final, "user": name.strip(), "token": token,
                           "matrix": last.get("matrix_name") or "(tidak diketahui)"}
        values = preview.values
        owner = {r: g.description for g in preview.projected.groups for r in g.rows}   # baris -> DESCRIPTION grup
        self._json(HTTPStatus.OK, {
            "summary": preview.summary,
            "changes": [{"row": c.row, "col": c.col, "old": c.old, "new": c.new,
                         "cell": f"{get_column_letter(c.col)}{c.row}", "description": owner.get(c.row, "")}
                        for c in values.changes[:MAX_CHANGES_SHOWN]],
            "changes_total": len(values.changes),
            "inserts": [{"description": o.description, "count": o.count} for o in preview.inserts],
            "skipped": [{"style": k.style, "reason": k.reason} for k in preview.skipped],
            "short_groups": values.short_groups, "unmatched_im_styles": values.unmatched_im_styles,
            "missing_styles": values.missing_styles,
            "warnings": values.warnings, "engine": preview.engine, "token": token})

    def _write_commit(self) -> None:
        body = self._read_json()
        if body is None:
            return
        token, password = body.get("token"), body.get("password")
        if not isinstance(token, str) or not token or not isinstance(password, str) or not password:
            self._error(HTTPStatus.BAD_REQUEST, "Token dan password wajib diisi")
            return
        if (busy := self._busy_message()) is not None:
            self._error(HTTPStatus.CONFLICT, busy)
            return
        state, settings = self.state, self.server.settings
        with state.wlock:
            plan = state.write
            if plan is None or plan["token"] is None:
                self._error(HTTPStatus.CONFLICT, "rencana sudah dipakai; buat rencana lagi")
                return
            if not hmac.compare_digest(token.encode(), plan["token"].encode()):
                self._error(HTTPStatus.FORBIDDEN, "Token rencana tidak cocok; buat rencana lagi")
                return
            plan["token"] = None   # sekali pakai
            state.write_running = True
        user = plan["user"]

        def job(progress: Callable[[str], None]) -> WriteReport:
            try:   # password hanya hidup di closure ini
                return commit_write(plan["preview"], plan["final"], settings, password, user,
                                    plan["matrix"], progress=progress)
            finally:
                with state.wlock:
                    state.write = None
                    state.write_running = False

        if not self.server.writer.start(user, job):   # balapan dengan sesi lain
            with state.wlock:
                plan["token"] = token
                state.write_running = False
            self._error(HTTPStatus.CONFLICT, self._busy_message() or "sedang ada penulisan oleh pengguna lain")
            return
        self._json(HTTPStatus.ACCEPTED, {"ok": True})

    def _write_preview_csv(self) -> None:
        """CSV semua perubahan rencana milik sesi ini (tabel di UI dibatasi 500 baris)."""
        with self.state.wlock:
            plan = self.state.write
        if plan is None:
            self._error(HTTPStatus.NOT_FOUND, "Belum ada rencana tulis; buat rencana dulu")
            return
        body = report_csv_bytes(report_rows(plan["preview"]))
        self._send(HTTPStatus.OK, body, "text/csv; charset=utf-8",
                   {"Content-Disposition": 'attachment; filename="rencana matrix2im.csv"'})

    def _download(self, name: str) -> None:
        last = self.state.last
        if last is None:
            self._error(HTTPStatus.NOT_FOUND, "Belum ada hasil; jalankan proses dulu")
            return
        data = last["raw"] if name == "raw.json" else last["final"]
        body = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        self._send(HTTPStatus.OK, body, "application/json; charset=utf-8",
                   {"Content-Disposition": f'attachment; filename="{name}"'})


class _UiServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, address: tuple[str, int], sessions: SessionStore, settings: Settings) -> None:
        super().__init__(address, _Handler)
        self.sessions = sessions
        self.settings = settings
        self.writer = WriteCoordinator()   # satu job tulis untuk semua sesi


def make_server(host: str, port: int, sessions: SessionStore, settings: Settings) -> _UiServer:
    """Buat server (belum berjalan); port 0 = pilih port bebas otomatis."""
    return _UiServer((host, port), sessions, settings)


def _stop_on_sigterm(signum: int, frame: object) -> None:
    """SIGTERM (mis. terminal ditutup / proses di-kill) diperlakukan seperti Ctrl+C supaya
    salinan Excel di folder kerja tetap dihapus."""
    raise KeyboardInterrupt


_LOOPBACK_HOSTS = ("127.0.0.1", "localhost", "::1")
_WILDCARD_HOSTS = ("0.0.0.0", "")


def serve(host: str = "127.0.0.1", port: int = 8765, open_browser: bool = True,
          settings: Settings | None = None) -> None:
    """Jalankan UI di http://<host>:<port>/ sampai Ctrl+C, lalu hapus folder kerja.

    Port yang terpakai hanya dialihkan ke port bebas untuk host loopback; server bersama
    yang pindah port diam-diam akan memutus bookmark rekan kerja, jadi itu error.
    """
    if threading.current_thread() is threading.main_thread():
        signal.signal(signal.SIGTERM, _stop_on_sigterm)
    settings = settings or Settings()
    sessions = SessionStore(Path(tempfile.mkdtemp(prefix="matrix2im-ui-")))
    try:
        try:
            server = make_server(host, port, sessions, settings)
        except OSError as exc:
            if host not in _LOOPBACK_HOSTS:
                raise Matrix2ImError(f"Port {port} sudah dipakai di {host}; hentikan program lain "
                                     "atau pilih port lain") from exc
            log.warning("Port %d sudah dipakai, memakai port lain yang bebas", port)
            server = make_server(host, 0, sessions, settings)
        actual_port = server.server_address[1]
        local_host = "127.0.0.1" if host in _WILDCARD_HOSTS else host
        url = f"http://{f'[{local_host}]' if ':' in local_host else local_host}:{actual_port}/"
        log.info("UI matrix2im berjalan di %s (tekan Ctrl+C untuk berhenti)", url)
        if host not in _LOOPBACK_HOSTS:
            ip = (_lan_address() if host in _WILDCARD_HOSTS else host) or "<alamat IP PC ini>"
            log.info("PC lain di jaringan membuka http://%s:%d/", f"[{ip}]" if ":" in ip else ip, actual_port)
        if open_browser:
            webbrowser.open(url)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            log.info("UI dihentikan")
        finally:
            server.server_close()
            if server.writer.status()["state"] == "running":   # jangan matikan penulisan di tengah jalan
                log.info("Menunggu penulisan IM MASTER selesai sebelum berhenti…")
                server.writer.wait(EXCEL_TIMEOUT_S + 120)
    finally:
        sessions.cleanup_all()

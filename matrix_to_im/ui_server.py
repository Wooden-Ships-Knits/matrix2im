"""UI web lokal untuk mencoba matrix2im: upload Excel di browser, lihat hasil kedua tahap.

Hanya memakai standard library (http.server) supaya tidak menambah dependensi. Server
sengaja hanya mendengarkan 127.0.0.1: file MATRIX berisi data produk internal. File yang
diupload disimpan sebagai salinan di folder kerja sementara, jadi file asli pengguna tidak
pernah disentuh, dan folder itu dihapus saat server berhenti.
"""
from __future__ import annotations

import json
import logging
import shutil
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

from .errors import Matrix2ImError
from .extract import extract
from .transform import transform
from .validate import read_im_header, validate

log = logging.getLogger("matrix2im.ui")

UPLOAD_KINDS = ("matrix", "im")
MAX_UPLOAD_BYTES = 600 * 1024 * 1024   # F27 = 125 MB; batas ini hanya mencegah upload tak wajar
_CHUNK = 1024 * 1024


class UiState:
    """State satu sesi UI: salinan file upload dan hasil run terakhir (untuk download)."""

    def __init__(self, workdir: Path) -> None:
        self.workdir = workdir
        self.workdir.mkdir(parents=True, exist_ok=True)
        self.uploads: dict[str, tuple[Path, str]] = {}   # kind -> (salinan, nama asli)
        self.last: dict | None = None
        self.lock = threading.Lock()   # satu run pada satu waktu; openpyxl memakan memori besar

    def cleanup(self) -> None:
        shutil.rmtree(self.workdir, ignore_errors=True)


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
        self.end_headers()
        self.wfile.write(body)

    def _json(self, status: int, payload: object, extra: dict[str, str] | None = None) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self._send(status, body, "application/json; charset=utf-8", extra)

    def _error(self, status: int, message: str) -> None:
        self._json(status, {"error": message})

    # ---------- routing ----------
    def do_GET(self) -> None:  # noqa: N802 - nama method bawaan http.server
        path = urllib.parse.urlparse(self.path).path
        if path in ("/", "/index.html"):
            page = resources.files("matrix_to_im").joinpath("ui/index.html").read_bytes()
            self._send(HTTPStatus.OK, page, "text/html; charset=utf-8")
        elif path in ("/api/download/raw.json", "/api/download/final.json"):
            self._download(path.rsplit("/", 1)[1])
        else:
            self._error(HTTPStatus.NOT_FOUND, "Halaman tidak ditemukan")

    def do_POST(self) -> None:  # noqa: N802
        path = urllib.parse.urlparse(self.path).path
        kind = path.removeprefix("/api/upload/") if path.startswith("/api/upload/") else None
        if kind in UPLOAD_KINDS:
            self._upload(kind)
        elif path == "/api/run":
            self._drain()
            self._run()
        else:
            self._drain()
            self._error(HTTPStatus.NOT_FOUND, "Halaman tidak ditemukan")

    def do_DELETE(self) -> None:  # noqa: N802
        path = urllib.parse.urlparse(self.path).path
        kind = path.removeprefix("/api/upload/") if path.startswith("/api/upload/") else None
        if kind not in UPLOAD_KINDS:
            self._error(HTTPStatus.NOT_FOUND, "Halaman tidak ditemukan")
            return
        state = self.server.state
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
        state = self.server.state
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
        if remaining:
            self._error(HTTPStatus.BAD_REQUEST, "Upload terputus sebelum selesai")
            return
        self._json(HTTPStatus.OK, {"name": name, "size": size})

    def _run(self) -> None:
        state = self.server.state
        with state.lock:
            matrix = state.uploads.get("matrix")
            im = state.uploads.get("im")
            if matrix is None:
                self._error(HTTPStatus.BAD_REQUEST, "Upload file MATRIX MASTER dulu")
                return
            try:
                result = run_pipeline(matrix[0], im[0] if im else None)
            except Matrix2ImError as exc:
                state.last = None
                self._error(HTTPStatus.UNPROCESSABLE_ENTITY, str(exc))
                return
            except Exception as exc:  # noqa: BLE001 - tampilkan di UI, bukan traceback
                state.last = None
                log.exception("Error tidak terduga saat memproses")
                self._error(HTTPStatus.INTERNAL_SERVER_ERROR, f"Error tidak terduga: {type(exc).__name__}: {exc}")
                return
            state.last = result
        files = {"matrix": matrix[1], "im": im[1] if im else None}
        self._json(HTTPStatus.OK, {"files": files, **result})

    def _download(self, name: str) -> None:
        last = self.server.state.last
        if last is None:
            self._error(HTTPStatus.NOT_FOUND, "Belum ada hasil; jalankan proses dulu")
            return
        data = last["raw"] if name == "raw.json" else last["final"]
        body = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        self._send(HTTPStatus.OK, body, "application/json; charset=utf-8",
                   {"Content-Disposition": f'attachment; filename="{name}"'})


class _UiServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, address: tuple[str, int], state: UiState) -> None:
        super().__init__(address, _Handler)
        self.state = state


def make_server(host: str, port: int, state: UiState) -> _UiServer:
    """Buat server (belum berjalan); port 0 = pilih port bebas otomatis."""
    return _UiServer((host, port), state)


def _stop_on_sigterm(signum: int, frame: object) -> None:
    """SIGTERM (mis. terminal ditutup / proses di-kill) diperlakukan seperti Ctrl+C supaya
    salinan Excel di folder kerja tetap dihapus."""
    raise KeyboardInterrupt


def serve(port: int = 8765, open_browser: bool = True) -> None:
    """Jalankan UI di http://127.0.0.1:<port>/ sampai Ctrl+C, lalu hapus folder kerja."""
    if threading.current_thread() is threading.main_thread():
        signal.signal(signal.SIGTERM, _stop_on_sigterm)
    state = UiState(Path(tempfile.mkdtemp(prefix="matrix2im-ui-")))
    try:
        try:
            server = make_server("127.0.0.1", port, state)
        except OSError:
            log.warning("Port %d sudah dipakai, memakai port lain yang bebas", port)
            server = make_server("127.0.0.1", 0, state)
        url = f"http://127.0.0.1:{server.server_address[1]}/"
        log.info("UI matrix2im berjalan di %s (tekan Ctrl+C untuk berhenti)", url)
        if open_browser:
            webbrowser.open(url)
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            log.info("UI dihentikan")
        finally:
            server.server_close()
    finally:
        state.cleanup()

"""Local HTTP server (stdlib only) serving the web UI and the workbook API.

Every file read/write goes through server.fileio (Python), never the browser.
"""
from __future__ import annotations

import gzip
import json
import mimetypes
import os
import string
import sys
import threading
import time
import traceback
from collections import OrderedDict
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

from . import fileio
from .units import DEFAULT_COL_PX, DEFAULT_ROW_PX

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
STATIC_DIR = os.path.join(SKILL_DIR, "assets", "web")
VERSION = "0.2.0"
EXTS = {".xlsx": "xlsx", ".xlsm": "xlsm", ".xls": "xls"}


class Config:
    root: str | None = None          # optional sandbox directory
    start: str = os.getcwd()         # initial directory of the file dialog
    allowed_hosts: set[str] = {"localhost", "127.0.0.1", "[::1]", "::1"}
    host_check: bool = True
    token: str | None = None         # API token (None = auth disabled, dev mode)
    last_activity: float = time.time()
    stream_clients: int = 0          # open SSE connections (keep the server alive)
    shutdown_cb = None               # set by xl.py: callable(force) -> None
    dirty_check = None               # callable() -> list of dirty doc names


class ApiError(Exception):
    def __init__(self, status, message):
        super().__init__(message)
        self.status, self.message = status, message




def _norm(path: str) -> str:
    if not path:
        raise ApiError(400, "경로가 비어 있습니다.")
    p = os.path.abspath(os.path.expanduser(path))
    if Config.root:
        root = os.path.abspath(Config.root)
        if os.path.commonpath([root, p]) != root:
            raise ApiError(403, f"허용된 루트({root}) 밖의 경로입니다.")
    return p


def _fmt_of(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext not in EXTS:
        raise ApiError(400, "xlsx / xlsm / xls 파일만 지원합니다.")
    return EXTS[ext]


def blank_workbook(name="통합 문서1.xlsx") -> dict:
    return {"name": name, "path": None, "format": "xlsx", "activeSheet": 0, "styles": [],
            "definedNames": {}, "warnings": [], "defaultFont": {"name": "맑은 고딕", "size": 11},
            "sheets": [{"name": "Sheet1", "origName": None, "state": "visible", "tabColor": None,
                        "defaultColW": DEFAULT_COL_PX, "defaultRowH": DEFAULT_ROW_PX,
                        "cols": {}, "rows": {}, "cells": [], "merges": [], "freeze": None,
                        "autoFilter": None, "charts": [], "showGrid": True}]}


# ----------------------------------------------------------------- API handlers

def api_info(_q):
    drives = []
    if os.name == "nt":
        drives = [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]
    start = Config.start if not Config.root else Config.root
    return {"start": os.path.abspath(start), "root": Config.root, "sep": os.sep,
            "drives": drives, "home": os.path.expanduser("~"), "python": sys.version.split()[0]}


def api_fs(q):
    path = _norm(q.get("path", [Config.root or Config.start])[0])
    if not os.path.isdir(path):
        raise ApiError(404, f"폴더가 없습니다: {path}")
    entries = []
    try:
        names = os.listdir(path)
    except PermissionError:
        raise ApiError(403, f"폴더 접근 권한이 없습니다: {path}")
    for name in names:
        if name.startswith(".") or name.startswith("~$"):
            continue  # hidden files and Office lock files
        full = os.path.join(path, name)
        try:
            is_dir = os.path.isdir(full)
            if not is_dir and os.path.splitext(name)[1].lower() not in EXTS:
                continue
            st = os.stat(full)
        except OSError:
            continue
        entries.append({"name": name, "path": full, "dir": is_dir,
                        "size": None if is_dir else st.st_size, "mtime": st.st_mtime})
    entries.sort(key=lambda e: (not e["dir"], e["name"].lower()))
    parent = os.path.dirname(path)
    if Config.root and os.path.commonpath([os.path.abspath(Config.root), parent]) != os.path.abspath(Config.root):
        parent = None
    return {"path": path, "parent": parent if parent != path else None, "entries": entries}


def read_workbook(path: str) -> dict:
    """Python 파일 읽기 함수: 경로 -> JSON 워크북 모델 (DRM-safe, fileio.read_bytes 사용)."""
    fmt = _fmt_of(path)
    if not os.path.isfile(path):
        raise ApiError(404, f"파일이 없습니다: {path}")
    data = fileio.read_bytes(path)
    if fmt == "xls":
        from .xls_io import read_xls
        model = read_xls(data)
    else:
        from .xlsx_io import read_xlsx
        if data[:2] != b"PK":
            raise ApiError(422, "xlsx(zip) 형식이 아닙니다. 파일이 DRM으로 암호화되어 있다면 "
                                "WEBEXCEL_READ_HOOK으로 복호화 함수를 지정하세요.")
        model = read_xlsx(data, keep_vba=(fmt == "xlsm"))
    model.update({"path": path, "name": os.path.basename(path), "format": fmt})
    return model


def write_workbook(path: str, fmt: str, model: dict, base: str | None) -> list[str]:
    """Python 파일 쓰기 함수: JSON 워크북 모델 -> 파일."""
    warns: list[str] = []
    if fmt == "xls":
        from .xls_io import write_xls
        data, warns = write_xls(model)
    else:
        from .xlsx_io import write_xlsx
        base_bytes = None
        if base and os.path.isfile(base) and EXTS.get(os.path.splitext(base)[1].lower()) in ("xlsx", "xlsm"):
            base_bytes = fileio.read_bytes(base)
            if base_bytes[:2] != b"PK":
                base_bytes = None
        keep_vba = fmt == "xlsm" and base_bytes is not None and base.lower().endswith(".xlsm")
        data = write_xlsx(model, base_bytes, keep_vba=keep_vba)
    fileio.write_bytes(path, data)
    return warns


def api_open(body):
    return read_workbook(_norm(body.get("path", "")))


def api_save(body):
    path = _norm(body.get("path", ""))
    fmt = _fmt_of(path)
    if os.path.isdir(path):
        raise ApiError(400, "폴더 경로에는 저장할 수 없습니다.")
    if not os.path.isdir(os.path.dirname(path)):
        raise ApiError(404, "저장할 폴더가 없습니다.")
    if os.path.exists(path) and not body.get("overwrite"):
        raise ApiError(409, "같은 이름의 파일이 이미 있습니다.")
    base = body.get("base")
    base = _norm(base) if base else None
    warns = write_workbook(path, fmt, body["workbook"], base)
    return {"ok": True, "path": path, "name": os.path.basename(path), "format": fmt,
            "warnings": warns, "savedAt": time.time()}


def api_new(_body):
    return blank_workbook()


# ----------------------------------------------------------------- routing

ROUTES: list = []  # (method, compiled regex, handler, opts)


def route(method, pattern, **opts):
    """Register handler(h, m, q, body) -> dict | bytes | None (None = handler wrote the response)."""
    import re

    def deco(fn):
        ROUTES.append((method, re.compile("^" + pattern + "$"), fn, opts))
        return fn
    return deco


@route("GET", "/api/health", auth=False)
def r_health(h, m, q, body):
    base = {"ok": True, "app": "web-excel"}
    if h.authorized():
        base.update({"version": VERSION, "pid": os.getpid(), "root": Config.root or Config.start})
    return base


@route("GET", "/api/info")
def r_info(h, m, q, body):
    return api_info(q)


@route("GET", "/api/fs")
def r_fs(h, m, q, body):
    return api_fs(q)


@route("POST", "/api/shutdown")
def r_shutdown(h, m, q, body):
    dirty = Config.dirty_check() if Config.dirty_check else []
    if dirty and not body.get("force"):
        raise ApiError(409, "저장되지 않은 문서가 있습니다: " + ", ".join(dirty))
    if Config.shutdown_cb:
        threading.Timer(0.2, Config.shutdown_cb).start()
    return {"ok": True, "dirty": dirty}


# ----------------------------------------------------------------- HTTP plumbing

class Handler(BaseHTTPRequestHandler):
    server_version = "web-excel/" + VERSION
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):  # quieter log
        sys.stderr.write("[web-excel] " + (fmt % args) + "\n")

    def _host_ok(self) -> bool:
        if not Config.host_check:
            return True
        host = (self.headers.get("Host") or "").strip()
        name = host.rsplit(":", 1)[0] if not host.startswith("[") else host.split("]")[0] + "]"
        return name in Config.allowed_hosts

    def _cookie_token(self):
        for part in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = part.strip().partition("=")
            if k == "wx_token":
                return v
        return None

    def authorized(self) -> bool:
        """CLI: X-WebExcel-Token header. Browser: SameSite=Strict cookie set when the page loads
        (+ the X-WebExcel header on POST, which forces a CORS preflight we never approve)."""
        if not Config.token:
            return self.command == "GET" or self.headers.get("X-WebExcel") == "1"
        import hmac
        hdr = self.headers.get("X-WebExcel-Token")
        if hdr and hmac.compare_digest(hdr, Config.token):
            return True
        ck = self._cookie_token()
        if ck and hmac.compare_digest(ck, Config.token):
            return self.command == "GET" or self.headers.get("X-WebExcel") == "1"
        return False

    def _send(self, status, body: bytes, ctype: str, extra=None):
        gz = "gzip" in (self.headers.get("Accept-Encoding") or "") and len(body) > 2048 and (
            ctype.startswith(("application/json", "text/", "application/javascript")))
        if gz:
            body = gzip.compress(body, 5)
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        if gz:
            self.send_header("Content-Encoding", "gzip")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _json(self, status, obj):
        self._send(status, json.dumps(obj, ensure_ascii=False, allow_nan=False, default=str)
                   .encode("utf-8"), "application/json; charset=utf-8")

    def _err(self, status, msg, **extra):
        self._json(status, {"error": msg, **extra})

    def _body(self) -> bytes:
        n = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(n) if n else b""

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        self._dispatch("GET")

    def do_POST(self):
        self._dispatch("POST")

    def do_DELETE(self):
        self._dispatch("DELETE")

    def _dispatch(self, method):
        Config.last_activity = time.time()
        if not self._host_ok():
            return self._err(403, "Host not allowed")
        url = urlparse(self.path)
        # always drain the body: leaving it unread desyncs the next request on a keep-alive socket
        raw = self._body() if method in ("POST", "DELETE", "PUT") else b""
        if not url.path.startswith("/api/"):
            return self._static(url.path) if method in ("GET", "HEAD") else self._err(405, "method not allowed")
        q = parse_qs(url.query)
        for meth, rx, fn, opts in ROUTES:
            if meth != (method if method != "HEAD" else "GET"):
                continue
            m = rx.match(url.path)
            if not m:
                continue
            if opts.get("auth", True) and not self.authorized():
                return self._err(401, "인증 토큰이 필요합니다.")
            try:
                body = json.loads(raw) if raw else {}
                res = fn(self, m, q, body)
                if res is None:
                    return None
                if isinstance(res, bytes):
                    return self._send(200, res, "application/json; charset=utf-8")
                return self._json(200, res)
            except ApiError as e:
                return self._err(e.status, e.message, **getattr(e, "extra", {}))
            except PermissionError as e:
                return self._err(403, f"권한이 없습니다(파일이 다른 프로그램에서 열려 있을 수 있음): {e}")
            except (BrokenPipeError, ConnectionResetError):
                return None
            except Exception as e:
                traceback.print_exc()
                return self._err(500, f"{type(e).__name__}: {e}")
        return self._err(404, "not found")

    def _static(self, path):
        if path in ("", "/"):
            path = "/index.html"
        full = os.path.abspath(os.path.join(STATIC_DIR, path.lstrip("/")))
        if os.path.commonpath([STATIC_DIR, full]) != STATIC_DIR or not os.path.isfile(full):
            return self._err(404, "not found")
        ctype = mimetypes.guess_type(full)[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype == "application/javascript":
            ctype += "; charset=utf-8"
        extra = {}
        if path == "/index.html" and Config.token:
            extra["Set-Cookie"] = f"wx_token={Config.token}; Path=/; HttpOnly; SameSite=Strict"
        with open(full, "rb") as fh:
            self._send(200, fh.read(), ctype, extra)


def serve(host="127.0.0.1", port=8765):
    mimetypes.add_type("application/javascript", ".js")
    mimetypes.add_type("text/css", ".css")
    httpd = ThreadingHTTPServer((host, port), Handler)
    httpd.daemon_threads = True
    return httpd

from . import docs_api  # noqa: E402,F401  (registers /api/docs routes)

"""Local HTTP server (stdlib only) serving the web UI and the workbook API.

Every file read/write goes through server.fileio (Python), never the browser.
"""
from __future__ import annotations

import gzip
import hmac
import json
import re
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
EXTS = {".xlsx": "xlsx", ".xlsm": "xlsm", ".xls": "xls", ".csv": "csv", ".tsv": "csv"}


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
        raise ApiError(400, "The path is empty.")
    p = os.path.abspath(os.path.expanduser(path))
    if Config.root:
        root = os.path.abspath(Config.root)
        if os.path.commonpath([root, p]) != root:
            raise ApiError(403, f"The path is outside the allowed root ({root}).")
    return p


def _fmt_of(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext not in EXTS:
        raise ApiError(400, "Only xlsx / xlsm / xls / csv / tsv files are supported.")
    return EXTS[ext]


# A new workbook's name and body font, as Excel in that UI language makes them ("ko": Korean Excel).
NEW_WORKBOOK = {"en": ("Book1.xlsx", "Calibri"), "ko": ("통합 문서1.xlsx", "맑은 고딕")}


def blank_workbook(name=None, lang=None) -> dict:
    default_name, font = NEW_WORKBOOK.get(lang or "en") or NEW_WORKBOOK["en"]
    return {"name": name or default_name, "path": None, "format": "xlsx", "activeSheet": 0, "styles": [],
            "definedNames": {}, "warnings": [], "defaultFont": {"name": font, "size": 11},
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
        raise ApiError(404, f"Folder not found: {path}")
    entries = []
    try:
        names = os.listdir(path)
    except PermissionError:
        raise ApiError(403, f"No permission to read the folder: {path}")
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
    """The Python file reader: path -> JSON workbook model (DRM-safe, through fileio.read_bytes)."""
    fmt = _fmt_of(path)
    if not os.path.isfile(path):
        raise ApiError(404, f"File not found: {path}")
    data = fileio.read_bytes(path)
    if fmt == "xls":
        try:
            from .xls_io import read_xls
        except ImportError:
            raise ApiError(422, "Opening .xls files needs the Python package xlrd: "
                                "python -m pip install -r scripts/requirements.txt (in the skill folder)")
        model = read_xls(data)
    elif fmt == "csv":
        from .csv_io import read_csv
        model = read_csv(data, path)
    else:
        from .xlsx_io import read_xlsx
        if data[:2] != b"PK":
            raise ApiError(422, "Not an xlsx (zip) file. If it is DRM-encrypted, "
                                "set MONOONOM_READ_HOOK to a decryption function.")
        model = read_xlsx(data, keep_vba=(fmt == "xlsm"))
    model.update({"path": path, "name": os.path.basename(path), "format": fmt})
    return model


def xlsx_target(path: str, fmt: str | None) -> tuple[str, str, bool]:
    """Saving never writes .xls. Returns (path, format, converted).

    report.xls -> report.xlsx (or report.xlsm when format 'xlsm' was asked for); converted=True.
    A path that already has an .xlsx/.xlsm extension keeps it and its format, whatever format was
    passed (e.g. a stale client sending 'xls'); that is not a conversion."""
    fmt = (fmt or "").lower() or None
    root, ext = os.path.splitext(path)
    ext = ext.lower()
    if ext in (".xls", ".csv", ".tsv"):  # a CSV keeps one sheet of values: save the workbook as .xlsx
        new = fmt if fmt in ("xlsx", "xlsm") else "xlsx"
        return root + "." + new, new, True
    if ext in (".xlsx", ".xlsm"):
        return path, ext[1:], False  # the extension decides; a stale client may still send 'xls'
    if fmt in ("xls", "csv"):
        return path + ".xlsx", "xlsx", True
    return path, fmt, False


def write_workbook(path: str, fmt: str, model: dict, base: str | None) -> list[str]:
    """The Python file writer: JSON workbook model -> file."""
    warns: list[str] = []
    if fmt == "xls" or path.lower().endswith(".xls"):
        # .xls is read-only (see xls_io); callers convert the target with xlsx_target() first
        raise ApiError(400, "Files are not saved as xls. Save as .xlsx.")
    if fmt == "csv" or path.lower().endswith((".csv", ".tsv")):  # only on an explicit save as .csv
        from .csv_io import write_csv
        data, warns = write_csv(model, path)
        fileio.write_bytes(path, data)
        return warns
    from .xlsx_io import write_xlsx
    base_bytes = None
    if base and os.path.isfile(base) and EXTS.get(os.path.splitext(base)[1].lower()) in ("xlsx", "xlsm"):
        base_bytes = fileio.read_bytes(base)
        if base_bytes[:2] != b"PK":
            base_bytes = None
    keep_vba = fmt == "xlsm" and base_bytes is not None and base.lower().endswith(".xlsm")
    data = write_xlsx({**model, "fileName": os.path.basename(path)}, base_bytes, keep_vba=keep_vba)
    fileio.write_bytes(path, data)
    return warns


def api_open(body):
    return read_workbook(_norm(body.get("path", "")))


def api_save(body):
    path, fmt, _ = xlsx_target(_norm(body.get("path", "")), None)
    fmt = fmt or _fmt_of(path)
    if os.path.isdir(path):
        raise ApiError(400, "Cannot save to a folder path.")
    if not os.path.isdir(os.path.dirname(path)):
        raise ApiError(404, "The folder to save in does not exist.")
    if os.path.exists(path) and not body.get("overwrite"):
        raise ApiError(409, "A file with this name already exists.")
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
    base = {"ok": True, "app": "monoonom"}
    if h.authorized():
        base.update({"version": VERSION, "pid": os.getpid(), "root": Config.root or Config.start})
    return base


@route("GET", "/api/info")
def r_info(h, m, q, body):
    return api_info(q)


@route("GET", "/api/fonts")
def r_fonts(h, m, q, body):
    """Fonts installed on this computer (server/fonts.py), for the preview's font picker."""
    from .fonts import list_fonts
    return list_fonts(refresh=(q.get("refresh") or [""])[0] == "1", lang=(q.get("lang") or ["en"])[0])


@route("GET", "/api/fs")
def r_fs(h, m, q, body):
    return api_fs(q)


@route("POST", "/api/shutdown")
def r_shutdown(h, m, q, body):
    dirty = Config.dirty_check() if Config.dirty_check else []
    if dirty and not body.get("force"):
        raise ApiError(409, "Unsaved documents: " + ", ".join(dirty))
    if Config.shutdown_cb:
        threading.Timer(0.2, Config.shutdown_cb).start()
    return {"ok": True, "dirty": dirty}


# ----------------------------------------------------------------- HTTP plumbing

class Handler(BaseHTTPRequestHandler):
    server_version = "monoonom/" + VERSION
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):  # quieter log; never write stream tickets to the log
        sys.stderr.write("[monoonom] " + _TICKET_RX.sub("ticket=***", fmt % args) + "\n")

    def _host_ok(self) -> bool:
        if not Config.host_check:
            return True
        host = (self.headers.get("Host") or "").strip()
        name = host.rsplit(":", 1)[0] if not host.startswith("[") else host.split("]")[0] + "]"
        return name in Config.allowed_hosts

    def cookie_name(self) -> str:
        # cookies are shared by all ports of a host, so each server (port) gets its own name
        return f"mn_token_{self.server.server_address[1]}"

    def _cookie_token(self):
        name = self.cookie_name()
        for part in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = part.strip().partition("=")
            if k == name:
                return v
        return None

    def authorized(self, q=None, allow_ticket=False) -> bool:
        """Who may call /api/*:
        - X-Monoonom-Token header: the CLI, and the page with the token from its URL (#t=...). Works in
          previews that embed the page in a cross-site iframe (VS Code, Tauri apps), where cookies are
          not sent. A custom header also forces a CORS preflight, which this server never approves.
        - mn_token_<port> cookie (SameSite=Strict, set when the page loads) for top-level tabs; POSTs
          must also carry the X-Monoonom header.
        - ?ticket= (single use, 60 s) for the event stream only: EventSource cannot send headers."""
        self.auth_reason = ""
        if not Config.token:
            return self.command == "GET" or self.headers.get("X-Monoonom") == "1"
        hdr = self.headers.get("X-Monoonom-Token")
        if hdr and same_secret(hdr, Config.token):
            return True
        # a wrong header (e.g. a token a tab kept from another server on this port) does not hide a
        # valid cookie
        ck = self._cookie_token()
        if ck and same_secret(ck, Config.token):
            if self.command == "GET" or self.headers.get("X-Monoonom") == "1":
                return True
        if allow_ticket and q and take_ticket((q.get("ticket") or [""])[0]):
            return True
        self.auth_reason = "token-mismatch" if hdr else "cookie-mismatch" if ck else "no-credentials"
        return False

    def auth_error(self):
        site = self.headers.get("Sec-Fetch-Site") or "-"
        sys.stderr.write(f"[monoonom] 401 {self.command} {urlparse(self.path).path} reason={self.auth_reason}"
                         f" sec-fetch-site={site}\n")
        if self.auth_reason == "token-mismatch":
            msg = ("The token does not match this server: it restarted with another token, or the URL belongs to another server. "
                   "Reopen the latest URL the agent gave you (with #t=…).")
        else:
            msg = ("No credentials. Open the URL the agent gave you exactly, including the #t=… part"
                   + (" (cookies are not sent when the preview shows the page inside another site)." if site == "cross-site" else "."))
        return self._err(401, msg, reason=self.auth_reason)

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
            if opts.get("auth", True) and not self.authorized(q, allow_ticket=opts.get("ticket", False)):
                return self.auth_error()
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
                return self._err(403, f"Permission denied (the file may be open in another program): {e}")
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
            extra["Set-Cookie"] = f"{self.cookie_name()}={Config.token}; Path=/; HttpOnly; SameSite=Strict"
        with open(full, "rb") as fh:
            self._send(200, fh.read(), ctype, extra)


_TICKET_RX = re.compile(r"ticket=[^&\s\"]+")


def same_secret(given: str, expected: str) -> bool:
    """Constant-time comparison that never raises: header values are latin-1 str and may contain
    non-ASCII characters, which hmac.compare_digest rejects for str."""
    return hmac.compare_digest(given.encode("latin-1", "replace"), expected.encode("latin-1", "replace"))


_TICKETS: dict = {}
_tickets_lock = threading.Lock()


def new_ticket(ttl=60.0) -> str:
    """Single-use credential for one EventSource connection (it cannot send the token header)."""
    import secrets
    t = secrets.token_urlsafe(18)
    now = time.time()
    with _tickets_lock:
        for k in [k for k, exp in _TICKETS.items() if exp < now]:
            del _TICKETS[k]
        _TICKETS[t] = now + ttl
    return t


def take_ticket(t: str) -> bool:
    if not t:
        return False
    with _tickets_lock:
        exp = _TICKETS.pop(t, None)
    return exp is not None and exp >= time.time()


def serve(host="127.0.0.1", port=8765):
    mimetypes.add_type("application/javascript", ".js")
    mimetypes.add_type("text/css", ".css")
    httpd = ThreadingHTTPServer((host, port), Handler)
    httpd.daemon_threads = True
    return httpd

from . import docs_api  # noqa: E402,F401  (registers /api/docs routes)
from . import agent_api  # noqa: E402,F401  (registers agent edit/read routes)

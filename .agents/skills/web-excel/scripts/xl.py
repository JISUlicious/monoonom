#!/usr/bin/env python3
"""web-excel skill CLI — single entry point for the agent (stdlib only + bundled _vendor).

Server lifecycle:
  xl.py serve [--port 8765] [--root DIR]      start (detached) or reuse the workspace server
  xl.py serve --foreground [--exit-with-parent]   run in this process (launch.json / dev)
  xl.py status | stop [--force] | launch-config [--port P]

Exit codes: 0 ok, 2 usage, 3 server not running, 4 conflict, 5 file/DRM error, 1 other.
"""
from __future__ import annotations

import argparse
import json
import os
import secrets
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "_vendor"))  # bundled pure-Python deps first
sys.path.insert(0, HERE)

DEFAULT_PORT = 8765
STATE_DIR = ".web-excel"
EXIT_USAGE, EXIT_NOSERVER, EXIT_CONFLICT, EXIT_FILE = 2, 3, 4, 5


class CliError(Exception):
    def __init__(self, msg, code=1):
        super().__init__(msg)
        self.code = code


# ----------------------------------------------------------------- workspace state

def workspace_root(arg=None) -> str:
    root = arg or os.environ.get("WEBEXCEL_ROOT") or os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()
    return os.path.abspath(os.path.expanduser(root))


def state_path(root) -> str:
    return os.path.join(root, STATE_DIR, "server.json")


def read_state(root):
    try:
        with open(state_path(root), encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def write_state(root, data):
    d = os.path.join(root, STATE_DIR)
    os.makedirs(d, exist_ok=True)
    gi = os.path.join(d, ".gitignore")
    if not os.path.exists(gi):
        with open(gi, "w") as fh:
            fh.write("*\n")  # never commit tokens, logs or backups
    fd, tmp = tempfile.mkstemp(dir=d, prefix=".server.", suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        json.dump(data, fh)
    try:
        os.chmod(tmp, 0o600)
    except OSError:
        pass
    os.replace(tmp, state_path(root))  # atomic: readers never see a half-written file


def clear_state(root, pid):
    st = read_state(root)
    if st and st.get("pid") == pid:
        try:
            os.remove(state_path(root))
        except OSError:
            pass


# ----------------------------------------------------------------- HTTP client

def request(st, method, path, body=None, timeout=30):
    url = f"http://127.0.0.1:{st['port']}{path}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("X-WebExcel-Token", st.get("token") or "")
    req.add_header("X-WebExcel", "1")
    if data is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as res:
            raw = res.read()
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read()).get("error", str(e))
        except Exception:
            msg = str(e)
        code = EXIT_CONFLICT if e.code == 409 else EXIT_FILE if e.code in (403, 404, 422) else 1
        raise CliError(f"{msg} (HTTP {e.code})", code)
    except (urllib.error.URLError, ConnectionError, socket.timeout) as e:
        raise CliError(f"서버에 연결할 수 없습니다: {e}", EXIT_NOSERVER)


def alive(st) -> bool:
    if not st:
        return False
    try:
        h = request(st, "GET", "/api/health", timeout=2)
        return bool(h.get("ok")) and h.get("pid") == st.get("pid")
    except CliError:
        return False


def connect(root):
    st = read_state(root)
    if not alive(st):
        raise CliError(f"web-excel 서버가 실행 중이 아닙니다. 먼저 `xl.py serve`를 실행하세요. (root={root})", EXIT_NOSERVER)
    return st


# ----------------------------------------------------------------- serve

def port_free(port) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(("127.0.0.1", port))
            return True
        except OSError:
            return False


def pick_port(preferred) -> int:
    for p in range(preferred, preferred + 50):
        if port_free(p):
            return p
    with socket.socket() as s:  # let the kernel choose
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def run_foreground(args):
    from server import app

    root = workspace_root(args.root)
    port = args.port if args.port else pick_port(DEFAULT_PORT)
    token = os.environ.get("WEBEXCEL_TOKEN") or secrets.token_urlsafe(24)
    app.Config.root = root if not args.no_sandbox else None
    app.Config.start = root
    app.Config.token = None if args.no_auth else token
    app.Config.allowed_hosts |= set(args.allow_host or [])
    try:
        httpd = app.serve("127.0.0.1", port)
    except OSError as e:
        sys.stderr.write(f"[web-excel] 127.0.0.1:{port}에 바인딩할 수 없습니다: {e}\n"
                         "에이전트 샌드박스가 루프백 바인딩을 막는 경우(예: Codex workspace-write) 승인 또는 전체 권한이 필요합니다.\n")
        return EXIT_USAGE
    pid = os.getpid()
    write_state(root, {"pid": pid, "port": port, "token": token, "root": root,
                       "url": f"http://localhost:{port}/", "started_at": time.time(),
                       "version": app.VERSION, "python": sys.executable})
    stop = threading.Event()

    def shutdown():
        if not stop.is_set():
            stop.set()
            threading.Thread(target=httpd.shutdown, daemon=True).start()

    app.Config.shutdown_cb = shutdown
    ppid = os.getppid()

    def watchdog():
        while not stop.wait(2):
            if args.exit_with_parent and os.name != "nt" and os.getppid() != ppid:
                sys.stderr.write("[web-excel] parent exited; shutting down\n")
                return shutdown()
            idle = time.time() - app.Config.last_activity
            if args.idle_timeout and app.Config.stream_clients == 0 and idle > args.idle_timeout:
                if app.Config.dirty_check and app.Config.dirty_check():
                    hook = getattr(app.Config, "recovery_cb", None)
                    if hook:
                        hook()
                sys.stderr.write(f"[web-excel] idle {int(idle)}s; shutting down\n")
                return shutdown()

    threading.Thread(target=watchdog, daemon=True).start()
    print(f"web-excel {app.VERSION} serving {root} at http://localhost:{port}/ (pid {pid})", flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
        clear_state(root, pid)
    return 0


def cmd_serve(args):
    if args.foreground:
        return run_foreground(args)
    root = workspace_root(args.root)
    st = read_state(root)
    if alive(st):
        print_status(st, reused=True)
        return 0
    port = pick_port(args.port or DEFAULT_PORT)
    os.makedirs(os.path.join(root, STATE_DIR), exist_ok=True)
    log = open(os.path.join(root, STATE_DIR, "server.log"), "ab")
    cmd = [sys.executable, os.path.abspath(__file__), "serve", "--foreground", "--port", str(port),
           "--root", root, "--idle-timeout", str(args.idle_timeout)]
    for h in args.allow_host or []:
        cmd += ["--allow-host", h]
    kw = {"stdin": subprocess.DEVNULL, "stdout": log, "stderr": log, "cwd": root}
    if os.name == "nt":
        kw["creationflags"] = 0x00000008 | 0x00000200 | 0x08000000  # DETACHED | NEW_GROUP | NO_WINDOW
    else:
        kw["start_new_session"] = True  # survive the agent's shell/session
    proc = subprocess.Popen(cmd, **kw)
    deadline = time.time() + 15
    while time.time() < deadline:
        time.sleep(0.2)
        st = read_state(root)
        if st and st.get("pid") == proc.pid and alive(st):
            print_status(st, reused=False)
            return 0
        if proc.poll() is not None:
            break
    tail = ""
    try:
        with open(os.path.join(root, STATE_DIR, "server.log"), "rb") as fh:
            tail = fh.read()[-1500:].decode("utf-8", "replace")
    except OSError:
        pass
    raise CliError("서버를 시작하지 못했습니다.\n" + tail, 1)


def print_status(st, reused=None, extra=None):
    tag = " (기존 서버 재사용)" if reused else " (새로 시작)" if reused is False else ""
    print(f"web-excel 서버 실행 중{tag}")
    print(f"  URL : {st['url']}   ← 미리보기 창에서 이 주소를 여세요")
    print(f"  root: {st['root']}")
    print(f"  pid : {st['pid']}")
    for line in extra or []:
        print("  " + line)


def cmd_status(args):
    root = workspace_root(args.root)
    st = read_state(root)
    if not alive(st):
        if args.brief:
            print(f"web-excel: not running (root={root})")
            return 0
        print(f"web-excel 서버가 실행 중이 아닙니다 (root={root}).")
        return EXIT_NOSERVER
    extra = []
    try:
        docs = request(st, "GET", "/api/docs").get("docs", [])
        for d in docs:
            extra.append(f"doc {d['docId']}: {d.get('path') or d['name']}  v{d['version']}"
                         f"{'  *수정됨*' if d.get('dirty') else ''}  clients={d.get('clients', 0)}")
    except CliError:
        pass
    if args.brief:
        print(f"web-excel: running at {st['url']} ({len(extra)} doc(s) open)")
        for e in extra:
            print("  " + e)
        return 0
    print_status(st, extra=extra)
    return 0


def cmd_stop(args):
    root = workspace_root(args.root)
    st = read_state(root)
    if not alive(st):
        print("실행 중인 서버가 없습니다.")
        clear_state(root, (st or {}).get("pid"))
        return 0
    res = request(st, "POST", "/api/shutdown", {"force": args.force})
    for _ in range(50):
        time.sleep(0.1)
        if not alive(st):
            break
    print("서버를 종료했습니다." + (f" (저장 안 된 문서 버림: {', '.join(res.get('dirty', []))})" if res.get("dirty") else ""))
    return 0


def cmd_launch_config(args):
    """Add/refresh a `.claude/launch.json` entry so Claude Code's preview owns the server (mode A)."""
    root = workspace_root(args.root)
    path = os.path.join(root, ".claude", "launch.json")
    cfg = {"version": "0.0.1", "configurations": []}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as fh:
            cfg = json.load(fh)
    confs = [c for c in cfg.get("configurations", []) if c.get("name") != "web-excel"]
    st = read_state(root)
    port = args.port or (st["port"] if alive(st) else pick_port(DEFAULT_PORT))
    py = "python" if os.name == "nt" else "python3"
    confs.append({"name": "web-excel", "runtimeExecutable": py,
                  "runtimeArgs": [os.path.abspath(__file__), "serve", "--foreground", "--port", str(port),
                                  "--root", root, "--exit-with-parent"],
                  "port": port})
    cfg["configurations"] = confs
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    print(f"{path}에 'web-excel' 미리보기 구성을 기록했습니다 (port {port}).")
    if alive(st):
        print(f"이미 {st['url']}에서 서버가 실행 중입니다. 미리보기가 같은 포트를 쓰므로 기존 서버를 먼저 `stop` 하거나 그대로 URL을 여세요.")
    return 0


# ----------------------------------------------------------------- main

def build_parser():
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--root", help="workspace root (default: $WEBEXCEL_ROOT, $CLAUDE_PROJECT_DIR or cwd)")
    ap = argparse.ArgumentParser(prog="xl.py", description="web-excel skill CLI")
    sub = ap.add_subparsers(dest="cmd", required=True)
    _orig = sub.add_parser
    sub.add_parser = lambda *a, **k: _orig(*a, parents=[common], **k)

    p = sub.add_parser("serve", help="start or reuse the workspace server")
    p.add_argument("--port", type=int, default=0)
    p.add_argument("--foreground", action="store_true")
    p.add_argument("--exit-with-parent", action="store_true")
    p.add_argument("--idle-timeout", type=int, default=3600, help="seconds; 0 = never")
    p.add_argument("--allow-host", action="append")
    p.add_argument("--no-auth", action="store_true", help=argparse.SUPPRESS)
    p.add_argument("--no-sandbox", action="store_true", help="allow paths outside the workspace root")
    p.set_defaults(fn=cmd_serve)

    p = sub.add_parser("status", help="server status and open documents")
    p.add_argument("--brief", action="store_true")
    p.set_defaults(fn=cmd_status)

    p = sub.add_parser("stop", help="stop the server")
    p.add_argument("--force", action="store_true")
    p.set_defaults(fn=cmd_stop)

    p = sub.add_parser("launch-config", help="write a .claude/launch.json preview entry")
    p.add_argument("--port", type=int, default=0)
    p.set_defaults(fn=cmd_launch_config)

    try:
        from cli_docs import register  # document commands (P2)
        register(sub)
    except ImportError:
        pass
    return ap


def main(argv=None):
    ap = build_parser()
    args = ap.parse_args(argv)
    try:
        return args.fn(args) or 0
    except CliError as e:
        print(f"오류: {e}", file=sys.stderr)
        return e.code
    except KeyboardInterrupt:
        return 130


if __name__ == "__main__":
    sys.exit(main())

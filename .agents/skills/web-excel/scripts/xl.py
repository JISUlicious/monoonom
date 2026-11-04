#!/usr/bin/env python3
"""web-excel skill CLI — single entry point for the agent (stdlib only; the server needs openpyxl).

Server lifecycle:
  xl.py serve [--port 8765] [--root DIR]      start (detached) or reuse the workspace server
  xl.py serve --foreground [--exit-with-parent]   run in this process (launch.json / dev)
  xl.py status | stop [--force] | launch-config [--port P]

Exit codes: 0 ok, 2 usage, 3 server not running, 4 conflict, 5 file/DRM error,
6 Python package missing (see scripts/requirements.txt), 1 other.
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
SKILL_DIR = os.path.dirname(HERE)
sys.path.insert(0, HERE)
REQUIREMENTS = os.path.join(HERE, "requirements.txt")

DEFAULT_PORT = 8765
STATE_DIR = ".web-excel"
EXIT_USAGE, EXIT_NOSERVER, EXIT_CONFLICT, EXIT_FILE, EXIT_DEPS = 2, 3, 4, 5, 6


class CliError(Exception):
    def __init__(self, msg, code=1):
        super().__init__(msg)
        self.code = code


# ----------------------------------------------------------------- dependencies

def install_hint() -> str:
    return f'"{sys.executable}" -m pip install -r "{REQUIREMENTS}"'


def check_deps(quiet=False) -> dict:
    """openpyxl is required (xlsx/xlsm); xlrd is optional (.xls). Raises CliError(EXIT_DEPS) if openpyxl is missing."""
    found = {}
    for name in ("openpyxl", "xlrd"):
        try:
            mod = __import__(name)
            found[name] = getattr(mod, "__version__", "?")
        except ImportError:
            found[name] = None
    if not found["openpyxl"]:
        raise CliError("Python 패키지 openpyxl이 없습니다. 설치한 뒤 다시 실행하세요:\n  " + install_hint(), EXIT_DEPS)
    if not quiet:
        if not found["openpyxl"].startswith("3.1."):
            print(f"주의: openpyxl {found['openpyxl']}은 검증된 3.1.x가 아닙니다. 문제가 있으면: {install_hint()}")
        if not found["xlrd"]:
            print(f"참고: xlrd가 없어 .xls 파일은 열 수 없습니다(.xlsx/.xlsm은 가능). 필요하면: {install_hint()}")
    return found


# ----------------------------------------------------------------- workspace state

def workspace_root(arg=None) -> str:
    """The user's workspace: --root, $WEBEXCEL_ROOT, $CLAUDE_PROJECT_DIR, else the current directory.

    Commands must run from the workspace, not from the skill folder: the server state (.web-excel/),
    the file sandbox and relative file paths all hang off this directory."""
    root = arg or os.environ.get("WEBEXCEL_ROOT") or os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()
    root = os.path.abspath(os.path.expanduser(root))
    skill = os.path.normcase(os.path.realpath(SKILL_DIR))
    real = os.path.normcase(os.path.realpath(root))
    if real == skill or real.startswith(skill + os.sep):
        raise CliError(
            f"작업 폴더가 skill 폴더 안({root})으로 잡혔습니다. xl.py는 사용자의 작업 폴더(프로젝트 폴더)에서 "
            f"실행하세요: cd <작업 폴더> && python \"{os.path.join(HERE, 'xl.py')}\" ... "
            "또는 --root <작업 폴더>를 지정하세요.", EXIT_USAGE)
    return root


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


def workspace_token(root) -> str:
    """The workspace's browser/API token, kept in .web-excel/token (0600) across server restarts, so a
    preview opened with #t=TOKEN keeps working after the server stops (idle timeout) and starts again."""
    d = os.path.join(root, STATE_DIR)
    path = token_path(root)
    os.makedirs(d, exist_ok=True)
    for _ in range(50):
        try:
            with open(path, encoding="utf-8") as fh:
                tok = fh.read().strip()
            if len(tok) >= 16:
                return tok
        except FileNotFoundError:
            # written to a temp file and linked into place: another process never sees it half-written,
            # and if two start at once, os.link fails for the loser, which then reads the winner's token
            tok = secrets.token_urlsafe(24)
            fd, tmp = tempfile.mkstemp(dir=d, prefix=".token.", suffix=".tmp")
            try:
                with os.fdopen(fd, "w", encoding="utf-8") as fh:
                    fh.write(tok)
                try:
                    os.chmod(tmp, 0o600)
                except OSError:
                    pass
                try:
                    os.link(tmp, path)
                    return tok
                except FileExistsError:
                    pass
                except OSError:  # file systems without hard links
                    if not os.path.exists(path):
                        os.replace(tmp, path)
                        return tok
            finally:
                if os.path.exists(tmp):
                    os.remove(tmp)
        time.sleep(0.05)  # a token file that is still empty or unreadable: retry
    raise CliError(f"토큰 파일을 읽을 수 없습니다: {path}")


def token_path(root) -> str:
    return os.path.join(root, STATE_DIR, "token")


def page_url(st) -> str:
    """The URL to give the user: with #t=TOKEN (servers started before tokens were in the URL only
    stored the bare URL)."""
    url = st.get("url") or ""
    if "#t=" in url or not st.get("token"):
        return url
    return (st.get("base_url") or url.split("#", 1)[0]) + "#t=" + st["token"]


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
        if os.name != "nt":  # match the server (allow_reuse_address): TIME_WAIT is not "in use"
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
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
    # keep the workspace's token across restarts so open previews (whose URL carries #t=TOKEN) keep working
    token = os.environ.get("WEBEXCEL_TOKEN") or workspace_token(root)
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
    base = f"http://localhost:{port}/"
    write_state(root, {"pid": pid, "port": port, "token": token, "root": root,
                       "base_url": base, "url": base + (f"#t={token}" if app.Config.token else ""),
                       "started_at": time.time(),
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
        check_deps(quiet=True)
        return run_foreground(args)
    check_deps()
    root = workspace_root(args.root)
    st = read_state(root)
    if args.new_token:
        if alive(st):
            raise CliError("서버가 실행 중입니다. 토큰을 바꾸려면 먼저 `xl.py stop` 하세요"
                           "(열린 미리보기는 새 주소로 다시 열어야 합니다).", EXIT_CONFLICT)
        try:
            os.remove(token_path(root))
        except FileNotFoundError:
            pass
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


URL_HINT = ("사용자에게 아래 주소를 한 글자도 바꾸지 말고 그대로 전달하세요 "
            "(끝의 #t= 토큰 포함. localhost 주소를 직접 만들거나 줄이지 말 것):")


def print_status(st, reused=None, extra=None):
    tag = " (기존 서버 재사용)" if reused else " (새로 시작)" if reused is False else ""
    print(f"web-excel 서버 실행 중{tag}")
    print(URL_HINT)
    print(page_url(st))
    print(f"  root: {st['root']}")
    print(f"  pid : {st['pid']}")
    for line in extra or []:
        print("  " + line)


def cmd_url(args):
    """Only the preview URL with its token, for copying verbatim."""
    st = read_state(workspace_root(args.root))
    if not alive(st):
        raise CliError("web-excel 서버가 실행 중이 아닙니다. 먼저 `xl.py serve`를 실행하세요.", EXIT_NOSERVER)
    if not args.bare:
        print(URL_HINT)
    print(page_url(st))
    return 0


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
        base = st.get("base_url") or (st.get("url") or "").split("#", 1)[0]
        print(f"web-excel: running at {base} ({len(extra)} doc(s) open). "
              "This bare address does not open workbooks: give the user the URL from `xl.py url`.")
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
    p.add_argument("--new-token", action="store_true",
                   help="issue a new workspace token (old preview URLs stop working); server must be stopped")
    p.set_defaults(fn=cmd_serve)

    p = sub.add_parser("url", help="the preview URL with its token (give it to the user verbatim)")
    p.add_argument("--bare", action="store_true", help="print only the URL")
    p.set_defaults(fn=cmd_url)

    p = sub.add_parser("status", help="server status and open documents")
    p.add_argument("--brief", action="store_true")
    p.set_defaults(fn=cmd_status)

    p = sub.add_parser("stop", help="stop the server")
    p.add_argument("--force", action="store_true")
    p.set_defaults(fn=cmd_stop)

    p = sub.add_parser("launch-config", help="write a .claude/launch.json preview entry")
    p.add_argument("--port", type=int, default=0)
    p.set_defaults(fn=cmd_launch_config)

    sys.modules.setdefault("xl", sys.modules[__name__])  # cli_docs imports helpers from this module
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

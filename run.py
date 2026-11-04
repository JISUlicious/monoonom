#!/usr/bin/env python3
"""web-excel 로컬 서버 실행.

    python run.py                      # http://127.0.0.1:8765
    python run.py --port 9000 --start ./samples --root D:\\work
"""
import argparse
import os
import sys
import webbrowser

from server.app import Config, serve


def main():
    ap = argparse.ArgumentParser(description="web-excel local server")
    ap.add_argument("--host", default=os.environ.get("WEBEXCEL_HOST", "127.0.0.1"))
    ap.add_argument("--port", type=int, default=int(os.environ.get("WEBEXCEL_PORT", "8765")))
    ap.add_argument("--start", default=os.environ.get("WEBEXCEL_START", os.getcwd()),
                    help="파일 열기 대화상자의 시작 폴더")
    ap.add_argument("--root", default=os.environ.get("WEBEXCEL_ROOT"),
                    help="이 폴더 밖의 파일 접근을 차단 (선택)")
    ap.add_argument("--allow-host", action="append", default=[],
                    help="추가로 허용할 Host 헤더 이름 (프록시/원격 미리보기용)")
    ap.add_argument("--no-host-check", action="store_true")
    ap.add_argument("--open", action="store_true", help="기본 브라우저로 열기")
    a = ap.parse_args()

    Config.start = os.path.abspath(a.start)
    Config.root = os.path.abspath(a.root) if a.root else None
    Config.allowed_hosts |= set(a.allow_host)
    Config.host_check = not a.no_host_check
    httpd = serve(a.host, a.port)
    url = f"http://{'127.0.0.1' if a.host in ('0.0.0.0', '') else a.host}:{a.port}/"
    print(f"web-excel running at {url}  (Ctrl+C to stop)", flush=True)
    if a.open:
        webbrowser.open(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())

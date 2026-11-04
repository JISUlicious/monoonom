# Troubleshooting

## "openpyxl이 없습니다" / exit code 6
The Python packages are not bundled. Install them for the same Python that runs `xl.py`:
`python -m pip install -r scripts/requirements.txt` (openpyxl required; xlrd only for `.xls`). Offline PCs:
on a connected machine run `pip download -r scripts/requirements.txt -d wheels`, copy `wheels/`, then
`pip install --no-index --find-links wheels -r scripts/requirements.txt`. If `.xls` files fail to open
with "xlrd가 필요합니다", only xlrd is missing.

## "서버에 연결할 수 없습니다" / exit code 3
Run `xl.py serve`. State lives in `<workspace>/.web-excel/server.json`; the log is `.web-excel/server.log`.

## Port already in use
`serve` picks the first free port from 8765 upward and prints it. For a fixed port: `xl.py serve --port 9000`.
In Claude Code desktop, rerun `xl.py launch-config` so `.claude/launch.json` matches the port.

## Cannot bind 127.0.0.1 (PermissionError / Operation not permitted)
The agent sandbox blocks loopback sockets. Codex CLI `workspace-write` is known to block it even with
network access enabled; the user must approve the command or run with full access.

## File opens as "not an xlsx (zip)" (HTTP 422)
The file is probably DRM-encrypted for this process. Point the server at a decrypt function:
`WEBEXCEL_READ_HOOK=package.module:function` (and `WEBEXCEL_WRITE_HOOK` for saving), then restart.

## Path outside the workspace (HTTP 403)
The server only reads/writes inside the workspace root. Restart with `--root DIR` or `--no-sandbox`.

## The preview shows an old page
Reload the preview. The page gets its auth cookie when it loads, so a server restart needs a reload.

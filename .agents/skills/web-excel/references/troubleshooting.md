# Troubleshooting

## "openpyxl이 없습니다" / exit code 6
The Python packages are not bundled. Install them for the same Python that runs `xl.py`:
`python -m pip install -r scripts/requirements.txt` (openpyxl required; xlrd only for `.xls`). Offline PCs:
on a connected machine run `pip download -r scripts/requirements.txt -d wheels`, copy `wheels/`, then
`pip install --no-index --find-links wheels -r scripts/requirements.txt`. If `.xls` files fail to open
with "xlrd가 필요합니다", only xlrd is missing.

## Preview shows the menus but no workbook opens ("인증 정보가 없습니다", HTTP 401)
The page was opened without its token. Open the URL that `serve`/`status` prints, including `#t=…`.
Previews that embed the page in an iframe of another site (VS Code Simple Browser, apps built with
Tauri/Electron webviews) do not send cookies, so the token in the URL is the only way in there.
"토큰이 이 서버와 다릅니다": the URL belongs to another server or workspace; use the current `status` URL.
`.web-excel/server.log` shows each 401 with `reason=` and `sec-fetch-site=` (`cross-site` = embedded page).
If a preview URL with the token leaked, `xl.py stop` then `xl.py serve --new-token` and reopen with the new URL.

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

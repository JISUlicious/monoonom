# Security and DRM

**English** | [한국어](security.ko.md) · [Back to README](../../README.md)

## Security

- The server binds to `127.0.0.1` only and checks the `Host` header (against DNS rebinding).
- Every API call needs the workspace's token: in the preview URL (`#t=…`, sent as a header) or in a per-port cookie for tabs opened without it. Changes also need a custom header, which cross-site requests cannot send without a CORS preflight.
- Files outside the workspace folder cannot be opened (`--no-sandbox` lifts this).
- Server state (`server.json`, token, log, backups) lives in `<workspace>/.monoonom/`; keep it out of version control.
- If a preview URL with its token leaked: `xl.py stop`, then `xl.py serve --new-token`, and reopen with the new URL.

## Network: what leaves the computer

MoNooNoM itself sends nothing to the internet: no telemetry, no update checks, no web fonts or CDN.

- **Server and CLI**: the server listens on `127.0.0.1` only, and `xl.py` talks only to it. The formula engine has no
  functions that fetch from the web (such as `WEBSERVICE`).
- **Preview page**: served with a Content-Security-Policy (`default-src 'self'; connect-src 'self'; form-action 'none'; …`),
  so the browser itself blocks any request from the page to another host: workbook data cannot leave through the page,
  even from injected script. Charts (Chart.js) and every other script ship with the package.
- **Links in cells**: Ctrl+click opens only `http`, `https` and `mailto` links, in a new tab without a `Referer`
  (`noreferrer`); a link to a place in the workbook (`#Sheet!A1`) is shown in place; other kinds (file paths,
  `javascript:`) are not opened. Opening a web link is a visit to that site, as in Excel.
- **The agent's model**: what the agent reads with `xl.py read` / `describe` becomes part of its conversation and goes
  to the agent's model provider, as anything else the agent reads. Limit what the agent reads (a range, not the whole
  sheet) when that matters.
- **Install time only**: `pip install` (PyPI), `npx skills add` (npm, GitHub) and the Claude Code plugin marketplace
  (GitHub) download the package; nothing is downloaded while it runs.
- **DRM hooks** (below) run code you configure; what that code sends is up to it.

## DRM environments

The browser never reads file contents; it only sends paths. Every read and write goes through
`scripts/server/fileio.py`, so a DRM agent that decrypts files for allowed processes (Python) works as is. Where a DRM
SDK must be called, plug in its functions:

```bash
MONOONOM_READ_HOOK=mypkg.drm:read_file     # (path) -> bytes
MONOONOM_WRITE_HOOK=mypkg.drm:write_file   # (path, data: bytes) -> None
```

Files are written in place (no temporary file and rename), for DRM systems that track a file's identity. A file that
is not a zip (xlsx) when read is reported with a hint to set the read hook. Dropping files onto the page is blocked on
purpose: files are opened through the server.

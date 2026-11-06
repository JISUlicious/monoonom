# Security and DRM

**English** | [한국어](security.ko.md) · [Back to README](../../README.md)

## Security

- The server binds to `127.0.0.1` only and checks the `Host` header (against DNS rebinding).
- Every API call needs the workspace's token: in the preview URL (`#t=…`, sent as a header) or in a per-port cookie for tabs opened without it. Changes also need a custom header, which cross-site requests cannot send without a CORS preflight.
- Files outside the workspace folder cannot be opened (`--no-sandbox` lifts this).
- Server state (`server.json`, token, log, backups) lives in `<workspace>/.monoonom/`; keep it out of version control.
- If a preview URL with its token leaked: `xl.py stop`, then `xl.py serve --new-token`, and reopen with the new URL.

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

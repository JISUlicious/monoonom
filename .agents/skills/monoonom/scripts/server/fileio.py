"""File byte I/O — the single place where workbook files touch the disk.

DRM environments
----------------
The browser never reads file bytes. It only sends paths; every read and write goes through this
module's Python functions (`read_bytes`, `write_bytes`). Where a DRM agent decrypts transparently
for allowed processes (python), this just works; where a DRM SDK call is needed, environment
variables swap in hooks:

    MONOONOM_READ_HOOK=mypkg.drm:read_file     # (path) -> bytes
    MONOONOM_WRITE_HOOK=mypkg.drm:write_file   # (path, data: bytes) -> None
"""
from __future__ import annotations

import importlib
import os


def _default_read(path: str) -> bytes:
    with open(path, "rb") as fh:
        return fh.read()


def _default_write(path: str, data: bytes) -> None:
    # Write in place (not tmp+rename): some DRM agents track the file identity/handle.
    with open(path, "wb") as fh:
        fh.write(data)


def _load_hook(env_name: str, default):
    # the name before the rename (web-excel -> monoonom) still works
    spec = os.environ.get(env_name) or os.environ.get(env_name.replace("MONOONOM_", "WEBEXCEL_"))
    if not spec:
        return default
    mod_name, _, fn_name = spec.partition(":")
    fn = getattr(importlib.import_module(mod_name), fn_name or "main")
    if not callable(fn):
        raise RuntimeError(f"{env_name}={spec} is not callable")
    return fn


_read_hook = _load_hook("MONOONOM_READ_HOOK", _default_read)
_write_hook = _load_hook("MONOONOM_WRITE_HOOK", _default_write)


def read_bytes(path: str) -> bytes:
    """Read a workbook file via Python (DRM-safe entry point)."""
    return _read_hook(path)


def write_bytes(path: str, data: bytes) -> None:
    """Write a workbook file via Python (DRM-safe entry point)."""
    _write_hook(path, data)

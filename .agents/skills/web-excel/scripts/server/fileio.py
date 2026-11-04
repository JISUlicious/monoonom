"""File byte I/O — the single place where workbook files touch the disk.

DRM 환경 대응
------------
브라우저는 파일 바이트를 절대 읽지 않습니다. 브라우저는 경로만 전달하고, 실제 읽기/쓰기는
항상 이 모듈의 파이썬 함수(`read_bytes`, `write_bytes`)를 통해 수행됩니다.
DRM 에이전트가 허용된 프로세스(python)에 대해서만 투명 복호화를 해주는 환경이라면
그대로 동작하며, 별도의 DRM SDK 호출이 필요하면 환경변수로 훅을 교체할 수 있습니다.

    WEBEXCEL_READ_HOOK=mypkg.drm:read_file     # (path) -> bytes
    WEBEXCEL_WRITE_HOOK=mypkg.drm:write_file   # (path, data: bytes) -> None
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
    spec = os.environ.get(env_name)
    if not spec:
        return default
    mod_name, _, fn_name = spec.partition(":")
    fn = getattr(importlib.import_module(mod_name), fn_name or "main")
    if not callable(fn):
        raise RuntimeError(f"{env_name}={spec} is not callable")
    return fn


_read_hook = _load_hook("WEBEXCEL_READ_HOOK", _default_read)
_write_hook = _load_hook("WEBEXCEL_WRITE_HOOK", _default_write)


def read_bytes(path: str) -> bytes:
    """Read a workbook file via Python (DRM-safe entry point)."""
    return _read_hook(path)


def write_bytes(path: str, data: bytes) -> None:
    """Write a workbook file via Python (DRM-safe entry point)."""
    _write_hook(path, data)

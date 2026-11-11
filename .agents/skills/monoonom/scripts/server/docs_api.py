"""HTTP endpoints for server-owned documents (P1): open/list/snapshot/ops/SSE/presence/calc/save."""
from __future__ import annotations

import json
import os
import queue
import time

from .app import ApiError, Config, EXTS, _fmt_of, _norm, blank_workbook, read_workbook, route, write_workbook, xlsx_target
from .session import DocSession, Registry, SessionError, backup_file

REGISTRY = Registry()
Config.dirty_check = REGISTRY.dirty_names


def state_dir():
    return os.path.join(Config.root or Config.start, ".monoonom")


Config.recovery_cb = lambda: REGISTRY.write_recovery(state_dir())


MEDIA_MAX = 10 * 1024 * 1024
IMAGE_TYPES = {"png": "image/png", "jpeg": "image/jpeg", "gif": "image/gif"}


def add_media(doc, data: bytes) -> dict:
    """Store picture bytes in the document: {"media": key, "format", "w", "h"} (pixels). PNG, JPEG and GIF
    are kept as they are; other formats Pillow reads (BMP, WebP, TIFF, ...) are converted to PNG."""
    import io
    from PIL import Image as PILImage, UnidentifiedImageError
    from .xlsx_io import media_key
    if not data or len(data) > MEDIA_MAX:
        raise ApiError(400, f"Images must be 1 byte to {MEDIA_MAX // (1024 * 1024)} MB.")
    try:
        img = PILImage.open(io.BytesIO(data))
        img.load()
    except (UnidentifiedImageError, OSError) as e:
        raise ApiError(400, f"Not an image this app can read: {e}")
    fmt = (img.format or "").lower()
    if fmt not in IMAGE_TYPES:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        data, fmt = buf.getvalue(), "png"
    key = media_key(data)
    with doc.lock:
        doc.media[key] = data
    return {"media": key, "format": fmt, "w": img.width, "h": img.height}


def _doc(m):
    try:
        return REGISTRY.get(m.group(1))
    except SessionError as e:
        raise ApiError(e.status, e.message)


def _wrap(fn):
    def inner(h, m, q, body):
        try:
            return fn(h, m, q, body)
        except SessionError as e:
            err = ApiError(e.status, e.message)
            err.extra = e.extra
            raise err
    return inner


def _register(doc, by=None, client=None):
    os.makedirs(state_dir(), exist_ok=True)
    doc.audit_path = os.path.join(state_dir(), "audit.log")
    doc.audit({"event": "open"})
    return REGISTRY.add(doc, by=by, client=client)


def open_document(path=None, new=False, name=None, by=None, client=None, lang=None) -> DocSession:
    if new:
        model = blank_workbook(name or None, lang)
        return _register(DocSession(model, None, "xlsx"), by, client)
    path = _norm(path)
    existing = REGISTRY.find_path(path)
    if existing:
        if by == "agent":
            REGISTRY.focus(existing, by, client)  # `xl.py open` again: show it in the previews
        else:
            existing.focused_at = time.time()
        return existing
    model = read_workbook(path)
    return _register(DocSession(model, path, model.get("format", "xlsx")), by, client)


@route("GET", "/api/docs")
def r_docs(h, m, q, body):
    return {"docs": REGISTRY.list()}


@route("POST", "/api/docs")
def r_docs_open(h, m, q, body):
    by = "agent" if body.get("actor") == "agent" else "user"
    doc = open_document(body.get("path"), bool(body.get("new")), body.get("name"), by=by, client=body.get("clientId"), lang=body.get("lang"))
    out = doc.summary()
    out["warnings"] = doc.warnings
    return out


@route("GET", r"/api/docs/([\w-]+)")
def r_doc(h, m, q, body):
    return _doc(m).summary()


@route("DELETE", r"/api/docs/([\w-]+)")
def r_doc_close(h, m, q, body):
    doc = _doc(m)
    if doc.dirty and not body.get("force"):
        raise ApiError(409, f"'{doc.name}' has unsaved changes.")
    doc.audit({"event": "close", "discarded": bool(doc.dirty)})
    REGISTRY.remove(doc.id)
    return {"ok": True}


@route("GET", r"/api/docs/([\w-]+)/snapshot")
def r_snapshot(h, m, q, body):
    return _doc(m).snapshot()


@route("POST", r"/api/docs/([\w-]+)/ops")
@_wrap
def r_ops(h, m, q, body):
    doc = _doc(m)
    ops = body.get("ops") or []
    if not isinstance(ops, list):
        raise ApiError(400, "ops must be an array.")
    v = doc.apply(ops, actor=body.get("actor", "user"), client_id=body.get("clientId"),
                  batch_id=body.get("batchId"), label=body.get("label", ""),
                  base_version=body.get("baseVersion"), structural=bool(body.get("structural")),
                  shifts=body.get("shifts"), transformable=body.get("transformable"))
    return {"version": v}


@route("POST", r"/api/docs/([\w-]+)/presence")
def r_presence(h, m, q, body):
    doc = _doc(m)
    cid = body.get("clientId") or "anon"
    with doc.lock:
        doc.presence[cid] = {"clientId": cid, "actor": body.get("actor", "user"), "sheet": body.get("sheet"),
                             "sel": body.get("sel"), "active": body.get("active"), "ts": time.time()}
        doc.presence = {k: v for k, v in doc.presence.items() if time.time() - v["ts"] < 600}
        doc.focused_at = time.time()  # someone is working in it
    doc.hub.publish("presence", {"docId": doc.id, **doc.presence[cid]})
    return {"ok": True}


@route("GET", r"/api/docs/([\w-]+)/presence")
def r_presence_get(h, m, q, body):
    doc = _doc(m)
    with doc.lock:
        items = sorted(doc.presence.values(), key=lambda p: -p["ts"])
    return {"presence": items}


@route("POST", r"/api/docs/([\w-]+)/calc")
def r_calc(h, m, q, body):
    doc = _doc(m)
    doc.set_calc(int(body.get("version", -1)), body.get("values") or [], body.get("source", "browser"),
                 bool(body.get("full")))
    return {"calcVersion": doc.calc_version}


@route("POST", r"/api/docs/([\w-]+)/save")
@_wrap
def r_save_doc(h, m, q, body):
    doc = _doc(m)
    path = _norm(body["path"]) if body.get("path") else doc.path
    if not path:
        raise ApiError(400, "A path to save to is required (new document).")
    fmt = body.get("format") or _fmt_of(path)
    if os.path.splitext(path)[1].lower() not in EXTS:
        path += "." + fmt
    path, fmt, converted = xlsx_target(path, fmt)
    from_xls = doc.format == "xls"
    if not os.path.isdir(os.path.dirname(path)):
        raise ApiError(404, "The folder to save in does not exist.")
    same = doc.path and os.path.normcase(path) == os.path.normcase(doc.path)
    if os.path.exists(path):
        if not same and not body.get("overwrite"):
            name = os.path.basename(path)
            err = ApiError(409, f"'{name}' already exists." + (
                f" An xls file is saved as {fmt}, which would overwrite that file. Save under another name to keep it." if converted else ""))
            err.extra = {"target": path, "targetName": name, "converted": converted, "format": fmt}
            raise err
        if same and doc.file_sig and doc.file_sig != _sig(path) and not body.get("force"):
            raise ApiError(409, "The file changed on disk after it was opened (another program saved it). Overwriting needs force.")
    backup = None
    if os.path.exists(path) and body.get("backup", True):
        try:
            backup = backup_file(path, state_dir())
        except OSError as e:
            raise ApiError(403, f"Cannot create a backup ({e}). Try again with backup=false.")
    # formula results are written into the file as cached values, so make them current first
    from .agent_api import ensure_calc
    calc = ensure_calc(doc, wait=2.0)
    warns = write_workbook(path, fmt, {**doc.file_model(), "media": doc.media}, doc.base_path)
    note = []
    if converted:
        note.append(f"Files are not saved as xls, so this was saved as '{os.path.basename(path)}' ({fmt}). The original .xls is unchanged.")
    if from_xls:  # also for an explicit `--as new.xlsx` / Save As from a document read from .xls
        note.append("Formulas read from xls were saved as their computed values only (xlrd cannot read xls formula text). "
                    "Re-enter the formulas of cells that need to recalculate.")
    warns = note + list(warns or [])
    if calc.get("status") == "stale":
        warns = list(warns or []) + ["Formula results may be out of date, so cached values in the file may be old (Excel recalculates on opening)."]
    with doc.lock:
        doc.path, doc.base_path, doc.format = path, path, fmt
        doc.name = os.path.basename(path)
        doc.file_sig = _sig(path)
        meta_ops = []
        for sh in doc.sheets:
            charts = []
            for i, ch in enumerate(sh["charts"]):
                ch = dict(ch)
                ch["orig"], ch["dirty"] = i, False
                charts.append(ch)
            # pivots created in the app are now in the file, after the ones it had (xlsx_io._create_pivot)
            pivots, n = [], sum(1 for p in sh.get("pivots") or [] if not p.get("create"))
            for p in sh.get("pivots") or []:
                p = dict(p)
                if p.pop("create", None):
                    p["idx"], n = n, n + 1
                pivots.append(p)
            meta_ops.append({"op": "setSheetMeta", "sheet": sh["id"], "meta": {"origName": sh["name"], "charts": charts, "pivots": pivots}})
        doc.apply(meta_ops, actor="system", label="save")
        doc.dirty = False
        doc.audit({"event": "save", "to": path, "format": fmt, "backup": backup})
    info = {"path": path, "name": doc.name, "format": fmt, "warnings": warns, "backup": backup,
            "version": doc.version, "savedAt": time.time()}
    doc.hub.publish("saved", {"docId": doc.id, **info})
    REGISTRY.hub.publish("saved", {"docId": doc.id, **info})
    return info


@route("POST", r"/api/docs/([\w-]+)/media")
@_wrap
def r_media_add(h, m, q, body):
    """Upload a picture (the preview's paste): {data: base64} -> {media, format, w, h}."""
    import base64
    doc = _doc(m)
    try:
        data = base64.b64decode(body.get("data") or "", validate=True)
    except ValueError:
        raise ApiError(400, "data must be base64.")
    return add_media(doc, data)


@route("GET", r"/api/docs/([\w-]+)/media/(\w+)")
@_wrap
def r_media_get(h, m, q, body):
    doc = _doc(m)
    data = doc.media.get(m.group(2))
    if data is None:
        raise ApiError(404, "Image not found.")
    from PIL import Image as PILImage
    import io
    fmt = (PILImage.open(io.BytesIO(data)).format or "png").lower()
    h._send(200, data, IMAGE_TYPES.get(fmt, "application/octet-stream"))
    return None


@route("POST", r"/api/docs/([\w-]+)/undo")
@_wrap
def r_undo(h, m, q, body):
    """Undo one batch (typically the agent's) by re-applying its recorded previous state."""
    doc = _doc(m)
    with doc.lock:
        entry = next((e for e in doc.log if e["batchId"] == body.get("batchId")), None)
        if not entry:
            raise ApiError(404, "That change (batch) was not found (it may have aged out of the history).")
        touched = _touched(entry["ops"])
        later = [e for e in doc.log if e["v"] > entry["v"] and _touched(e["ops"]) & touched]
        if later and not body.get("force"):
            raise ApiError(409, f"The same cells / sheets changed {len(later)} more times after this change. Undoing anyway needs force.")
        v = doc.apply(entry["undo"], actor=body.get("actor", "agent"), label=f"undo: {entry['label']}",
                      structural=entry["structural"],
                      shifts=[{**x, "count": -x["count"]} for x in reversed(entry.get("shifts", []))],
                      transformable=entry.get("transformable"))
    return {"version": v}


def _touched(ops):
    out = set()
    for op in ops:
        if op.get("op") == "setCells":
            out |= {(op["sheet"], r, c) for r, c, _ in op["cells"]}
        elif op.get("op") == "setSheetMeta":
            out |= {(op["sheet"], "meta", k) for k in op["meta"]}
        else:
            out.add(("workbook", op.get("op")))
    return out


def _sig(path):
    try:
        st = os.stat(path)
        return [st.st_mtime_ns, st.st_size]
    except OSError:
        return None


# ------------------------------------------------------------------ SSE

def _sse(h, hub, initial):
    h.send_response(200)
    h.send_header("Content-Type", "text/event-stream; charset=utf-8")
    h.send_header("Cache-Control", "no-store")
    h.send_header("X-Accel-Buffering", "no")
    h.end_headers()
    h.close_connection = True
    q = hub.subscribe()
    Config.stream_clients += 1

    def send(event, data):
        payload = json.dumps(data, ensure_ascii=False, default=str)
        h.wfile.write(f"event: {event}\ndata: {payload}\n\n".encode("utf-8"))
        h.wfile.flush()

    try:
        h.wfile.write(b"retry: 1500\n\n")
        for event, data in initial:
            send(event, data)
        while True:
            try:
                event, data = q.get(timeout=15)
            except queue.Empty:
                h.wfile.write(b": ping\n\n")
                h.wfile.flush()
                Config.last_activity = time.time()
                continue
            send(event, data)
            if event == "closed":
                break
    except (BrokenPipeError, ConnectionResetError, OSError):
        pass
    finally:
        hub.unsubscribe(q)
        Config.stream_clients -= 1
        Config.last_activity = time.time()
    return None


@route("GET", r"/api/docs/([\w-]+)/events")
def r_events(h, m, q, body):
    doc = _doc(m)
    since = int((q.get("since") or ["-1"])[0])
    initial = []
    if since >= 0:
        evs = doc.events_since(since)
        if evs is None:
            initial.append(("resync", {"version": doc.version}))
        else:
            initial += [("ops", e) for e in evs]
    initial.append(("hello", {"version": doc.version, "dirty": doc.dirty}))
    return _sse(h, doc.hub, initial)


@route("GET", "/api/events")
def r_global_events(h, m, q, body):
    return _sse(h, REGISTRY.hub, [("docs", {"docs": REGISTRY.list()})])


# ------------------------------------------------------------------ multiplexed stream (one per window)
# Browsers allow ~6 HTTP/1.1 connections per host across all tabs, so a window keeps ONE event
# stream and subscribes documents over POST instead of opening an EventSource per document.
import threading as _th
import queue as _queue

STREAMS: dict = {}
_streams_lock = _th.Lock()


@route("POST", "/api/stream/ticket")
def r_stream_ticket(h, m, q, body):
    """Single-use ticket for GET /api/stream (EventSource cannot send the X-Monoonom-Token header)."""
    from .app import new_ticket
    return {"ticket": new_ticket()}


@route("GET", "/api/stream", ticket=True)
def r_stream(h, m, q, body):
    cid = (q.get("client") or [""])[0] or "anon"
    sq = _queue.Queue(maxsize=20000)
    with _streams_lock:
        old = STREAMS.get(cid)
        STREAMS[cid] = {"q": sq, "docs": set()}
    if old:
        old["q"].put_nowait(("bye", {}))
    REGISTRY.hub.subscribe(sq)
    try:
        return _sse_queue(h, sq, [("ready", {"docs": REGISTRY.list()})])
    finally:
        REGISTRY.hub.unsubscribe(sq)
        with _streams_lock:
            st = STREAMS.get(cid)
            if st and st["q"] is sq:
                for did in st["docs"]:
                    d = REGISTRY.docs.get(did)
                    if d:
                        d.hub.unsubscribe(sq)
                del STREAMS[cid]


@route("POST", "/api/stream/subscribe")
def r_stream_sub(h, m, q, body):
    with _streams_lock:
        st = STREAMS.get(body.get("client"))
    if not st:
        raise ApiError(404, "No stream. Open /api/stream first.")
    doc = REGISTRY.get(body["docId"])
    since = int(body.get("since", -1))
    with doc.lock:  # replay + subscribe atomically: no gap, no duplicate
        evs = doc.events_since(since) if since >= 0 else []
        if evs is None:
            st["q"].put_nowait(("resync", {"docId": doc.id, "version": doc.version}))
        else:
            for e in evs:
                st["q"].put_nowait(("ops", e))
        doc.hub.subscribe(st["q"])
        st["docs"].add(doc.id)
    return {"version": doc.version}


@route("POST", "/api/stream/unsubscribe")
def r_stream_unsub(h, m, q, body):
    with _streams_lock:
        st = STREAMS.get(body.get("client"))
    if st:
        d = REGISTRY.docs.get(body.get("docId"))
        if d:
            d.hub.unsubscribe(st["q"])
        st["docs"].discard(body.get("docId"))
    return {"ok": True}


def _sse_queue(h, sq, initial):
    h.send_response(200)
    h.send_header("Content-Type", "text/event-stream; charset=utf-8")
    h.send_header("Cache-Control", "no-store")
    h.send_header("X-Accel-Buffering", "no")
    h.end_headers()
    h.close_connection = True
    Config.stream_clients += 1

    def send(event, data):
        h.wfile.write(f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False, default=str)}\n\n".encode("utf-8"))
        h.wfile.flush()

    try:
        h.wfile.write(b"retry: 1500\n\n")
        for event, data in initial:
            send(event, data)
        while True:
            try:
                event, data = sq.get(timeout=15)
            except _queue.Empty:
                h.wfile.write(b": ping\n\n")
                h.wfile.flush()
                Config.last_activity = time.time()
                continue
            if event == "bye":
                break
            send(event, data)
    except (BrokenPipeError, ConnectionResetError, OSError):
        pass
    finally:
        Config.stream_clients -= 1
        Config.last_activity = time.time()
    return None

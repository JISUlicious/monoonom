"""Document commands for xl.py (P2). Output is compact text by default, --json for raw results."""
from __future__ import annotations

import json
import os
import sys

from xl import EXIT_USAGE, CliError, connect, request, workspace_root


def in_root(args, p):
    """Relative paths are relative to the workspace root (not the CLI's current folder)."""
    p = os.path.expanduser(p)
    return p if os.path.isabs(p) else os.path.join(workspace_root(args.root), p)


def _read_text(args, path):
    p = path if os.path.exists(path) else in_root(args, path)
    try:
        with open(p, encoding="utf-8") as fh:
            return fh.read()
    except OSError as e:
        raise CliError(f"Cannot read the file: {path} ({e.strerror})", EXIT_USAGE)


def out_json(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=1, default=str))


def pick_doc(st, want):
    docs = request(st, "GET", "/api/docs")["docs"]
    if want:
        w = os.path.normcase(os.path.abspath(want)) if (os.sep in want or "/" in want) or want.endswith((".xlsx", ".xls", ".xlsm")) else None
        for d in docs:
            if d["docId"] == want or d["name"] == want or (w and d.get("path") and os.path.normcase(d["path"]) == w):
                return d
        raise CliError(f"No open document matches '{want}'. Check with `xl.py docs` or open it with `xl.py open PATH`.", EXIT_USAGE)
    if len(docs) == 1:
        return docs[0]
    if not docs:
        raise CliError("No document is open. Open one with `xl.py open PATH`.", EXIT_USAGE)
    raise CliError("Several documents are open. Choose one with --doc ID|name|path:\n" +
                   "\n".join(f"  {d['docId']}  {d['name']}" for d in docs), EXIT_USAGE)


def ctx(args):
    st = connect(workspace_root(args.root))
    return st, pick_doc(st, getattr(args, "doc", None))


def edit(args, edits, label=None):
    st, d = ctx(args)
    res = request(st, "POST", f"/api/docs/{d['docId']}/edit", {"edits": edits, "label": label})
    if getattr(args, "json", False):
        return out_json(res)
    batches = [b for b in res.get("batches", [res]) if b.get("batchId")]
    if not batches:
        print(f"v{res['version']}  no change  " + "; ".join(res.get("summary", [])))
        return 0
    for b in batches:
        print(f"v{b['version']}  batch={b['batchId']}  " + "; ".join(b.get("summary", [])))
    print("(undo: " + ", ".join(f"xl.py undo {b['batchId']}" for b in reversed(batches)) + ")")
    return 0


def literal(s):
    """CLI token -> JSON value: numbers/booleans stay strings so the server parses them like Excel input."""
    return s


def fmt_cell(v):
    if v is None:
        return ""
    if isinstance(v, float):
        v = float(f"{v:.15g}")
        return str(int(v)) if v.is_integer() else repr(v)
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return str(v).replace("\t", " ").replace("\n", "\\n")


# ------------------------------------------------------------------ commands

def cmd_ls(args):
    st = connect(workspace_root(args.root))
    res = request(st, "GET", "/api/fs" + (f"?path={_q(in_root(args, args.path))}" if args.path else ""))
    if args.json:
        return out_json(res)
    print(res["path"])
    for e in res["entries"]:
        print(f"  {'[dir] ' if e['dir'] else ''}{e['name']}" + ("" if e["dir"] else f"  ({e['size']:,} B)"))
    return 0


def _q(s):
    from urllib.parse import quote
    return quote(s)


def cmd_open(args):
    st = connect(workspace_root(args.root))
    res = request(st, "POST", "/api/docs", {"path": in_root(args, args.path), "actor": "agent"})
    if args.json:
        return out_json(res)
    print(f"Opened: {res['docId']}  {res['name']}  sheets: {', '.join(res['sheets'])}  (shown as a tab in the preview)")
    print("  Get the preview URL (with its token) from `xl.py url` and pass it on unchanged.")
    for w in res.get("warnings", []):
        print(f"  warning: {w}")
    return 0


def cmd_new(args):
    st = connect(workspace_root(args.root))
    res = request(st, "POST", "/api/docs", {"new": True, "name": args.name, "lang": args.lang, "actor": "agent"})
    print(f"New document: {res['docId']}  {res['name']}")
    return 0


def cmd_docs(args):
    st = connect(workspace_root(args.root))
    docs = request(st, "GET", "/api/docs")["docs"]
    if args.json:
        return out_json(docs)
    if not docs:
        print("No document is open.")
    for d in docs:
        print(f"{d['docId']}  {d['name']}  v{d['version']}{'  *modified*' if d['dirty'] else ''}  windows {d['clients']}  {d.get('path') or '(not saved)'}")
    return 0


def cmd_describe(args):
    st, d = ctx(args)
    res = request(st, "GET", f"/api/docs/{d['docId']}/describe")
    if args.json:
        return out_json(res)
    print(f"{res['name']}  ({res['path'] or 'not saved'})  v{res['version']}{'  *modified*' if res['dirty'] else ''}"
          f"  calc: {res['calc']['status']}{' (' + res['calc']['source'] + ')' if res['calc'].get('source') else ''}")
    if res["calc"].get("note"):
        print(f"  warning: {res['calc']['note']}")
    for sh in res["sheets"]:
        print(f"\n[{sh['name']}]{' (active)' if sh['active'] else ''}{' (hidden)' if sh['state'] != 'visible' else ''}"
              f"  used {sh['used'] or 'none'}  cells {sh['cells']}  formulas {sh['formulas']}"
              f"{'  freeze ' + sh['freeze'] if sh['freeze'] else ''}{'  filter ' + sh['autoFilter'] if sh['autoFilter'] else ''}"
              f"{'  merges ' + str(sh['merges']) if sh['merges'] else ''}")
        if sh.get("unsupported"):
            u = sh["unsupported"]
            names = sorted({n for x in u for n in x["functions"]})
            kept = sum(1 for x in u if x["kept"])
            print(f"  warning: functions this app cannot compute: {', '.join(names)} — {len(u)} cells "
                  f"({', '.join(x['cell'] for x in u[:8])}{' …' if len(u) > 8 else ''}). "
                  f"values are the file's last results; {kept} keep their original formula when saved. Edited ones are computed only by Excel.")
        for t in sh.get("tables", [])[:5]:
            print(f"  table {t['range']}: header row {t['headerRow']}" + (f", total rows {t['totalRows']}" if t.get("totalRows") else "")
                  + "  [" + " | ".join(fmt_cell(v) for v in t["header"][:12]) + "]")
        if sh.get("header") is not None:
            print(f"  header (row {sh['headerRow']}): " + " | ".join(fmt_cell(v) for v in sh["header"]))
            print("  column types: " + ", ".join(f"{k}={v}" for k, v in sh["columnTypes"].items()))
        for p in sh.get("pivots", []):
            print(f"  pivot table '{p['name']}' {p['range']} (read-only): source {p['source']}  rows {', '.join(p['rows']) or '-'}"
                  f"  columns {', '.join(p['cols']) or '-'}  values {', '.join(p['values']) or '-'}" + (f"  filters {', '.join(p['filters'])}" if p["filters"] else ""))
        for ch in sh["charts"]:
            print(f"  chart {ch['id']}: {ch['type']} '{ch['title'] or ''}' @{ch['at']}  data {', '.join(x or '' for x in ch['data'])}"
                  + (f"  (pivot chart of '{ch['pivot']}')" if ch.get("pivot") else ""))
    if res["definedNames"]:
        print("\nDefined names: " + ", ".join(f"{k}={v}" for k, v in res["definedNames"].items()))
    return 0


def cmd_read(args):
    st, d = ctx(args)
    from urllib.parse import urlencode
    qs = urlencode({"ref": args.range, "what": args.what, "max": args.max})
    res = request(st, "GET", f"/api/docs/{d['docId']}/range?{qs}")
    if args.json:
        return out_json(res)
    calc = res["calc"]
    print(f"{res['sheet']}!{res['ref'] or args.range}  ({args.what}"
          f"{', calc ' + calc['status'] if calc['status'] not in ('none', 'skipped') else ''})")
    if calc.get("note"):
        print(f"warning: {calc['note']}")
    if not res["ref"]:
        print("(empty range)")
        return 0
    from server.a1 import col_name, parse_range
    r1, c1, _, _ = parse_range(res["ref"].split(":")[0])
    width = max((len(r) for r in res["rows"]), default=0)
    print("\t" + "\t".join(col_name(c1 + j) for j in range(width)))
    for i, row in enumerate(res["rows"]):
        print(f"{r1 + i + 1}\t" + "\t".join(fmt_cell(v) for v in row))
    if res.get("truncated"):
        print(f"... truncated. Rest: {res['truncated']}")
    return 0


def _values_from(args):
    if args.json_values is not None:
        return json.loads(args.json_values)
    if args.tsv is not None:
        text = sys.stdin.read() if args.tsv == "-" else _read_text(args, args.tsv)
        return [line.split("\t") for line in text.rstrip("\n").split("\n")]
    if not args.values:
        raise CliError("Give values: positional arguments, --data '[[...]]' or --tsv FILE|-", EXIT_USAGE)
    vals = [literal(v) for v in args.values]
    return [[v] for v in vals] if args.down else [vals]


def cmd_write(args):
    return edit(args, [{"op": "set", "range": args.range, "values": _values_from(args), "parse": not args.raw}])


def cmd_fill(args):
    return edit(args, [{"op": "set", "range": args.range, "values": [[args.value]], "parse": not args.raw}])


def cmd_clear(args):
    return edit(args, [{"op": "clear", "range": args.range, "what": args.what}])


def _kv(items):
    out = {}
    for it in items or []:
        k, _, v = it.partition("=")
        if not _:
            raise CliError(f"Formats are key=value: {it}", EXIT_USAGE)
        lv = v.lower()
        out[k] = True if lv == "true" else False if lv == "false" else (float(v) if _is_num(v) and k in ("fs", "ind", "rot") else v)
    return out


def _is_num(v):
    try:
        float(v)
        return True
    except ValueError:
        return False


def cmd_format(args):
    e = {"op": "format", "range": args.range, "set": _kv(args.set), "unset": args.unset or []}
    if args.border:
        kind, _, rest = args.border.partition(":")
        style, _, color = rest.partition(":")
        e["border"] = {"kind": kind, **({"style": style} if style else {}), **({"color": color} if color else {})}
    return edit(args, [e])


def cmd_merge(args):
    e = {"op": args.cmd, "range": args.range}
    if args.cmd == "merge":
        e.update(across=args.across, center=not args.no_center)
    return edit(args, [e])


def cmd_width(args):
    return edit(args, [{"op": "colWidth", "range": args.cols, "width": args.chars}])


def cmd_height(args):
    return edit(args, [{"op": "rowHeight", "range": args.rows, "height": args.points}])


def cmd_hide(args):
    return edit(args, [{"op": args.cmd, "range": args.range}])


def cmd_freeze(args):
    return edit(args, [{"op": "freeze", "sheet": args.sheet, "cell": None if args.cell.lower() in ("none", "off") else args.cell}])


def cmd_filter(args):
    if args.range.lower() == "clear":
        return edit(args, [{"op": "filterClear", "sheet": args.sheet}])
    if args.range.lower() == "reapply":
        return edit(args, [{"op": "filterReapply", "sheet": args.sheet}])
    if args.column:
        e = {"op": "filterColumn", "sheet": args.sheet, "column": args.column}
        if args.values is not None:
            e["values"] = [v for v in args.values.split("|")] if args.values else []
            e["blank"] = args.blank
        custom = []
        for cond in args.where or []:
            m = __import__("re").match(r"\s*(>=|<=|<>|!=|==|=|>|<|contains|notContains|beginsWith|endsWith)\s*(.*)$", cond)
            if not m:
                raise CliError(f"Condition format: '>150', 'contains Seoul', …: {cond}", EXIT_USAGE)
            custom.append({"op": m.group(1), "val": m.group(2).strip()})
        if custom:
            e["custom"], e["and"] = custom, not args.any
        return edit(args, [e])
    return edit(args, [{"op": "filter", "sheet": args.sheet, "range": None if args.range.lower() == "off" else args.range}])


def cmd_sheet(args):
    if args.action == "add":
        return edit(args, [{"op": "addSheet", "name": args.names[0] if args.names else None}])
    if args.action == "rename":
        if len(args.names) != 2:
            raise CliError("Usage: sheet rename OLD NEW", EXIT_USAGE)
        return edit(args, [{"op": "renameSheet", "sheet": args.names[0], "name": args.names[1]}])
    if args.action == "rm":
        return edit(args, [{"op": "deleteSheet", "sheet": args.names[0]}])
    if args.action == "move":
        return edit(args, [{"op": "moveSheet", "sheet": args.names[0], "to": int(args.names[1]) - 1 if args.names[1].isdigit() else args.names[1]}])
    if args.action == "copy":
        return edit(args, [{"op": "duplicateSheet", "sheet": args.names[0], "name": args.names[1] if len(args.names) > 1 else None}])
    if args.action in ("hide", "unhide"):
        return edit(args, [{"op": args.action + "Sheet", "sheet": args.names[0]}])
    if args.action == "color":
        return edit(args, [{"op": "tabColor", "sheet": args.names[0], "color": None if len(args.names) < 2 or args.names[1] == "none" else args.names[1]}])
    if args.action == "gridlines":
        return edit(args, [{"op": "gridlines", "sheet": args.names[0], "show": len(args.names) < 2 or args.names[1] not in ("off", "false", "0")}])
    raise CliError("sheet add|rename|rm|move|copy|hide|unhide|color|gridlines", EXIT_USAGE)


def cmd_pivot(args):
    st, d = ctx(args)
    if args.action == "ls":
        res = request(st, "GET", f"/api/docs/{d['docId']}/describe")
        if args.json:
            return out_json([dict(p, sheet=sh["name"]) for sh in res["sheets"] for p in sh.get("pivots", [])])
        found = False
        for sh in res["sheets"]:
            for p in sh.get("pivots", []):
                found = True
                print(f"{sh['name']}  '{p['name']}'  {p['range']}  source {p['source']}  rows {', '.join(p['rows']) or '-'}  "
                      f"columns {', '.join(p['cols']) or '-'}  values {', '.join(p['values']) or '-'}" + (f"  filters {', '.join(p['filters'])}" if p["filters"] else ""))
        if not found:
            print("No pivot tables.")
        return 0
    if args.action == "add":
        split = lambda s: [x.strip() for x in (s or "").split(",") if x.strip()]  # noqa: E731
        body = {"source": args.name, "rows": split(args.rows), "cols": split(args.cols), "filters": split(args.filters),
                "values": [dict(zip(("field", "fn"), (v.split(":", 1) + ["sum"])[:2])) for v in split(args.values)],
                "at": args.at, "name": args.title, "lang": args.lang}
        res = request(st, "POST", f"/api/docs/{d['docId']}/pivots", body)
        if args.json:
            return out_json(res)
        print(f"v{res['version']}  batch={res['batchId']}  pivot table '{res['name']}' → {res['sheet']}!{res['range']} (source {res['source']})")
        print("  After saving, Excel rebuilds it from the source when it opens the file (refresh).")
        print(f"(undo: xl.py undo {res['batchId']})")
        return 0
    res = request(st, "POST", f"/api/docs/{d['docId']}/pivots/refresh", {"name": args.name, "sheet": args.sheet})
    if args.json:
        return out_json(res)
    for p in res["pivots"]:
        print(f"{p['sheet']}  '{p['name']}'  {p['range']}: {p['changed']} values changed"
              + (f"; pivot charts {', '.join(p['charts'])} show them" if p.get("charts") else ""))
    for n in res.get("notes", []):
        print("  note: " + n)
    if res.get("batchId"):
        print(f"v{res['version']}  batch={res['batchId']}  (undo: xl.py undo {res['batchId']})")
    return 0


def cmd_chart(args):
    if args.action == "add" and args.pivot:
        e = {"op": "chart", "pivot": args.pivot, "pivotSheet": args.sheet, "type": args.type, "title": args.title,
             "at": args.at, "legend": args.legend, "width": args.width, "height": args.height,
             "xTitle": args.x_title, "yTitle": args.y_title}
        return edit(args, [{k: v for k, v in e.items() if v is not None}])
    if args.action == "add":
        e = {"op": "chart", "data": args.target, "type": args.type, "title": args.title, "at": args.at,
             "sheet": args.sheet, "legend": args.legend}
        if args.width:
            e["width"] = args.width
        if args.height:
            e["height"] = args.height
        return edit(args, [{k: v for k, v in e.items() if v is not None}])
    if args.action == "rm":
        return edit(args, [{"op": "deleteChart", "id": args.target, "sheet": args.sheet}])
    if args.action == "set":
        st = {k: v for k, v in (("type", args.type_set), ("title", args.title), ("at", args.at), ("legend", args.legend),
                                ("width", args.width), ("height", args.height), ("data", args.data), ("xTitle", args.x_title),
                                ("yTitle", args.y_title),
                                ("varyColors", None if args.vary_colors is None else args.vary_colors == "on"),
                                ("switchRowCol", True if args.switch else None)) if v is not None}
        return edit(args, [{"op": "chartUpdate", "id": args.target, "sheet": args.sheet, "set": st}])
    st, d = ctx(args)
    res = request(st, "GET", f"/api/docs/{d['docId']}/describe")
    for sh in res["sheets"]:
        for ch in sh["charts"]:
            pv = f"  pivot chart of '{ch['pivot']}'" if ch.get("pivot") else ""
            print(f"{sh['name']}  {ch['id']}  {ch['type']}  '{ch['title'] or ''}'  @{ch['at']}{pv}")
            for i, s in enumerate(ch.get("series", [])):
                name = s.get("nameRef") or (f"'{s['name']}'" if s.get("name") else "-")
                kind = f" ({s['type']})" if s.get("type") else ""
                print(f"    series {i + 1}{kind}: name {name}  X {s.get('x') or '-'}  Y {s.get('y') or '-'}")
    return 0


def cmd_rows_cols(args):
    axis = "Rows" if args.cmd == "rows" else "Cols"
    op = ("insert" if args.action == "insert" else "delete") + axis
    e = {"op": op, "at": args.at, "sheet": args.sheet}
    if args.count:
        e["count"] = args.count
    return edit(args, [{k: v for k, v in e.items() if v is not None}])


def cmd_sort(args):
    e = {"op": "sort", "range": args.range, "by": args.by or []}
    if args.header is not None:
        e["header"] = args.header
    return edit(args, [e])


def cmd_series(args):
    return edit(args, [{"op": "fillSeries", "source": args.source, "target": args.target, "sheet": args.sheet}])


def cmd_copy(args):
    e = {"op": "move" if args.cmd == "move" else "copy", "from": args.src, "to": args.dst, "sheet": args.sheet}
    if args.cmd == "copy":
        e.update(what=args.what, transpose=args.transpose)
    return edit(args, [e])


def cmd_cells(args):
    return edit(args, [{"op": "insertCells" if args.action == "insert" else "deleteCells", "range": args.range,
                        "shift": args.shift, "sheet": args.sheet}])


def cmd_dedupe(args):
    e = {"op": "removeDuplicates", "range": args.range, "sheet": args.sheet}
    if args.columns:
        e["columns"] = args.columns.split(",")
    if args.header is not None:
        e["header"] = args.header
    return edit(args, [e])


def cmd_find(args):
    st, d = ctx(args)
    from urllib.parse import urlencode
    qs = {"q": args.text, "in": args.where, "case": int(args.match_case), "whole": int(args.whole), "max": args.max}
    if args.sheet:
        qs["sheet"] = args.sheet
    res = request(st, "GET", f"/api/docs/{d['docId']}/find?{urlencode(qs)}")
    if args.json:
        return out_json(res)
    if not res["matches"]:
        print("Nothing found.")
    for mt in res["matches"]:
        print(f"{mt['sheet']}!{mt['cell']}\t{fmt_cell(mt['value'])}")
    if res["truncated"]:
        print(f"... stopped at {args.max} (raise with --max)")
    return 0


def cmd_replace(args):
    return edit(args, [{"op": "replace", "find": args.text, "with": args.replacement, "in": args.where,
                        "matchCase": args.match_case, "whole": args.whole, "all": args.all, "sheet": args.sheet}])


def cmd_comment(args):
    return edit(args, [{"op": "comment", "cell": args.cell, "text": args.text}])


def cmd_link(args):
    return edit(args, [{"op": "link", "cell": args.cell, "url": args.url, "text": args.text}])


def cmd_autofit(args):
    return edit(args, [{"op": "autofit", "range": args.range, "sheet": args.sheet}])


def cmd_copy_format(args):
    return edit(args, [{"op": "copyFormat", "from": args.src, "to": args.dst, "sheet": args.sheet}])


def cmd_decimals(args):
    return edit(args, [{"op": "decimals", "range": args.range, "delta": args.delta}])


def cmd_goto(args):
    st, d = ctx(args)
    res = request(st, "POST", f"/api/docs/{d['docId']}/goto", {"range": args.range, "sheet": args.sheet})
    print(f"Moved the preview to {res['sheet']}!{res['range']}" + ("" if res["clients"] else " (no preview window open)"))
    return 0


def cmd_view(args):
    st, d = ctx(args)
    body = {}
    if args.zoom:
        body["zoom"] = args.zoom / 100
    if args.formulas is not None:
        body["showFormulas"] = args.formulas
    if args.follow:
        body["follow"] = args.follow == "on"
    res = request(st, "POST", f"/api/docs/{d['docId']}/view", body)
    print("Preview view settings applied" + ("" if res["clients"] else " (no preview window open)"))
    return 0


def cmd_apply(args):
    text = sys.stdin.read() if args.file == "-" else _read_text(args, args.file)
    edits = json.loads(text)
    return edit(args, edits if isinstance(edits, list) else edits.get("edits", [edits]), args.label)


def cmd_changes(args):
    st, d = ctx(args)
    q = f"?since={args.since}" + (f"&actor={args.actor}" if args.actor else "")
    res = request(st, "GET", f"/api/docs/{d['docId']}/changes{q}")
    if args.json:
        return out_json(res)
    if not res["changes"]:
        print(f"No changes since v{args.since} (now v{res['version']})")
    for c in res["changes"]:
        who = {"user": "user", "agent": "agent", "system": "system"}.get(c["actor"], c["actor"])
        print(f"v{c['v']} {c['time']} {who}: {c['label'] or ''}  [{c['batchId']}]")
        for x in c["cells"] + c["meta"]:
            print(f"    {x}")
    print(f"now v{res['version']}{'  *modified*' if res['dirty'] else ''}")
    return 0


def cmd_fonts(args):
    """Installed fonts (read on this computer; no server needed), optionally filtered."""
    from server.fonts import list_fonts
    res = list_fonts()
    q = (args.search or "").lower()
    fonts = [f for f in res["fonts"] if not q or any(q in n.lower() for n in f["names"])]
    if args.json:
        return out_json({"fonts": fonts, "files": res["files"]})
    for f in fonts:
        print(f["name"] + (f"  ({', '.join(n for n in f['names'] if n != f['name'])})" if len(f["names"]) > 1 else ""))
    print(f"— {len(fonts)} fonts" + (f" containing '{args.search}'" if q else "") + f", read from {res['files']} files")
    return 0


def cmd_selection(args):
    st, d = ctx(args)
    res = request(st, "GET", f"/api/docs/{d['docId']}/selection")["selection"]
    if args.json:
        return out_json(res)
    if not res:
        print("No selection in the preview (no window open).")
    for p in res:
        print(f"{p['sheet']}!{p['range']}  (active cell {p['active']}, {p['secondsAgo']} s ago)")
    return 0


def cmd_undo(args):
    st, d = ctx(args)
    res = request(st, "POST", f"/api/docs/{d['docId']}/undo", {"batchId": args.batch, "force": args.force})
    print(f"Undone → v{res['version']}")
    return 0


def cmd_save(args):
    st, d = ctx(args)
    body = {"overwrite": args.overwrite, "force": args.force, "backup": not args.no_backup}
    if args.as_path:
        body["path"] = in_root(args, args.as_path)
    if args.format:
        body["format"] = args.format
    res = request(st, "POST", f"/api/docs/{d['docId']}/save", body)
    print(f"Saved: {res['path']}" + (f"  (backup: {res['backup']})" if res.get("backup") else ""))
    for w in res.get("warnings", []):
        print(f"  warning: {w}")
    return 0


def cmd_close(args):
    st, d = ctx(args)
    request(st, "DELETE", f"/api/docs/{d['docId']}", {"force": args.force})
    print(f"Closed: {d['name']}")
    return 0


# ------------------------------------------------------------------ registration

def register(sub):
    def add(name, fn, help_, doc=True, json_=True):
        p = sub.add_parser(name, help=help_)
        if doc:
            p.add_argument("--doc", help="docId, file name or path (omit when one document is open)")
        if json_:
            p.add_argument("--json", action="store_true", help="print the result as JSON")
        p.set_defaults(fn=fn)
        return p

    p = add("ls", cmd_ls, "list Excel files in the workspace", doc=False); p.add_argument("path", nargs="?")
    p = add("open", cmd_open, "open a file (shown as a tab in the preview)", doc=False); p.add_argument("path")
    p = add("new", cmd_new, "new workbook", doc=False, json_=False); p.add_argument("--name")
    p.add_argument("--lang", choices=["en", "ko"], default="en", help="Excel UI language for the default name and body font (en: Book1, Calibri; ko: Korean Excel's)")
    add("docs", cmd_docs, "list open documents", doc=False)
    add("describe", cmd_describe, "summary: sheets, used ranges, headers, column types, charts")
    p = add("read", cmd_read, "read a range (e.g. 'Sheet1!A1:D20')"); p.add_argument("range")
    p.add_argument("--what", choices=["values", "text", "raw", "formulas", "all"], default="values", help="values: dates as YYYY-MM-DD, text: as displayed (number format applied), raw: serial numbers, formulas, all"); p.add_argument("--max", type=int, default=200)
    p = add("write", cmd_write, "write values / formulas (a leading = makes a formula)"); p.add_argument("range"); p.add_argument("values", nargs="*")
    p.add_argument("--data", dest="json_values", metavar="JSON", help="2-D JSON array, e.g. '[[1,2],[3,4]]'"); p.add_argument("--tsv", metavar="FILE|-")
    p.add_argument("--down", action="store_true", help="write positional values downwards"); p.add_argument("--raw", action="store_true", help="do not parse strings")
    p = add("fill", cmd_fill, "fill a range with one value / formula (relative references shift)"); p.add_argument("range"); p.add_argument("value"); p.add_argument("--raw", action="store_true")
    p = add("clear", cmd_clear, "clear"); p.add_argument("range"); p.add_argument("--what", choices=["contents", "formats", "comments", "all"], default="contents")
    p = add("format", cmd_format, "format (e.g. b=true bg=#FFF2CC nf=#,##0)"); p.add_argument("range"); p.add_argument("set", nargs="*")
    p.add_argument("--unset", action="append"); p.add_argument("--border", help="all|outline|inside|top|bottom|left|right|none[:style[:#color]]")
    for name in ("merge", "unmerge"):
        p = add(name, cmd_merge, "merge cells (--across: per row)" if name == "merge" else "unmerge"); p.add_argument("range")
        if name == "merge":
            p.add_argument("--across", action="store_true"); p.add_argument("--no-center", action="store_true")
    p = add("width", cmd_width, "column width (characters)"); p.add_argument("cols"); p.add_argument("chars", type=float)
    p = add("height", cmd_height, "row height (points)"); p.add_argument("rows"); p.add_argument("points", type=float)
    for name in ("hide", "unhide"):
        p = add(name, cmd_hide, "hide rows / columns" if name == "hide" else "unhide"); p.add_argument("range", help="e.g. B:C or 3:5")
    p = add("freeze", cmd_freeze, "freeze panes (e.g. B2, none)"); p.add_argument("cell"); p.add_argument("--sheet")
    p = add("filter", cmd_filter, "autofilter: RANGE (on) | off | clear | reapply | col C --where '>150' | col C --values 'a|b'")
    p.add_argument("range", help="range, off, clear, reapply or col"); p.add_argument("column", nargs="?", help="column letter with col")
    p.add_argument("--values", help="values to show (| separated, as displayed)"); p.add_argument("--blank", action="store_true", help="also show blanks")
    p.add_argument("--where", action="append", help="condition: '>150', '<=10', 'contains Seoul', 'beginsWith A' (repeatable)"); p.add_argument("--any", action="store_true", help="OR the conditions")
    p.add_argument("--sheet")
    p = add("sheet", cmd_sheet, "sheets: add NAME | rename OLD NEW | rm NAME | move NAME POS | copy NAME [NEW] | hide|unhide NAME | color NAME #RRGGBB|none | gridlines NAME on|off")
    p.add_argument("action", choices=["add", "rename", "rm", "move", "copy", "hide", "unhide", "color", "gridlines"]); p.add_argument("names", nargs="*")
    p = add("pivot", cmd_pivot, "pivot tables: ls | refresh [NAME] | add SOURCE --rows F --values F:sum [--cols F] [--filters F] [--at 'S!A3']")
    p.add_argument("action", choices=["ls", "refresh", "add"]); p.add_argument("name", nargs="?", help="refresh: pivot name, add: source range")
    p.add_argument("--sheet"); p.add_argument("--rows", help="add: row fields (comma separated)"); p.add_argument("--cols", help="add: column field (one)")
    p.add_argument("--values", help="add: value field:function (sum, count, average, max, min, product, countNums, stdDev, stdDevp, var, varp), comma separated")
    p.add_argument("--filters", help="add: report filter fields"); p.add_argument("--at", help="add: top-left cell, e.g. 'Pivot!A3' (a missing sheet is created)")
    p.add_argument("--title", help="add: pivot name")
    p.add_argument("--lang", choices=["en", "ko"], default="en", help="add: label language (Grand Total, Sum of …), as Excel in that UI language writes them")
    p = add("chart", cmd_chart, "charts: add DATA | add --pivot NAME | set ID --title … | rm ID | ls"); p.add_argument("action", choices=["add", "set", "rm", "ls"]); p.add_argument("target", nargs="?")
    p.add_argument("--type", default="bar", help="bar, barH, barStacked, line, area, pie, doughnut, scatter, radar")
    p.add_argument("--to-type", dest="type_set", help="set: change the chart type"); p.add_argument("--data", help="set: change the data range")
    p.add_argument("--x-title"); p.add_argument("--y-title")
    p.add_argument("--vary-colors", choices=["on", "off"], help="one-series chart: a color per category (legend lists the categories)")
    p.add_argument("--switch", action="store_true", help="set: switch row/column (swap series and categories)")
    p.add_argument("--title"); p.add_argument("--at", help="top-left cell"); p.add_argument("--sheet"); p.add_argument("--legend")
    p.add_argument("--width", type=int); p.add_argument("--height", type=int)
    p.add_argument("--pivot", help="add: a pivot chart of this pivot table (its data, linked to it); --sheet: the pivot's sheet if names repeat")
    for name in ("rows", "cols"):
        p = add(name, cmd_rows_cols, "insert / delete rows (e.g. rows insert 5 --count 2, rows delete 5:7)" if name == "rows" else "insert / delete columns (e.g. cols insert C, cols delete C:D)")
        p.add_argument("action", choices=["insert", "delete"]); p.add_argument("at", help="5 / 5:7" if name == "rows" else "C / C:D")
        p.add_argument("--count", type=int); p.add_argument("--sheet")
    p = add("sort", cmd_sort, "sort a range (e.g. sort A1:F50 --by C:desc --by A)"); p.add_argument("range")
    p.add_argument("--by", action="append", help="column[:asc|desc], repeatable"); p.add_argument("--header", dest="header", action="store_true", default=None)
    p.add_argument("--no-header", dest="header", action="store_false")
    p = add("series", cmd_series, "fill a series (fill handle): series A2:A3 A2:A20"); p.add_argument("source"); p.add_argument("target"); p.add_argument("--sheet")
    p = add("copy", cmd_copy, "copy and paste: copy A1:C5 E1 [--what values|formats|formulas] [--transpose]"); p.add_argument("src"); p.add_argument("dst")
    p.add_argument("--what", choices=["all", "values", "formats", "formulas"], default="all"); p.add_argument("--transpose", action="store_true"); p.add_argument("--sheet")
    p = add("move", cmd_copy, "cut and paste (move): move A1:C5 E1"); p.add_argument("src"); p.add_argument("dst"); p.add_argument("--sheet")
    p = add("cells", cmd_cells, "insert / delete cells (shift): cells insert B2:C3 --shift down"); p.add_argument("action", choices=["insert", "delete"]); p.add_argument("range")
    p.add_argument("--shift", choices=["right", "down", "left", "up"]); p.add_argument("--sheet")
    p = add("dedupe", cmd_dedupe, "remove duplicates"); p.add_argument("range"); p.add_argument("--columns", help="key columns, e.g. A,C"); p.add_argument("--sheet")
    p.add_argument("--header", dest="header", action="store_true", default=None); p.add_argument("--no-header", dest="header", action="store_false")
    for name in ("find", "replace"):
        p = add(name, cmd_find if name == "find" else cmd_replace, "find (* ? wildcards)" if name == "find" else "replace")
        p.add_argument("text")
        if name == "replace":
            p.add_argument("replacement"); p.add_argument("--all", action="store_true", help="all sheets")
        else:
            p.add_argument("--max", type=int, default=200)
        p.add_argument("--in", dest="where", choices=["values", "formulas"], default="values"); p.add_argument("--match-case", action="store_true")
        p.add_argument("--whole", action="store_true", help="match the whole cell"); p.add_argument("--sheet")
    p = add("comment", cmd_comment, "note: comment B2 'text' (no text deletes it)"); p.add_argument("cell"); p.add_argument("text", nargs="?")
    p = add("link", cmd_link, "hyperlink: link B2 https://… [--text label]"); p.add_argument("cell"); p.add_argument("url", nargs="?"); p.add_argument("--text")
    p = add("autofit", cmd_autofit, "autofit column widths / row heights: autofit A:D | autofit 1:10"); p.add_argument("range"); p.add_argument("--sheet")
    p = add("format-copy", cmd_copy_format, "copy formats (format painter): format-copy A1:B1 D1:H1"); p.add_argument("src"); p.add_argument("dst"); p.add_argument("--sheet")
    p = add("decimals", cmd_decimals, "increase / decrease decimals: decimals B2:B9 +1"); p.add_argument("range"); p.add_argument("delta", type=int)
    p = add("goto", cmd_goto, "move the user's preview to a sheet / range"); p.add_argument("range"); p.add_argument("--sheet")
    p = add("view", cmd_view, "preview display: --zoom 150, --formulas/--no-formulas", json_=False); p.add_argument("--zoom", type=int)
    p.add_argument("--formulas", dest="formulas", action="store_true", default=None); p.add_argument("--no-formulas", dest="formulas", action="store_false")
    p.add_argument("--follow", choices=["on", "off"], help="whether the preview follows where the agent works")
    p = add("apply", cmd_apply, "apply a list of edits (JSON file / -) in one batch"); p.add_argument("file"); p.add_argument("--label")
    p = add("changes", cmd_changes, "change history (user / agent)"); p.add_argument("--since", type=int, default=0); p.add_argument("--actor", choices=["user", "agent"])
    add("selection", cmd_selection, "the range the user has selected in the preview")
    p = add("fonts", cmd_fonts, "fonts installed on this computer (names for format ff=)", doc=False); p.add_argument("--search", help="text in the name")
    p = add("undo", cmd_undo, "undo an agent batch", json_=False); p.add_argument("batch"); p.add_argument("--force", action="store_true")
    p = add("save", cmd_save, "save (Python writes the file; backup before overwriting)", json_=False); p.add_argument("--as", dest="as_path")
    p.add_argument("--format", choices=["xlsx", "xlsm"]); p.add_argument("--overwrite", action="store_true")
    p.add_argument("--force", action="store_true", help="overwrite even if the file changed on disk"); p.add_argument("--no-backup", action="store_true")
    p = add("close", cmd_close, "close the document", json_=False); p.add_argument("--force", action="store_true")

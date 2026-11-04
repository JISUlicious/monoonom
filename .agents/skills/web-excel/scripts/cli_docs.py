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
        raise CliError(f"파일을 읽을 수 없습니다: {path} ({e.strerror})", EXIT_USAGE)


def out_json(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=1, default=str))


def pick_doc(st, want):
    docs = request(st, "GET", "/api/docs")["docs"]
    if want:
        w = os.path.normcase(os.path.abspath(want)) if (os.sep in want or "/" in want) or want.endswith((".xlsx", ".xls", ".xlsm")) else None
        for d in docs:
            if d["docId"] == want or d["name"] == want or (w and d.get("path") and os.path.normcase(d["path"]) == w):
                return d
        raise CliError(f"열린 문서 중 '{want}'이(가) 없습니다. `xl.py docs`로 확인하거나 `xl.py open PATH`로 여세요.", EXIT_USAGE)
    if len(docs) == 1:
        return docs[0]
    if not docs:
        raise CliError("열린 문서가 없습니다. `xl.py open PATH`로 여세요.", EXIT_USAGE)
    raise CliError("문서가 여러 개 열려 있습니다. --doc ID|이름|경로 로 지정하세요:\n" +
                   "\n".join(f"  {d['docId']}  {d['name']}" for d in docs), EXIT_USAGE)


def ctx(args):
    st = connect(workspace_root(args.root))
    return st, pick_doc(st, getattr(args, "doc", None))


def edit(args, edits, label=None):
    st, d = ctx(args)
    res = request(st, "POST", f"/api/docs/{d['docId']}/edit", {"edits": edits, "label": label})
    if getattr(args, "json", False):
        return out_json(res)
    print(f"v{res['version']}  batch={res.get('batchId')}  " + "; ".join(res.get("summary", [])))
    print(f"(되돌리기: xl.py undo {res.get('batchId')})")
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
        print(f"  {'[폴더] ' if e['dir'] else ''}{e['name']}" + ("" if e["dir"] else f"  ({e['size']:,} B)"))
    return 0


def _q(s):
    from urllib.parse import quote
    return quote(s)


def cmd_open(args):
    st = connect(workspace_root(args.root))
    res = request(st, "POST", "/api/docs", {"path": in_root(args, args.path)})
    if args.json:
        return out_json(res)
    print(f"열림: {res['docId']}  {res['name']}  시트: {', '.join(res['sheets'])}  (미리보기에 탭으로 표시됨)")
    for w in res.get("warnings", []):
        print(f"  주의: {w}")
    return 0


def cmd_new(args):
    st = connect(workspace_root(args.root))
    res = request(st, "POST", "/api/docs", {"new": True, "name": args.name})
    print(f"새 문서: {res['docId']}  {res['name']}")
    return 0


def cmd_docs(args):
    st = connect(workspace_root(args.root))
    docs = request(st, "GET", "/api/docs")["docs"]
    if args.json:
        return out_json(docs)
    if not docs:
        print("열린 문서가 없습니다.")
    for d in docs:
        print(f"{d['docId']}  {d['name']}  v{d['version']}{'  *수정됨*' if d['dirty'] else ''}  창 {d['clients']}개  {d.get('path') or '(저장 안 됨)'}")
    return 0


def cmd_describe(args):
    st, d = ctx(args)
    res = request(st, "GET", f"/api/docs/{d['docId']}/describe")
    if args.json:
        return out_json(res)
    print(f"{res['name']}  ({res['path'] or '저장 안 됨'})  v{res['version']}{'  *수정됨*' if res['dirty'] else ''}"
          f"  계산: {res['calc']['status']}{' (' + res['calc']['source'] + ')' if res['calc'].get('source') else ''}")
    if res["calc"].get("note"):
        print(f"  주의: {res['calc']['note']}")
    for sh in res["sheets"]:
        print(f"\n[{sh['name']}]{' (활성)' if sh['active'] else ''}{' (숨김)' if sh['state'] != 'visible' else ''}"
              f"  사용 범위 {sh['used'] or '없음'}  셀 {sh['cells']}  수식 {sh['formulas']}"
              f"{'  틀 고정 ' + sh['freeze'] if sh['freeze'] else ''}{'  필터 ' + sh['autoFilter'] if sh['autoFilter'] else ''}"
              f"{'  병합 ' + str(sh['merges']) if sh['merges'] else ''}")
        if sh.get("header") is not None:
            print(f"  머리글(행 {sh['headerRow']}): " + " | ".join(fmt_cell(v) for v in sh["header"]))
            print("  열 유형: " + ", ".join(f"{k}={v}" for k, v in sh["columnTypes"].items()))
        for ch in sh["charts"]:
            print(f"  차트 {ch['id']}: {ch['type']} '{ch['title'] or ''}' @{ch['at']}  데이터 {', '.join(x or '' for x in ch['data'])}")
    if res["definedNames"]:
        print("\n이름 정의: " + ", ".join(f"{k}={v}" for k, v in res["definedNames"].items()))
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
          f"{', 계산 ' + calc['status'] if calc['status'] not in ('none', 'skipped') else ''})")
    if calc.get("note"):
        print(f"주의: {calc['note']}")
    if not res["ref"]:
        print("(빈 범위)")
        return 0
    from server.a1 import col_name, parse_range
    r1, c1, _, _ = parse_range(res["ref"].split(":")[0])
    width = max((len(r) for r in res["rows"]), default=0)
    print("\t" + "\t".join(col_name(c1 + j) for j in range(width)))
    for i, row in enumerate(res["rows"]):
        print(f"{r1 + i + 1}\t" + "\t".join(fmt_cell(v) for v in row))
    if res.get("truncated"):
        print(f"... 잘림. 나머지: {res['truncated']}")
    return 0


def _values_from(args):
    if args.json_values is not None:
        return json.loads(args.json_values)
    if args.tsv is not None:
        text = sys.stdin.read() if args.tsv == "-" else _read_text(args, args.tsv)
        return [line.split("\t") for line in text.rstrip("\n").split("\n")]
    if not args.values:
        raise CliError("값을 주세요: 위치 인자, --data '[[...]]', 또는 --tsv FILE|-", EXIT_USAGE)
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
            raise CliError(f"서식은 key=value 형식이어야 합니다: {it}", EXIT_USAGE)
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
    return edit(args, [{"op": args.cmd, "range": args.range}])


def cmd_width(args):
    return edit(args, [{"op": "colWidth", "range": args.cols, "width": args.chars}])


def cmd_height(args):
    return edit(args, [{"op": "rowHeight", "range": args.rows, "height": args.points}])


def cmd_hide(args):
    return edit(args, [{"op": args.cmd, "range": args.range}])


def cmd_freeze(args):
    return edit(args, [{"op": "freeze", "sheet": args.sheet, "cell": None if args.cell.lower() in ("none", "off") else args.cell}])


def cmd_filter(args):
    return edit(args, [{"op": "filter", "sheet": args.sheet, "range": None if args.range.lower() == "off" else args.range}])


def cmd_sheet(args):
    if args.action == "add":
        return edit(args, [{"op": "addSheet", "name": args.names[0] if args.names else None}])
    if args.action == "rename":
        if len(args.names) != 2:
            raise CliError("사용법: sheet rename 기존이름 새이름", EXIT_USAGE)
        return edit(args, [{"op": "renameSheet", "sheet": args.names[0], "name": args.names[1]}])
    if args.action == "rm":
        return edit(args, [{"op": "deleteSheet", "sheet": args.names[0]}])
    raise CliError("sheet add|rename|rm", EXIT_USAGE)


def cmd_chart(args):
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
    st, d = ctx(args)
    res = request(st, "GET", f"/api/docs/{d['docId']}/describe")
    for sh in res["sheets"]:
        for ch in sh["charts"]:
            print(f"{sh['name']}  {ch['id']}  {ch['type']}  '{ch['title'] or ''}'  @{ch['at']}  {', '.join(x or '' for x in ch['data'])}")
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
        print(f"v{args.since} 이후 변경 없음 (현재 v{res['version']})")
    for c in res["changes"]:
        who = {"user": "사용자", "agent": "에이전트", "system": "시스템"}.get(c["actor"], c["actor"])
        print(f"v{c['v']} {c['time']} {who}: {c['label'] or ''}  [{c['batchId']}]")
        for x in c["cells"] + c["meta"]:
            print(f"    {x}")
    print(f"현재 v{res['version']}{'  *수정됨*' if res['dirty'] else ''}")
    return 0


def cmd_selection(args):
    st, d = ctx(args)
    res = request(st, "GET", f"/api/docs/{d['docId']}/selection")["selection"]
    if args.json:
        return out_json(res)
    if not res:
        print("미리보기에서 선택 중인 영역이 없습니다(창이 열려 있지 않음).")
    for p in res:
        print(f"{p['sheet']}!{p['range']}  (활성 셀 {p['active']}, {p['secondsAgo']}초 전)")
    return 0


def cmd_undo(args):
    st, d = ctx(args)
    res = request(st, "POST", f"/api/docs/{d['docId']}/undo", {"batchId": args.batch, "force": args.force})
    print(f"되돌림 → v{res['version']}")
    return 0


def cmd_save(args):
    st, d = ctx(args)
    body = {"overwrite": args.overwrite, "force": args.force, "backup": not args.no_backup}
    if args.as_path:
        body["path"] = in_root(args, args.as_path)
    if args.format:
        body["format"] = args.format
    res = request(st, "POST", f"/api/docs/{d['docId']}/save", body)
    print(f"저장됨: {res['path']}" + (f"  (백업: {res['backup']})" if res.get("backup") else ""))
    for w in res.get("warnings", []):
        print(f"  주의: {w}")
    return 0


def cmd_close(args):
    st, d = ctx(args)
    request(st, "DELETE", f"/api/docs/{d['docId']}", {"force": args.force})
    print(f"닫음: {d['name']}")
    return 0


# ------------------------------------------------------------------ registration

def register(sub):
    def add(name, fn, help_, doc=True, json_=True):
        p = sub.add_parser(name, help=help_)
        if doc:
            p.add_argument("--doc", help="docId, 파일 이름 또는 경로 (문서가 하나면 생략)")
        if json_:
            p.add_argument("--json", action="store_true", help="결과를 JSON으로")
        p.set_defaults(fn=fn)
        return p

    p = add("ls", cmd_ls, "작업 경로의 Excel 파일 목록", doc=False); p.add_argument("path", nargs="?")
    p = add("open", cmd_open, "파일 열기(미리보기에 탭으로 표시)", doc=False); p.add_argument("path")
    p = add("new", cmd_new, "새 통합 문서", doc=False, json_=False); p.add_argument("--name")
    add("docs", cmd_docs, "열린 문서 목록", doc=False)
    add("describe", cmd_describe, "시트·사용 범위·머리글·열 유형·차트 요약")
    p = add("read", cmd_read, "범위 읽기 (예: 'Sheet1!A1:D20')"); p.add_argument("range")
    p.add_argument("--what", choices=["values", "raw", "formulas", "all"], default="values", help="values: 날짜는 YYYY-MM-DD로, raw: 일련번호 그대로"); p.add_argument("--max", type=int, default=200)
    p = add("write", cmd_write, "값/수식 쓰기 (=로 시작하면 수식)"); p.add_argument("range"); p.add_argument("values", nargs="*")
    p.add_argument("--data", dest="json_values", metavar="JSON", help="2차원 배열 JSON, 예: '[[1,2],[3,4]]'"); p.add_argument("--tsv", metavar="FILE|-")
    p.add_argument("--down", action="store_true", help="위치 인자 값을 세로로"); p.add_argument("--raw", action="store_true", help="문자열을 해석하지 않음")
    p = add("fill", cmd_fill, "범위를 한 값/수식으로 채우기(상대 참조 이동)"); p.add_argument("range"); p.add_argument("value"); p.add_argument("--raw", action="store_true")
    p = add("clear", cmd_clear, "지우기"); p.add_argument("range"); p.add_argument("--what", choices=["contents", "formats", "all"], default="contents")
    p = add("format", cmd_format, "서식 (예: b=true bg=#FFF2CC nf=#,##0)"); p.add_argument("range"); p.add_argument("set", nargs="*")
    p.add_argument("--unset", action="append"); p.add_argument("--border", help="all|outline|inside|top|bottom|left|right|none[:style[:#color]]")
    for name in ("merge", "unmerge"):
        p = add(name, cmd_merge, "셀 병합" if name == "merge" else "병합 해제"); p.add_argument("range")
    p = add("width", cmd_width, "열 너비(문자 수)"); p.add_argument("cols"); p.add_argument("chars", type=float)
    p = add("height", cmd_height, "행 높이(포인트)"); p.add_argument("rows"); p.add_argument("points", type=float)
    for name in ("hide", "unhide"):
        p = add(name, cmd_hide, "행/열 숨기기" if name == "hide" else "숨기기 취소"); p.add_argument("range", help="예: B:C 또는 3:5")
    p = add("freeze", cmd_freeze, "틀 고정 (예: B2, none)"); p.add_argument("cell"); p.add_argument("--sheet")
    p = add("filter", cmd_filter, "자동 필터 켜기(범위) / off"); p.add_argument("range"); p.add_argument("--sheet")
    p = add("sheet", cmd_sheet, "시트 add NAME | rename OLD NEW | rm NAME"); p.add_argument("action", choices=["add", "rename", "rm"]); p.add_argument("names", nargs="*")
    p = add("chart", cmd_chart, "차트 add DATA | rm ID | ls"); p.add_argument("action", choices=["add", "rm", "ls"]); p.add_argument("target", nargs="?")
    p.add_argument("--type", default="bar", help="bar, barH, barStacked, line, area, pie, doughnut, scatter, radar")
    p.add_argument("--title"); p.add_argument("--at", help="왼쪽 위 셀"); p.add_argument("--sheet"); p.add_argument("--legend")
    p.add_argument("--width", type=int); p.add_argument("--height", type=int)
    p = add("apply", cmd_apply, "편집 목록(JSON 파일/-)을 한 번에 적용"); p.add_argument("file"); p.add_argument("--label")
    p = add("changes", cmd_changes, "변경 이력(사용자/에이전트)"); p.add_argument("--since", type=int, default=0); p.add_argument("--actor", choices=["user", "agent"])
    add("selection", cmd_selection, "사용자가 미리보기에서 선택 중인 범위")
    p = add("undo", cmd_undo, "에이전트 batch 되돌리기", json_=False); p.add_argument("batch"); p.add_argument("--force", action="store_true")
    p = add("save", cmd_save, "저장 (Python이 파일 기록, 덮어쓰기 전 백업)", json_=False); p.add_argument("--as", dest="as_path")
    p.add_argument("--format", choices=["xlsx", "xlsm", "xls"]); p.add_argument("--overwrite", action="store_true")
    p.add_argument("--force", action="store_true", help="디스크에서 바뀐 파일도 덮어쓰기"); p.add_argument("--no-backup", action="store_true")
    p = add("close", cmd_close, "문서 닫기", json_=False); p.add_argument("--force", action="store_true")

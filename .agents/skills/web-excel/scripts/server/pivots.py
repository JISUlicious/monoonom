"""Pivot tables (step 1: safeguards). Sheet meta "pivots" (from xlsx_io): each pivot's output range
`ref`, its `source` ('Sheet'!A1:D13), fields, and `idx` in the file. Shared by the edit guard
(session), structural edits (agent_api) and describe."""
from __future__ import annotations

from . import a1 as A


def ranges(sh):
    """[(pivot, (r1, c1, r2, c2))] of the pivots on a sheet."""
    out = []
    for p in sh.get("pivots") or []:
        try:
            out.append((p, A.parse_range(p["ref"])))
        except (KeyError, ValueError):
            pass
    return out


def at(sh, r, c):
    """The pivot whose output covers cell (r, c), or None."""
    for p, (r1, c1, r2, c2) in ranges(sh):
        if r1 <= r <= r2 and c1 <= c <= c2:
            return p
    return None


def overlapping(sh, g):
    """The pivots whose output intersects range g = (r1, c1, r2, c2)."""
    return [p for p, (r1, c1, r2, c2) in ranges(sh) if r1 <= g[2] and g[0] <= r2 and c1 <= g[3] and g[1] <= c2]


def blocked_message(p):
    return (f"피벗 테이블 '{p['name']}'({p['ref']}) 영역은 바꿀 수 없습니다. 피벗 결과는 원본 데이터"
            f"({p.get('source') or '?'})에서 만들어지므로, 원본을 고치세요(저장하면 Excel이 열 때 새로 고칩니다).")


def shift_problem(sh, axis, at_, count):
    """Why inserting/deleting rows or columns here would break a pivot (Excel refuses too), else None.
    Inserting at the pivot's first row/column or before it moves the pivot; inside it, or deleting any
    of its rows/columns, would cut through it."""
    for p, (r1, c1, r2, c2) in ranges(sh):
        lo, hi = (r1, r2) if axis == "row" else (c1, c2)
        what = "행" if axis == "row" else "열"
        if count > 0 and lo < at_ <= hi:
            return f"피벗 테이블 '{p['name']}'({p['ref']}) 가운데에는 {what}을 삽입할 수 없습니다. 피벗 위나 아래(왼쪽·오른쪽)에 삽입하세요."
        if count < 0 and at_ <= hi and lo <= at_ - count - 1:
            return f"피벗 테이블 '{p['name']}'({p['ref']})의 {what}은 삭제할 수 없습니다."
    return None

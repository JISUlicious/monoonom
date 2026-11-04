// AutoFilter (Excel-compatible values/custom criteria) and sorting.
import { key, MAX_COLS } from './address.js';
import { formatValue } from './numfmt.js';
import { translateFormula, refsWithPositions } from './formula.js';

const TOTAL_LABEL = /^\s*(합\s*계|총\s*계|소\s*계|계|누\s*계|평\s*균|total|totals|grand total|sum|subtotal|average)\s*$/i;
const AGG_FN = /\b(SUM|SUBTOTAL|AVERAGE|COUNT|COUNTA|MAX|MIN|AGGREGATE)\s*\(/i;
import { cmp } from './engine.js';

const OPS = [
  ['equal', '같음'], ['notEqual', '같지 않음'], ['greaterThan', '보다 큼'], ['greaterThanOrEqual', '크거나 같음'],
  ['lessThan', '보다 작음'], ['lessThanOrEqual', '작거나 같음'], ['contains', '포함'], ['notContains', '포함하지 않음'],
  ['beginsWith', '시작 문자'], ['endsWith', '끝 문자'],
];

/** Convert UI op to the stored Excel customFilter (contains => "*x*" etc.). */
function toExcel(op, val) {
  switch (op) {
    case 'contains': return { op: 'equal', val: `*${val}*` };
    case 'notContains': return { op: 'notEqual', val: `*${val}*` };
    case 'beginsWith': return { op: 'equal', val: `${val}*` };
    case 'endsWith': return { op: 'equal', val: `*${val}` };
    default: return { op, val };
  }
}
function fromExcel(cf) {
  const v = String(cf.val ?? '');
  if (/^\*.+\*$/.test(v)) return { op: cf.op === 'notEqual' ? 'notContains' : 'contains', val: v.slice(1, -1) };
  if (cf.op === 'equal' && /^[^*]+\*$/.test(v)) return { op: 'beginsWith', val: v.slice(0, -1) };
  if (cf.op === 'equal' && /^\*[^*]+$/.test(v)) return { op: 'endsWith', val: v.slice(1) };
  return { op: cf.op || 'equal', val: v };
}

function wild(pat) {
  return new RegExp('^' + pat.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
}

function customTest(cf, value, text) {
  const v = String(cf.val ?? '');
  const num = Number(v);
  const critNum = v.trim() !== '' && !isNaN(num);
  const isNum = typeof value === 'number' && critNum;
  // Excel: a numeric comparison never matches text cells (and vice versa)
  if (/^(greater|less)/.test(cf.op) && critNum !== (typeof value === 'number')) return false;
  const c = isNum ? (value < num ? -1 : value > num ? 1 : 0) : cmp(text.toLowerCase(), v.toLowerCase());
  switch (cf.op) {
    case 'equal': return /[*?]/.test(v) ? wild(v).test(text) : isNum ? c === 0 : text.toLowerCase() === v.toLowerCase();
    case 'notEqual': return /[*?]/.test(v) ? !wild(v).test(text) : isNum ? c !== 0 : text.toLowerCase() !== v.toLowerCase();
    case 'greaterThan': return c > 0;
    case 'greaterThanOrEqual': return c >= 0;
    case 'lessThan': return c < 0;
    case 'lessThanOrEqual': return c <= 0;
  }
  return true;
}

export class Filters {
  constructor(app) {
    this.app = app;
    this.menu = null;
    document.addEventListener('pointerdown', (e) => { if (this.menu && !this.menu.contains(e.target)) this.close(); }, true);
  }

  displayText(sh, r, c) {
    const wb = this.app.doc.wb;
    const v = wb.value(sh, r, c);
    if (v === null || v === undefined || v === '') return '';
    return formatValue(v, sh.get(r, c)?.s?.nf).text;
  }

  /** contiguous data region around a cell (Ctrl+A / filter / sort region) */
  currentRegion(sh, r, c) {
    const filled = (rr, cc) => { if (rr < 0 || cc < 0) return false; const x = sh.cells.get(key(rr, cc)); return !!x && (x.f || (x.v !== undefined && x.v !== null && x.v !== '')); };
    const g = { r1: r, c1: c, r2: r, c2: c };
    const rowHas = (rr) => { if (rr < 0) return false; for (let cc = Math.max(0, g.c1 - 1); cc <= g.c2 + 1; cc++) if (filled(rr, cc)) return true; return false; };
    const colHas = (cc) => { if (cc < 0 || cc >= MAX_COLS) return false; for (let rr = Math.max(0, g.r1 - 1); rr <= g.r2 + 1; rr++) if (filled(rr, cc)) return true; return false; };
    for (let changed = true; changed;) {
      changed = false;
      while (g.r1 > 0 && rowHas(g.r1 - 1)) { g.r1--; changed = true; }
      while (g.r2 <= sh.maxRow && rowHas(g.r2 + 1)) { g.r2++; changed = true; }
      while (g.c1 > 0 && colHas(g.c1 - 1)) { g.c1--; changed = true; }
      while (g.c2 <= sh.maxCol && colHas(g.c2 + 1)) { g.c2++; changed = true; }
    }
    return g;
  }

  /** a row under the data that summarises it (합계/Total label, or SUM/AVERAGE… over the rows above) */
  isTotalRow(sh, r, g) {
    const first = sh.get(r, g.c1)?.v;
    if (typeof first === 'string' && TOTAL_LABEL.test(first)) return true;
    for (let c = g.c1; c <= g.c2; c++) {
      const f = sh.get(r, c)?.f;
      if (!f || !AGG_FN.test(f)) continue;
      for (const { ref } of refsWithPositions(f)) {
        if (ref.kind !== 'cell') continue;
        const b = ref.b || ref.a, lo = Math.min(ref.a.r, b.r), hi = Math.max(ref.a.r, b.r);
        if (hi < r && hi - lo >= 1 && lo >= g.r1) return true;
      }
    }
    return false;
  }

  /** header + data rows of a table: the region without trailing total rows */
  tableBody(sh, g) {
    const t = { ...g };
    while (t.r2 > t.r1 + 1 && this.isTotalRow(sh, t.r2, t)) t.r2--;
    return t;
  }

  /** the table at a cell (for filter / sort / chart / remove duplicates started from one cell) */
  tableAt(sh, r, c) { return this.tableBody(sh, this.currentRegion(sh, r, c)); }

  toggle() {
    const app = this.app, wb = app.doc.wb, sh = wb.sheet;
    wb.begin('필터');
    wb.touchMeta(sh);
    if (sh.autoFilter) {
      sh.autoFilter = null;
      sh.filterHidden.clear();
    } else {
      const sel = app.view().sel;
      let range = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? this.tableAt(sh, sel.r1, sel.c1) : { ...sel, r2: Math.min(sel.r2, sh.maxRow), c2: Math.min(sel.c2, sh.maxCol) };
      if (range.r2 === range.r1) range.r2 = range.r1 + 1;
      sh.autoFilter = { range, cols: {} };
    }
    wb.commit();
    app.render();
  }

  clearAll() {
    const wb = this.app.doc.wb, sh = wb.sheet;
    if (!sh.autoFilter) return;
    wb.begin('필터 지우기'); wb.touchMeta(sh);
    sh.autoFilter.cols = {}; sh.filterHidden.clear();
    wb.commit(); this.app.render();
  }

  reapply() {
    const wb = this.app.doc.wb, sh = wb.sheet;
    if (!sh.autoFilter) return;
    wb.begin('필터 다시 적용'); wb.touchMeta(sh); this.apply(sh, { extend: true }); wb.commit(); this.app.render();
  }

  /** recompute filterHidden from criteria (extends the range down over contiguous data) */
  apply(sh, { extend = false } = {}) {
    const af = sh.autoFilter;
    if (!af) return;
    const g = af.range;
    if (extend) { // only on explicit "reapply": pick up rows typed directly below the range
      const rowFilled = (r) => { for (let c = g.c1; c <= g.c2; c++) if (sh.cells.has(key(r, c))) return true; return false; };
      while (g.r2 < sh.maxRow && rowFilled(g.r2 + 1)) g.r2++;
      g.r2 = this.tableBody(sh, g).r2; // but not a totals row under the data
    }
    sh.filterHidden.clear();
    const entries = Object.entries(af.cols).map(([k, e]) => [g.c1 + +k, e]);
    if (!entries.length) return;
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      for (const [c, e] of entries) {
        const text = this.displayText(sh, r, c);
        const value = this.app.doc.wb.value(sh, r, c);
        let ok = true;
        if (e.values) ok = text === '' ? !!e.blank : e.values.some((x) => x.toLowerCase() === text.toLowerCase());
        if (ok && e.custom?.length) {
          const res = e.custom.map((cf) => customTest(cf, value, text));
          ok = e.and ? res.every(Boolean) : res.some(Boolean);
        }
        if (!ok) { sh.filterHidden.add(r); break; }
      }
    }
  }

  setColumn(c, entry) {
    const wb = this.app.doc.wb, sh = wb.sheet, af = sh.autoFilter;
    wb.begin('필터'); wb.touchMeta(sh);
    const k = String(c - af.range.c1);
    if (entry) af.cols[k] = entry; else delete af.cols[k];
    this.apply(sh);
    wb.commit();
    this.app.render();
  }

  // ------------------------------------------------------------------ sort
  /** sort rows of range by keys [{c, asc}] ; header row excluded if hasHeader */
  sort(sh, range, keys, hasHeader = false) {
    const wb = this.app.doc.wb;
    const r1 = range.r1 + (hasHeader ? 1 : 0), r2 = Math.min(range.r2, sh.maxRow);
    if (r2 <= r1) return;
    const rows = [];
    for (let r = r1; r <= r2; r++) {
      if (sh.filterHidden.has(r)) continue; // like Excel: sort visible rows only
      rows.push({ r, vals: keys.map((k) => wb.value(sh, r, k.c)) });
    }
    const rank = (v) => (v === null || v === undefined || v === '' ? 4 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
    rows.sort((a, b) => {
      for (let i = 0; i < keys.length; i++) {
        const x = a.vals[i], y = b.vals[i];
        const rx = rank(x), ry = rank(y);
        if (rx === 4 || ry === 4) { if (rx !== ry) return rx === 4 ? 1 : -1; continue; } // blanks last
        let d = rx !== ry ? rx - ry : rx === 1 ? String(x).localeCompare(String(y), 'ko', { sensitivity: 'base', numeric: false }) : x < y ? -1 : x > y ? 1 : 0;
        if (!keys[i].asc) d = -d;
        if (d) return d;
      }
      return a.r - b.r;
    });
    const targets = rows.map((x) => x.r).sort((a, b) => a - b);
    const snapshot = new Map();
    for (const { r } of rows) for (let c = range.c1; c <= range.c2; c++) snapshot.set(key(r, c), sh.cells.get(key(r, c)));
    wb.begin('정렬');
    wb.markStructural();
    rows.forEach((row, i) => {
      const dst = targets[i];
      for (let c = range.c1; c <= range.c2; c++) {
        const src = snapshot.get(key(row.r, c));
        let cell = src ? { ...src } : null;
        if (cell?.f) { cell.f = translateFormula(cell.f, dst - row.r, 0); delete cell.cached; }
        wb.setCell(sh, dst, c, cell);
      }
    });
    wb.commit();
    this.app.render();
  }

  sortSelection(asc) {
    const app = this.app, sh = app.doc.wb.sheet, sel = app.view().sel;
    let range, hasHeader = false, col = app.view().active.c;
    if (sh.autoFilter && sel.r1 >= sh.autoFilter.range.r1 && sel.r1 <= sh.autoFilter.range.r2 && col >= sh.autoFilter.range.c1 && col <= sh.autoFilter.range.c2) {
      range = this.tableBody(sh, sh.autoFilter.range); hasHeader = true;
    } else if (sel.r1 === sel.r2 && sel.c1 === sel.c2) {
      range = this.tableAt(sh, sel.r1, sel.c1);
      hasHeader = this.guessHeader(sh, range);
    } else range = { ...sel, r2: Math.min(sel.r2, sh.maxRow), c2: Math.min(sel.c2, sh.maxCol) };
    this.sort(sh, range, [{ c: col, asc }], hasHeader);
  }

  guessHeader(sh, range) {
    if (range.r2 <= range.r1) return false;
    const wb = this.app.doc.wb;
    let textTop = 0, numBelow = 0;
    for (let c = range.c1; c <= range.c2; c++) {
      if (typeof wb.value(sh, range.r1, c) === 'string') textTop++;
      if (typeof wb.value(sh, range.r1 + 1, c) === 'number') numBelow++;
    }
    return textTop > 0 && (numBelow > 0 || textTop === range.c2 - range.c1 + 1);
  }

  // ------------------------------------------------------------------ dropdown UI
  close() { this.menu?.remove(); this.menu = null; }

  openMenu(c, x, y) {
    this.close();
    const app = this.app, sh = app.doc.wb.sheet, af = sh.autoFilter;
    if (!af) return;
    const k = String(c - af.range.c1);
    const cur = af.cols[k] || {};
    // distinct values (from rows not hidden by *other* columns)
    const others = { ...af.cols }; delete others[k];
    const counts = new Map();
    let hasBlank = false;
    const g = af.range;
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      const t = this.displayText(sh, r, c);
      if (t === '') hasBlank = true; else counts.set(t, (counts.get(t) || 0) + 1);
    }
    const vals = [...counts.keys()].sort((a, b) => {
      const na = parseFloat(a.replace(/[,₩$%]/g, '')), nb = parseFloat(b.replace(/[,₩$%]/g, ''));
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return a.localeCompare(b, 'ko');
    });
    const checked = new Set(cur.values ? cur.values.map((v) => v.toLowerCase()) : vals.map((v) => v.toLowerCase()));
    const blankOn = cur.values ? !!cur.blank : true;
    const custom = (cur.custom || []).map(fromExcel);

    const m = document.createElement('div');
    m.className = 'filter-menu';
    m.innerHTML = `
      <button class="fm-item" data-act="asc">↑ 텍스트 오름차순 정렬</button>
      <button class="fm-item" data-act="desc">↓ 텍스트 내림차순 정렬</button>
      <hr>
      <button class="fm-item" data-act="clear" ${af.cols[k] ? '' : 'disabled'}>⨯ 이 열의 필터 해제</button>
      <details class="fm-custom" ${custom.length ? 'open' : ''}><summary>조건 필터…</summary>
        <div class="fm-row"><select data-cf="op0">${OPS.map(([v, l]) => `<option value="${v}"${custom[0]?.op === v ? ' selected' : ''}>${l}</option>`).join('')}</select><input data-cf="v0" value="${escapeAttr(custom[0]?.val ?? '')}" placeholder="값"></div>
        <div class="fm-row"><label><input type="radio" name="andor" value="and" ${cur.and !== false ? 'checked' : ''}> 그리고</label><label><input type="radio" name="andor" value="or" ${cur.and === false ? 'checked' : ''}> 또는</label></div>
        <div class="fm-row"><select data-cf="op1"><option value="">(없음)</option>${OPS.map(([v, l]) => `<option value="${v}"${custom[1]?.op === v ? ' selected' : ''}>${l}</option>`).join('')}</select><input data-cf="v1" value="${escapeAttr(custom[1]?.val ?? '')}" placeholder="값"></div>
      </details>
      <input class="fm-search" placeholder="검색">
      <div class="fm-list">
        <label class="fm-all"><input type="checkbox" data-all> (모두 선택)</label>
        ${vals.map((v) => `<label><input type="checkbox" data-v="${escapeAttr(v)}" ${checked.has(v.toLowerCase()) ? 'checked' : ''}> ${escapeHtml(v)}</label>`).join('')}
        ${hasBlank ? `<label><input type="checkbox" data-blank ${blankOn ? 'checked' : ''}> (필드 값 없음)</label>` : ''}
      </div>
      <div class="fm-btns"><button class="btn primary" data-act="ok">확인</button><button class="btn" data-act="cancel">취소</button></div>`;
    document.body.append(m);
    const vw = window.innerWidth, vh = window.innerHeight;
    m.style.left = Math.min(x, vw - m.offsetWidth - 8) + 'px';
    m.style.top = Math.min(y + 4, vh - m.offsetHeight - 8) + 'px';
    this.menu = m;
    const boxes = () => [...m.querySelectorAll('.fm-list input[data-v], .fm-list input[data-blank]')].filter((b) => b.closest('label').style.display !== 'none');
    const all = m.querySelector('[data-all]');
    const syncAll = () => { const b = boxes(); all.checked = b.every((x) => x.checked); all.indeterminate = !all.checked && b.some((x) => x.checked); };
    syncAll();
    all.addEventListener('change', () => { boxes().forEach((b) => { b.checked = all.checked; }); });
    m.querySelector('.fm-list').addEventListener('change', (e) => { if (!e.target.dataset.all === undefined) return; if (e.target !== all) syncAll(); });
    m.querySelector('.fm-search').addEventListener('input', (e) => {
      const q = e.target.value.toLowerCase();
      for (const lab of m.querySelectorAll('.fm-list label:not(.fm-all)')) lab.style.display = lab.textContent.toLowerCase().includes(q) ? '' : 'none';
      syncAll();
    });
    m.addEventListener('click', (e) => {
      const act = e.target.closest('[data-act]')?.dataset.act;
      if (!act) return;
      if (act === 'cancel') return this.close();
      if (act === 'asc' || act === 'desc') { this.close(); return this.sort(sh, this.tableBody(sh, af.range), [{ c, asc: act === 'asc' }], true); }
      if (act === 'clear') { this.close(); return this.setColumn(c, null); }
      if (act === 'ok') {
        const entry = {};
        const op0 = m.querySelector('[data-cf=op0]').value, v0 = m.querySelector('[data-cf=v0]').value;
        const op1 = m.querySelector('[data-cf=op1]').value, v1 = m.querySelector('[data-cf=v1]').value;
        const det = m.querySelector('.fm-custom');
        if (det.open && v0 !== '') {
          entry.custom = [toExcel(op0, v0)];
          if (op1 && v1 !== '') entry.custom.push(toExcel(op1, v1));
          entry.and = m.querySelector('input[name=andor]:checked').value === 'and';
        } else {
          const all2 = [...m.querySelectorAll('.fm-list input[data-v]')];
          const sel = all2.filter((b) => b.checked).map((b) => b.dataset.v);
          const blank = m.querySelector('[data-blank]');
          const allOn = sel.length === all2.length && (!blank || blank.checked);
          if (!allOn) { entry.values = sel; if (blank?.checked) entry.blank = true; }
        }
        this.close();
        this.setColumn(c, Object.keys(entry).length ? entry : null);
      }
    });
    m.querySelector('.fm-search').focus();
  }
}

function escapeHtml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }

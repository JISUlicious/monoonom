// Workbook / Sheet model with transactional undo-redo and (de)serialization.
import { key, keyRow, keyCol, MAX_COLS, MAX_ROWS, parseRange, rangeA1, a1, parseA1 } from './address.js';
import { translateFormula, adjustFormula, renameSheetInFormula } from './formula.js';
import { Engine } from './engine.js';

let uid = 0;
export const newId = (p = 'id') => `${p}${Date.now().toString(36)}${(uid++).toString(36)}${Math.random().toString(36).slice(2, 6)}`;

// ---- interned style objects (cells reference the same frozen object)
const styleIntern = new Map();
export function internStyle(st) {
  if (!st) return null;
  const clean = {};
  for (const k of Object.keys(st).sort()) {
    const v = st[k];
    if (v === undefined || v === null || v === false || v === '') continue;
    clean[k] = v;
  }
  if (!Object.keys(clean).length) return null;
  const k = JSON.stringify(clean);
  let s = styleIntern.get(k);
  if (!s) { s = Object.freeze(clean); styleIntern.set(k, s); }
  return s;
}

export class Sheet {
  constructor(name) {
    this.id = newId('s');
    this.name = name;
    this.origName = null;
    this.cells = new Map();      // key -> {v, f, t, s, h, cm, a, cached}
    this.colW = new Map();       // c -> px
    this.rowH = new Map();       // r -> px
    this.hiddenCols = new Set();
    this.hiddenRows = new Set(); // manual
    this.filterHidden = new Set();
    this.merges = [];            // [{r1,c1,r2,c2}]
    this.freeze = { r: 0, c: 0 };
    this.autoFilter = null;      // {range:{...}, cols:{relIdx:{values?:[], blank?, custom?:[{op,val}], and?}}}
    this.charts = [];
    this.state = 'visible';
    this.tabColor = null;
    this.showGrid = true;
    this.defaultColW = 64;
    this.defaultRowH = 20;
    this.maxRow = 0;
    this.maxCol = 0;
    this.calc = new Map();
    this._geomVersion = 0;
  }
  isRowHidden(r) { return this.hiddenRows.has(r) || this.filterHidden.has(r); }
  isColHidden(c) { return this.hiddenCols.has(c); }
  get(r, c) { return this.cells.get(key(r, c)); }
  updateExtent() {
    let mr = 0, mc = 0;
    for (const k of this.cells.keys()) { const r = keyRow(k), c = keyCol(k); if (r > mr) mr = r; if (c > mc) mc = c; }
    this.maxRow = mr; this.maxCol = mc;
  }
  mergeAt(r, c) {
    for (const m of this.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    return null;
  }
  metaSnapshot() {
    return structuredClone({
      name: this.name, colW: [...this.colW], rowH: [...this.rowH], hiddenCols: [...this.hiddenCols],
      hiddenRows: [...this.hiddenRows], filterHidden: [...this.filterHidden], merges: this.merges,
      freeze: this.freeze, autoFilter: this.autoFilter, charts: this.charts, state: this.state,
      tabColor: this.tabColor, showGrid: this.showGrid, defaultColW: this.defaultColW, defaultRowH: this.defaultRowH,
    });
  }
  applyMeta(m) {
    m = structuredClone(m);
    this.name = m.name; this.colW = new Map(m.colW); this.rowH = new Map(m.rowH);
    this.hiddenCols = new Set(m.hiddenCols); this.hiddenRows = new Set(m.hiddenRows);
    this.filterHidden = new Set(m.filterHidden); this.merges = m.merges; this.freeze = m.freeze;
    this.autoFilter = m.autoFilter; this.charts = m.charts; this.state = m.state; this.tabColor = m.tabColor;
    this.showGrid = m.showGrid; this.defaultColW = m.defaultColW; this.defaultRowH = m.defaultRowH;
    this._geomVersion++;
  }
}

/** Cell objects are treated as immutable: every change replaces the object. */
function cloneCell(c) { return c ? { ...c } : null; }

export class Workbook {
  constructor() {
    this.id = newId('wb');
    this.docId = newId('doc');
    this.name = '통합 문서1.xlsx';
    this.path = null;
    this.format = 'xlsx';
    this.basePath = null; // original file for patch-saving
    this.sheets = [];
    this.active = 0;
    this.definedNames = {};
    this.defaultFont = { name: '맑은 고딕', size: 11 };
    this.dirty = false;
    this.undoStack = [];
    this.redoStack = [];
    this.txn = null;
    this.listeners = new Set();
    this.engine = new Engine(this);
    this.version = 0;
  }

  get sheet() { return this.sheets[this.active]; }
  on(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }
  emit(ev) { for (const fn of this.listeners) fn(ev); }

  // ------------------------------------------------------------ transactions
  begin(label) {
    if (this.txn) { this.txn.depth++; return; }
    this.txn = { label, depth: 1, cells: new Map(), meta: new Map(), structure: null };
  }
  _rec(sheet) {
    let m = this.txn.cells.get(sheet);
    if (!m) { m = new Map(); this.txn.cells.set(sheet, m); }
    return m;
  }
  setCell(sheet, r, c, cell) {
    if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) return;
    const k = key(r, c);
    const auto = !this.txn;
    if (auto) this.begin('edit');
    const rec = this._rec(sheet);
    if (!rec.has(k)) rec.set(k, cloneCell(sheet.cells.get(k)));
    if (cell && (cell.v !== undefined && cell.v !== null && cell.v !== '' || cell.f || cell.s || cell.h || cell.cm)) {
      const n = {};
      for (const p of ['v', 'f', 't', 's', 'h', 'cm', 'a', 'cached']) if (cell[p] !== undefined && cell[p] !== null && cell[p] !== '') n[p] = cell[p];
      if (cell.v === '') delete n.v;
      sheet.cells.set(k, n);
      if (r > sheet.maxRow) sheet.maxRow = r;
      if (c > sheet.maxCol) sheet.maxCol = c;
    } else {
      sheet.cells.delete(k);
    }
    if (auto) this.commit();
  }
  patchCell(sheet, r, c, patch) {
    const cur = sheet.cells.get(key(r, c)) || {};
    const n = { ...cur, ...patch };
    for (const k of Object.keys(patch)) if (patch[k] === undefined || patch[k] === null) delete n[k];
    this.setCell(sheet, r, c, n);
  }
  touchMeta(sheet) {
    const auto = !this.txn;
    if (auto) throw new Error('touchMeta outside transaction');
    if (!this.txn.meta.has(sheet)) this.txn.meta.set(sheet, sheet.metaSnapshot());
    sheet._geomVersion++;
  }
  touchStructure() {
    if (!this.txn.structure) this.txn.structure = { sheets: [...this.sheets], active: this.active };
  }
  commit() {
    const t = this.txn;
    if (!t) return;
    if (--t.depth > 0) return;
    this.txn = null;
    const entry = { label: t.label, cells: [], meta: [], structure: null };
    for (const [sheet, m] of t.cells) for (const [k, before] of m) {
      const after = cloneCell(sheet.cells.get(k));
      if (JSON.stringify(before) !== JSON.stringify(after)) entry.cells.push({ sheet, k, before, after });
    }
    for (const [sheet, before] of t.meta) entry.meta.push({ sheet, before, after: sheet.metaSnapshot() });
    if (t.structure) entry.structure = { before: t.structure, after: { sheets: [...this.sheets], active: this.active } };
    if (!entry.cells.length && !entry.meta.length && !entry.structure) return;
    this.undoStack.push(entry);
    if (this.undoStack.length > 300) this.undoStack.shift();
    this.redoStack = [];
    this.changed(entry, 'do');
  }
  _apply(entry, dir) {
    const side = dir === 'undo' ? 'before' : 'after';
    if (entry.structure) {
      this.sheets = [...entry.structure[side].sheets];
      this.active = Math.min(entry.structure[side].active, this.sheets.length - 1);
    }
    for (const m of entry.meta) m.sheet.applyMeta(m[side]);
    const touched = new Set();
    for (const ch of entry.cells) {
      const v = ch[side];
      if (v) ch.sheet.cells.set(ch.k, { ...v }); else ch.sheet.cells.delete(ch.k);
      touched.add(ch.sheet);
    }
    for (const s of touched) s.updateExtent();
  }
  undo() {
    const e = this.undoStack.pop();
    if (!e) return false;
    this._apply(e, 'undo');
    this.redoStack.push(e);
    this.changed(e, 'undo');
    return e;
  }
  redo() {
    const e = this.redoStack.pop();
    if (!e) return false;
    this._apply(e, 'redo');
    this.undoStack.push(e);
    this.changed(e, 'redo');
    return e;
  }
  changed(entry, kind) {
    this.dirty = true;
    this.version++;
    this.recalc();
    this.emit({ type: 'change', entry, kind });
  }
  recalc() {
    const t0 = performance.now();
    this.engine.recalcAll();
    this.lastCalcMs = performance.now() - t0;
  }

  // ------------------------------------------------------------ queries
  value(sheet, r, c) {
    try { return this.engine.cellValue(sheet, r, c); } catch { this.engine.depth = 0; this.engine.evaluating.clear(); return { err: '#CALC!' }; }
  }
  sheetByName(name) { const lc = name.toLowerCase(); return this.sheets.find((s) => s.name.toLowerCase() === lc); }
  uniqueSheetName(base = 'Sheet') {
    let i = 1;
    const names = new Set(this.sheets.map((s) => s.name.toLowerCase()));
    if (base !== 'Sheet' && !names.has(base.toLowerCase())) return base;
    while (names.has(`${base}${i}`.toLowerCase()) || names.has(`${base} (${i})`.toLowerCase())) i++;
    return base === 'Sheet' ? `${base}${i}` : `${base} (${i})`;
  }

  // ------------------------------------------------------------ sheet operations
  addSheet(name, at = this.sheets.length) {
    this.begin('시트 추가');
    this.touchStructure();
    const s = new Sheet(name || this.uniqueSheetName());
    this.sheets.splice(at, 0, s);
    this.active = at;
    this.commit();
    return s;
  }
  deleteSheet(idx) {
    if (this.sheets.length <= 1) return false;
    this.begin('시트 삭제');
    this.touchStructure();
    this.sheets.splice(idx, 1);
    this.active = Math.max(0, Math.min(this.active, this.sheets.length - 1));
    this.commit();
    return true;
  }
  renameSheet(idx, name) {
    name = name.trim();
    if (!name || /[\\/?*[\]:]/.test(name) || name.length > 31) throw new Error('시트 이름에 \\ / ? * [ ] : 를 사용할 수 없고 31자 이하여야 합니다.');
    if (this.sheets.some((s, i) => i !== idx && s.name.toLowerCase() === name.toLowerCase())) throw new Error('같은 이름의 시트가 이미 있습니다.');
    const sh = this.sheets[idx];
    const old = sh.name;
    if (old === name) return;
    this.begin('시트 이름 바꾸기');
    this.touchMeta(sh);
    sh.name = name;
    // update references in every formula and chart
    for (const s of this.sheets) {
      for (const [k, cell] of s.cells) {
        if (cell.f && cell.f.toLowerCase().includes(old.toLowerCase())) {
          const nf = renameSheetInFormula(cell.f, old, name);
          if (nf !== cell.f) this.setCell(s, keyRow(k), keyCol(k), { ...cell, f: nf });
        }
      }
      if (s.charts.length) {
        this.touchMeta(s);
        for (const ch of s.charts) for (const ser of ch.series) for (const p of ['val', 'cat', 'nameRef']) {
          if (ser[p]) ser[p] = renameSheetInFormula(ser[p], old, name);
        }
      }
    }
    for (const [n, v] of Object.entries(this.definedNames)) this.definedNames[n] = renameSheetInFormula(v, old, name);
    this.commit();
  }
  moveSheet(from, to) {
    if (to < 0 || to >= this.sheets.length || from === to) return;
    this.begin('시트 이동');
    this.touchStructure();
    const [s] = this.sheets.splice(from, 1);
    this.sheets.splice(to, 0, s);
    this.active = to;
    this.commit();
  }
  duplicateSheet(idx) {
    const src = this.sheets[idx];
    this.begin('시트 복사');
    this.touchStructure();
    const s = new Sheet(this.uniqueSheetName(src.name));
    s.applyMeta({ ...src.metaSnapshot(), name: s.name });
    s.charts.forEach((c) => { c.id = newId('ch'); c.orig = null; c.dirty = true; });
    for (const [k, c] of src.cells) s.cells.set(k, { ...c });
    s.updateExtent();
    this.sheets.splice(idx + 1, 0, s);
    this.active = idx + 1;
    this.commit();
    return s;
  }

  // ------------------------------------------------------------ structural (rows / cols)
  /** Insert (count>0) or delete (count<0) rows or columns at `at` on `sheet`. */
  shift(sheet, axis, at, count) {
    this.begin(count > 0 ? (axis === 'row' ? '행 삽입' : '열 삽입') : (axis === 'row' ? '행 삭제' : '열 삭제'));
    this.touchMeta(sheet);
    const isRow = axis === 'row';
    const moved = [];
    for (const [k, cell] of sheet.cells) {
      const r = keyRow(k), c = keyCol(k);
      const p = isRow ? r : c;
      if (p < at) continue;
      moved.push([r, c, cell]);
    }
    for (const [r, c] of moved) this.setCell(sheet, r, c, null);
    for (const [r, c, cell] of moved) {
      const p = isRow ? r : c;
      if (count < 0 && p < at - count) continue; // deleted
      const nr = isRow ? r + count : r, nc = isRow ? c : c + count;
      if (nr >= MAX_ROWS || nc >= MAX_COLS) continue;
      this.setCell(sheet, nr, nc, cell);
    }
    // formulas everywhere
    for (const s of this.sheets) {
      for (const [k, cell] of [...s.cells]) {
        if (!cell.f) continue;
        const nf = adjustFormula(cell.f, s.name, sheet.name, axis, at, count);
        if (nf !== cell.f) this.setCell(s, keyRow(k), keyCol(k), { ...cell, f: nf });
      }
      for (const ch of s.charts) {
        this.touchMeta(s);
        for (const ser of ch.series) for (const p of ['val', 'cat', 'nameRef']) {
          if (ser[p]) { const n = adjustFormula('=' + ser[p], s.name, sheet.name, axis, at, count).slice(1); if (n !== ser[p]) { ser[p] = n; ch.dirty = true; } }
        }
      }
    }
    const shiftMap = (map) => {
      const n = new Map();
      for (const [i, v] of map) {
        if (i < at) n.set(i, v);
        else if (count > 0) n.set(i + count, v);
        else if (i >= at - count) n.set(i + count, v);
      }
      return n;
    };
    const shiftSet = (set) => new Set([...shiftMap(new Map([...set].map((x) => [x, 1]))).keys()]);
    if (isRow) { sheet.rowH = shiftMap(sheet.rowH); sheet.hiddenRows = shiftSet(sheet.hiddenRows); sheet.filterHidden = shiftSet(sheet.filterHidden); }
    else { sheet.colW = shiftMap(sheet.colW); sheet.hiddenCols = shiftSet(sheet.hiddenCols); }
    const k1 = isRow ? 'r1' : 'c1', k2 = isRow ? 'r2' : 'c2';
    const adjRange = (g) => {
      const n = { ...g };
      if (count > 0) { if (n[k1] >= at) n[k1] += count; if (n[k2] >= at) n[k2] += count; return n; }
      const end = at - count - 1;
      if (n[k1] >= at && n[k2] <= end) return null;
      n[k1] = n[k1] > end ? n[k1] + count : n[k1] >= at ? at : n[k1];
      n[k2] = n[k2] > end ? n[k2] + count : n[k2] >= at ? at - 1 : n[k2];
      return n[k2] < n[k1] ? null : n;
    };
    sheet.merges = sheet.merges.map(adjRange).filter((m) => m && (m.r2 > m.r1 || m.c2 > m.c1));
    if (sheet.autoFilter) {
      const nr = adjRange(sheet.autoFilter.range);
      if (!nr) { sheet.autoFilter = null; sheet.filterHidden.clear(); } else sheet.autoFilter.range = nr;
    }
    for (const ch of sheet.charts) {
      const p = isRow ? 'r' : 'c';
      if (ch.anchor[p] >= at) ch.anchor[p] = Math.max(at, ch.anchor[p] + count);
    }
    this.commit();
  }

  // ------------------------------------------------------------ range helpers
  /** Copy a block (array of rows of cell objects or null) into sheet at (r,c), translating formulas. */
  pasteBlock(sheet, r0, c0, block, { srcR = r0, srcC = c0, translate = true, only = 'all' } = {}) {
    for (let i = 0; i < block.length; i++) {
      for (let j = 0; j < block[i].length; j++) {
        const src = block[i][j];
        const r = r0 + i, c = c0 + j;
        if (r >= MAX_ROWS || c >= MAX_COLS) continue;
        if (only === 'all') {
          let cell = src ? { ...src } : null;
          if (cell && cell.f && translate) cell.f = translateFormula(cell.f, r - (srcR + i), c - (srcC + j));
          if (cell) { delete cell.cached; delete cell.a; }
          this.setCell(sheet, r, c, cell);
        } else if (only === 'values') {
          const v = src ? (src.f ? src._value : src.v) : null;
          const cur = sheet.get(r, c);
          this.setCell(sheet, r, c, { ...(cur || {}), v: v && v.err ? v.err : v, t: v && v.err ? 'e' : undefined, f: undefined });
        } else if (only === 'formats') {
          this.patchCell(sheet, r, c, { s: src ? src.s : null });
        } else if (only === 'formulas') {
          const cur = sheet.get(r, c);
          let cell = src ? { ...src, s: cur?.s } : (cur?.s ? { s: cur.s } : null);
          if (cell && cell.f && translate) cell.f = translateFormula(cell.f, r - (srcR + i), c - (srcC + j));
          this.setCell(sheet, r, c, cell);
        }
      }
    }
  }

  setStyle(sheet, g, fn) {
    // fn(oldStyle) -> newStyle object; applied across the range (including empty cells)
    const big = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > 2e5;
    if (big) {
      // whole rows/columns: only apply to used part
      g = { ...g, r2: Math.min(g.r2, Math.max(sheet.maxRow, g.r1 + 199)), c2: Math.min(g.c2, Math.max(sheet.maxCol, g.c1 + 49)) };
    }
    this.begin('서식');
    for (let r = g.r1; r <= g.r2; r++) {
      if (sheet.isRowHidden(r) && sheet.filterHidden.has(r)) continue;
      for (let c = g.c1; c <= g.c2; c++) {
        const cur = sheet.get(r, c);
        const ns = internStyle(fn(cur?.s ? { ...cur.s } : {}, r, c));
        if ((cur?.s || null) !== ns) this.patchCell(sheet, r, c, { s: ns });
      }
    }
    this.commit();
  }

  // ------------------------------------------------------------ (de)serialization
  static fromJSON(j) {
    const wb = new Workbook();
    wb.name = j.name || wb.name;
    wb.path = j.path || null;
    wb.format = j.format || 'xlsx';
    wb.basePath = j.path || null;
    wb.definedNames = j.definedNames || {};
    wb.defaultFont = j.defaultFont || wb.defaultFont;
    wb.warnings = j.warnings || [];
    if (j.docId) wb.docId = j.docId;
    const styles = (j.styles || []).map(internStyle);
    for (const sj of j.sheets) {
      const s = new Sheet(sj.name);
      if (sj.id) s.id = sj.id;
      s.origName = sj.origName ?? null;
      s.state = sj.state || 'visible';
      s.tabColor = sj.tabColor || null;
      s.showGrid = sj.showGrid !== false;
      s.defaultColW = sj.defaultColW || 64;
      s.defaultRowH = sj.defaultRowH || 20;
      for (const [c, e] of Object.entries(sj.cols || {})) {
        if (e.w !== undefined) s.colW.set(+c, e.w);
        if (e.hidden) s.hiddenCols.add(+c);
      }
      for (const [r, e] of Object.entries(sj.rows || {})) {
        if (e.h !== undefined) s.rowH.set(+r, e.h);
        if (e.hidden) s.hiddenRows.add(+r);
      }
      for (const [r, c, it] of sj.cells || []) {
        const cell = {};
        if (it.f) { cell.f = it.f.startsWith('=') ? it.f : '=' + it.f; if (it.v !== undefined) cell.cached = it.t === 'e' ? { err: it.v } : it.v; }
        else if (it.v !== undefined && it.v !== null) { cell.v = it.v; if (it.t === 'e') cell.t = 'e'; }
        if (typeof it.s === 'number' && styles[it.s]) cell.s = styles[it.s];
        else if (it.s && typeof it.s === 'object') cell.s = internStyle(it.s);
        if (it.h) cell.h = it.h;
        if (it.cm) cell.cm = it.cm;
        if (it.a) cell.a = it.a;
        if (Object.keys(cell).length) s.cells.set(key(r, c), cell);
      }
      s.merges = (sj.merges || []).map((m) => (typeof m === 'string' ? parseRange(m) : m)).filter(Boolean);
      if (sj.freeze) {
        const p = typeof sj.freeze === 'string' ? parseA1(sj.freeze) : sj.freeze;
        if (p) s.freeze = { r: p.r, c: p.c };
      }
      if (sj.autoFilter && sj.autoFilter.ref) {
        s.autoFilter = { range: parseRange(sj.autoFilter.ref), cols: sj.autoFilter.cols || {} };
      }
      s.charts = (sj.charts || []).map((c) => ({ id: newId('ch'), ...c }));
      s.updateExtent();
      wb.sheets.push(s);
    }
    if (!wb.sheets.length) wb.sheets.push(new Sheet('Sheet1'));
    wb.active = Math.min(j.activeSheet || 0, wb.sheets.length - 1);
    if (wb.sheets[wb.active].state !== 'visible') wb.active = Math.max(0, wb.sheets.findIndex((s) => s.state === 'visible'));
    wb.recalc();
    // filter state derived from the file: rows hidden inside a filtered range belong to the filter
    for (const s of wb.sheets) {
      if (s.autoFilter && Object.keys(s.autoFilter.cols).length) {
        for (const r of [...s.hiddenRows]) if (r > s.autoFilter.range.r1 && r <= s.autoFilter.range.r2) { s.hiddenRows.delete(r); s.filterHidden.add(r); }
      }
    }
    return wb;
  }

  toJSON({ forSave = false } = {}) {
    const styles = [], sidx = new Map();
    const styleId = (st) => {
      if (!st) return undefined;
      let i = sidx.get(st);
      if (i === undefined) { i = styles.length; styles.push(st); sidx.set(st, i); }
      return i;
    };
    const sheets = this.sheets.map((s) => {
      const cols = {}, rows = {};
      for (const [c, w] of s.colW) (cols[c] ||= {}).w = w;
      for (const c of s.hiddenCols) (cols[c] ||= {}).hidden = true;
      for (const [r, h] of s.rowH) (rows[r] ||= {}).h = h;
      for (const r of s.hiddenRows) (rows[r] ||= {}).hidden = true;
      for (const r of s.filterHidden) (rows[r] ||= {}).hidden = true;
      const cells = [];
      const keys = [...s.cells.keys()].sort((a, b) => a - b);
      for (const k of keys) {
        const cell = s.cells.get(k);
        const it = {};
        if (cell.f) {
          it.f = cell.f;
          if (forSave) {
            const cv = s.calc.get(k);
            if (cv !== undefined && cv !== null) it.cv = cv && cv.err ? cv.err : cv;
          } else if (cell.cached !== undefined) it.v = cell.cached && cell.cached.err ? cell.cached.err : cell.cached;
          if (cell.a) it.a = cell.a;
        } else if (cell.v !== undefined) { it.v = cell.v; if (cell.t === 'e') it.t = 'e'; }
        const si = styleId(cell.s);
        if (si !== undefined) it.s = si;
        if (cell.h) it.h = cell.h;
        if (cell.cm) it.cm = cell.cm;
        cells.push([keyRow(k), keyCol(k), it]);
      }
      return {
        id: s.id, name: s.name, origName: s.origName, state: s.state, tabColor: s.tabColor, showGrid: s.showGrid,
        defaultColW: s.defaultColW, defaultRowH: s.defaultRowH, cols, rows, cells,
        merges: s.merges.map(rangeA1),
        freeze: s.freeze && (s.freeze.r || s.freeze.c) ? a1(s.freeze.r, s.freeze.c) : null,
        autoFilter: s.autoFilter ? { ref: rangeA1(s.autoFilter.range), cols: s.autoFilter.cols } : null,
        charts: s.charts.map((c) => { const o = { ...c }; delete o.id; return o; }),
      };
    });
    return {
      docId: this.docId, name: this.name, path: this.path, format: this.format, activeSheet: this.active,
      styles, definedNames: this.definedNames, defaultFont: this.defaultFont, sheets,
    };
  }
}

export function blankWorkbook(name) {
  const wb = new Workbook();
  if (name) wb.name = name;
  wb.sheets.push(new Sheet('Sheet1'));
  wb.recalc();
  return wb;
}

// Formula evaluation engine + Excel function library.
import { parse, FormulaError } from './formula.js';
import { key, MAX_COLS, MAX_ROWS, parseSheetRef, a1, colName } from './address.js';
import { formatValue, serialToParts, dateToSerial, parseInput } from './numfmt.js';

// ------------------------------------------------------------------ value helpers
export class Err { constructor(err) { this.err = err; } toString() { return this.err; } }
const E = {
  DIV0: new Err('#DIV/0!'), VALUE: new Err('#VALUE!'), REF: new Err('#REF!'), NAME: new Err('#NAME?'),
  NUM: new Err('#NUM!'), NA: new Err('#N/A'), NULL: new Err('#NULL!'), CALC: new Err('#CALC!'),
};
const errOf = (s) => Object.values(E).find((e) => e.err === s) || new Err(s);
export const isErr = (v) => v instanceof Err || (v && typeof v === 'object' && typeof v.err === 'string' && !v.kind);
class Throw { constructor(e) { this.e = e; } }
class Deep { constructor(sh, k) { this.sh = sh; this.k = k; } }
const MAX_DEPTH = 150;
const fail = (e) => { throw new Throw(e); };

class Range {
  constructor(sheet, r1, c1, r2, c2) { this.kind = 'range'; this.sheet = sheet; this.r1 = r1; this.c1 = c1; this.r2 = r2; this.c2 = c2; }
  get rows() { return this.r2 - this.r1 + 1; }
  get cols() { return this.c2 - this.c1 + 1; }
}
class Arr {
  constructor(rows) { this.kind = 'array'; this.data = rows; }
  get rows() { return this.data.length; }
  get cols() { return this.data[0] ? this.data[0].length : 0; }
}
const isRange = (v) => v instanceof Range;
const isArr = (v) => v instanceof Arr;

// ------------------------------------------------------------------ engine

export class Engine {
  constructor(wb) {
    this.wb = wb;
    this.astCache = new Map();
    this.evaluating = new Set();
    this.circular = new Set();
    this.depth = 0;
    this.functions = FUNCTIONS;
  }

  sheetByName(name) {
    if (name == null) return null;
    const lc = name.toLowerCase();
    return this.wb.sheets.find((s) => s.name.toLowerCase() === lc) || null;
  }

  ast(formula) {
    let a = this.astCache.get(formula);
    if (a === undefined) {
      try { a = parse(formula); } catch (e) { if (e instanceof RangeError) throw e; a = { t: 'parseError', msg: e.message }; }
      if (this.astCache.size > 50000) this.astCache.clear();
      this.astCache.set(formula, a);
    }
    return a;
  }

  /**
   * Full recalculation. Recursion depth is bounded: when a dependency chain gets deeper than
   * MAX_DEPTH the innermost pending cell is evaluated first as a new root (iterative deepening),
   * so long chains (e.g. running totals over 100k rows) never overflow the JS stack.
   */
  recalcAll() {
    this.circular.clear();
    this.depth = 0;
    const pending = [];
    for (const sh of this.wb.sheets) {
      sh.calc = new Map();
      for (const [k, cell] of sh.cells) if (cell.f) pending.push([sh, k]);
    }
    const stack = [];
    for (let i = this.wb.sheets.length - 1; i >= 0; i--) {
      const sh = this.wb.sheets[i];
      const ks = pending.filter((p) => p[0] === sh).map((p) => p[1]).sort((a, b) => b - a);
      for (const k of ks) stack.push([sh, k]);
    }
    let guard = pending.length * 4 + 1000;
    while (stack.length) {
      const [sh, k] = stack[stack.length - 1];
      if (sh.calc.has(k)) { stack.pop(); continue; }
      if (--guard < 0) { sh.calc.set(k, E.CALC); stack.pop(); continue; }
      try {
        this.cellValue(sh, Math.floor(k / MAX_COLS), k % MAX_COLS);
        stack.pop();
      } catch (e) {
        this.evaluating.clear();
        this.depth = 0;
        if (e instanceof Deep) stack.push([e.sh, e.k]);
        else if (e instanceof RangeError) { sh.calc.set(k, E.CALC); stack.pop(); }
        else throw e;
      }
    }
  }

  /** Value of a cell (computed for formulas). Blank -> null. */
  cellValue(sh, r, c) {
    const k = key(r, c);
    const cell = sh.cells.get(k);
    if (!cell) return null;
    if (!cell.f) return cell.t === 'e' ? errOf(cell.v) : (cell.v ?? null);
    if (sh.calc && sh.calc.has(k)) return sh.calc.get(k);
    const id = sh.id + ':' + k;
    if (this.evaluating.has(id)) { this.circular.add(`${sh.name}!${a1(r, c)}`); return 0; }
    if (this.depth >= MAX_DEPTH) throw new Deep(sh, k);
    this.evaluating.add(id);
    this.depth++;
    let v;
    try {
      v = this.evalFormula(cell.f, sh, r, c, cell);
    } finally { this.evaluating.delete(id); this.depth--; }
    if (!sh.calc) sh.calc = new Map();
    sh.calc.set(k, v);
    return v;
  }

  evalFormula(formula, sh, r, c, cell) {
    const ast = this.ast(formula);
    if (ast.t === 'parseError') return E.NAME;
    const ctx = { sh, r, c, engine: this };
    let v;
    try { v = this.ev(ast, ctx); } catch (e) {
      if (e instanceof Throw) v = e.e; else if (e instanceof RangeError || e instanceof Deep) throw e; else { console.error(e); v = E.VALUE; }
    }
    v = this.finalize(v, ctx);
    if (v === E.NAME && cell && cell.cached !== undefined && cell.cached !== null) {
      // unsupported function: fall back to the value Excel last calculated
      return cell.cached;
    }
    return v;
  }

  finalize(v, ctx) {
    if (isRange(v)) {
      if (v.rows === 1 && v.cols === 1) return this.cellValue(v.sheet, v.r1, v.c1) ?? 0;
      // implicit intersection
      if (v.cols === 1 && ctx.r >= v.r1 && ctx.r <= v.r2) return this.cellValue(v.sheet, ctx.r, v.c1) ?? 0;
      if (v.rows === 1 && ctx.c >= v.c1 && ctx.c <= v.c2) return this.cellValue(v.sheet, v.r1, ctx.c) ?? 0;
      return this.cellValue(v.sheet, v.r1, v.c1) ?? 0;
    }
    if (isArr(v)) { const x = v.data[0]?.[0]; return x === undefined || x === null ? 0 : x; }
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number' && !isFinite(v)) return E.NUM;
    return v;
  }

  /** Evaluate formula text in the context of a sheet (used by charts / names). */
  evalText(text, sh) {
    const ast = this.ast(text);
    if (ast.t === 'parseError') return E.NAME;
    try { return this.ev(ast, { sh, r: 0, c: 0, engine: this }); } catch (e) { if (e instanceof Throw) return e.e; return E.VALUE; }
  }

  ev(n, ctx) {
    switch (n.t) {
      case 'num': case 'str': case 'bool': return n.v;
      case 'err': return errOf(n.v);
      case 'missing': return null;
      case 'paren': return this.ev(n.e, ctx);
      case 'array': return new Arr(n.rows.map((r) => r.map((x) => (x && x.err ? errOf(x.err) : x))));
      case 'ref': {
        const ref = n.ref;
        const sh = ref.sheet != null ? this.sheetByName(ref.sheet) : ctx.sh;
        if (!sh) return E.REF;
        const b = ref.b || ref.a;
        return new Range(sh, Math.min(ref.a.r, b.r), Math.min(ref.a.c, b.c), Math.max(ref.a.r, b.r), Math.max(ref.a.c, b.c));
      }
      case 'name': {
        const nm = this.wb.definedNames && (this.wb.definedNames[n.v] ?? Object.entries(this.wb.definedNames)
          .find(([k]) => k.toLowerCase() === n.v.toLowerCase())?.[1]);
        if (nm == null) return E.NAME;
        const id = 'name:' + n.v;
        if (this.evaluating.has(id)) return E.REF;
        this.evaluating.add(id);
        try { return this.ev(this.ast('=' + String(nm).replace(/^=/, '')), ctx); } finally { this.evaluating.delete(id); }
      }
      case 'range': {
        const a = this.ev(n.a, ctx), b = this.ev(n.b, ctx);
        if (!isRange(a) || !isRange(b)) return E.VALUE;
        if (a.sheet !== b.sheet) return E.REF;
        return new Range(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
      }
      case 'neg': return this.unary(this.ev(n.e, ctx), (x) => -x);
      case 'pct': return this.unary(this.ev(n.e, ctx), (x) => x / 100);
      case 'bin': return this.binary(n.op, this.ev(n.a, ctx), this.ev(n.b, ctx), ctx);
      case 'fn': {
        const def = this.functions[n.name];
        if (!def) return E.NAME;
        if (def.lazy) return def.fn(n.args, ctx, this);
        const args = n.args.map((a) => this.ev(a, ctx));
        if (def.raw) return def.fn(args, ctx, this);
        return def.fn(args, ctx, this);
      }
      default: return E.VALUE;
    }
  }

  // --- coercions
  deref(v) { // scalar from range (top-left) or array
    if (isRange(v)) return v.rows === 1 && v.cols === 1 ? this.cellValue(v.sheet, v.r1, v.c1) : this.finalizeScalar(v);
    if (isArr(v)) return v.data[0][0];
    return v;
  }
  finalizeScalar(v) { return E.VALUE; }

  toGrid(v) { // 2D values from range/array/scalar
    if (isRange(v)) {
      const sh = v.sheet;
      const r2 = Math.min(v.r2, Math.max(sh.maxRow ?? 0, v.r1)), c2 = Math.min(v.c2, Math.max(sh.maxCol ?? 0, v.c1));
      const out = [];
      for (let r = v.r1; r <= r2; r++) {
        const row = [];
        for (let c = v.c1; c <= c2; c++) row.push(this.cellValue(sh, r, c));
        out.push(row);
      }
      return out;
    }
    if (isArr(v)) return v.data;
    return [[v]];
  }

  /** iterate values of a range lazily over existing cells only */
  *rangeValues(v) {
    if (isRange(v)) {
      const sh = v.sheet;
      const r2 = Math.min(v.r2, sh.maxRow ?? 0), c2 = Math.min(v.c2, sh.maxCol ?? 0);
      const area = (r2 - v.r1 + 1) * (c2 - v.c1 + 1);
      if (area > 4 * sh.cells.size && sh.cells.size < 200000) {
        const ks = [];
        for (const k of sh.cells.keys()) {
          const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS;
          if (r >= v.r1 && r <= r2 && c >= v.c1 && c <= c2) ks.push(k);
        }
        ks.sort((x, y) => x - y);
        for (const k of ks) yield this.cellValue(sh, Math.floor(k / MAX_COLS), k % MAX_COLS);
        return;
      }
      for (let r = v.r1; r <= r2; r++) for (let c = v.c1; c <= c2; c++) yield this.cellValue(sh, r, c);
      return;
    }
    if (isArr(v)) { for (const row of v.data) for (const x of row) yield x; return; }
    yield v;
  }

  num(v) {
    v = this.deref(v);
    if (isErr(v)) fail(v);
    if (typeof v === 'number') return v;
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
      const p = parseInput(v);
      if (typeof p.v === 'number') return p.v;
      fail(E.VALUE);
    }
    fail(E.VALUE);
  }
  str(v) {
    v = this.deref(v);
    if (isErr(v)) fail(v);
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return formatValue(v, 'General').text;
    return String(v);
  }
  bool(v) {
    v = this.deref(v);
    if (isErr(v)) fail(v);
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (v === null || v === undefined) return false;
    if (typeof v === 'string') { const u = v.toUpperCase(); if (u === 'TRUE') return true; if (u === 'FALSE') return false; }
    fail(E.VALUE);
  }
  int(v) { return Math.trunc(this.num(v)); }

  unary(v, f) {
    if (isRange(v) || isArr(v)) return new Arr(this.toGrid(v).map((row) => row.map((x) => {
      if (isErr(x)) return x;
      try { return f(this.num(x)); } catch (e) { if (e instanceof Throw) return e.e; throw e; }
    })));
    if (isErr(v)) return v;
    return f(this.num(v));
  }

  binary(op, a, b, ctx) {
    const aMulti = (isRange(a) && (a.rows > 1 || a.cols > 1)) || isArr(a);
    const bMulti = (isRange(b) && (b.rows > 1 || b.cols > 1)) || isArr(b);
    if (aMulti || bMulti) {
      const A = aMulti ? this.toGridFull(a) : [[this.deref(a)]];
      const B = bMulti ? this.toGridFull(b) : [[this.deref(b)]];
      const rows = Math.max(A.length, B.length), cols = Math.max(A[0].length, B[0].length);
      const pick = (G, r, c) => {
        const rr = G.length === 1 ? 0 : r, cc = G[0].length === 1 ? 0 : c;
        return G[rr] && G[rr][cc] !== undefined ? G[rr][cc] : E.NA;
      };
      const out = [];
      for (let r = 0; r < rows; r++) {
        const row = [];
        for (let c = 0; c < cols; c++) row.push(this.scalarOp(op, pick(A, r, c), pick(B, r, c)));
        out.push(row);
      }
      return new Arr(out);
    }
    return this.scalarOp(op, this.deref(a), this.deref(b));
  }

  toGridFull(v) { // like toGrid but without clamping (keeps shape for element-wise ops)
    if (isRange(v)) {
      const sh = v.sheet;
      const r2 = Math.min(v.r2, v.r1 + Math.max(0, (sh.maxRow ?? 0) - v.r1) ), c2 = Math.min(v.c2, v.c1 + Math.max(0, (sh.maxCol ?? 0) - v.c1));
      const out = [];
      for (let r = v.r1; r <= r2; r++) { const row = []; for (let c = v.c1; c <= c2; c++) row.push(this.cellValue(sh, r, c)); out.push(row); }
      return out.length ? out : [[null]];
    }
    return this.toGrid(v);
  }

  scalarOp(op, x, y) {
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    try {
      switch (op) {
        case '+': return this.num(x) + this.num(y);
        case '-': return this.num(x) - this.num(y);
        case '*': return this.num(x) * this.num(y);
        case '/': { const d = this.num(y); if (d === 0) return E.DIV0; return this.num(x) / d; }
        case '^': { const r = Math.pow(this.num(x), this.num(y)); return isFinite(r) ? r : E.NUM; }
        case '&': return this.str(x) + this.str(y);
        default: return compareOp(op, cmp(x, y));
      }
    } catch (e) { if (e instanceof Throw) return e.e; throw e; }
  }
}

function compareOp(op, c) {
  switch (op) {
    case '=': return c === 0; case '<>': return c !== 0; case '<': return c < 0;
    case '>': return c > 0; case '<=': return c <= 0; case '>=': return c >= 0;
  }
  return E.VALUE;
}

/** Excel comparison: numbers < text < logicals; blank acts as 0/""/FALSE; text case-insensitive. */
export function cmp(x, y) {
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : -1);
  if (x === null || x === undefined) x = typeof y === 'string' ? '' : typeof y === 'boolean' ? false : 0;
  if (y === null || y === undefined) y = typeof x === 'string' ? '' : typeof x === 'boolean' ? false : 0;
  const rx = rank(x), ry = rank(y);
  if (rx !== ry) return rx < ry ? -1 : 1;
  if (rx === 1) { const a = x.toLowerCase(), b = y.toLowerCase(); return a < b ? -1 : a > b ? 1 : 0; }
  if (rx === 2) return x === y ? 0 : x ? 1 : -1;
  if (x === y) return 0;
  const a = parseFloat(x.toPrecision(15)), b = parseFloat(y.toPrecision(15));
  return a < b ? -1 : a > b ? 1 : 0;
}

function wildcardRe(pat) {
  let re = '';
  for (let i = 0; i < pat.length; i++) {
    const ch = pat[i];
    if (ch === '~' && i + 1 < pat.length) { re += pat[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
    if (ch === '*') re += '.*'; else if (ch === '?') re += '.'; else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp('^' + re + '$', 'i');
}

/** COUNTIF-style criteria -> predicate */
export function criteria(crit) {
  if (typeof crit === 'number') return (v) => typeof v === 'number' && v === crit || (typeof v === 'string' && v !== '' && Number(v) === crit);
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit === null || crit === undefined) return (v) => v === null || v === '';
  const s = String(crit);
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s);
  const op = m[1] || '=';
  const rhsS = m[2];
  const p = parseInput(rhsS);
  const rhsNum = typeof p.v === 'number' ? p.v : null;
  const rhsBool = /^(true|false)$/i.test(rhsS) ? rhsS.toUpperCase() === 'TRUE' : null;
  if (op === '=' || op === '<>') {
    let test;
    if (rhsS === '') test = (v) => v === null || v === undefined || v === '';
    else if (rhsNum !== null) test = (v) => typeof v === 'number' ? v === rhsNum : (typeof v === 'string' && v.trim() !== '' && parseInput(v).v === rhsNum);
    else if (rhsBool !== null) test = (v) => v === rhsBool;
    else if (/[*?~]/.test(rhsS)) { const re = wildcardRe(rhsS); test = (v) => typeof v === 'string' && re.test(v); }
    else { const lc = rhsS.toLowerCase(); test = (v) => typeof v === 'string' && v.toLowerCase() === lc; }
    return op === '=' ? test : (v) => !test(v);
  }
  if (rhsNum !== null) return (v) => typeof v === 'number' && compareOp(op, v < rhsNum ? -1 : v > rhsNum ? 1 : 0);
  return (v) => typeof v === 'string' && compareOp(op, cmp(v, rhsS));
}

// ------------------------------------------------------------------ function library helpers
const FUNCTIONS = {};
const def = (names, fn, opts = {}) => { for (const n of names.split(' ')) FUNCTIONS[n] = { fn, ...opts }; };

/** numbers from args, Excel aggregate semantics: refs -> only numbers; direct args -> coerced */
function nums(eng, args, { countBool = false, countText = false } = {}) {
  const out = [];
  for (const a of args) {
    if (isRange(a) || isArr(a)) {
      for (const v of eng.rangeValues(a)) {
        if (isErr(v)) fail(v);
        if (typeof v === 'number') out.push(v);
        else if (countBool && typeof v === 'boolean') out.push(v ? 1 : 0);
        else if (countText && typeof v === 'string') out.push(0);
      }
    } else if (a === null) {
      continue;
    } else {
      if (isErr(a)) fail(a);
      if (typeof a === 'string' && countText && isNaN(Number(a))) { out.push(0); continue; }
      out.push(eng.num(a));
    }
  }
  return out;
}
const flat = (eng, args) => { const out = []; for (const a of args) for (const v of eng.rangeValues(a)) out.push(v); return out; };
const wrap = (fn) => (args, ctx, eng) => { try { return fn(args, ctx, eng); } catch (e) { if (e instanceof Throw) return e.e; throw e; } };
const optNum = (eng, v, d) => (v === undefined || v === null ? d : eng.num(v));
const sum = (a) => a.reduce((s, x) => s + x, 0);
const mean = (a) => (a.length ? sum(a) / a.length : fail(E.DIV0));
const variance = (a, sample) => {
  if (a.length < (sample ? 2 : 1)) fail(E.DIV0);
  const m = mean(a); return sum(a.map((x) => (x - m) ** 2)) / (a.length - (sample ? 1 : 0));
};
const round = (x, d) => { const f = Math.pow(10, d); const y = Math.abs(x) * f; const r = Math.round(y + 1e-9 * Math.max(1, y)) / f; return Math.sign(x) * r; };
const fixFloat = (x) => (Number.isFinite(x) ? parseFloat(x.toPrecision(15)) : x);

// --- logical / control (lazy)
def('IF', (a, ctx, eng) => {
  let c = eng.ev(a[0], ctx);
  if (isRange(c) && (c.rows > 1 || c.cols > 1)) c = eng.finalize(c, ctx);
  c = eng.deref(c);
  if (isErr(c)) return c;
  let t;
  try { t = eng.bool(c); } catch (e) { if (e instanceof Throw) return e.e; throw e; }
  if (t) return a.length > 1 ? (a[1].t === 'missing' ? 0 : eng.ev(a[1], ctx)) : true;
  return a.length > 2 ? (a[2].t === 'missing' ? 0 : eng.ev(a[2], ctx)) : false;
}, { lazy: true });
def('IFS', wrap((a, ctx, eng) => {
  for (let i = 0; i + 1 < a.length; i += 2) if (eng.bool(eng.ev(a[i], ctx))) return eng.ev(a[i + 1], ctx);
  return E.NA;
}), { lazy: true });
def('IFERROR', (a, ctx, eng) => { const v = eng.finalize(eng.ev(a[0], ctx), ctx); return isErr(v) ? eng.ev(a[1], ctx) : eng.ev(a[0], ctx); }, { lazy: true });
def('IFNA', (a, ctx, eng) => { const v = eng.finalize(eng.ev(a[0], ctx), ctx); return isErr(v) && v.err === '#N/A' ? eng.ev(a[1], ctx) : eng.ev(a[0], ctx); }, { lazy: true });
def('CHOOSE', wrap((a, ctx, eng) => { const i = eng.int(eng.ev(a[0], ctx)); if (i < 1 || i >= a.length) return E.VALUE; return eng.ev(a[i], ctx); }), { lazy: true });
def('SWITCH', wrap((a, ctx, eng) => {
  const v = eng.deref(eng.ev(a[0], ctx));
  if (isErr(v)) return v;
  let i = 1;
  for (; i + 1 < a.length; i += 2) if (cmp(v, eng.deref(eng.ev(a[i], ctx))) === 0) return eng.ev(a[i + 1], ctx);
  return i < a.length ? eng.ev(a[i], ctx) : E.NA;
}), { lazy: true });
def('LET', wrap((a, ctx, eng) => {
  // LET(name1, value1, ..., calc) — implemented via temporary defined names
  const saved = eng.wb.definedNames; const names = { ...(saved || {}) };
  eng.wb.definedNames = names;
  try {
    for (let i = 0; i + 1 < a.length; i += 2) {
      if (a[i].t !== 'name') return E.VALUE;
      const v = eng.finalize(eng.ev(a[i + 1], ctx), ctx);
      names[a[i].v] = typeof v === 'string' ? `"${v.replace(/"/g, '""')}"` : isErr(v) ? v.err : String(v).toUpperCase();
    }
    return eng.ev(a[a.length - 1], ctx);
  } finally { eng.wb.definedNames = saved; }
}), { lazy: true });
def('AND', wrap((a, _c, eng) => { const v = flat(eng, a).filter((x) => x !== null && typeof x !== 'string'); if (!v.length) return E.VALUE; return v.every((x) => eng.bool(x)); }));
def('OR', wrap((a, _c, eng) => { const v = flat(eng, a).filter((x) => x !== null && typeof x !== 'string'); if (!v.length) return E.VALUE; return v.some((x) => eng.bool(x)); }));
def('XOR', wrap((a, _c, eng) => flat(eng, a).filter((x) => x !== null && typeof x !== 'string').reduce((s, x) => s !== eng.bool(x), false)));
def('NOT', wrap((a, _c, eng) => !eng.bool(a[0])));
def('TRUE', () => true);
def('FALSE', () => false);

// --- math
def('SUM', wrap((a, _c, eng) => fixFloat(sum(nums(eng, a)))));
def('PRODUCT', wrap((a, _c, eng) => nums(eng, a).reduce((p, x) => p * x, 1)));
def('SUMSQ', wrap((a, _c, eng) => sum(nums(eng, a).map((x) => x * x))));
def('ABS', wrap((a, _c, e) => Math.abs(e.num(a[0]))));
def('SIGN', wrap((a, _c, e) => Math.sign(e.num(a[0]))));
def('INT', wrap((a, _c, e) => Math.floor(e.num(a[0]))));
def('TRUNC', wrap((a, _c, e) => { const d = optNum(e, a[1], 0); const f = 10 ** d; return Math.trunc(e.num(a[0]) * f) / f; }));
def('ROUND', wrap((a, _c, e) => round(e.num(a[0]), e.int(a[1] ?? 0))));
def('ROUNDUP', wrap((a, _c, e) => { const x = e.num(a[0]), f = 10 ** e.int(a[1] ?? 0); return Math.sign(x) * Math.ceil(fixFloat(Math.abs(x) * f)) / f; }));
def('ROUNDDOWN', wrap((a, _c, e) => { const x = e.num(a[0]), f = 10 ** e.int(a[1] ?? 0); return Math.sign(x) * Math.floor(fixFloat(Math.abs(x) * f)) / f; }));
def('MROUND', wrap((a, _c, e) => { const x = e.num(a[0]), m = e.num(a[1]); if (m === 0) return 0; if (x * m < 0) return E.NUM; return fixFloat(Math.round(x / m) * m); }));
def('CEILING CEILING.PRECISE ISO.CEILING', wrap((a, _c, e) => { const x = e.num(a[0]), s = optNum(e, a[1], 1); if (s === 0) return 0; return fixFloat(Math.ceil(fixFloat(x / s)) * s); }));
def('FLOOR FLOOR.PRECISE', wrap((a, _c, e) => { const x = e.num(a[0]), s = optNum(e, a[1], 1); if (s === 0) return E.DIV0; return fixFloat(Math.floor(fixFloat(x / s)) * s); }));
def('CEILING.MATH', wrap((a, _c, e) => { const x = e.num(a[0]), s = Math.abs(optNum(e, a[1], 1)) || 1, m = optNum(e, a[2], 0); return fixFloat((x < 0 && m ? -Math.ceil(fixFloat(-x / s)) : Math.ceil(fixFloat(x / s))) * s); }));
def('FLOOR.MATH', wrap((a, _c, e) => { const x = e.num(a[0]), s = Math.abs(optNum(e, a[1], 1)) || 1, m = optNum(e, a[2], 0); return fixFloat((x < 0 && m ? -Math.floor(fixFloat(-x / s)) : Math.floor(fixFloat(x / s))) * s); }));
def('EVEN', wrap((a, _c, e) => { const x = e.num(a[0]); const v = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -v : v; }));
def('ODD', wrap((a, _c, e) => { const x = e.num(a[0]); let v = Math.ceil(Math.abs(x)); if (v % 2 === 0) v += 1; return x < 0 ? -v : v; }));
def('MOD', wrap((a, _c, e) => { const n = e.num(a[0]), d = e.num(a[1]); if (d === 0) return E.DIV0; return fixFloat(n - d * Math.floor(n / d)); }));
def('QUOTIENT', wrap((a, _c, e) => { const d = e.num(a[1]); if (d === 0) return E.DIV0; return Math.trunc(e.num(a[0]) / d); }));
def('POWER', wrap((a, _c, e) => { const r = Math.pow(e.num(a[0]), e.num(a[1])); return isFinite(r) ? r : E.NUM; }));
def('SQRT', wrap((a, _c, e) => { const x = e.num(a[0]); return x < 0 ? E.NUM : Math.sqrt(x); }));
def('SQRTPI', wrap((a, _c, e) => Math.sqrt(e.num(a[0]) * Math.PI)));
def('EXP', wrap((a, _c, e) => Math.exp(e.num(a[0]))));
def('LN', wrap((a, _c, e) => { const x = e.num(a[0]); return x <= 0 ? E.NUM : Math.log(x); }));
def('LOG', wrap((a, _c, e) => { const x = e.num(a[0]), b = optNum(e, a[1], 10); return x <= 0 || b <= 0 ? E.NUM : Math.log(x) / Math.log(b); }));
def('LOG10', wrap((a, _c, e) => { const x = e.num(a[0]); return x <= 0 ? E.NUM : Math.log10(x); }));
def('PI', () => Math.PI);
def('RAND', () => Math.random());
def('RANDBETWEEN', wrap((a, _c, e) => { const lo = Math.ceil(e.num(a[0])), hi = Math.floor(e.num(a[1])); return lo + Math.floor(Math.random() * (hi - lo + 1)); }));
def('FACT', wrap((a, _c, e) => { let n = Math.floor(e.num(a[0])); if (n < 0) return E.NUM; let r = 1; while (n > 1) r *= n--; return r; }));
def('COMBIN', wrap((a, _c, e) => { const n = Math.floor(e.num(a[0])), k = Math.floor(e.num(a[1])); if (k < 0 || n < k) return E.NUM; let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); }));
def('PERMUT', wrap((a, _c, e) => { const n = Math.floor(e.num(a[0])), k = Math.floor(e.num(a[1])); if (k < 0 || n < k) return E.NUM; let r = 1; for (let i = 0; i < k; i++) r *= n - i; return r; }));
const _gcd = (x, y) => (y ? _gcd(y, x % y) : x);
def('GCD', wrap((a, _c, e) => nums(e, a).map(Math.floor).reduce(_gcd)));
def('LCM', wrap((a, _c, e) => nums(e, a).map(Math.floor).reduce((x, y) => (x && y ? (x * y) / _gcd(x, y) : 0))));
for (const [n, f] of Object.entries({ SIN: Math.sin, COS: Math.cos, TAN: Math.tan, ASIN: Math.asin, ACOS: Math.acos, ATAN: Math.atan, SINH: Math.sinh, COSH: Math.cosh, TANH: Math.tanh })) {
  def(n, wrap((a, _c, e) => { const r = f(e.num(a[0])); return isNaN(r) ? E.NUM : r; }));
}
def('ATAN2', wrap((a, _c, e) => { const x = e.num(a[0]), y = e.num(a[1]); if (!x && !y) return E.DIV0; return Math.atan2(y, x); }));
def('DEGREES', wrap((a, _c, e) => (e.num(a[0]) * 180) / Math.PI));
def('RADIANS', wrap((a, _c, e) => (e.num(a[0]) * Math.PI) / 180));
def('SUMPRODUCT', wrap((a, _c, eng) => {
  const grids = a.map((x) => eng.toGridFull(x));
  const R = grids[0].length, C = grids[0][0].length;
  if (grids.some((g) => g.length !== R || g[0].length !== C)) return E.VALUE;
  let s = 0;
  for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
    let p = 1;
    for (const g of grids) { const v = g[r][c]; if (isErr(v)) return v; p *= typeof v === 'number' ? v : typeof v === 'boolean' && grids.length === 1 ? 0 : 0; }
    s += p;
  }
  return fixFloat(s);
}));

function ifsPairs(eng, a, start) {
  const pairs = [];
  for (let i = start; i + 1 < a.length; i += 2) pairs.push([eng.toGridFull(a[i]), criteria(eng.deref(a[i + 1]))]);
  return pairs;
}
function matchMask(pairs, R, C) {
  return (r, c) => pairs.every(([g, p]) => p(g[r]?.[c] ?? null));
}
function shapeOf(eng, v) { return isRange(v) ? [v.rows, v.cols] : isArr(v) ? [v.rows, v.cols] : [1, 1]; }
function cellAt(eng, v, r, c) {
  if (isRange(v)) return eng.cellValue(v.sheet, v.r1 + r, v.c1 + c);
  if (isArr(v)) return v.data[r]?.[c] ?? null;
  return v;
}
function condAgg(eng, critRange, crit, valRange, kind) {
  const [R, C] = shapeOf(eng, critRange);
  const pred = criteria(eng.deref(crit));
  const vr = valRange ?? critRange;
  const sh = isRange(critRange) ? critRange.sheet : null;
  const maxR = sh ? Math.min(R, (sh.maxRow ?? 0) - critRange.r1 + 1) : R;
  const out = [];
  for (let r = 0; r < Math.max(0, maxR); r++) for (let c = 0; c < C; c++) {
    if (pred(cellAt(eng, critRange, r, c))) {
      const v = cellAt(eng, vr, r, c);
      if (kind === 'count') out.push(1);
      else if (typeof v === 'number') out.push(v);
    }
  }
  if (!sh || R !== maxR) { /* blank rows beyond data */
    if (sh && R > maxR && pred(null) && kind === 'count') out.push(...Array(Math.max(0, (R - Math.max(0, maxR)) * C)).fill(1));
  }
  return out;
}
def('SUMIF', wrap((a, _c, eng) => fixFloat(sum(condAgg(eng, a[0], a[1], a[2] ?? null, 'sum')))));
def('COUNTIF', wrap((a, _c, eng) => condAgg(eng, a[0], a[1], null, 'count').length));
def('AVERAGEIF', wrap((a, _c, eng) => mean(condAgg(eng, a[0], a[1], a[2] ?? null, 'sum'))));
function multiIf(eng, valRange, a, start) {
  const pairs = [];
  for (let i = start; i + 1 < a.length; i += 2) pairs.push([a[i], criteria(eng.deref(a[i + 1]))]);
  const ref = valRange ?? pairs[0][0];
  const [R, C] = shapeOf(eng, ref);
  const sh = isRange(ref) ? ref.sheet : null;
  const maxR = sh ? Math.max(0, Math.min(R, (sh.maxRow ?? 0) - ref.r1 + 1)) : R;
  for (const [g] of pairs) { const [r2, c2] = shapeOf(eng, g); if (r2 !== R || c2 !== C) fail(E.VALUE); }
  const out = [];
  for (let r = 0; r < maxR; r++) for (let c = 0; c < C; c++) {
    if (pairs.every(([g, p]) => p(cellAt(eng, g, r, c)))) out.push(valRange ? cellAt(eng, valRange, r, c) : 1);
  }
  if (!valRange && R > maxR && pairs.every(([, p]) => p(null))) for (let i = 0; i < (R - maxR) * C; i++) out.push(1);
  return out;
}
def('SUMIFS', wrap((a, _c, eng) => fixFloat(sum(multiIf(eng, a[0], a, 1).filter((x) => typeof x === 'number')))));
def('COUNTIFS', wrap((a, _c, eng) => multiIf(eng, null, a, 0).length));
def('AVERAGEIFS', wrap((a, _c, eng) => mean(multiIf(eng, a[0], a, 1).filter((x) => typeof x === 'number'))));
def('MAXIFS', wrap((a, _c, eng) => { const v = multiIf(eng, a[0], a, 1).filter((x) => typeof x === 'number'); return v.length ? Math.max(...v) : 0; }));
def('MINIFS', wrap((a, _c, eng) => { const v = multiIf(eng, a[0], a, 1).filter((x) => typeof x === 'number'); return v.length ? Math.min(...v) : 0; }));

// --- statistics
def('AVERAGE', wrap((a, _c, e) => mean(nums(e, a))));
def('AVERAGEA', wrap((a, _c, e) => mean(nums(e, a, { countBool: true, countText: true }))));
def('COUNT', wrap((a, _c, eng) => { let n = 0; for (const x of a) { if (isRange(x) || isArr(x)) { for (const v of eng.rangeValues(x)) if (typeof v === 'number') n++; } else if (typeof x === 'number' || typeof x === 'boolean' || (typeof x === 'string' && !isNaN(parseFloat(x)))) n++; } return n; }));
def('COUNTA', wrap((a, _c, eng) => { let n = 0; for (const x of a) { if (isRange(x) || isArr(x)) { for (const v of eng.rangeValues(x)) if (v !== null && v !== undefined && v !== '') n++; else if (v === '') n++; } else if (x !== null) n++; } return n; }));
def('COUNTBLANK', wrap((a, _c, eng) => { const r = a[0]; if (!isRange(r)) return E.VALUE; let n = 0; for (let i = r.r1; i <= r.r2; i++) for (let j = r.c1; j <= r.c2; j++) { const v = eng.cellValue(r.sheet, i, j); if (v === null || v === '') n++; } return n; }));
def('MAX', wrap((a, _c, e) => { const v = nums(e, a); return v.length ? Math.max(...v) : 0; }));
def('MIN', wrap((a, _c, e) => { const v = nums(e, a); return v.length ? Math.min(...v) : 0; }));
def('MAXA', wrap((a, _c, e) => { const v = nums(e, a, { countBool: true, countText: true }); return v.length ? Math.max(...v) : 0; }));
def('MINA', wrap((a, _c, e) => { const v = nums(e, a, { countBool: true, countText: true }); return v.length ? Math.min(...v) : 0; }));
const median = (v) => { if (!v.length) fail(E.NUM); const s = [...v].sort((x, y) => x - y); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
def('MEDIAN', wrap((a, _c, e) => median(nums(e, a))));
def('MODE MODE.SNGL', wrap((a, _c, e) => { const cnt = new Map(); let best = null, bc = 1; for (const x of nums(e, a)) { const n = (cnt.get(x) || 0) + 1; cnt.set(x, n); if (n > bc) { bc = n; best = x; } } return best === null ? E.NA : best; }));
def('STDEV STDEV.S', wrap((a, _c, e) => Math.sqrt(variance(nums(e, a), true))));
def('STDEVP STDEV.P', wrap((a, _c, e) => Math.sqrt(variance(nums(e, a), false))));
def('VAR VAR.S', wrap((a, _c, e) => variance(nums(e, a), true)));
def('VARP VAR.P', wrap((a, _c, e) => variance(nums(e, a), false)));
def('LARGE', wrap((a, _c, e) => { const v = nums(e, [a[0]]).sort((x, y) => y - x); const k = e.int(a[1]); return k < 1 || k > v.length ? E.NUM : v[k - 1]; }));
def('SMALL', wrap((a, _c, e) => { const v = nums(e, [a[0]]).sort((x, y) => x - y); const k = e.int(a[1]); return k < 1 || k > v.length ? E.NUM : v[k - 1]; }));
def('RANK RANK.EQ', wrap((a, _c, e) => { const x = e.num(a[0]); const v = nums(e, [a[1]]); const asc = optNum(e, a[2], 0); if (!v.includes(x)) return E.NA; return 1 + v.filter((y) => (asc ? y < x : y > x)).length; }));
def('RANK.AVG', wrap((a, _c, e) => { const x = e.num(a[0]); const v = nums(e, [a[1]]); const asc = optNum(e, a[2], 0); const n = v.filter((y) => y === x).length; if (!n) return E.NA; return 1 + v.filter((y) => (asc ? y < x : y > x)).length + (n - 1) / 2; }));
const pctInc = (v, p) => { if (!v.length || p < 0 || p > 1) fail(E.NUM); const s = [...v].sort((x, y) => x - y); const h = (s.length - 1) * p; const lo = Math.floor(h); return s[lo] + (h - lo) * ((s[lo + 1] ?? s[lo]) - s[lo]); };
const pctExc = (v, p) => { const s = [...v].sort((x, y) => x - y); const h = (s.length + 1) * p - 1; if (h < 0 || h > s.length - 1) fail(E.NUM); const lo = Math.floor(h); return s[lo] + (h - lo) * ((s[lo + 1] ?? s[lo]) - s[lo]); };
def('PERCENTILE PERCENTILE.INC', wrap((a, _c, e) => pctInc(nums(e, [a[0]]), e.num(a[1]))));
def('PERCENTILE.EXC', wrap((a, _c, e) => pctExc(nums(e, [a[0]]), e.num(a[1]))));
def('QUARTILE QUARTILE.INC', wrap((a, _c, e) => { const q = e.int(a[1]); if (q < 0 || q > 4) return E.NUM; return pctInc(nums(e, [a[0]]), q / 4); }));
def('QUARTILE.EXC', wrap((a, _c, e) => { const q = e.int(a[1]); if (q < 1 || q > 3) return E.NUM; return pctExc(nums(e, [a[0]]), q / 4); }));
function pairedNums(eng, x, y) {
  const X = [...eng.rangeValues(x)], Y = [...eng.rangeValues(y)];
  const px = [], py = [];
  for (let i = 0; i < Math.min(X.length, Y.length); i++) if (typeof X[i] === 'number' && typeof Y[i] === 'number') { px.push(X[i]); py.push(Y[i]); }
  return [px, py];
}
def('CORREL PEARSON', wrap((a, _c, e) => { const [x, y] = pairedNums(e, a[0], a[1]); const mx = mean(x), my = mean(y); let sxy = 0, sxx = 0, syy = 0; x.forEach((v, i) => { sxy += (v - mx) * (y[i] - my); sxx += (v - mx) ** 2; syy += (y[i] - my) ** 2; }); return sxx && syy ? sxy / Math.sqrt(sxx * syy) : E.DIV0; }));
def('SLOPE', wrap((a, _c, e) => { const [y, x] = pairedNums(e, a[0], a[1]); const mx = mean(x), my = mean(y); let n = 0, d = 0; x.forEach((v, i) => { n += (v - mx) * (y[i] - my); d += (v - mx) ** 2; }); return d ? n / d : E.DIV0; }));
def('INTERCEPT', wrap((a, _c, e) => { const [y, x] = pairedNums(e, a[0], a[1]); const mx = mean(x), my = mean(y); let n = 0, d = 0; x.forEach((v, i) => { n += (v - mx) * (y[i] - my); d += (v - mx) ** 2; }); return d ? my - (n / d) * mx : E.DIV0; }));
def('FORECAST FORECAST.LINEAR', wrap((a, _c, e) => { const t = e.num(a[0]); const [y, x] = pairedNums(e, a[1], a[2]); const mx = mean(x), my = mean(y); let n = 0, d = 0; x.forEach((v, i) => { n += (v - mx) * (y[i] - my); d += (v - mx) ** 2; }); if (!d) return E.DIV0; const b = n / d; return my - b * mx + b * t; }));
def('COVARIANCE.P COVAR', wrap((a, _c, e) => { const [x, y] = pairedNums(e, a[0], a[1]); const mx = mean(x), my = mean(y); return sum(x.map((v, i) => (v - mx) * (y[i] - my))) / x.length; }));
def('COVARIANCE.S', wrap((a, _c, e) => { const [x, y] = pairedNums(e, a[0], a[1]); if (x.length < 2) return E.DIV0; const mx = mean(x), my = mean(y); return sum(x.map((v, i) => (v - mx) * (y[i] - my))) / (x.length - 1); }));
def('GEOMEAN', wrap((a, _c, e) => { const v = nums(e, a); if (!v.length || v.some((x) => x <= 0)) return E.NUM; return Math.exp(sum(v.map(Math.log)) / v.length); }));
def('AVEDEV', wrap((a, _c, e) => { const v = nums(e, a); const m = mean(v); return sum(v.map((x) => Math.abs(x - m))) / v.length; }));
const erf = (x) => { const t = 1 / (1 + 0.3275911 * Math.abs(x)); const y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x); return x >= 0 ? y : -y; };
def('NORM.S.DIST', wrap((a, _c, e) => { const z = e.num(a[0]); return e.bool(a[1] ?? true) ? 0.5 * (1 + erf(z / Math.SQRT2)) : Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI); }));
def('NORM.DIST NORMDIST', wrap((a, _c, e) => { const x = e.num(a[0]), m = e.num(a[1]), s = e.num(a[2]); if (s <= 0) return E.NUM; const z = (x - m) / s; return e.bool(a[3]) ? 0.5 * (1 + erf(z / Math.SQRT2)) : Math.exp(-z * z / 2) / (s * Math.sqrt(2 * Math.PI)); }));
def('SUBTOTAL', wrap((a, ctx, eng) => {
  const code = eng.int(a[0]) % 100;
  const map = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };
  const fn = map[code];
  if (!fn) return E.VALUE;
  // SUBTOTAL ignores rows hidden by filter (and manually hidden when code >= 100)
  const hideManual = eng.int(a[0]) >= 100;
  const args = a.slice(1).map((r) => {
    if (!isRange(r)) return r;
    const sh = r.sheet; const rows = [];
    for (let i = r.r1; i <= Math.min(r.r2, sh.maxRow ?? 0); i++) {
      if (sh.filterHidden?.has(i) || (hideManual && sh.isRowHidden?.(i))) continue;
      const row = []; for (let j = r.c1; j <= r.c2; j++) row.push(eng.cellValue(sh, i, j)); rows.push(row);
    }
    return new Arr(rows.length ? rows : [[null]]);
  });
  return FUNCTIONS[fn].fn(args, ctx, eng);
}));
def('AGGREGATE', wrap((a, ctx, eng) => {
  const code = eng.int(a[0]);
  const map = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S', 11: 'VAR.P', 12: 'MEDIAN', 13: 'MODE.SNGL', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC' };
  const fn = map[code]; if (!fn) return E.VALUE;
  const vals = flat(eng, [a[2]]).filter((v) => typeof v === 'number');
  return FUNCTIONS[fn].fn([new Arr([vals.length ? vals : [null]]), ...a.slice(3)], ctx, eng);
}));

// --- text
def('CONCATENATE', wrap((a, _c, e) => a.map((x) => e.str(x)).join('')));
def('CONCAT', wrap((a, _c, e) => flat(e, a).map((x) => (isErr(x) ? fail(x) : e.str(x))).join('')));
def('TEXTJOIN', wrap((a, _c, e) => { const d = e.str(a[0]); const ign = e.bool(a[1]); return flat(e, a.slice(2)).map((x) => (isErr(x) ? fail(x) : e.str(x))).filter((x) => !ign || x !== '').join(d); }));
def('LEFT', wrap((a, _c, e) => { const n = optNum(e, a[1], 1); if (n < 0) return E.VALUE; return [...e.str(a[0])].slice(0, n).join(''); }));
def('RIGHT', wrap((a, _c, e) => { const n = optNum(e, a[1], 1); if (n < 0) return E.VALUE; const s = [...e.str(a[0])]; return n === 0 ? '' : s.slice(-n).join(''); }));
def('MID', wrap((a, _c, e) => { const st = e.int(a[1]), n = e.int(a[2]); if (st < 1 || n < 0) return E.VALUE; return [...e.str(a[0])].slice(st - 1, st - 1 + n).join(''); }));
def('LEN', wrap((a, _c, e) => [...e.str(a[0])].length));
def('LENB', wrap((a, _c, e) => [...e.str(a[0])].reduce((n, ch) => n + (ch.charCodeAt(0) > 255 ? 2 : 1), 0)));
def('UPPER', wrap((a, _c, e) => e.str(a[0]).toUpperCase()));
def('LOWER', wrap((a, _c, e) => e.str(a[0]).toLowerCase()));
def('PROPER', wrap((a, _c, e) => e.str(a[0]).toLowerCase().replace(/(^|[^A-Za-zÀ-￿])([a-z])/g, (m, p, c) => p + c.toUpperCase())));
def('TRIM', wrap((a, _c, e) => e.str(a[0]).replace(/ +/g, ' ').trim()));
def('CLEAN', wrap((a, _c, e) => e.str(a[0]).replace(/[\x00-\x1f]/g, '')));
def('SUBSTITUTE', wrap((a, _c, e) => {
  const s = e.str(a[0]), o = e.str(a[1]), n = e.str(a[2]);
  if (!o) return s;
  if (a[3] === undefined || a[3] === null) return s.split(o).join(n);
  const k = e.int(a[3]); let idx = -1;
  for (let i = 0; i < k; i++) { idx = s.indexOf(o, idx + 1); if (idx < 0) return s; }
  return s.slice(0, idx) + n + s.slice(idx + o.length);
}));
def('REPLACE', wrap((a, _c, e) => { const s = e.str(a[0]); const st = e.int(a[1]), n = e.int(a[2]); return s.slice(0, st - 1) + e.str(a[3]) + s.slice(st - 1 + n); }));
def('FIND', wrap((a, _c, e) => { const st = optNum(e, a[2], 1); const i = e.str(a[1]).indexOf(e.str(a[0]), st - 1); return i < 0 ? E.VALUE : i + 1; }));
def('SEARCH', wrap((a, _c, e) => {
  const st = optNum(e, a[2], 1); const pat = e.str(a[0]); const s = e.str(a[1]);
  const re = new RegExp(wildcardRe(pat).source.slice(1, -1), 'i');
  const m = re.exec(s.slice(st - 1));
  return m ? m.index + st : E.VALUE;
}));
def('REPT', wrap((a, _c, e) => { const n = e.int(a[1]); return n < 0 ? E.VALUE : e.str(a[0]).repeat(n); }));
def('EXACT', wrap((a, _c, e) => e.str(a[0]) === e.str(a[1])));
def('CHAR', wrap((a, _c, e) => String.fromCharCode(e.int(a[0]))));
def('UNICHAR', wrap((a, _c, e) => String.fromCodePoint(e.int(a[0]))));
def('CODE UNICODE', wrap((a, _c, e) => { const s = e.str(a[0]); return s ? s.codePointAt(0) : E.VALUE; }));
def('T', wrap((a, _c, e) => { const v = e.deref(a[0]); return typeof v === 'string' ? v : ''; }));
def('N', wrap((a, _c, e) => { const v = e.deref(a[0]); return typeof v === 'number' ? v : typeof v === 'boolean' ? +v : isErr(v) ? v : 0; }));
def('VALUE', wrap((a, _c, e) => { const v = e.deref(a[0]); if (typeof v === 'number') return v; const p = parseInput(e.str(v)); return typeof p.v === 'number' ? p.v : E.VALUE; }));
def('NUMBERVALUE', wrap((a, _c, e) => { let s = e.str(a[0]); const dec = a[1] != null ? e.str(a[1]) : '.', grp = a[2] != null ? e.str(a[2]) : ','; s = s.split(grp).join('').replace(dec, '.').replace(/\s/g, ''); let pct = 0; while (s.endsWith('%')) { s = s.slice(0, -1); pct++; } const v = Number(s); return isNaN(v) ? E.VALUE : v / 100 ** pct; }));
def('TEXT', wrap((a, _c, e) => { const v = e.deref(a[0]); if (isErr(v)) return v; let x = v; if (typeof v === 'string') { const p = parseInput(v); if (typeof p.v === 'number') x = p.v; } return formatValue(x, e.str(a[1])).text; }));
def('FIXED', wrap((a, _c, e) => { const d = optNum(e, a[1], 2); const nc = a[2] != null && e.bool(a[2]); const v = round(e.num(a[0]), d); return v.toLocaleString('en-US', { minimumFractionDigits: Math.max(0, d), maximumFractionDigits: Math.max(0, d), useGrouping: !nc }); }));
def('DOLLAR', wrap((a, _c, e) => { const d = optNum(e, a[1], 2); return formatValue(e.num(a[0]), `"₩"#,##0${d > 0 ? '.' + '0'.repeat(d) : ''}`).text; }));
def('TEXTBEFORE', wrap((a, _c, e) => { const s = e.str(a[0]), d = e.str(a[1]); const n = optNum(e, a[2], 1); let i = -1; if (n > 0) { for (let k = 0; k < n; k++) { i = s.indexOf(d, i + 1); if (i < 0) return E.NA; } } else { i = s.length; for (let k = 0; k < -n; k++) { i = s.lastIndexOf(d, i - 1); if (i < 0) return E.NA; } } return s.slice(0, i); }));
def('TEXTAFTER', wrap((a, _c, e) => { const s = e.str(a[0]), d = e.str(a[1]); const n = optNum(e, a[2], 1); let i = -1; if (n > 0) { for (let k = 0; k < n; k++) { i = s.indexOf(d, i + 1); if (i < 0) return E.NA; } } else { i = s.length; for (let k = 0; k < -n; k++) { i = s.lastIndexOf(d, i - 1); if (i < 0) return E.NA; } } return s.slice(i + d.length); }));

// --- information
def('ISBLANK', (a, _c, e) => { const v = e.deref(a[0]); return v === null || v === undefined; });
def('ISNUMBER', (a, _c, e) => typeof e.deref(a[0]) === 'number');
def('ISTEXT', (a, _c, e) => typeof e.deref(a[0]) === 'string');
def('ISNONTEXT', (a, _c, e) => typeof e.deref(a[0]) !== 'string');
def('ISLOGICAL', (a, _c, e) => typeof e.deref(a[0]) === 'boolean');
def('ISERROR', (a, _c, e) => isErr(e.deref(a[0])));
def('ISERR', (a, _c, e) => { const v = e.deref(a[0]); return isErr(v) && v.err !== '#N/A'; });
def('ISNA', (a, _c, e) => { const v = e.deref(a[0]); return isErr(v) && v.err === '#N/A'; });
def('ISREF', (a) => isRange(a[0]));
def('ISEVEN', wrap((a, _c, e) => Math.floor(Math.abs(e.num(a[0]))) % 2 === 0));
def('ISODD', wrap((a, _c, e) => Math.floor(Math.abs(e.num(a[0]))) % 2 === 1));
def('ISFORMULA', (a) => { const r = a[0]; if (!isRange(r)) return E.VALUE; return !!r.sheet.cells.get(key(r.r1, r.c1))?.f; });
def('FORMULATEXT', (a) => { const r = a[0]; if (!isRange(r)) return E.VALUE; return r.sheet.cells.get(key(r.r1, r.c1))?.f ?? E.NA; });
def('NA', () => E.NA);
def('ERROR.TYPE', (a, _c, e) => { const v = e.deref(a[0]); const order = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']; return isErr(v) ? order.indexOf(v.err) + 1 || E.NA : E.NA; });
def('TYPE', (a, _c, e) => { const v = a[0]; if (isArr(v) || (isRange(v) && (v.rows > 1 || v.cols > 1))) return 64; const x = e.deref(v); return typeof x === 'number' || x === null ? 1 : typeof x === 'string' ? 2 : typeof x === 'boolean' ? 4 : 16; });
def('SHEET', (a, ctx, e) => { const r = a[0]; const sh = isRange(r) ? r.sheet : typeof r === 'string' ? e.sheetByName(r) : ctx.sh; return sh ? e.wb.sheets.indexOf(sh) + 1 : E.NA; });
def('SHEETS', (_a, _c, e) => e.wb.sheets.length);

// --- lookup & reference
def('ROW', (a, ctx) => (a.length && isRange(a[0]) ? a[0].r1 + 1 : ctx.r + 1));
def('COLUMN', (a, ctx) => (a.length && isRange(a[0]) ? a[0].c1 + 1 : ctx.c + 1));
def('ROWS', (a, _c, e) => (isRange(a[0]) || isArr(a[0]) ? a[0].rows : 1));
def('COLUMNS', (a, _c, e) => (isRange(a[0]) || isArr(a[0]) ? a[0].cols : 1));
def('ADDRESS', wrap((a, _c, e) => {
  const r = e.int(a[0]), c = e.int(a[1]), abs = optNum(e, a[2], 1);
  const ra = abs === 1 || abs === 2, ca = abs === 1 || abs === 3;
  let s = (ca ? '$' : '') + colName(c - 1) + (ra ? '$' : '') + r;
  if (a[4] != null) s = `${/^[\w.]+$/.test(e.str(a[4])) ? e.str(a[4]) : `'${e.str(a[4])}'`}!${s}`;
  return s;
}));
def('INDIRECT', wrap((a, ctx, e) => {
  const p = parseSheetRef(e.str(a[0]));
  if (!p) return E.REF;
  const sh = p.sheet != null ? e.sheetByName(p.sheet) : ctx.sh;
  if (!sh) return E.REF;
  return new Range(sh, p.range.r1, p.range.c1, p.range.r2, p.range.c2);
}));
def('OFFSET', wrap((a, _c, e) => {
  const r = a[0]; if (!isRange(r)) return E.VALUE;
  const dr = e.int(a[1]), dc = e.int(a[2]);
  const h = a[3] != null ? e.int(a[3]) : r.rows, w = a[4] != null ? e.int(a[4]) : r.cols;
  const r1 = r.r1 + dr, c1 = r.c1 + dc;
  if (r1 < 0 || c1 < 0 || h < 1 || w < 1 || r1 + h > MAX_ROWS || c1 + w > MAX_COLS) return E.REF;
  return new Range(r.sheet, r1, c1, r1 + h - 1, c1 + w - 1);
}));
def('INDEX', wrap((a, _c, e) => {
  const src = a[0];
  let row = a[1] != null ? e.int(a[1]) : 0, col = a[2] != null ? e.int(a[2]) : 0;
  const [R, C] = isRange(src) || isArr(src) ? [src.rows, src.cols] : [1, 1];
  if (a[2] == null && R === 1 && C > 1) { col = row; row = 1; }
  if (row < 0 || col < 0 || row > R || col > C) return E.REF;
  if (isRange(src)) {
    const r1 = row ? src.r1 + row - 1 : src.r1, r2 = row ? r1 : src.r2;
    const c1 = col ? src.c1 + col - 1 : src.c1, c2 = col ? c1 : src.c2;
    return new Range(src.sheet, r1, c1, r2, c2);
  }
  if (isArr(src)) return src.data[(row || 1) - 1][(col || 1) - 1];
  return src;
}));
function lookupVector(eng, v) {
  // returns [values[], (i) => position info]
  if (isRange(v)) {
    const vertical = v.cols === 1 || v.rows > 1;
    const n = vertical ? v.rows : v.cols;
    const sh = v.sheet;
    const lim = vertical ? Math.min(n, Math.max(0, (sh.maxRow ?? 0) - v.r1 + 1)) : Math.min(n, Math.max(0, (sh.maxCol ?? 0) - v.c1 + 1));
    const vals = [];
    for (let i = 0; i < lim; i++) vals.push(vertical ? eng.cellValue(sh, v.r1 + i, v.c1) : eng.cellValue(sh, v.r1, v.c1 + i));
    return { vals, vertical, n };
  }
  if (isArr(v)) { const vertical = v.cols === 1; return { vals: vertical ? v.data.map((r) => r[0]) : v.data[0], vertical, n: vertical ? v.rows : v.cols }; }
  return { vals: [v], vertical: true, n: 1 };
}
function findMatch(vals, x, mode, wildcard = true) {
  // mode 0 exact, 1 largest <= (sorted asc, binary search), -1 smallest >= (sorted desc)
  if (mode === 0) {
    if (typeof x === 'string' && wildcard && /[*?~]/.test(x)) { const re = wildcardRe(x); return vals.findIndex((v) => typeof v === 'string' && re.test(v)); }
    return vals.findIndex((v) => v !== null && v !== undefined && !isErr(v) && typeof v === typeof x && cmp(v, x) === 0);
  }
  if (mode === 1) {
    let lo = 0, hi = vals.length - 1, ans = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const v = vals[mid];
      if (v === null || v === undefined || typeof v !== typeof x) {
        // skip incompatible types: linear fallback
        let j = mid; while (j >= lo && (vals[j] === null || typeof vals[j] !== typeof x)) j--;
        if (j < lo) { lo = mid + 1; continue; }
        if (cmp(vals[j], x) <= 0) { ans = j; lo = mid + 1; } else hi = j - 1;
        continue;
      }
      if (cmp(v, x) <= 0) { ans = mid; lo = mid + 1; } else hi = mid - 1;
    }
    return ans;
  }
  let ans = -1;
  for (let i = 0; i < vals.length; i++) { const v = vals[i]; if (typeof v !== typeof x) continue; if (cmp(v, x) >= 0) ans = i; else break; }
  return ans;
}
def('MATCH', wrap((a, _c, e) => {
  const x = e.deref(a[0]); if (isErr(x)) return x;
  const mode = a[2] != null ? Math.sign(e.num(a[2])) : 1;
  const { vals } = lookupVector(e, a[1]);
  const i = findMatch(vals, x, mode);
  return i < 0 ? E.NA : i + 1;
}));
def('XMATCH', wrap((a, _c, e) => {
  const x = e.deref(a[0]); const mm = optNum(e, a[2], 0); const sm = optNum(e, a[3], 1);
  let { vals } = lookupVector(e, a[1]);
  const idx = xfind(vals, x, mm, sm);
  return idx < 0 ? E.NA : idx + 1;
}));
function xfind(vals, x, mm, sm) {
  const order = [...vals.keys()]; if (sm < 0) order.reverse();
  let best = -1;
  const re = mm === 2 && typeof x === 'string' ? wildcardRe(x) : null;
  for (const i of order) {
    const v = vals[i];
    if (v === null || v === undefined) continue;
    if (re) { if (typeof v === 'string' && re.test(v)) return i; continue; }
    if (typeof v === typeof x && cmp(v, x) === 0) return i;
    if (mm === -1 && typeof v === typeof x && cmp(v, x) < 0 && (best < 0 || cmp(v, vals[best]) > 0)) best = i;
    if (mm === 1 && typeof v === typeof x && cmp(v, x) > 0 && (best < 0 || cmp(v, vals[best]) < 0)) best = i;
  }
  return best;
}
def('XLOOKUP', wrap((a, _c, e) => {
  const x = e.deref(a[0]); if (isErr(x)) return x;
  const { vals, vertical } = lookupVector(e, a[1]);
  const mm = optNum(e, a[4], 0), sm = optNum(e, a[5], 1);
  const i = xfind(vals, x, mm, sm);
  if (i < 0) return a[3] !== undefined && a[3] !== null ? a[3] : E.NA;
  const ret = a[2];
  if (isRange(ret)) return vertical ? new Range(ret.sheet, ret.r1 + i, ret.c1, ret.r1 + i, ret.c2) : new Range(ret.sheet, ret.r1, ret.c1 + i, ret.r2, ret.c1 + i);
  if (isArr(ret)) return vertical ? ret.data[i]?.[0] ?? E.NA : ret.data[0][i] ?? E.NA;
  return ret;
}));
def('VLOOKUP', wrap((a, _c, e) => {
  const x = e.deref(a[0]); if (isErr(x)) return x;
  const t = a[1]; const col = e.int(a[2]); const approx = a[3] == null ? true : e.bool(a[3]);
  if (!isRange(t) && !isArr(t)) return E.VALUE;
  if (col < 1 || col > t.cols) return col < 1 ? E.VALUE : E.REF;
  const first = isRange(t) ? new Range(t.sheet, t.r1, t.c1, t.r2, t.c1) : new Arr(t.data.map((r) => [r[0]]));
  const { vals } = lookupVector(e, first);
  const i = findMatch(vals, x, approx ? 1 : 0);
  if (i < 0) return E.NA;
  return isRange(t) ? e.cellValue(t.sheet, t.r1 + i, t.c1 + col - 1) : t.data[i][col - 1];
}));
def('HLOOKUP', wrap((a, _c, e) => {
  const x = e.deref(a[0]); if (isErr(x)) return x;
  const t = a[1]; const row = e.int(a[2]); const approx = a[3] == null ? true : e.bool(a[3]);
  if (!isRange(t) && !isArr(t)) return E.VALUE;
  if (row < 1 || row > t.rows) return row < 1 ? E.VALUE : E.REF;
  const first = isRange(t) ? new Range(t.sheet, t.r1, t.c1, t.r1, t.c2) : new Arr([t.data[0]]);
  const { vals } = lookupVector(e, first);
  const i = findMatch(vals, x, approx ? 1 : 0);
  if (i < 0) return E.NA;
  return isRange(t) ? e.cellValue(t.sheet, t.r1 + row - 1, t.c1 + i) : t.data[row - 1][i];
}));
def('LOOKUP', wrap((a, _c, e) => {
  const x = e.deref(a[0]);
  const { vals } = lookupVector(e, a[1]);
  const i = findMatch(vals, x, 1);
  if (i < 0) return E.NA;
  const res = a[2] ?? a[1];
  const rv = lookupVector(e, res).vals;
  return rv[i] ?? null;
}));
def('TRANSPOSE', wrap((a, _c, e) => { const g = e.toGrid(a[0]); return new Arr(g[0].map((_, c) => g.map((r) => r[c]))); }));
def('SEQUENCE', wrap((a, _c, e) => { const R = e.int(a[0]), C = optNum(e, a[1], 1), s = optNum(e, a[2], 1), st = optNum(e, a[3], 1); const out = []; for (let r = 0; r < R; r++) { const row = []; for (let c = 0; c < C; c++) row.push(s + (r * C + c) * st); out.push(row); } return new Arr(out); }));
def('UNIQUE', wrap((a, _c, e) => { const g = e.toGrid(a[0]); const seen = new Set(); const out = []; for (const r of g) { const k = JSON.stringify(r.map((v) => (typeof v === 'string' ? v.toLowerCase() : v))); if (!seen.has(k)) { seen.add(k); out.push(r); } } return new Arr(out); }));
def('SORT', wrap((a, _c, e) => { const g = e.toGrid(a[0]).map((r) => [...r]); const idx = optNum(e, a[1], 1) - 1, ord = optNum(e, a[2], 1); g.sort((x, y) => cmp(x[idx], y[idx]) * ord); return new Arr(g); }));
def('FILTER', wrap((a, _c, e) => { const g = e.toGrid(a[0]); const m = e.toGrid(a[1]); const out = g.filter((_, i) => { const v = m[i]?.[0]; return v === true || (typeof v === 'number' && v !== 0); }); return out.length ? new Arr(out) : (a[2] !== undefined ? a[2] : E.CALC); }));
def('HYPERLINK', wrap((a, _c, e) => (a[1] != null ? e.deref(a[1]) : e.str(a[0]))));

// --- date & time
const today = () => { const d = new Date(); return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); };
const dnum = (e, v) => { const x = e.deref(v); if (typeof x === 'string') { const p = parseInput(x); if (typeof p.v === 'number') return p.v; fail(E.VALUE); } const n = e.num(x); if (n < 0) fail(E.NUM); return n; };
def('TODAY', () => today());
def('NOW', () => { const d = new Date(); return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()); });
def('DATE', wrap((a, _c, e) => { let y = e.int(a[0]); if (y < 1900) y += 1900; const v = dateToSerial(y, e.int(a[1]), e.int(a[2])); return v < 0 ? E.NUM : Math.round(v); }));
def('TIME', wrap((a, _c, e) => { const s = e.int(a[0]) * 3600 + e.int(a[1]) * 60 + e.int(a[2]); return (((s % 86400) + 86400) % 86400) / 86400; }));
def('YEAR', wrap((a, _c, e) => serialToParts(dnum(e, a[0])).y));
def('MONTH', wrap((a, _c, e) => serialToParts(dnum(e, a[0])).m));
def('DAY', wrap((a, _c, e) => serialToParts(dnum(e, a[0])).d));
def('HOUR', wrap((a, _c, e) => serialToParts(dnum(e, a[0])).H));
def('MINUTE', wrap((a, _c, e) => serialToParts(dnum(e, a[0])).M));
def('SECOND', wrap((a, _c, e) => serialToParts(dnum(e, a[0])).S));
def('WEEKDAY', wrap((a, _c, e) => {
  const wd = serialToParts(dnum(e, a[0])).wd; const t = optNum(e, a[1], 1);
  if (t === 1 || t === 17) return wd + 1; if (t === 2 || t === 11) return ((wd + 6) % 7) + 1; if (t === 3) return (wd + 6) % 7;
  if (t >= 12 && t <= 16) return ((wd - (t - 10) + 7) % 7) + 1;
  return E.NUM;
}));
def('WEEKNUM', wrap((a, _c, e) => {
  const s = Math.floor(dnum(e, a[0])); const t = optNum(e, a[1], 1);
  const p = serialToParts(s); const jan1 = dateToSerial(p.y, 1, 1);
  const start = t === 2 || t === 11 ? 1 : 0; const wd1 = (serialToParts(jan1).wd - start + 7) % 7;
  return Math.floor((s - jan1 + wd1) / 7) + 1;
}));
def('ISOWEEKNUM', wrap((a, _c, e) => { const s = Math.floor(dnum(e, a[0])); const p = serialToParts(s); const d = new Date(Date.UTC(p.y, p.m - 1, p.d)); const day = (d.getUTCDay() + 6) % 7; d.setUTCDate(d.getUTCDate() - day + 3); const fy = new Date(Date.UTC(d.getUTCFullYear(), 0, 4)); return 1 + Math.round(((d - fy) / 86400000 - 3 + ((fy.getUTCDay() + 6) % 7)) / 7); }));
def('EDATE', wrap((a, _c, e) => { const p = serialToParts(dnum(e, a[0])); const m = p.m + e.int(a[1]); const last = serialToParts(dateToSerial(p.y, m + 1, 0)).d; return Math.round(dateToSerial(p.y, m, Math.min(p.d, last))); }));
def('EOMONTH', wrap((a, _c, e) => { const p = serialToParts(dnum(e, a[0])); return Math.round(dateToSerial(p.y, p.m + e.int(a[1]) + 1, 0)); }));
def('DAYS', wrap((a, _c, e) => Math.floor(dnum(e, a[0])) - Math.floor(dnum(e, a[1]))));
def('DAYS360', wrap((a, _c, e) => { const s = serialToParts(dnum(e, a[0])), t = serialToParts(dnum(e, a[1])); let d1 = s.d === 31 ? 30 : s.d; let d2 = t.d === 31 && d1 >= 30 ? 30 : t.d; return (t.y - s.y) * 360 + (t.m - s.m) * 30 + (d2 - d1); }));
def('DATEDIF', wrap((a, _c, e) => {
  const s = Math.floor(dnum(e, a[0])), t = Math.floor(dnum(e, a[1])); if (s > t) return E.NUM;
  const u = e.str(a[2]).toUpperCase(); const p = serialToParts(s), q = serialToParts(t);
  let months = (q.y - p.y) * 12 + q.m - p.m; if (q.d < p.d) months--;
  switch (u) {
    case 'D': return t - s; case 'M': return months; case 'Y': return Math.floor(months / 12);
    case 'YM': return months % 12;
    case 'MD': { let d = q.d - p.d; if (d < 0) d += serialToParts(dateToSerial(q.y, q.m, 0)).d; return d; }
    case 'YD': { let ys = dateToSerial(q.y, p.m, p.d); if (ys > t) ys = dateToSerial(q.y - 1, p.m, p.d); return Math.round(t - ys); }
  }
  return E.NUM;
}));
function holidays(e, h) { return new Set(h == null ? [] : [...e.rangeValues(h)].filter((x) => typeof x === 'number').map(Math.floor)); }
function isWorkday(s, hol, weekend = [0, 6]) { const wd = serialToParts(s).wd; return !weekend.includes(wd) && !hol.has(s); }
def('NETWORKDAYS', wrap((a, _c, e) => { let s = Math.floor(dnum(e, a[0])), t = Math.floor(dnum(e, a[1])); const sign = s > t ? -1 : 1; if (s > t) [s, t] = [t, s]; const hol = holidays(e, a[2]); let n = 0; for (let d = s; d <= t; d++) if (isWorkday(d, hol)) n++; return n * sign; }));
def('WORKDAY', wrap((a, _c, e) => { let d = Math.floor(dnum(e, a[0])); let n = e.int(a[1]); const hol = holidays(e, a[2]); const st = Math.sign(n); while (n !== 0) { d += st; if (isWorkday(d, hol)) n -= st; } return d; }));
def('DATEVALUE', wrap((a, _c, e) => { const p = parseInput(e.str(a[0])); return typeof p.v === 'number' && p.nf && /y|d/.test(p.nf) ? Math.floor(p.v) : E.VALUE; }));
def('TIMEVALUE', wrap((a, _c, e) => { const p = parseInput(e.str(a[0])); return typeof p.v === 'number' && p.nf ? p.v - Math.floor(p.v) : E.VALUE; }));
def('YEARFRAC', wrap((a, _c, e) => { const s = dnum(e, a[0]), t = dnum(e, a[1]); const b = optNum(e, a[2], 0); const d = Math.abs(t - s); if (b === 1) return d / 365.25; if (b === 2) return d / 360; if (b === 3) return d / 365; return FUNCTIONS.DAYS360.fn([Math.min(s, t), Math.max(s, t)], null, e) / 360; }));

// --- financial
def('PMT', wrap((a, _c, e) => { const r = e.num(a[0]), n = e.num(a[1]), pv = e.num(a[2]), fv = optNum(e, a[3], 0), t = optNum(e, a[4], 0); if (r === 0) return -(pv + fv) / n; const f = Math.pow(1 + r, n); return -(r * (pv * f + fv)) / ((1 + r * t) * (f - 1)); }));
def('FV', wrap((a, _c, e) => { const r = e.num(a[0]), n = e.num(a[1]), p = e.num(a[2]), pv = optNum(e, a[3], 0), t = optNum(e, a[4], 0); if (r === 0) return -(pv + p * n); const f = Math.pow(1 + r, n); return -(pv * f + (p * (1 + r * t) * (f - 1)) / r); }));
def('PV', wrap((a, _c, e) => { const r = e.num(a[0]), n = e.num(a[1]), p = e.num(a[2]), fv = optNum(e, a[3], 0), t = optNum(e, a[4], 0); if (r === 0) return -(fv + p * n); const f = Math.pow(1 + r, n); return -(fv + (p * (1 + r * t) * (f - 1)) / r) / f; }));
def('NPER', wrap((a, _c, e) => { const r = e.num(a[0]), p = e.num(a[1]), pv = e.num(a[2]), fv = optNum(e, a[3], 0), t = optNum(e, a[4], 0); if (r === 0) return -(pv + fv) / p; const z = p * (1 + r * t) / r; return Math.log((z - fv) / (z + pv)) / Math.log(1 + r); }));
def('NPV', wrap((a, _c, e) => { const r = e.num(a[0]); return nums(e, a.slice(1)).reduce((s, v, i) => s + v / Math.pow(1 + r, i + 1), 0); }));
def('IRR', wrap((a, _c, e) => { const v = nums(e, [a[0]]); let r = optNum(e, a[1], 0.1); for (let i = 0; i < 100; i++) { let f = 0, df = 0; v.forEach((c, k) => { f += c / Math.pow(1 + r, k); df -= (k * c) / Math.pow(1 + r, k + 1); }); const nr = r - f / df; if (!isFinite(nr)) return E.NUM; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } return E.NUM; }));
def('RATE', wrap((a, _c, e) => { const n = e.num(a[0]), p = e.num(a[1]), pv = e.num(a[2]), fv = optNum(e, a[3], 0), t = optNum(e, a[4], 0); let r = optNum(e, a[5], 0.1); const F = (x) => { const f = Math.pow(1 + x, n); return pv * f + p * (1 + x * t) * (f - 1) / x + fv; }; for (let i = 0; i < 100; i++) { const h = 1e-7; const d = (F(r + h) - F(r - h)) / (2 * h); const nr = r - F(r) / d; if (!isFinite(nr)) return E.NUM; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } return E.NUM; }));

export const FUNCTION_NAMES = Object.keys(FUNCTIONS).sort();
export { E as Errors, Range, Arr, FormulaError };

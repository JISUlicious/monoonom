// Excel formula tokenizer, parser and reference rewriting.
// Grammar/precedence follow Excel: ':' > unary '-' > '%' > '^' > '* /' > '+ -' > '&' > comparisons.
import { colIndex, colName, MAX_COLS, MAX_ROWS } from './address.js';

const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_À-￿][\w.À-￿]*))!`;
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d+)`;
const REF_RE = new RegExp(String.raw`^(?:${SHEET})?(?:${CELL}(?::${CELL})?|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)(\d+):(\$?)(\d+))`);
const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'];

export class FormulaError extends Error {}

function refFromMatch(m) {
  const sheet = m[1] != null ? m[1].replace(/''/g, "'") : (m[2] ?? null);
  if (m[4]) {
    const a = { cAbs: !!m[3], c: colIndex(m[4]), rAbs: !!m[5], r: parseInt(m[6], 10) - 1 };
    const b = m[8] ? { cAbs: !!m[7], c: colIndex(m[8]), rAbs: !!m[9], r: parseInt(m[10], 10) - 1 } : null;
    if (a.c >= MAX_COLS || a.r < 0 || a.r >= MAX_ROWS || (b && (b.c >= MAX_COLS || b.r < 0 || b.r >= MAX_ROWS))) return null;
    return { kind: 'cell', sheet, a, b };
  }
  if (m[12]) {
    const a = { cAbs: !!m[11], c: colIndex(m[12]), rAbs: false, r: 0 };
    const b = { cAbs: !!m[13], c: colIndex(m[14]), rAbs: false, r: MAX_ROWS - 1 };
    if (a.c >= MAX_COLS || b.c >= MAX_COLS) return null;
    return { kind: 'col', sheet, a, b };
  }
  const a = { rAbs: !!m[15], r: parseInt(m[16], 10) - 1, cAbs: false, c: 0 };
  const b = { rAbs: !!m[17], r: parseInt(m[18], 10) - 1, cAbs: false, c: MAX_COLS - 1 };
  if (a.r < 0 || b.r < 0) return null;
  return { kind: 'row', sheet, a, b };
}

export function tokenize(src) {
  const s = src.startsWith('=') ? src.slice(1) : src;
  const off = src.length - s.length;
  const toks = [];
  let i = 0;
  const push = (t) => { toks.push(t); };
  while (i < s.length) {
    const ch = s[i];
    const start = i;
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '"') {
      let j = i + 1, v = '';
      for (;;) {
        if (j >= s.length) throw new FormulaError('문자열이 닫히지 않았습니다');
        if (s[j] === '"') { if (s[j + 1] === '"') { v += '"'; j += 2; continue; } break; }
        v += s[j++];
      }
      i = j + 1;
      push({ type: 'str', v, s: start + off, e: i + off });
      continue;
    }
    if (ch === '#') {
      const err = ERRORS.find((e) => s.slice(i, i + e.length).toUpperCase() === e);
      if (err) { i += err.length; push({ type: 'err', v: err, s: start + off, e: i + off }); continue; }
      if (/^#REF!/i.test(s.slice(i))) { i += 5; push({ type: 'err', v: '#REF!', s: start + off, e: i + off }); continue; }
    }
    // number
    let m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(s.slice(i));
    if (m && !(/^\d+:\$?\d/.test(s.slice(i)))) {
      // could be a row range like 1:3 — handled below
      const prev = toks[toks.length - 1];
      if (!(prev && prev.type === 'op' && prev.v === ':' && false)) {
        i += m[0].length;
        push({ type: 'num', v: parseFloat(m[0]), s: start + off, e: i + off });
        continue;
      }
    }
    // reference (with optional sheet)
    if (/[A-Za-z_$'À-￿\d]/.test(ch)) {
      m = REF_RE.exec(s.slice(i));
      if (m) {
        const after = s[i + m[0].length] ?? '';
        if (!/[\w(.À-￿!]/.test(after)) {
          const ref = refFromMatch(m);
          if (ref) {
            i += m[0].length;
            push({ type: 'ref', ref, s: start + off, e: i + off });
            continue;
          }
        }
      }
      // sheet!#REF!
      const sm = new RegExp(`^${SHEET}#REF!`).exec(s.slice(i));
      if (sm) { i += sm[0].length; push({ type: 'err', v: '#REF!', s: start + off, e: i + off }); continue; }
      const im = /^[A-Za-z_\\À-￿][\w.À-￿]*/.exec(s.slice(i));
      if (im) {
        i += im[0].length;
        let j = i;
        while (s[j] === ' ') j++;
        const name = im[0];
        if (s[j] === '(') { push({ type: 'fn', v: name.toUpperCase().replace(/^_XLFN\.(_XLWS\.)?|^_XLWS\./, ''), s: start + off, e: i + off }); i = j; continue; }
        const up = name.toUpperCase();
        if (up === 'TRUE' || up === 'FALSE') { push({ type: 'bool', v: up === 'TRUE', s: start + off, e: i + off }); continue; }
        push({ type: 'name', v: name, s: start + off, e: i + off });
        continue;
      }
    }
    const two = s.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') { i += 2; push({ type: 'op', v: two, s: start + off, e: i + off }); continue; }
    if ('+-*/^&=<>%:'.includes(ch)) { i++; push({ type: 'op', v: ch, s: start + off, e: i + off }); continue; }
    if ('(),;{}'.includes(ch)) { i++; push({ type: ch, s: start + off, e: i + off }); continue; }
    if (ch === '!') { i++; push({ type: 'op', v: '!', s: start + off, e: i + off }); continue; }
    throw new FormulaError(`알 수 없는 문자: ${ch}`);
  }
  return toks;
}

// ------------------------------------------------------------------ parser

export function parse(src) {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (type) => {
    const t = next();
    if (!t || t.type !== type) throw new FormulaError(`'${type}'가 필요합니다`);
    return t;
  };
  const isOp = (t, v) => t && t.type === 'op' && t.v === v;

  function primary() {
    const t = next();
    if (!t) throw new FormulaError('수식이 불완전합니다');
    switch (t.type) {
      case 'num': return { t: 'num', v: t.v };
      case 'str': return { t: 'str', v: t.v };
      case 'bool': return { t: 'bool', v: t.v };
      case 'err': return { t: 'err', v: t.v };
      case 'ref': return { t: 'ref', ref: t.ref };
      case 'name': return { t: 'name', v: t.v };
      case '(': { const e = expr(0); expect(')'); return { t: 'paren', e }; }
      case '{': return arrayConst();
      case 'fn': {
        expect('(');
        const args = [];
        if (peek() && peek().type === ')') { next(); return { t: 'fn', name: t.v, args }; }
        for (;;) {
          if (peek() && (peek().type === ',' || peek().type === ')')) args.push({ t: 'missing' });
          else args.push(expr(0));
          const sep = next();
          if (!sep) throw new FormulaError('괄호가 닫히지 않았습니다');
          if (sep.type === ')') break;
          if (sep.type !== ',' && sep.type !== ';') throw new FormulaError('인수 구분자가 필요합니다');
        }
        return { t: 'fn', name: t.v, args };
      }
      case 'op':
        if (t.v === '-' || t.v === '+') {
          const e = unaryOperand();
          return t.v === '-' ? { t: 'neg', e } : e;
        }
        break;
      default: break;
    }
    throw new FormulaError('잘못된 수식입니다');
  }

  function arrayConst() {
    const rows = [[]];
    for (;;) {
      const t = next();
      if (!t) throw new FormulaError('배열이 닫히지 않았습니다');
      if (t.type === '}') break;
      if (t.type === ',') continue;
      if (t.type === ';') { rows.push([]); continue; }
      let v;
      if (t.type === 'op' && t.v === '-') { const n = next(); v = -n.v; }
      else if (t.type === 'num' || t.type === 'str' || t.type === 'bool') v = t.v;
      else if (t.type === 'err') v = { err: t.v };
      else throw new FormulaError('배열 상수 오류');
      rows[rows.length - 1].push(v);
    }
    return { t: 'array', rows };
  }

  // range operator binds tightest
  function rangeExpr() {
    let left = primary();
    while (isOp(peek(), ':')) { next(); left = { t: 'range', a: left, b: primary() }; }
    return left;
  }
  function unaryOperand() {
    // unary minus binds tighter than ^ in Excel (-2^2 = 4)
    let e = rangeExpr();
    while (isOp(peek(), '%')) { next(); e = { t: 'pct', e }; }
    return e;
  }
  const BIN = { '^': [6], '*': [5], '/': [5], '+': [4], '-': [4], '&': [3],
    '=': [2], '<>': [2], '<': [2], '>': [2], '<=': [2], '>=': [2] };

  function expr(minPrec) {
    let left = unaryOperand();
    for (;;) {
      const t = peek();
      if (!t || t.type !== 'op' || !BIN[t.v]) break;
      const [prec, assoc] = BIN[t.v];
      if (prec < minPrec) break;
      next();
      const right = expr(assoc === 'R' ? prec : prec + 1);
      left = { t: 'bin', op: t.v, a: left, b: right };
    }
    return left;
  }

  if (!toks.length) throw new FormulaError('빈 수식');
  const ast = expr(0);
  if (p < toks.length) throw new FormulaError('수식 끝에 불필요한 항목이 있습니다');
  return ast;
}

// ------------------------------------------------------------------ serialization / rewriting

export function quoteSheetName(name) {
  return /^[A-Za-z_À-￿][\w.À-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    && !/^(TRUE|FALSE)$/i.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

export function refToString(ref) {
  const sh = ref.sheet != null ? quoteSheetName(ref.sheet) + '!' : '';
  const cp = (p) => (p.cAbs ? '$' : '') + colName(p.c);
  const rp = (p) => (p.rAbs ? '$' : '') + (p.r + 1);
  if (ref.kind === 'col') return `${sh}${cp(ref.a)}:${cp(ref.b)}`;
  if (ref.kind === 'row') return `${sh}${rp(ref.a)}:${rp(ref.b)}`;
  return sh + cp(ref.a) + rp(ref.a) + (ref.b ? ':' + cp(ref.b) + rp(ref.b) : '');
}

/** Rewrite every reference token in a formula with fn(ref) -> ref | null (null => #REF!). */
export function rewriteRefs(formula, fn) {
  let toks;
  try { toks = tokenize(formula); } catch { return formula; }
  let out = '', last = 0;
  for (const t of toks) {
    if (t.type !== 'ref') continue;
    const nr = fn(structuredClone(t.ref));
    out += formula.slice(last, t.s);
    if (nr === null) out += (t.ref.sheet != null ? quoteSheetName(t.ref.sheet) + '!' : '') + '#REF!';
    else out += refToString(nr);
    last = t.e;
  }
  return out + formula.slice(last);
}

/** Copy/fill semantics: shift relative parts by (dr, dc). */
export function translateFormula(formula, dr, dc) {
  if (!dr && !dc) return formula;
  return rewriteRefs(formula, (ref) => {
    for (const p of [ref.a, ref.b]) {
      if (!p) continue;
      if (ref.kind !== 'row' && !p.cAbs) p.c += dc;
      if (ref.kind !== 'col' && !p.rAbs) p.r += dr;
      if (p.c < 0 || p.r < 0 || p.c >= MAX_COLS || p.r >= MAX_ROWS) return null;
    }
    return ref;
  });
}

/**
 * Structural change: insert (count>0) or delete (count<0) rows/cols at `at` on sheet `target`.
 * `here` is the name of the sheet that owns the formula.
 */
export function adjustFormula(formula, here, target, axis, at, count) {
  const lc = (x) => (x ?? '').toLowerCase();
  return rewriteRefs(formula, (ref) => {
    const sheet = ref.sheet ?? here;
    if (lc(sheet) !== lc(target)) return ref;
    if ((axis === 'row' && ref.kind === 'col') || (axis === 'col' && ref.kind === 'row')) return ref;
    const k = axis === 'row' ? 'r' : 'c';
    const pts = ref.b ? [ref.a, ref.b] : [ref.a];
    if (count > 0) {
      for (const p of pts) if (p[k] >= at) p[k] += count;
      return ref;
    }
    const n = -count, end = at + n - 1;
    if (!ref.b) {
      if (ref.a[k] >= at && ref.a[k] <= end) return null;
      if (ref.a[k] > end) ref.a[k] -= n;
      return ref;
    }
    let lo = Math.min(ref.a[k], ref.b[k]), hi = Math.max(ref.a[k], ref.b[k]);
    if (lo >= at && hi <= end) return null;
    const nlo = lo > end ? lo - n : lo >= at ? at : lo;
    const nhi = hi > end ? hi - n : hi >= at ? at - 1 : hi;
    const [pa, pb] = ref.a[k] <= ref.b[k] ? [ref.a, ref.b] : [ref.b, ref.a];
    pa[k] = nlo; pb[k] = nhi;
    return ref;
  });
}

export function renameSheetInFormula(formula, oldName, newName) {
  const lc = oldName.toLowerCase();
  return rewriteRefs(formula, (ref) => {
    if (ref.sheet != null && ref.sheet.toLowerCase() === lc) ref.sheet = newName;
    return ref;
  });
}

/** F4 cycling for the reference at caret position. Returns {text, caret}. */
export function cycleAbsolute(text, caret) {
  let toks;
  try { toks = tokenize(text); } catch { return { text, caret }; }
  const t = toks.find((x) => x.type === 'ref' && caret >= x.s && caret <= x.e);
  if (!t) return { text, caret };
  const ref = t.ref;
  const step = (p) => {
    const state = (p.cAbs ? 2 : 0) + (p.rAbs ? 1 : 0); // 0 A1, 3 $A$1, 1 A$1, 2 $A1
    const order = [0, 3, 1, 2];
    const ns = order[(order.indexOf(state) + 1) % 4];
    p.cAbs = !!(ns & 2); p.rAbs = !!(ns & 1);
  };
  step(ref.a);
  if (ref.b) { ref.b.cAbs = ref.a.cAbs; ref.b.rAbs = ref.a.rAbs; }
  const s = refToString(ref);
  return { text: text.slice(0, t.s) + s + text.slice(t.e), caret: t.s + s.length };
}

/** For editor highlighting: list of refs with positions. */
export function refsWithPositions(text) {
  try {
    return tokenize(text).filter((t) => t.type === 'ref').map((t) => ({ ref: t.ref, s: t.s, e: t.e }));
  } catch { return []; }
}

export { ERRORS };

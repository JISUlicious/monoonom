// In-cell editor + formula bar: IME-safe typing, Enter/Edit/Point modes, F4, autocomplete.
import { key, MAX_ROWS, MAX_COLS, rangeA1, parseA1 } from './address.js';
import { tokenize, cycleAbsolute, refsWithPositions, refToString, quoteSheetName, parse } from './formula.js';
import { formatValue, isDateFormat, parseInput } from './numfmt.js';
import { FUNCTION_NAMES } from './engine.js';
import { FN_HELP } from './fnhelp.js';

const REF_COLORS = ['#1a73e8', '#d93025', '#7b1fa2', '#188038', '#e37400', '#00838f', '#c2185b'];

export class Editor {
  constructor(app, host, fxInput) {
    this.app = app;
    this.host = host;
    this.ta = document.createElement('textarea');
    this.ta.className = 'cell-editor idle';
    this.ta.spellcheck = false;
    this.ta.setAttribute('autocomplete', 'off');
    this.ta.setAttribute('aria-label', '셀 입력');
    host.append(this.ta);
    this.fx = fxInput;
    this.ac = document.createElement('div');
    this.ac.className = 'autocomplete hidden';
    document.body.append(this.ac);
    this.hint = document.createElement('div');
    this.hint.className = 'fn-hint hidden';
    document.body.append(this.hint);
    this.active = false;
    this.mode = 'enter';
    this.target = null;
    this.point = null; // {s, e, anchor:{r,c}, cur:{r,c}} reference being pointed
    this.acItems = [];
    this.acIndex = 0;
    this.source = this.ta;
    this.bind();
  }

  get grid() { return this.app.grid; }

  // ------------------------------------------------------------------ lifecycle
  editTextOf(sh, r, c) {
    const cell = sh.cells.get(key(r, c));
    if (!cell) return '';
    if (cell.f) return cell.f;
    const v = cell.v;
    if (v === undefined || v === null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') {
      const nf = cell.s?.nf;
      if (nf && isDateFormat(nf)) {
        const hasTime = /[hs]/i.test(nf.replace(/"[^"]*"/g, ''));
        const hasDate = /[yd]/i.test(nf.replace(/"[^"]*"/g, '')) || /m{3,}/.test(nf);
        const code = hasDate && hasTime ? 'yyyy-mm-dd h:mm:ss' : hasTime ? 'h:mm:ss' : 'yyyy-mm-dd';
        return formatValue(v, code).text;
      }
      if (nf && /%/.test(nf.replace(/"[^"]*"/g, ''))) return String(parseFloat((v * 100).toPrecision(15))) + '%';
      return String(parseFloat(v.toPrecision(15)));
    }
    if (typeof v === 'string' && (/^[=+-]/.test(v) && v.length > 1 || typeof parseInput(v).v !== 'string')) return "'" + v;
    return String(v);
  }

  start({ mode = 'enter', text, source } = {}) {
    if (this.active) return;
    const v = this.app.view();
    const sh = this.app.doc.wb.sheet;
    if (this.app.isProtectedEdit?.()) return;
    this.active = true;
    this.mode = mode;
    this.target = { sheet: sh, r: v.active.r, c: v.active.c };
    const mg = sh.mergeAt(v.active.r, v.active.c);
    if (mg) { this.target.r = mg.r1; this.target.c = mg.c1; }
    this.original = this.editTextOf(sh, this.target.r, this.target.c);
    this.source = source || this.ta;
    if (text !== undefined) this.setText(text);
    else if (mode === 'edit') this.setText(this.original);
    this.ta.classList.remove('idle');
    this.style();
    this.reposition();
    this.grid.render();
    this.app.onEditState?.(true);
    if (this.source === this.ta) {
      this.ta.focus();
      if (mode === 'edit') this.ta.setSelectionRange(this.ta.value.length, this.ta.value.length);
    }
    this.onInput();
  }

  setText(t) { this.ta.value = t; this.fx.value = t; }
  get text() { return this.source.value; }

  style() {
    const { sheet, r, c } = this.target;
    const st = sheet.cells.get(key(r, c))?.s || {};
    const z = this.grid.zoom;
    const f = this.grid.fontFor(st);
    this.ta.style.font = f.css;
    this.ta.style.fontSize = f.size * z + 'px';
    this.ta.style.color = st.fc || '#000';
    this.ta.style.background = st.bg || '#fff';
    this.ta.style.textAlign = st.ha === 'center' ? 'center' : st.ha === 'right' ? 'right' : 'left';
    this.ta.style.whiteSpace = st.wrap ? 'pre-wrap' : 'pre';
    this.wrap = !!st.wrap;
  }

  reposition() {
    if (!this.app.doc) return;
    if (!this.active) {
      // idle: park the (invisible) textarea on the active cell so IME candidate windows appear there
      const v = this.app.view();
      const r = this.grid.cellRect(v.active.r, v.active.c);
      this.ta.style.left = r.x + 'px'; this.ta.style.top = r.y + 'px';
      this.ta.style.width = '2px'; this.ta.style.height = r.h + 'px';
      return;
    }
    if (this.app.doc.wb.sheet !== this.target.sheet) { this.ta.style.left = '-9999px'; return; }
    const rc = this.grid.cellRect(this.target.r, this.target.c);
    const hostW = this.host.clientWidth, hostH = this.host.clientHeight;
    this.ta.style.left = rc.x + 'px';
    this.ta.style.top = rc.y + 'px';
    this.ta.style.minWidth = rc.w + 'px';
    this.ta.style.minHeight = rc.h + 'px';
    this.ta.style.width = rc.w + 'px';
    this.ta.style.height = rc.h + 'px';
    const maxW = Math.max(rc.w, hostW - rc.x - 18);
    if (!this.wrap) this.ta.style.width = Math.min(maxW, Math.max(rc.w, this.ta.scrollWidth + 4)) + 'px';
    this.ta.style.height = Math.min(hostH - rc.y - 18, Math.max(rc.h, this.ta.scrollHeight)) + 'px';
  }

  isFormula() { return this.active && this.text.startsWith('=') && this.text.length >= 1; }

  /** Can a click/arrow insert a reference at the caret? */
  isFormulaPointable() {
    if (!this.isFormula()) return false;
    const el = this.source;
    const pos = el.selectionStart;
    if (this.point && this.point.e === pos) return true;
    if (this.mode === 'edit' && !this.point && el.selectionEnd === pos && pos < el.value.length && false) return false;
    const before = el.value.slice(0, pos).replace(/\s+$/, '');
    return /[=(,;+\-*/^&<>:]$/.test(before);
  }

  /** insert/replace the pointed reference (from mouse) */
  pointRange(range, isStart, sheet) {
    const el = this.source;
    const sameSheet = sheet === this.target.sheet;
    let ref = rangeA1(range);
    if (range.r1 === 0 && range.r2 === MAX_ROWS - 1) ref = `${ref.replace(/\d+/g, '')}`.replace(/^(\w+):?(\w*)$/, (m, a, b) => `${a}:${b || a}`);
    else if (range.c1 === 0 && range.c2 === MAX_COLS - 1) ref = `${range.r1 + 1}:${range.r2 + 1}`;
    if (!sameSheet) ref = quoteSheetName(sheet.name) + '!' + ref;
    let s = el.selectionStart, e = el.selectionEnd;
    if (this.point && (!isStart || this.point.e === s)) { s = this.point.s; e = this.point.e; }
    const t = el.value;
    const nt = t.slice(0, s) + ref + t.slice(e);
    this.setText(nt);
    el.setSelectionRange(s + ref.length, s + ref.length);
    this.point = { s, e: s + ref.length, anchor: { r: range.r1, c: range.c1 }, cur: { r: range.r2, c: range.c2 }, sheet };
    el.focus();
    this.onInput(true);
  }

  /** Shift+click while pointing: extend the reference at the caret (the one just pointed, or one typed
   * right before the caret such as `=SUM(B2`) from its first cell to the clicked cell. */
  pointExtend(r, c, sheet) {
    if (!this.point) {
      const el = this.source, pos = el.selectionStart;
      const m = /(\$?[A-Za-z]{1,3}\$?\d+)$/.exec(el.value.slice(0, pos));
      const a = m && parseA1(m[1].replace(/\$/g, ''));
      if (!a) return false;
      this.point = { s: pos - m[1].length, e: pos, anchor: { r: a.r, c: a.c }, cur: { r: a.r, c: a.c }, sheet: this.target.sheet };
    }
    const a = this.point.anchor;
    this.pointRange({ r1: Math.min(a.r, r), c1: Math.min(a.c, c), r2: Math.max(a.r, r), c2: Math.max(a.c, c) }, false, sheet);
    this.point.anchor = a; this.point.cur = { r, c };
    return true;
  }

  /** Ctrl/Cmd+click while pointing: keep the current reference and start another one after a comma. */
  pointAddSeparator() {
    const el = this.source, pos = this.point ? this.point.e : el.selectionStart;
    const before = el.value.slice(0, pos).replace(/\s+$/, '');
    if (/[=(,;]$/.test(before)) { this.point = null; return; } // nothing to separate from yet
    const t = el.value;
    this.setText(t.slice(0, pos) + ',' + t.slice(pos));
    el.setSelectionRange(pos + 1, pos + 1);
    this.point = null;
  }

  pointMove(dr, dc, extend) {
    const sh = this.app.doc.wb.sheet;
    let p = this.point;
    let anchor, cur;
    if (!p) { anchor = { r: this.target.r, c: this.target.c }; cur = { ...anchor }; }
    else { anchor = p.anchor; cur = p.cur; }
    const nc = { r: Math.min(MAX_ROWS - 1, Math.max(0, cur.r + dr)), c: Math.min(MAX_COLS - 1, Math.max(0, cur.c + dc)) };
    const na = extend ? anchor : nc;
    const range = { r1: Math.min(na.r, nc.r), c1: Math.min(na.c, nc.c), r2: Math.max(na.r, nc.r), c2: Math.max(na.c, nc.c) };
    this.pointRange(range, !p, sh);
    this.point.anchor = na; this.point.cur = nc;
    this.grid.scrollTo(nc.r, nc.c);
  }

  commit({ move = null, fillSelection = false } = {}) {
    if (!this.active) return true;
    let text = this.text;
    // unchanged text is committed as is: a formula from the file may use syntax this parser cannot
    // read (e.g. LAMBDA(x,x*2)(A1)) and must not trap the user in the editor
    if (text.startsWith('=') && text.length > 1 && text !== this.original) {
      // auto-close parentheses like Excel
      let depth = 0, inStr = false;
      for (const ch of text) { if (ch === '"') inStr = !inStr; else if (!inStr) { if (ch === '(') depth++; else if (ch === ')') depth--; } }
      if (inStr) text += '"';
      while (depth-- > 0) text += ')';
      try { parse(text); } catch (e) {
        this.app.toast(`수식에 문제가 있습니다: ${e.message}`, 'error');
        return false;
      }
      text = normalizeFormula(text);
    }
    const { sheet, r, c } = this.target;
    this.finish();
    if (text !== this.original || fillSelection) this.app.setCellInput(sheet, r, c, text, { fillSelection });
    if (move) this.app.moveActive(move.dr, move.dc, { withinSelection: true });
    return true;
  }

  cancel() { if (!this.active) return; this.finish(); }

  finish() {
    this.active = false;
    this.point = null;
    this.ta.value = '';
    this.ta.classList.add('idle');
    this.ta.style.width = '2px';
    this.grid.refHighlights = [];
    this.hideAc();
    this.hint.classList.add('hidden');
    if (this.target && this.app.doc.wb.sheet !== this.target.sheet) {
      this.app.activateSheet(this.app.doc.wb.sheets.indexOf(this.target.sheet), { keepEditor: true });
    }
    this.app.onEditState?.(false);
    this.app.syncFormulaBar();
    this.app.focusGrid();
    this.grid.render();
  }

  // ------------------------------------------------------------------ events
  bind() {
    const ta = this.ta;
    ta.addEventListener('compositionstart', () => { if (!this.active) this.start({ mode: 'enter' }); });
    ta.addEventListener('input', () => {
      if (!this.active) { this.start({ mode: 'enter', text: ta.value }); return; }
      if (this.source !== ta) return;
      this.fx.value = ta.value;
      this.point = null;
      this.onInput();
    });
    ta.addEventListener('keydown', (e) => this.onKey(e, ta));
    this.fx.addEventListener('focus', () => {
      if (!this.active) this.start({ mode: 'edit', source: this.fx });
      else this.source = this.fx;
    });
    this.fx.addEventListener('input', () => {
      if (!this.active) this.start({ mode: 'edit', source: this.fx });
      this.source = this.fx;
      this.ta.value = this.fx.value;
      this.point = null;
      this.onInput();
    });
    this.fx.addEventListener('keydown', (e) => this.onKey(e, this.fx));
    this.ac.addEventListener('mousedown', (e) => {
      const it = e.target.closest('[data-fn]');
      if (it) { e.preventDefault(); this.acceptAc(it.dataset.fn); }
    });
  }

  onKey(e, el) {
    if (!this.active) return; // idle keys are handled by the app shortcut layer
    if (e.isComposing || e.keyCode === 229) return;
    const k = e.key;
    const acOpen = !this.ac.classList.contains('hidden');
    if (acOpen && (k === 'ArrowDown' || k === 'ArrowUp')) {
      e.preventDefault(); this.acIndex = (this.acIndex + (k === 'ArrowDown' ? 1 : -1) + this.acItems.length) % this.acItems.length; this.renderAc(); return;
    }
    if (acOpen && k === 'Tab') { e.preventDefault(); this.acceptAc(this.acItems[this.acIndex]); return; }
    if (acOpen && k === 'Escape') { e.preventDefault(); this.hideAc(); return; }
    if (k === 'Enter') {
      e.preventDefault();
      if (e.altKey) { this.insertAtCaret(el, '\n'); return; }
      if (e.ctrlKey || e.metaKey) { this.commit({ fillSelection: true }); return; }
      this.commit({ move: { dr: e.shiftKey ? -1 : 1, dc: 0 } });
      return;
    }
    if (k === 'Tab') { e.preventDefault(); this.commit({ move: { dr: 0, dc: e.shiftKey ? -1 : 1 } }); return; }
    if (k === 'Escape') { e.preventDefault(); this.cancel(); return; }
    if (k === 'F4') {
      e.preventDefault();
      const res = cycleAbsolute(el.value, el.selectionStart);
      this.setText(res.text); el.setSelectionRange(res.caret, res.caret); this.onInput(); return;
    }
    if (k === 'F2') { e.preventDefault(); this.mode = this.mode === 'edit' ? 'enter' : 'edit'; this.app.setStatus(this.mode === 'edit' ? '편집' : '입력'); return; }
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) {
      const [dr, dc] = arrows[k];
      if (this.isFormulaPointable() && (this.mode === 'enter' || this.point)) { e.preventDefault(); this.pointMove(dr, dc, e.shiftKey); return; }
      if (this.mode === 'enter' && el === this.ta && !e.shiftKey) { e.preventDefault(); this.commit({ move: { dr, dc } }); return; }
      return;
    }
    if ((e.ctrlKey || e.metaKey) && (k === 'z' || k === 'Z') && !e.shiftKey) { if (el.value === this.original) { e.preventDefault(); this.cancel(); } return; }
    if ((e.ctrlKey || e.metaKey) && k === ';') { e.preventDefault(); const d = new Date(); this.insertAtCaret(el, `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`); }
  }

  insertAtCaret(el, s) {
    const a = el.selectionStart, b = el.selectionEnd;
    el.value = el.value.slice(0, a) + s + el.value.slice(b);
    el.setSelectionRange(a + s.length, a + s.length);
    if (el === this.ta) this.fx.value = el.value; else this.ta.value = el.value;
    this.onInput();
  }

  onInput(fromPoint = false) {
    this.reposition();
    const text = this.text;
    // reference highlights
    if (text.startsWith('=')) {
      const refs = refsWithPositions(text);
      const wb = this.app.doc.wb;
      this.grid.refHighlights = refs.map((x, i) => {
        const sh = x.ref.sheet ? wb.sheetByName(x.ref.sheet) : this.target.sheet;
        const b = x.ref.b || x.ref.a;
        return { sheet: sh, color: REF_COLORS[i % REF_COLORS.length], range: { r1: Math.min(x.ref.a.r, b.r), c1: Math.min(x.ref.a.c, b.c), r2: Math.max(x.ref.a.r, b.r), c2: Math.max(x.ref.a.c, b.c) } };
      }).filter((h) => h.sheet);
      if (!fromPoint) this.updateAc();
      this.updateHint();
    } else { this.grid.refHighlights = []; this.hideAc(); this.hint.classList.add('hidden'); }
    this.grid.render();
  }

  // ------------------------------------------------------------------ autocomplete & hints
  updateAc() {
    const el = this.source;
    const before = el.value.slice(0, el.selectionStart);
    const m = /(?:^=|[=(,;+\-*/^&<>:\s])([A-Za-z][A-Za-z0-9.]*)$/.exec(before);
    if (!m) return this.hideAc();
    const pre = m[1].toUpperCase();
    const items = FUNCTION_NAMES.filter((n) => n.startsWith(pre)).slice(0, 12);
    if (!items.length || (items.length === 1 && items[0] === pre)) return this.hideAc();
    this.acItems = items; this.acIndex = 0; this.acPrefix = m[1];
    this.renderAc();
  }
  renderAc() {
    const rect = this.source.getBoundingClientRect();
    this.ac.innerHTML = this.acItems.map((n, i) => `<div class="ac-item${i === this.acIndex ? ' on' : ''}" data-fn="${n}"><b>${n}</b><span>${FN_HELP[n]?.[2] || ''}</span></div>`).join('');
    this.ac.style.left = rect.left + 'px';
    this.ac.style.top = rect.bottom + 2 + 'px';
    this.ac.classList.remove('hidden');
  }
  hideAc() { this.ac.classList.add('hidden'); this.acItems = []; }
  acceptAc(name) {
    const el = this.source;
    const pos = el.selectionStart;
    const start = pos - this.acPrefix.length;
    el.value = el.value.slice(0, start) + name + '(' + el.value.slice(pos);
    const np = start + name.length + 1;
    el.setSelectionRange(np, np);
    if (el === this.ta) this.fx.value = el.value; else this.ta.value = el.value;
    this.hideAc();
    this.onInput(true);
  }
  updateHint() {
    const el = this.source;
    const before = el.value.slice(0, el.selectionStart);
    // find the innermost open function call and which argument the caret is in
    let depth = 0, name = null, argIdx = 0, inStr = false;
    for (let i = before.length - 1; i >= 0; i--) {
      const ch = before[i];
      if (ch === '"') inStr = !inStr;
      if (inStr) continue;
      if (ch === ')' || ch === '}') depth++;
      else if (ch === '{') depth--;
      else if (ch === '(') {
        if (depth === 0) { const m = /([A-Za-z][A-Za-z0-9.]*)$/.exec(before.slice(0, i)); name = m ? m[1].toUpperCase() : null; break; }
        depth--;
      } else if (ch === ',' && depth === 0) argIdx++;
    }
    const h = name && FN_HELP[name];
    if (!h || !this.ac.classList.contains('hidden')) { this.hint.classList.add('hidden'); return; }
    this.hint.innerHTML = renderSignature(h[1], argIdx) + `<div class="fn-hint-desc">${escHtml(h[2] || '')}</div>`;
    const rect = this.source.getBoundingClientRect();
    this.hint.style.left = rect.left + 'px';
    this.hint.style.top = rect.bottom + 2 + 'px';
    this.hint.classList.remove('hidden');
  }
}

const escHtml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Parameters of 'NAME(a, [b], [c, d, ...], ...)', split on top-level commas only. */
export function signatureParams(sig) {
  const inner = sig.slice(sig.indexOf('(') + 1, sig.lastIndexOf(')'));
  const out = [];
  let depth = 0, cur = '';
  for (const ch of inner) {
    if (ch === '[') depth++;
    else if (ch === ']') depth--;
    if (ch === ',' && depth === 0) { out.push(cur.trim()); cur = ''; } else cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}

/** The parameter the argument at argIdx stands for. Repeating arguments ('...') are named like Excel
 * does: SUM's 3rd argument is [number3], SUMIFS's 4th is criteria_range2. */
export function paramForArg(params, argIdx) {
  const ell = params[params.length - 1] === '...';
  const base = ell ? params.slice(0, -1) : params;
  if (argIdx < base.length || !ell) return { list: base, idx: Math.min(argIdx, base.length - 1), extra: null };
  const num = (p) => { const m = /(\d+)\]?$/.exec(p); return m ? +m[1] : null; };
  const maxN = Math.max(...base.map((p) => num(p) ?? -1));
  let k = base.length; while (k > 0 && num(base[k - 1]) === maxN) k--;
  const unit = base.slice(k);
  if (!unit.length || maxN < 0) return { list: base, idx: base.length - 1, extra: null };
  const off = argIdx - k, rep = Math.floor(off / unit.length), p = unit[off % unit.length];
  const name = p.replace(/^\[|\]$/g, '').replace(/\d+$/, String(maxN + rep));
  return { list: base, idx: -1, extra: `[${name}]` };
}

function renderSignature(sig, argIdx) {
  const name = sig.slice(0, sig.indexOf('('));
  const params = signatureParams(sig);
  const { list, idx, extra } = paramForArg(params, argIdx);
  const parts = list.map((p, i) => (i === idx ? `<b>${escHtml(p)}</b>` : escHtml(p)));
  if (extra) parts.push(`<b>${escHtml(extra)}</b>`);
  if (params[params.length - 1] === '...') parts.push('...');
  return `${escHtml(name)}(${parts.join(', ')})`;
}

/** Uppercase function names and references (Excel normalizes formula text). */
export function normalizeFormula(text) {
  let toks;
  try { toks = tokenize(text); } catch { return text; }
  let out = '', last = 0;
  for (const t of toks) {
    if (t.type === 'fn' || t.type === 'ref' || t.type === 'bool') {
      out += text.slice(last, t.s);
      const raw = text.slice(t.s, t.e);
      out += t.type === 'ref' ? refToString(t.ref) : raw.toUpperCase();
      last = t.e;
    }
  }
  return out + text.slice(last);
}

// Application controller: documents/tabs/windows, commands, ribbon, keyboard shortcuts.
import { key, keyRow, keyCol, MAX_ROWS, MAX_COLS, a1, rangeA1, parseRange, parseSheetRef, colName, normRange, inRange } from './address.js';
import { Workbook, blankWorkbook, internStyle, newId } from './model.js';
import { translateFormula, quoteSheetName } from './formula.js';
import { formatValue, parseInput, isDateFormat, dateToSerial } from './numfmt.js';
import { Grid, ptToPx } from './grid.js';
import { Editor } from './editor.js';
import { Clipboard } from './clipboard.js';
import { Filters } from './filter.js';
import { ChartLayer } from './charts.js';
import { DocSync, EventStream } from './docsync.js';
import { api } from './api.js';
import { geom } from './geometry.js';
import * as D from './dialogs.js';

const LISTS = [
  ['일', '월', '화', '수', '목', '금', '토'],
  ['일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일'],
  ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['1분기', '2분기', '3분기', '4분기'], ['갑', '을', '병', '정', '무', '기', '경', '신', '임', '계'],
];

class App {
  constructor() {
    this.docs = [];
    this.di = -1;
    this.$ = (s) => document.querySelector(s);
    this.gridHost = this.$('#grid-host');
    this.grid = new Grid(this, this.gridHost);
    this.editor = new Editor(this, this.gridHost, this.$('#fx-input'));
    this.clipboard = new Clipboard(this);
    this.filters = new Filters(this);
    this.charts = new ChartLayer(this, this.grid);
    this.clientId = newId('c');
    this.stream = new EventStream(this);
    this.lastAction = null;
    this.painter = null;
    this.bindUI();
  }

  // ------------------------------------------------------------------ docs & views
  get doc() { return this.docs[this.di]; }
  view(sheet) {
    const doc = this.doc;
    if (!doc) return null;
    sheet = sheet || doc.wb.sheet;
    let v = doc.views.get(sheet.id);
    if (!v) {
      v = { sel: { r1: 0, c1: 0, r2: 0, c2: 0 }, active: { r: 0, c: 0 }, anchor: { r: 0, c: 0 }, zoom: 1, showFormulas: false, scrollTop: 0, scrollLeft: 0, extR: 0, extC: 0 };
      doc.views.set(sheet.id, v);
    }
    return v;
  }

  /** Show a server-owned document in this window (doc = {wb, views, sync}). */
  addDoc(wb, { activate = true, version = 0 } = {}) {
    const doc = { wb, views: new Map(), id: newId('d') };
    wb.on((ev) => this.onWbChange(doc, ev));
    doc.sync = new DocSync(this, doc, version);
    this.docs.push(doc);
    doc.sync.connect();
    doc.sync.scheduleCalc();
    if (activate) this.activateDoc(this.docs.length - 1);
    if (wb.warnings?.length) this.toast(wb.warnings.join('\n'), 'warn', 8000);
    return doc;
  }

  /** Attach a document that is open on the server (by id). */
  async attachDoc(docId, { activate = true } = {}) {
    const i = this.docs.findIndex((d) => d.wb.docId === docId);
    if (i >= 0) { if (activate) this.activateDoc(i); return this.docs[i]; }
    this._attaching ||= new Map();
    if (!this._attaching.has(docId)) {
      this._attaching.set(docId, (async () => {
        try {
          const snap = await api.snapshot(docId);
          const wb = Workbook.fromJSON(snap);
          wb.docId = snap.docId; wb.dirty = !!snap.dirty; wb.basePath = snap.path;
          return this.addDoc(wb, { activate: false, version: snap.version });
        } finally { this._attaching.delete(docId); }
      })());
    }
    const doc = await this._attaching.get(docId);
    if (doc && (activate || !this.doc || this.di < 0)) this.activateDoc(this.docs.indexOf(doc));
    return doc;
  }

  /** Replace the in-memory workbook with a fresh server snapshot (after a conflict). */
  replaceDoc(doc, snap) {
    const wb = Workbook.fromJSON(snap);
    wb.docId = snap.docId; wb.dirty = !!snap.dirty; wb.basePath = snap.path;
    wb.on((ev) => this.onWbChange(doc, ev));
    doc.wb = wb;
    if (doc === this.doc) { this.restoreScroll(); this.render(); } else this.renderTabs();
  }

  /** Remove a document from this window only. */
  detachDoc(doc) {
    const i = this.docs.indexOf(doc);
    if (i < 0) return;
    doc.sync.close();
    this.docs.splice(i, 1);
    if (!this.docs.length) { this.di = -1; this.newDoc(); return; }
    if (i === this.di || this.di >= this.docs.length) { this.di = -1; this.activateDoc(Math.min(i, this.docs.length - 1)); }
    else { if (i < this.di) this.di--; this.renderTabs(); }
  }

  onDocClosedRemotely(doc) {
    if (!this.docs.includes(doc)) return;
    this.toast(`'${doc.wb.name}' 문서가 닫혔습니다.`);
    this.detachDoc(doc);
  }

  saveScroll() {
    const v = this.view();
    if (v) { v.scrollTop = this.grid.scroller.scrollTop; v.scrollLeft = this.grid.scroller.scrollLeft; }
  }
  restoreScroll() {
    const v = this.view();
    this.grid.updateExtent(this.grid.metrics());
    this.grid.scroller.scrollTop = v.scrollTop; this.grid.scroller.scrollLeft = v.scrollLeft;
  }

  activateDoc(i) {
    if (this.editor.active && !this.editor.commit()) return;
    if (this.doc) this.saveScroll();
    this.di = i;
    this.clipboard.state && this.clipboard.state.wb !== this.doc.wb && null;
    this.restoreScroll();
    this.render();
    document.title = `${this.doc.wb.name} - Web Excel`;
    this.focusGrid();
  }

  activateSheet(i, { keepEditor = false } = {}) {
    const wb = this.doc.wb;
    if (i < 0 || i >= wb.sheets.length) return;
    if (!keepEditor && this.editor.active) {
      if (this.editor.isFormulaPointable() || this.editor.isFormula()) { /* pointing across sheets */ }
      else if (!this.editor.commit()) return;
    }
    this.saveScroll();
    wb.active = i;
    this.restoreScroll();
    this.render();
    if (!this.editor.active) this.focusGrid();
  }

  async closeDoc(i) {
    const doc = this.docs[i];
    if (doc.wb.dirty) {
      const ans = await D.confirmBox(`'${doc.wb.name}'의 변경 내용을 저장하시겠습니까?`, { yes: '저장', no: '저장 안 함', cancel: true });
      if (ans === null) return;
      if (ans === 'yes') { this.di = i; if (!(await this.save())) return; }
    }
    try { await api.closeDoc(doc.wb.docId, true); } catch (e) { if (e.status !== 404) { this.toast(e.message, 'error'); return; } }
    this.detachDoc(doc);
  }

  nextBlankName() {
    let n = 1;
    while (this.docs.some((d) => d.wb.name === `통합 문서${n}.xlsx`)) n++;
    return `통합 문서${n}.xlsx`;
  }

  onWbChange(doc, ev) {
    if (ev.type !== 'change') return;
    if (ev.entry) doc.sync.send(ev.entry);
    doc.sync.scheduleCalc();
    if (doc === this.doc) { this.render(); }
    else this.renderTabs();
  }

  /** ops from another client (agent CLI or another window) were applied */
  onRemoteChange(doc, ev, boxes) {
    if (doc !== this.doc) { this.renderTabs(); return; }
    const now = performance.now();
    this.grid.flashes = (this.grid.flashes || []).filter((f) => f.until > now).concat(boxes.map((b) => ({ ...b, until: now + 1600 })));
    if (boxes.length) setTimeout(() => this.grid.render(), 1700);
    const who = ev.actor === 'agent' ? '에이전트' : ev.actor === 'system' ? '' : '다른 창';
    if (who && (boxes.length || ev.label)) {
      const where = boxes.slice(0, 3).map((b) => `${b.sheet.name}!${rangeA1(b)}`).join(', ');
      this.setStatus(`${who}: ${ev.label || '변경'}${where ? ` (${where})` : ''}`);
    }
    this.render();
  }

  // ------------------------------------------------------------------ rendering
  render() {
    if (!this.doc) return;
    this.grid.render();
    this.renderTabs();
    this.renderSheetTabs();
    this.syncFormulaBar();
    this.updateStatus();
    this.updateRibbon();
  }

  renderTabs() {
    const bar = this.$('#doc-tabs');
    bar.innerHTML = this.docs.map((d, i) => `<div class="doc-tab${i === this.di ? ' on' : ''}" data-i="${i}" title="${D.esc(d.wb.path || d.wb.name)}">
      <span class="dt-ic">${d.wb.format === 'xls' ? '📗' : '📊'}</span><span class="dt-name">${D.esc(d.wb.name)}</span>${d.wb.dirty ? '<span class="dt-dirty" title="저장되지 않은 변경 내용">●</span>' : ''}<button class="dt-x" data-close="${i}" title="닫기 (Ctrl+F4)">×</button></div>`).join('')
      + '<button class="doc-new" data-cmd="new" title="새 통합 문서 (Ctrl+N)">＋</button>';
    if (this.doc) document.title = `${this.doc.wb.dirty ? '● ' : ''}${this.doc.wb.name} - Web Excel`;
  }

  renderSheetTabs() {
    const wb = this.doc.wb;
    const box = this.$('#sheet-tabs');
    box.innerHTML = wb.sheets.map((s, i) => s.state !== 'visible' ? '' : `<div class="sheet-tab${i === wb.active ? ' on' : ''}" data-i="${i}" draggable="true" style="${s.tabColor ? `--tab:${s.tabColor}` : ''}">${D.esc(s.name)}</div>`).join('');
  }

  syncFormulaBar() {
    if (!this.doc) return;
    const v = this.view(), sh = this.doc.wb.sheet;
    const sel = v.sel;
    const single = sel.r1 === sel.r2 && sel.c1 === sel.c2;
    const mg = sh.mergeAt(v.active.r, v.active.c);
    let nb = a1(v.active.r, v.active.c);
    if (!single && !(mg && mg.r1 === sel.r1 && mg.c1 === sel.c1 && mg.r2 === sel.r2 && mg.c2 === sel.c2) && this.grid.drag) nb = `${sel.r2 - sel.r1 + 1}R x ${sel.c2 - sel.c1 + 1}C`;
    const nameBox = this.$('#name-box');
    if (document.activeElement !== nameBox) nameBox.value = nb;
    if (!this.editor.active) this.$('#fx-input').value = this.editor.editTextOf(sh, mg ? mg.r1 : v.active.r, mg ? mg.c1 : v.active.c);
  }

  updateStatus() {
    const wb = this.doc.wb, sh = wb.sheet, sel = this.view().sel;
    const g = { ...sel, r2: Math.min(sel.r2, sh.maxRow), c2: Math.min(sel.c2, sh.maxCol) };
    let count = 0, ncount = 0, sum = 0, min = Infinity, max = -Infinity;
    const area = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);
    const visit = (r, c) => {
      if (sh.isRowHidden(r)) return;
      const v = wb.value(sh, r, c);
      if (v === null || v === undefined || v === '') return;
      count++;
      if (typeof v === 'number') { ncount++; sum += v; if (v < min) min = v; if (v > max) max = v; }
    };
    if (area > 1 && g.r2 >= g.r1 && g.c2 >= g.c1) {
      if (area > 4 * sh.cells.size) { for (const k of sh.cells.keys()) { const r = keyRow(k), c = keyCol(k); if (inRange(g, r, c)) visit(r, c); } }
      else for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) visit(r, c);
    }
    const nf = sh.get(this.view().active.r, this.view().active.c)?.s?.nf;
    const f = (x) => formatValue(parseFloat(x.toPrecision(15)), nf && !isDateFormat(nf) ? nf : 'General').text;
    const stats = this.$('#status-stats');
    stats.innerHTML = count > 1 ? (ncount ? `<span>평균: ${f(sum / ncount)}</span><span>개수: ${count}</span><span>숫자 셀 수: ${ncount}</span><span>최소값: ${f(min)}</span><span>최대값: ${f(max)}</span><span>합계: ${f(sum)}</span>` : `<span>개수: ${count}</span>`) : '';
    const circ = wb.engine.circular.size ? `<span class="warn">순환 참조: ${[...wb.engine.circular].slice(0, 3).join(', ')}</span>` : '';
    const filt = sh.autoFilter && sh.filterHidden.size ? `<span>${(sh.autoFilter.range.r2 - sh.autoFilter.range.r1) - sh.filterHidden.size}/${sh.autoFilter.range.r2 - sh.autoFilter.range.r1}개 레코드 표시</span>` : '';
    this.$('#status-msg').innerHTML = (this.statusText ? `<span>${D.esc(this.statusText)}</span>` : '<span>준비</span>') + circ + filt;
    this.$('#zoom-label').textContent = Math.round(this.view().zoom * 100) + '%';
    this.$('#zoom-range').value = Math.round(this.view().zoom * 100);
  }

  setStatus(t) { this.statusText = t; this.updateStatus(); clearTimeout(this._st); this._st = setTimeout(() => { this.statusText = ''; this.updateStatus(); }, 5000); }

  updateRibbon() {
    const v = this.view(), sh = this.doc.wb.sheet;
    const st = sh.get(v.active.r, v.active.c)?.s || {};
    const on = (cmd, b) => document.querySelectorAll(`[data-cmd="${cmd}"]`).forEach((el) => el.classList.toggle('on', !!b));
    on('bold', st.b); on('italic', st.i); on('underline', st.u); on('strike', st.st);
    on('alignLeft', st.ha === 'left'); on('alignCenter', st.ha === 'center'); on('alignRight', st.ha === 'right');
    on('vTop', st.va === 'top'); on('vMiddle', st.va === 'center'); on('vBottom', !st.va || st.va === 'bottom');
    on('wrap', st.wrap); on('filter', !!sh.autoFilter); on('gridlines', sh.showGrid !== false); on('showFormulas', v.showFormulas);
    on('mergeCenter', !!sh.mergeAt(v.active.r, v.active.c));
    const fn = this.$('#font-name'), fs = this.$('#font-size'), nf = this.$('#num-format');
    if (document.activeElement !== fn) fn.value = st.ff || this.doc.wb.defaultFont.name;
    if (document.activeElement !== fs) fs.value = st.fs || this.doc.wb.defaultFont.size;
    if (document.activeElement !== nf) {
      const code = st.nf || 'General';
      const opt = [...nf.options].find((o) => o.value === code);
      if (opt) nf.value = code; else { nf.value = '__custom'; nf.querySelector('[value=__custom]').textContent = `사용자 지정 (${code})`; }
    }
    this.$('#undo-btn').disabled = !this.doc.wb.undoStack.length;
    this.$('#redo-btn').disabled = !this.doc.wb.redoStack.length;
    this.$('#undo-btn').title = this.doc.wb.undoStack.length ? `실행 취소: ${this.doc.wb.undoStack.at(-1).label} (Ctrl+Z)` : '실행 취소 (Ctrl+Z)';
  }

  // ------------------------------------------------------------------ selection
  expandMerges(g) {
    const sh = this.doc.wb.sheet;
    if (!sh.merges.length) return g;
    let changed = true;
    g = { ...g };
    while (changed) {
      changed = false;
      for (const m of sh.merges) {
        if (m.r2 >= g.r1 && m.r1 <= g.r2 && m.c2 >= g.c1 && m.c1 <= g.c2) {
          if (m.r1 < g.r1) { g.r1 = m.r1; changed = true; } if (m.r2 > g.r2) { g.r2 = m.r2; changed = true; }
          if (m.c1 < g.c1) { g.c1 = m.c1; changed = true; } if (m.c2 > g.c2) { g.c2 = m.c2; changed = true; }
        }
      }
    }
    return g;
  }
  select(r, c, { scroll = true } = {}) {
    const v = this.view();
    r = Math.max(0, Math.min(MAX_ROWS - 1, r)); c = Math.max(0, Math.min(MAX_COLS - 1, c));
    v.active = { r, c }; v.anchor = { r, c };
    v.sel = this.expandMerges({ r1: r, c1: c, r2: r, c2: c });
    if (scroll) this.grid.scrollTo(r, c);
    this.afterSelect();
  }
  selectRange(g, active = null, { anchor = null, scroll = true, expandMerges = true } = {}) {
    const v = this.view();
    g = normRange(g);
    v.sel = expandMerges ? this.expandMerges(g) : g;
    if (active) v.active = { ...active };
    else v.active = { r: g.r1, c: g.c1 };
    v.anchor = anchor ? { ...anchor } : { ...v.active };
    if (scroll) this.grid.scrollTo(v.active.r, v.active.c);
    this.afterSelect();
  }
  extendTo(r, c) {
    const v = this.view();
    r = Math.max(0, Math.min(MAX_ROWS - 1, r)); c = Math.max(0, Math.min(MAX_COLS - 1, c));
    v.sel = this.expandMerges(normRange({ r1: v.anchor.r, c1: v.anchor.c, r2: r, c2: c }));
    v.ext = { r, c };
    this.grid.scrollTo(r, c);
    this.afterSelect();
  }
  afterSelect() {
    this.doc?.sync?.presence(this.doc.wb.sheet, this.view());
    this.grid.render();
    this.syncFormulaBar();
    this.updateStatus();
    this.updateRibbon();
    this.editor.reposition();
  }
  onSelectionDone() {
    if (this.painter) {
      const p = this.painter; this.painter = p.sticky ? p : null;
      this.applyPainter(p);
      if (!this.painter) document.body.classList.remove('painting');
    }
  }

  /** move active cell; within a multi-cell selection Enter/Tab cycles inside it (Excel behaviour) */
  moveActive(dr, dc, { withinSelection = false } = {}) {
    const v = this.view(), sh = this.doc.wb.sheet, sel = v.sel;
    const multi = !(sel.r1 === sel.r2 && sel.c1 === sel.c2) && !(() => { const m = sh.mergeAt(sel.r1, sel.c1); return m && m.r1 === sel.r1 && m.c1 === sel.c1 && m.r2 === sel.r2 && m.c2 === sel.c2; })();
    if (withinSelection && multi) {
      let { r, c } = v.active;
      const step = () => {
        if (dr) { r += dr; if (r > sel.r2) { r = sel.r1; c = c + 1 > sel.c2 ? sel.c1 : c + 1; } if (r < sel.r1) { r = sel.r2; c = c - 1 < sel.c1 ? sel.c2 : c - 1; } }
        else { c += dc; if (c > sel.c2) { c = sel.c1; r = r + 1 > sel.r2 ? sel.r1 : r + 1; } if (c < sel.c1) { c = sel.c2; r = r - 1 < sel.r1 ? sel.r2 : r - 1; } }
      };
      step();
      let guard = 0;
      while ((sh.isRowHidden(r) || sh.isColHidden(c)) && guard++ < 10000) step();
      v.active = { r, c };
      this.grid.scrollTo(r, c); this.afterSelect();
      return;
    }
    let { r, c } = v.active;
    const mg = sh.mergeAt(r, c);
    if (mg) { if (dr > 0) r = mg.r2; if (dc > 0) c = mg.c2; if (dr < 0) r = mg.r1; if (dc < 0) c = mg.c1; }
    r += dr; c += dc;
    while (dr && r > 0 && r < MAX_ROWS - 1 && sh.isRowHidden(r)) r += dr;
    while (dc && c > 0 && c < MAX_COLS - 1 && sh.isColHidden(c)) c += dc;
    this.select(r, c);
  }

  /** Ctrl+Arrow target */
  edge(r, c, dr, dc) {
    const wb = this.doc.wb, sh = wb.sheet;
    const filled = (rr, cc) => { const cell = sh.cells.get(key(rr, cc)); if (!cell) return false; if (cell.f) return true; return cell.v !== undefined && cell.v !== null && cell.v !== ''; };
    const inb = (rr, cc) => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS;
    let nr = r + dr, nc = c + dc;
    if (!inb(nr, nc)) return { r, c };
    const hidden = (rr, cc) => sh.isRowHidden(rr) || sh.isColHidden(cc);
    if (filled(r, c) && filled(nr, nc)) {
      while (inb(nr + dr, nc + dc) && (filled(nr + dr, nc + dc) || hidden(nr + dr, nc + dc))) { nr += dr; nc += dc; }
      while (hidden(nr, nc) && (nr !== r || nc !== c)) { nr -= dr; nc -= dc; }
      return { r: nr, c: nc };
    }
    const limR = dr > 0 ? sh.maxRow : 0, limC = dc > 0 ? sh.maxCol : 0;
    while (inb(nr, nc) && !filled(nr, nc)) {
      if ((dr > 0 && nr > limR) || (dc > 0 && nc > limC)) return { r: dr > 0 ? MAX_ROWS - 1 : nr, c: dc > 0 ? MAX_COLS - 1 : nc };
      nr += dr; nc += dc;
    }
    if (!inb(nr, nc)) return { r: Math.max(0, Math.min(MAX_ROWS - 1, nr - dr)), c: Math.max(0, Math.min(MAX_COLS - 1, nc - dc)) };
    return { r: nr, c: nc };
  }

  focusGrid() { if (!this.editor.active) this.editor.ta.focus({ preventScroll: true }); }

  // ------------------------------------------------------------------ cell editing
  setCellInput(sh, r, c, text, { fillSelection = false, inTxn = false } = {}) {
    const wb = this.doc.wb;
    if (!inTxn) wb.begin('입력');
    const write = (rr, cc, srcText) => {
      const cur = sh.get(rr, cc);
      const base = { s: cur?.s, h: cur?.h, cm: cur?.cm };
      if (srcText.startsWith('=') && srcText.length > 1) { wb.setCell(sh, rr, cc, { ...base, f: srcText }); return; }
      if (srcText.startsWith("'")) { wb.setCell(sh, rr, cc, { ...base, v: srcText.slice(1) }); return; }
      if (srcText === '') { wb.setCell(sh, rr, cc, base.s || base.cm ? base : null); return; }
      const p = parseInput(srcText);
      let s = base.s;
      if (p.nf && (!s?.nf || (isDateFormat(p.nf) !== isDateFormat(s.nf) && !/%/.test(s.nf)))) s = internStyle({ ...(s || {}), nf: p.nf });
      if (typeof p.v === 'number' && s?.nf && /%/.test(s.nf) && !/%/.test(srcText) && !p.nf) p.v = p.v / 100; // typing 5 into a % cell
      wb.setCell(sh, rr, cc, { ...base, s, v: p.v, t: p.t });
      if (typeof p.v === 'string' && p.v.includes('\n') && !s?.wrap) wb.setCell(sh, rr, cc, { ...base, v: p.v, s: internStyle({ ...(s || {}), wrap: true }) });
    };
    if (fillSelection) {
      const sel = this.view().sel;
      for (let rr = sel.r1; rr <= Math.min(sel.r2, sel.r1 + 100000); rr++) for (let cc = sel.c1; cc <= sel.c2; cc++) {
        if (sh.isRowHidden(rr)) continue;
        const t = text.startsWith('=') ? translateFormula(text, rr - r, cc - c) : text;
        write(rr, cc, t);
      }
    } else write(r, c, text);
    // grow row height for multi-line content (rows without explicit height)
    const cell = sh.get(r, c);
    if (cell && (cell.s?.wrap || (typeof cell.v === 'string' && cell.v.includes('\n')))) this.autofitRowsIn([r], true);
    if (!inTxn) wb.commit();
    this.lastAction = null;
  }

  // ------------------------------------------------------------------ style commands
  selection() { return this.view().sel; }
  styleRange(fn, label = '서식') {
    const wb = this.doc.wb, sh = wb.sheet;
    wb.begin(label);
    wb.setStyle(sh, this.selection(), fn);
    wb.commit();
    this.lastAction = () => this.styleRange(fn, label);
  }
  toggleStyle(prop) {
    const sh = this.doc.wb.sheet, v = this.view();
    const cur = sh.get(v.active.r, v.active.c)?.s?.[prop];
    this.styleRange((s) => ({ ...s, [prop]: cur ? null : true }));
  }
  setStyleProp(prop, value) { this.styleRange((s) => ({ ...s, [prop]: value })); }

  applyStylePatch(patch, { merge = undefined, border = null } = {}) {
    const wb = this.doc.wb;
    wb.begin('셀 서식');
    this.styleRange((s) => { const n = { ...s }; for (const [k, v] of Object.entries(patch)) { if (v === null || v === undefined) delete n[k]; else n[k] = v; } return n; });
    if (border) this.setBorders(border.kind, border.s, border.c);
    if (merge !== undefined) {
      const sel = this.selection(), sh = wb.sheet;
      const isMerged = !!sh.mergeAt(sel.r1, sel.c1);
      if (merge && !isMerged && (sel.r2 > sel.r1 || sel.c2 > sel.c1)) this.merge('merge', true);
      if (!merge && isMerged) this.merge('unmerge', true);
    }
    wb.commit();
  }

  setBorders(kind, style = 'thin', color = null) {
    const sel = this.selection();
    const sh = this.doc.wb.sheet;
    const r2 = Math.min(sel.r2, Math.max(sh.maxRow, sel.r1 + 200)), c2 = Math.min(sel.c2, Math.max(sh.maxCol, sel.c1 + 50));
    const b = { s: style, ...(color && color !== '#000000' ? { c: color } : {}) };
    this.styleRange((s, r, c) => {
      const n = { ...s };
      const top = r === sel.r1, bottom = r === r2, left = c === sel.c1, right = c === c2;
      switch (kind) {
        case 'none': delete n.bt; delete n.bb; delete n.bl; delete n.br; break;
        case 'all': n.bt = n.bb = n.bl = n.br = b; break;
        case 'outline': if (top) n.bt = b; if (bottom) n.bb = b; if (left) n.bl = b; if (right) n.br = b; break;
        case 'thickOutline': { const t = { ...b, s: 'medium' }; if (top) n.bt = t; if (bottom) n.bb = t; if (left) n.bl = t; if (right) n.br = t; break; }
        case 'inside': if (!top) n.bt = b; if (!bottom) n.bb = b; if (!left) n.bl = b; if (!right) n.br = b; break;
        case 'top': if (top) n.bt = b; break;
        case 'bottom': if (bottom) n.bb = b; break;
        case 'left': if (left) n.bl = b; break;
        case 'right': if (right) n.br = b; break;
        case 'doubleBottom': if (bottom) n.bb = { ...b, s: 'double' }; break;
        case 'thickBottom': if (bottom) n.bb = { ...b, s: 'medium' }; break;
        case 'topBottom': if (top) n.bt = b; if (bottom) n.bb = b; break;
        default: break;
      }
      return n;
    }, '테두리');
  }

  changeDecimals(delta) {
    const wb = this.doc.wb, sh = wb.sheet, v = this.view();
    const val = wb.value(sh, v.active.r, v.active.c);
    this.styleRange((s) => {
      let nf = s.nf && s.nf !== 'General' ? s.nf : null;
      if (!nf) {
        const dec = typeof val === 'number' ? (String(parseFloat(val.toPrecision(10))).split('.')[1] || '').length : 0;
        nf = dec ? '0.' + '0'.repeat(dec) : '0';
      }
      const parts = nf.split(';').map((p) => {
        if (delta > 0) return /0\.0*/.test(p) ? p.replace(/0\.(0*)/, (m, z) => '0.' + z + '0') : p.replace(/0(?![0.])/, '0.0');
        return p.replace(/0\.(0*)/, (m, z) => (z.length > 1 ? '0.' + z.slice(1) : '0'));
      });
      return { ...s, nf: parts.join(';') };
    }, '자릿수');
  }

  merge(mode, inTxn = false) {
    const wb = this.doc.wb, sh = wb.sheet, sel = this.selection();
    if (!inTxn) wb.begin('셀 병합');
    wb.touchMeta(sh);
    const overlaps = (m) => m.r2 >= sel.r1 && m.r1 <= sel.r2 && m.c2 >= sel.c1 && m.c1 <= sel.c2;
    sh.merges = sh.merges.filter((m) => !overlaps(m));
    if (mode !== 'unmerge') {
      const regions = mode === 'across' ? Array.from({ length: sel.r2 - sel.r1 + 1 }, (_, i) => ({ r1: sel.r1 + i, r2: sel.r1 + i, c1: sel.c1, c2: sel.c2 })) : [sel];
      for (const g of regions) {
        if (g.r1 === g.r2 && g.c1 === g.c2) continue;
        // keep only the upper-left value (Excel)
        for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
          if (r === g.r1 && c === g.c1) continue;
          const cell = sh.get(r, c);
          if (cell && (cell.v !== undefined || cell.f)) wb.setCell(sh, r, c, cell.s ? { s: cell.s } : null);
        }
        sh.merges.push(g);
        if (mode === 'center') { const cell = sh.get(g.r1, g.c1); wb.patchCell(sh, g.r1, g.c1, { s: internStyle({ ...(cell?.s || {}), ha: 'center' }) }); }
      }
    }
    if (!inTxn) wb.commit();
    this.selectRange(sel, this.view().active);
  }

  async mergeCommand(mode) {
    const sh = this.doc.wb.sheet, sel = this.selection();
    if (mode !== 'unmerge' && sh.mergeAt(sel.r1, sel.c1) && mode === 'center') return this.merge('unmerge');
    if (mode !== 'unmerge') {
      let n = 0;
      for (let r = sel.r1; r <= Math.min(sel.r2, sh.maxRow); r++) for (let c = sel.c1; c <= Math.min(sel.c2, sh.maxCol); c++) { const cell = sh.get(r, c); if (cell && (cell.v !== undefined || cell.f)) n++; }
      if (n > 1 && mode !== 'across') {
        const ok = await D.confirmBox('셀을 병합하면 왼쪽 위의 값만 남고 나머지 값은 삭제됩니다.', { yes: '확인', no: '취소' });
        if (ok !== 'yes') return;
      }
    }
    this.merge(mode);
  }

  // ------------------------------------------------------------------ painter
  startPainter(sticky = false) {
    const sh = this.doc.wb.sheet, sel = this.selection();
    const block = [];
    for (let r = sel.r1; r <= Math.min(sel.r2, sel.r1 + 500); r++) { const row = []; for (let c = sel.c1; c <= Math.min(sel.c2, sel.c1 + 200); c++) row.push(sh.get(r, c)?.s || null); block.push(row); }
    this.painter = { block, sticky };
    document.body.classList.add('painting');
    this.setStatus('서식을 붙여넣을 셀을 선택하세요. (Esc: 취소)');
  }
  applyPainter(p) {
    const { block } = p;
    const sel = this.selection();
    this.styleRange((s, r, c) => block[(r - sel.r1) % block.length][(c - sel.c1) % block[0].length] || {}, '서식 복사');
  }

  // ------------------------------------------------------------------ clear / fill / sums
  clear(kind = 'contents') {
    const wb = this.doc.wb, sh = wb.sheet, sel = this.selection();
    const g = { ...sel, r2: Math.min(sel.r2, sh.maxRow), c2: Math.min(sel.c2, sh.maxCol) };
    wb.begin(kind === 'all' ? '모두 지우기' : kind === 'formats' ? '서식 지우기' : '내용 지우기');
    const keys = [];
    if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > sh.cells.size) { for (const k of sh.cells.keys()) if (inRange(g, keyRow(k), keyCol(k))) keys.push(k); }
    else for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (sh.cells.has(key(r, c))) keys.push(key(r, c));
    for (const k of keys) {
      const r = keyRow(k), c = keyCol(k);
      if (sh.filterHidden.has(r)) continue;
      const cell = sh.cells.get(k);
      if (kind === 'all') wb.setCell(sh, r, c, null);
      else if (kind === 'formats') wb.setCell(sh, r, c, { ...cell, s: null });
      else if (kind === 'comments') wb.setCell(sh, r, c, { ...cell, cm: null });
      else wb.setCell(sh, r, c, cell.s || cell.cm || cell.h ? { s: cell.s, cm: cell.cm } : null);
    }
    if (kind === 'all') {
      const ch = sh.charts.filter((c) => !(c.anchor.r >= g.r1 && c.anchor.r <= g.r2 && c.anchor.c >= g.c1 && c.anchor.c <= g.c2));
      if (ch.length !== sh.charts.length) { wb.touchMeta(sh); sh.charts = ch; }
    }
    wb.commit();
  }

  /** fill handle: extend src into target (series detection) */
  fill(src, target) {
    const wb = this.doc.wb, sh = wb.sheet;
    const vertical = target.r1 !== src.r1 || target.r2 !== src.r2;
    wb.begin('채우기');
    if (target.r1 >= src.r1 && target.r2 <= src.r2 && target.c1 >= src.c1 && target.c2 <= src.c2) {
      // shrinking the selection with the handle clears cells (Excel)
      for (let r = target.r1; r <= target.r2; r++) for (let c = target.c1; c <= target.c2; c++) { /* keep */ }
    }
    const lines = vertical ? range(src.c1, src.c2) : range(src.r1, src.r2);
    for (const line of lines) {
      const seq = (vertical ? range(src.r1, src.r2) : range(src.c1, src.c2)).map((i) => ({ i, cell: vertical ? sh.get(i, line) : sh.get(line, i) }));
      const start = vertical ? src.r1 : src.c1, len = seq.length;
      const tStart = vertical ? target.r1 : target.c1, tEnd = vertical ? target.r2 : target.c2;
      const gen = makeSeries(seq.map((x) => x.cell), wb);
      for (let p = tStart; p <= tEnd; p++) {
        if (p >= start && p < start + len) continue;
        const k = p - start;
        const idx = ((k % len) + len) % len;
        const base = seq[idx].cell;
        let cell;
        if (base?.f) cell = { ...base, f: translateFormula(base.f, vertical ? k - idx : 0, vertical ? 0 : k - idx) };
        else cell = gen(k, base);
        if (cell) { delete cell.cached; delete cell.a; }
        if (vertical) { if (!sh.isRowHidden(p)) wb.setCell(sh, p, line, cell); }
        else wb.setCell(sh, line, p, cell);
      }
    }
    wb.commit();
    this.selectRange(normRange({ r1: Math.min(src.r1, target.r1), c1: Math.min(src.c1, target.c1), r2: Math.max(src.r2, target.r2), c2: Math.max(src.c2, target.c2) }), this.view().active);
  }

  fillDir(dir) {
    const wb = this.doc.wb, sh = wb.sheet, sel = this.selection();
    wb.begin(dir === 'down' ? '아래로 채우기' : '오른쪽으로 채우기');
    if (dir === 'down') {
      const r0 = sel.r1 === sel.r2 ? sel.r1 - 1 : sel.r1;
      if (r0 < 0) { wb.commit(); return; }
      for (let c = sel.c1; c <= Math.min(sel.c2, sh.maxCol); c++) {
        const src = sh.get(r0, c);
        for (let r = r0 + 1; r <= Math.min(sel.r2, Math.max(sh.maxRow, sel.r2)); r++) {
          if (sh.isRowHidden(r)) continue;
          wb.setCell(sh, r, c, src ? { ...src, f: src.f ? translateFormula(src.f, r - r0, 0) : undefined, cached: undefined } : null);
        }
      }
    } else {
      const c0 = sel.c1 === sel.c2 ? sel.c1 - 1 : sel.c1;
      if (c0 < 0) { wb.commit(); return; }
      for (let r = sel.r1; r <= Math.min(sel.r2, sh.maxRow); r++) {
        const src = sh.get(r, c0);
        for (let c = c0 + 1; c <= sel.c2; c++) wb.setCell(sh, r, c, src ? { ...src, f: src.f ? translateFormula(src.f, 0, c - c0) : undefined, cached: undefined } : null);
      }
    }
    wb.commit();
  }

  autoSum(fn = 'SUM') {
    const wb = this.doc.wb, sh = wb.sheet, v = this.view(), sel = v.sel;
    const isNum = (r, c) => typeof wb.value(sh, r, c) === 'number';
    if (sel.r1 === sel.r2 && sel.c1 === sel.c2) {
      const { r, c } = v.active;
      let r1 = r - 1;
      while (r1 >= 0 && (isNum(r1, c) || (r1 === r - 1 && false))) r1--;
      let text;
      if (r1 < r - 1) text = `=${fn}(${a1(r1 + 1, c)}:${a1(r - 1, c)})`;
      else {
        let c1 = c - 1;
        while (c1 >= 0 && isNum(r, c1)) c1--;
        text = c1 < c - 1 ? `=${fn}(${a1(r, c1 + 1)}:${a1(r, c - 1)})` : `=${fn}()`;
      }
      this.editor.start({ mode: 'edit', text });
      const pos = text.length - 1;
      this.editor.ta.setSelectionRange(text.endsWith('()') ? pos : text.indexOf('(') + 1, pos);
      return;
    }
    wb.begin('자동 합계');
    const rowTarget = Math.min(sel.r2, sh.maxRow) + 1;
    for (let c = sel.c1; c <= Math.min(sel.c2, sh.maxCol); c++) wb.setCell(sh, rowTarget, c, { ...(sh.get(rowTarget, c) || {}), f: `=${fn}(${a1(sel.r1, c)}:${a1(rowTarget - 1, c)})` });
    wb.commit();
  }

  insertFunction(name) {
    if (!this.editor.active) this.editor.start({ mode: 'edit', text: '=' + name + '()' });
    else this.editor.insertAtCaret(this.editor.source, name + '()');
    const el = this.editor.source; const p = el.value.lastIndexOf(')'); el.setSelectionRange(p, p);
    this.editor.onInput(true);
  }

  insertNow(kind) {
    const d = new Date();
    const v = this.view();
    const sh = this.doc.wb.sheet;
    const serial = kind === 'date' ? dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) : (d.getHours() * 3600 + d.getMinutes() * 60) / 86400;
    const wb = this.doc.wb;
    wb.begin(kind === 'date' ? '날짜 입력' : '시간 입력');
    const cur = sh.get(v.active.r, v.active.c);
    wb.setCell(sh, v.active.r, v.active.c, { ...(cur || {}), f: undefined, v: serial, s: internStyle({ ...(cur?.s || {}), nf: kind === 'date' ? 'yyyy-mm-dd' : 'h:mm' }) });
    wb.commit();
  }

  // ------------------------------------------------------------------ rows / cols
  fullRows() { const s = this.selection(); return s.c1 === 0 && s.c2 === MAX_COLS - 1; }
  fullCols() { const s = this.selection(); return s.r1 === 0 && s.r2 === MAX_ROWS - 1; }

  insert(axis) {
    const wb = this.doc.wb, sh = wb.sheet, sel = this.selection();
    if (axis === 'row') wb.shift(sh, 'row', sel.r1, sel.r2 - sel.r1 + 1 > 10000 ? 1 : sel.r2 - sel.r1 + 1);
    else wb.shift(sh, 'col', sel.c1, sel.c2 - sel.c1 + 1 > 1000 ? 1 : sel.c2 - sel.c1 + 1);
    this.lastAction = () => this.insert(axis);
  }
  remove(axis) {
    const wb = this.doc.wb, sh = wb.sheet, sel = this.selection();
    if (axis === 'row') wb.shift(sh, 'row', sel.r1, -(Math.min(sel.r2, Math.max(sh.maxRow, sel.r1)) - sel.r1 + 1));
    else wb.shift(sh, 'col', sel.c1, -(Math.min(sel.c2, Math.max(sh.maxCol, sel.c1)) - sel.c1 + 1));
    this.select(sel.r1, sel.c1);
    this.lastAction = () => this.remove(axis);
  }
  async insertDialog(del = false) {
    if (this.fullRows()) return del ? this.remove('row') : this.insert('row');
    if (this.fullCols()) return del ? this.remove('col') : this.insert('col');
    const res = await D.modal({
      title: del ? '삭제' : '삽입', width: 300,
      body: `<div class="radio-col">${del ? '<label><input type="radio" name="ins" value="left" checked> 셀을 왼쪽으로 밀기</label><label><input type="radio" name="ins" value="up"> 셀을 위로 밀기</label>' : '<label><input type="radio" name="ins" value="right" checked> 셀을 오른쪽으로 밀기</label><label><input type="radio" name="ins" value="down"> 셀을 아래로 밀기</label>'}<label><input type="radio" name="ins" value="row"> 행 전체</label><label><input type="radio" name="ins" value="col"> 열 전체</label></div>`,
      onOk: (el) => el.querySelector('input[name=ins]:checked').value,
    }).done;
    if (!res) return;
    if (res === 'row' || res === 'col') return del ? this.remove(res) : this.insert(res);
    this.shiftCells(res);
  }
  /** move cells inside the selection band (references are not rewritten — like a cut/paste) */
  shiftCells(dir) {
    const wb = this.doc.wb, sh = wb.sheet, sel = this.selection();
    const h = sel.r2 - sel.r1 + 1, w = sel.c2 - sel.c1 + 1;
    wb.begin('셀 이동');
    wb.markStructural();
    const moves = [];
    for (const [k, cell] of sh.cells) {
      const r = keyRow(k), c = keyCol(k);
      if ((dir === 'right' || dir === 'left') && r >= sel.r1 && r <= sel.r2 && c >= sel.c1) moves.push([r, c, cell]);
      if ((dir === 'down' || dir === 'up') && c >= sel.c1 && c <= sel.c2 && r >= sel.r1) moves.push([r, c, cell]);
    }
    for (const [r, c] of moves) wb.setCell(sh, r, c, null);
    for (const [r, c, cell] of moves) {
      let nr = r, nc = c;
      if (dir === 'right') nc += w; if (dir === 'down') nr += h;
      if (dir === 'left') { if (c <= sel.c2) continue; nc -= w; }
      if (dir === 'up') { if (r <= sel.r2) continue; nr -= h; }
      wb.setCell(sh, nr, nc, cell);
    }
    wb.commit();
  }

  resize(axis, targets, size) {
    const wb = this.doc.wb, sh = wb.sheet;
    wb.begin(axis === 'col' ? '열 너비' : '행 높이');
    wb.touchMeta(sh);
    for (const t of targets) {
      if (axis === 'col') { if (size <= 0) { sh.hiddenCols.add(t); } else { sh.hiddenCols.delete(t); sh.colW.set(t, size); } }
      else { if (size <= 0) sh.hiddenRows.add(t); else { sh.hiddenRows.delete(t); sh.rowH.set(t, size); } }
    }
    wb.commit();
  }

  async sizeDialog(axis) {
    const sh = this.doc.wb.sheet, sel = this.selection(), g = geom(sh);
    const cur = axis === 'col' ? ((g.cols.size(sel.c1) - 5) / 7).toFixed(2) : (g.rows.size(sel.r1) * 0.75).toFixed(2);
    const v = await D.promptBox(axis === 'col' ? '열 너비 (문자 수):' : '행 높이 (포인트):', cur, axis === 'col' ? '열 너비' : '행 높이');
    if (v === null || isNaN(parseFloat(v))) return;
    const px = axis === 'col' ? Math.round(parseFloat(v) * 7 + 5) : Math.round(parseFloat(v) / 0.75);
    const targets = axis === 'col' ? range(sel.c1, Math.min(sel.c2, sel.c1 + 1000)) : range(sel.r1, Math.min(sel.r2, sel.r1 + 100000));
    this.resize(axis, targets, px);
  }

  autofitCols(c) {
    const sel = this.selection();
    const cols = this.fullCols() && c >= sel.c1 && c <= sel.c2 ? range(sel.c1, sel.c2) : [c];
    const sh = this.doc.wb.sheet;
    const widths = cols.map((cc) => {
      let w = 0;
      for (const [k] of sh.cells) {
        if (keyCol(k) !== cc) continue;
        const r = keyRow(k);
        if (sh.mergeAt(r, cc)) continue;
        const d = this.grid.display(sh, r, cc);
        if (!d?.text) continue;
        const f = this.grid.fontFor(d.st);
        const lines = String(d.text).split('\n');
        for (const ln of lines) w = Math.max(w, this.grid.measure(f.css, ln));
        if (sh.autoFilter && r === sh.autoFilter.range.r1) w += 16;
      }
      return w ? Math.ceil(w + 10) : sh.defaultColW;
    });
    const wb = this.doc.wb;
    wb.begin('열 너비 자동 맞춤'); wb.touchMeta(sh);
    cols.forEach((cc, i) => { sh.colW.set(cc, widths[i]); sh.hiddenCols.delete(cc); });
    wb.commit();
  }

  autofitRows(r) {
    const sel = this.selection();
    const rows = this.fullRows() && r >= sel.r1 && r <= sel.r2 ? range(sel.r1, sel.r2) : [r];
    this.autofitRowsIn(rows, false);
  }
  autofitRowsIn(rows, onlyGrow) {
    const wb = this.doc.wb, sh = wb.sheet, g = geom(sh);
    const heights = new Map();
    for (const [k] of sh.cells) {
      const r = keyRow(k);
      if (!rows.includes(r)) continue;
      const c = keyCol(k);
      const d = this.grid.display(sh, r, c);
      const f = this.grid.fontFor(d?.st);
      let lines = 1;
      if (d?.text) lines = d.st?.wrap ? this.grid.wrapLines(f.css, d.text, Math.max(4, g.cols.size(c) - 6)).length : String(d.text).split('\n').length;
      heights.set(r, Math.max(heights.get(r) || 0, Math.ceil(lines * f.size * 1.25 + 5)));
    }
    const own = !wb.txn;
    if (own) wb.begin('행 높이 자동 맞춤');
    wb.touchMeta(sh);
    for (const r of rows) {
      const h = Math.max(sh.defaultRowH, heights.get(r) || sh.defaultRowH);
      if (onlyGrow && h <= g.rows.size(r)) continue;
      if (h === sh.defaultRowH) sh.rowH.delete(r); else sh.rowH.set(r, h);
    }
    if (own) wb.commit();
  }

  hide(axis, hidden) {
    const wb = this.doc.wb, sh = wb.sheet, sel = this.selection();
    wb.begin(hidden ? '숨기기' : '숨기기 취소'); wb.touchMeta(sh);
    if (axis === 'row') for (let r = sel.r1; r <= Math.min(sel.r2, Math.max(sh.maxRow + 1, sel.r1 + 1000)); r++) hidden ? sh.hiddenRows.add(r) : sh.hiddenRows.delete(r);
    else for (let c = sel.c1; c <= Math.min(sel.c2, Math.max(sh.maxCol + 1, sel.c1 + 200)); c++) hidden ? sh.hiddenCols.add(c) : sh.hiddenCols.delete(c);
    if (!hidden && axis === 'row' && sel.r1 > 0) { for (let r = sel.r1 - 1; r >= 0 && sh.hiddenRows.has(r); r--) sh.hiddenRows.delete(r); }
    if (!hidden && axis === 'col' && sel.c1 > 0) { for (let c = sel.c1 - 1; c >= 0 && sh.hiddenCols.has(c); c--) sh.hiddenCols.delete(c); }
    wb.commit();
  }

  freeze(mode) {
    const wb = this.doc.wb, sh = wb.sheet, v = this.view();
    wb.begin('틀 고정'); wb.touchMeta(sh);
    if (mode === 'none' || ((sh.freeze.r || sh.freeze.c) && mode === 'toggle')) sh.freeze = { r: 0, c: 0 };
    else if (mode === 'row') sh.freeze = { r: 1, c: 0 };
    else if (mode === 'col') sh.freeze = { r: 0, c: 1 };
    else sh.freeze = { r: v.active.r, c: v.active.c };
    wb.commit();
    this.grid.scroller.scrollTop = 0; this.grid.scroller.scrollLeft = 0;
    this.render();
  }

  setZoom(z) {
    z = Math.max(0.25, Math.min(4, z));
    const v = this.view();
    const ratio = z / v.zoom;
    v.zoom = z;
    this.grid.updateExtent(this.grid.metrics());
    this.grid.scroller.scrollTop *= ratio; this.grid.scroller.scrollLeft *= ratio;
    this.grid.measureCache.clear();
    this.render();
    if (this.editor.active) { this.editor.style(); this.editor.reposition(); }
  }

  // ------------------------------------------------------------------ sheets
  async renameSheetPrompt(i) {
    const wb = this.doc.wb;
    const name = await D.promptBox('시트 이름:', wb.sheets[i].name, '이름 바꾸기');
    if (name === null) return;
    try { wb.renameSheet(i, name); } catch (e) { this.toast(e.message, 'error'); }
  }
  async deleteSheet(i) {
    const wb = this.doc.wb;
    if (wb.sheets.filter((s) => s.state === 'visible').length <= 1) return this.toast('통합 문서에는 하나 이상의 표시된 시트가 있어야 합니다.', 'error');
    const sh = wb.sheets[i];
    if (sh.cells.size && (await D.confirmBox(`'${sh.name}' 시트를 삭제하시겠습니까? 이 작업은 실행 취소할 수 있습니다.`, { yes: '삭제', no: '취소' })) !== 'yes') return;
    wb.deleteSheet(i);
  }
  async unhideSheet() {
    const wb = this.doc.wb;
    const hidden = wb.sheets.map((s, i) => [s, i]).filter(([s]) => s.state !== 'visible');
    if (!hidden.length) return this.toast('숨겨진 시트가 없습니다.');
    const res = await D.modal({ title: '시트 숨기기 취소', width: 320, body: `<select class="inp" size="6" id="uh">${hidden.map(([s, i]) => `<option value="${i}">${D.esc(s.name)}</option>`).join('')}</select>`, onOk: (el) => el.querySelector('#uh').value }).done;
    if (res === null || res === '') return;
    const sh = wb.sheets[+res];
    wb.begin('시트 숨기기 취소'); wb.touchMeta(sh); sh.state = 'visible'; wb.commit();
    this.activateSheet(+res);
  }
  hideSheet(i) {
    const wb = this.doc.wb;
    if (wb.sheets.filter((s) => s.state === 'visible').length <= 1) return this.toast('표시된 시트가 하나뿐이면 숨길 수 없습니다.', 'error');
    const sh = wb.sheets[i];
    wb.begin('시트 숨기기'); wb.touchMeta(sh); sh.state = 'hidden'; wb.commit();
    if (wb.active === i) this.activateSheet(wb.sheets.findIndex((s) => s.state === 'visible'));
  }
  tabColor(i, anchor) {
    const wb = this.doc.wb, sh = wb.sheets[i];
    D.colorPopup(anchor, { noneLabel: '색 없음', onPick: (c) => { wb.begin('탭 색'); wb.touchMeta(sh); sh.tabColor = c; wb.commit(); } });
  }

  // ------------------------------------------------------------------ files
  async newDoc() {
    try {
      const res = await api.newDoc(this.nextBlankName());
      await this.attachDoc(res.docId);
    } catch (e) { this.toast(`새 통합 문서를 만들 수 없습니다: ${e.message}`, 'error'); }
  }

  async openFile(path = null) {
    if (!path) {
      const res = await D.fileDialog({ mode: 'open' });
      if (!res) return;
      path = res.path;
    }
    this.showBusy(`여는 중… ${path}`);
    try {
      const res = await api.openDoc(path);
      // replace an untouched blank workbook (Excel behaviour)
      const blank = this.docs.length === 1 && !this.doc.wb.path && !this.doc.wb.dirty && this.doc.wb.sheets.length === 1 && !this.doc.wb.sheet.cells.size && this.doc.wb.docId !== res.docId ? this.doc : null;
      const doc = await this.attachDoc(res.docId);
      if (blank && doc) { api.closeDoc(blank.wb.docId).catch(() => {}); this.detachDoc(blank); this.activateDoc(this.docs.indexOf(doc)); }
      D.addRecent(path);
      if (doc) this.setStatus(`열기 완료 (${doc.wb.sheets.length}개 시트, 계산 ${doc.wb.lastCalcMs?.toFixed(0)}ms)`);
    } catch (e) {
      this.toast(`파일을 열 수 없습니다: ${e.message}`, 'error', 8000);
    } finally { this.hideBusy(); }
  }

  async save() {
    if (this.editor.active && !this.editor.commit()) return false;
    const wb = this.doc.wb;
    if (!wb.path) return this.saveAs();
    return this.writeTo(wb.path, wb.format, true);
  }

  async saveAs() {
    if (this.editor.active && !this.editor.commit()) return false;
    const wb = this.doc.wb;
    const startDir = wb.path ? wb.path.replace(/[\\/][^\\/]*$/, '') : null;
    const res = await D.fileDialog({ mode: 'save', name: wb.name.replace(/\.(xlsx|xlsm|xls)$/i, ''), format: wb.format, startDir });
    if (!res) return false;
    if (res.format === 'xls' && wb.format !== 'xls') {
      const hasCharts = wb.sheets.some((s) => s.charts.length);
      const ok = await D.confirmBox(`xls(Excel 97-2003) 형식은 65,536행 x 256열, 56색 팔레트로 제한되며${hasCharts ? ' 차트와' : ''} 자동 필터는 저장되지 않습니다. 계속하시겠습니까?`, { yes: '계속', no: '취소' });
      if (ok !== 'yes') return false;
    }
    return this.writeTo(res.path, res.format, false);
  }

  /** Save through the server (Python writes the file). Wait for pending edits first. */
  async writeTo(path, format, overwrite) {
    const doc = this.doc, wb = doc.wb;
    this.showBusy('저장 중…');
    try {
      for (let i = 0; i < 50 && doc.sync.pending.size; i++) await new Promise((r) => setTimeout(r, 100));
      await doc.sync.reportCalc(); // formulas' computed values are stored as cached results (xls)
      const opts = { path, format, overwrite };
      let res;
      for (;;) {
        try { res = await api.saveDoc(wb.docId, opts); break; } catch (e) {
          if (e.status !== 409) throw e;
          this.hideBusy();
          const changed = /디스크/.test(e.message);
          const ans = await D.confirmBox(changed ? `${e.message}\n그래도 덮어쓰시겠습니까?` : `'${path.split(/[\\/]/).pop()}'이(가) 이미 있습니다. 바꾸시겠습니까?`, { yes: '바꾸기', no: '아니요' });
          if (ans !== 'yes') return false;
          if (changed) opts.force = true; else opts.overwrite = true;
          this.showBusy('저장 중…');
        }
      }
      Object.assign(wb, { path: res.path, name: res.name, format: res.format, basePath: res.path, dirty: false });
      D.addRecent(res.path);
      this.renderTabs();
      this.setStatus(`저장됨: ${res.path}${res.backup ? ' (이전 파일 백업됨)' : ''}`);
      if (res.warnings?.length) this.toast(res.warnings.join('\n'), 'warn', 9000);
      return true;
    } catch (e) {
      this.toast(`저장하지 못했습니다: ${e.message}`, 'error', 9000);
      return false;
    } finally { this.hideBusy(); }
  }

  // ------------------------------------------------------------------ windows
  newWindow(docIdx = this.di) {
    const doc = this.docs[docIdx];
    const w = window.open(`${location.pathname}?doc=${encodeURIComponent(doc.wb.docId)}`, '_blank');
    if (!w) this.toast(`팝업이 차단되었습니다. 다른 창/브라우저에서 ${location.origin}/?doc=${doc.wb.docId} 를 여세요.`, 'warn', 9000);
  }
  async windowsMenu(anchor) {
    let docs = [];
    try { docs = (await api.docs()).docs; } catch { /* offline */ }
    const items = [['새 창 (현재 통합 문서)', () => this.newWindow()], null,
      ...docs.map((d) => {
        const i = this.docs.findIndex((x) => x.wb.docId === d.docId);
        return [`${i === this.di ? '✓ ' : '　'}${d.name}${d.dirty ? ' ●' : ''} — 연결 ${d.clients}`, () => (i >= 0 ? this.activateDoc(i) : this.attachDoc(d.docId))];
      })];
    const r = anchor.getBoundingClientRect();
    this.menu(r.left, r.bottom + 2, items);
  }
  onWindowsChanged(docs) { const m = this.$('#win-count'); if (m && docs) m.textContent = String(docs.reduce((n, d) => Math.max(n, d.clients || 0), 1)); }

  // ------------------------------------------------------------------ charts
  chartDialog(model, type) { return D.chartDialog(this, model, type); }
  saveChart(m, isNew) {
    const wb = this.doc.wb, sh = wb.sheet;
    wb.begin(isNew ? '차트 삽입' : '차트 편집'); wb.touchMeta(sh);
    if (isNew) sh.charts.push(m);
    else { const i = sh.charts.findIndex((c) => c.id === m.id); if (i >= 0) sh.charts[i] = m; }
    wb.commit();
    this.charts.select(m.id);
  }
  quickChart(type) {
    const wb = this.doc.wb, sh = wb.sheet;
    let g = this.selection();
    if (g.r1 === g.r2 && g.c1 === g.c2) g = this.filters.currentRegion(sh, g.r1, g.c1);
    if (!sh.cells.size) return this.toast('차트를 만들 데이터를 먼저 선택하세요.', 'warn');
    const preset = { barH: { type: 'bar', barDir: 'bar' }, barStacked: { type: 'bar', grouping: 'stacked' } }[type] || { type };
    const m = ChartLayer.fromRange(wb, sh, g, preset.type);
    Object.assign(m, preset);
    const geo = geom(sh);
    m.anchor.c = Math.min(MAX_COLS - 10, geo.cols.at(geo.cols.pos(g.c2 + 1) + 20));
    this.saveChart(m, true);
  }

  // ------------------------------------------------------------------ comments / links
  async editComment() {
    const sh = this.doc.wb.sheet, v = this.view(), cell = sh.get(v.active.r, v.active.c);
    const res = await D.modal({ title: '메모', width: 360, body: `<textarea class="inp" id="cmt" rows="5" autofocus>${D.esc(cell?.cm || '')}</textarea>`, onOk: (el) => el.querySelector('#cmt').value, buttons: [['확인', 'ok', true], ['메모 삭제', 'del'], ['취소', 'cancel']] }).done;
    if (res === null) return;
    this.doc.wb.patchCell(sh, v.active.r, v.active.c, { cm: res === 'del' ? null : res || null });
  }
  async editLink() {
    const sh = this.doc.wb.sheet, v = this.view(), cell = sh.get(v.active.r, v.active.c);
    const url = await D.promptBox('주소 (https://…):', cell?.h || 'https://', '하이퍼링크');
    if (url === null) return;
    const wb = this.doc.wb;
    wb.begin('하이퍼링크');
    const st = cell?.s || {};
    wb.patchCell(sh, v.active.r, v.active.c, { h: url && url !== 'https://' ? url : null, v: cell?.v ?? (url || undefined), s: url ? internStyle({ ...st, fc: st.fc || '#0563C1', u: st.u || true }) : st });
    wb.commit();
  }
  hoverComment(info) {
    const tip = this.$('#comment-tip');
    if (!info) { tip.hidden = true; return; }
    tip.textContent = info.text; tip.hidden = false;
    tip.style.left = info.x + 14 + 'px'; tip.style.top = info.y + 10 + 'px';
  }

  // ------------------------------------------------------------------ menus / toasts
  menu(x, y, items) {
    document.querySelector('.ctx-menu')?.remove();
    const m = document.createElement('div');
    m.className = 'ctx-menu';
    m.innerHTML = items.map((it, i) => (it === null ? '<hr>' : `<button data-i="${i}" ${it[2]?.disabled ? 'disabled' : ''}><span>${D.esc(it[0])}</span>${it[2]?.key ? `<kbd>${it[2].key}</kbd>` : ''}</button>`)).join('');
    document.body.append(m);
    m.style.left = Math.min(x, window.innerWidth - m.offsetWidth - 4) + 'px';
    m.style.top = Math.min(y, window.innerHeight - m.offsetHeight - 4) + 'px';
    const close = () => { m.remove(); document.removeEventListener('pointerdown', out, true); document.removeEventListener('keydown', esc, true); };
    const out = (e) => { if (!m.contains(e.target)) close(); };
    const esc = (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(); this.focusGrid(); } };
    setTimeout(() => { document.addEventListener('pointerdown', out, true); document.addEventListener('keydown', esc, true); }, 0);
    m.addEventListener('click', (e) => { const b = e.target.closest('button[data-i]'); if (!b) return; close(); items[+b.dataset.i][1](); });
    return m;
  }

  contextMenu(x, y, type) {
    const rows = this.fullRows(), cols = this.fullCols();
    const items = [
      ['잘라내기', () => this.clipboard.copyCommand(true), { key: 'Ctrl+X' }], ['복사', () => this.clipboard.copyCommand(false), { key: 'Ctrl+C' }],
      ['붙여넣기', () => this.clipboard.pasteCommand(), { key: 'Ctrl+V' }], ['선택하여 붙여넣기…', () => D.pasteSpecialDialog(this), { key: 'Ctrl+Alt+V' }], null,
      rows ? ['행 삽입', () => this.insert('row'), { key: 'Ctrl++' }] : cols ? ['열 삽입', () => this.insert('col'), { key: 'Ctrl++' }] : ['삽입…', () => this.insertDialog(false), { key: 'Ctrl++' }],
      rows ? ['행 삭제', () => this.remove('row'), { key: 'Ctrl+-' }] : cols ? ['열 삭제', () => this.remove('col'), { key: 'Ctrl+-' }] : ['삭제…', () => this.insertDialog(true), { key: 'Ctrl+-' }],
      ['내용 지우기', () => this.clear('contents'), { key: 'Delete' }], null,
      ['필터', () => this.filters.toggle(), { key: 'Ctrl+Shift+L' }], ['오름차순 정렬', () => this.filters.sortSelection(true)], ['내림차순 정렬', () => this.filters.sortSelection(false)], null,
      ['메모 삽입/편집', () => this.editComment(), { key: 'Shift+F2' }], ['셀 서식…', () => D.formatCellsDialog(this), { key: 'Ctrl+1' }], ['하이퍼링크…', () => this.editLink(), { key: 'Ctrl+K' }],
    ];
    if (rows) items.push(null, ['행 높이…', () => this.sizeDialog('row')], ['숨기기', () => this.hide('row', true), { key: 'Ctrl+9' }], ['숨기기 취소', () => this.hide('row', false), { key: 'Ctrl+Shift+9' }]);
    if (cols) items.push(null, ['열 너비…', () => this.sizeDialog('col')], ['숨기기', () => this.hide('col', true), { key: 'Ctrl+0' }], ['숨기기 취소', () => this.hide('col', false), { key: 'Ctrl+Shift+0' }]);
    this.menu(x, y, items);
  }

  toast(msg, type = 'info', ms = 4000) {
    const box = this.$('#toasts');
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.textContent = msg;
    t.addEventListener('click', () => t.remove());
    box.append(t);
    setTimeout(() => t.remove(), ms);
  }
  tooltip(text, x, y) {
    const t = this.$('#tooltip');
    if (!text) { t.hidden = true; return; }
    t.textContent = text; t.hidden = false; t.style.left = x + 'px'; t.style.top = y + 'px';
  }
  showBusy(t) { const b = this.$('#busy'); b.querySelector('span').textContent = t; b.hidden = false; }
  hideBusy() { this.$('#busy').hidden = true; }
  onEditState(on) { document.body.classList.toggle('editing', on); this.setStatus(on ? (this.editor.mode === 'edit' ? '편집' : '입력') : ''); }

  // ------------------------------------------------------------------ commands
  get commands() {
    const C = {
      new: () => this.newDoc(), open: () => this.openFile(), save: () => this.save(), saveAs: () => this.saveAs(), close: () => this.closeDoc(this.di),
      undo: () => { if (this.editor.active) this.editor.cancel(); const e = this.doc.wb.undo(); if (e) this.setStatus(`실행 취소: ${e.label}`); },
      redo: () => { const e = this.doc.wb.redo(); if (e) this.setStatus(`다시 실행: ${e.label}`); },
      cut: () => this.clipboard.copyCommand(true), copy: () => this.clipboard.copyCommand(false), paste: () => this.clipboard.pasteCommand(), pasteSpecial: () => D.pasteSpecialDialog(this),
      painter: () => this.startPainter(false),
      bold: () => this.toggleStyle('b'), italic: () => this.toggleStyle('i'), underline: () => this.toggleStyle('u'), strike: () => this.toggleStyle('st'),
      fontName: (v) => this.setStyleProp('ff', v === this.doc.wb.defaultFont.name ? null : v),
      fontSize: (v) => { const n = parseFloat(v); if (n > 0) this.setStyleProp('fs', n); },
      fontGrow: () => this.styleRange((s) => ({ ...s, fs: nextSize(s.fs || this.doc.wb.defaultFont.size, 1) })),
      fontShrink: () => this.styleRange((s) => ({ ...s, fs: nextSize(s.fs || this.doc.wb.defaultFont.size, -1) })),
      fontColor: (v) => this.setStyleProp('fc', v), fillColor: (v) => this.setStyleProp('bg', v),
      alignLeft: () => this.setStyleProp('ha', 'left'), alignCenter: () => this.setStyleProp('ha', 'center'), alignRight: () => this.setStyleProp('ha', 'right'),
      vTop: () => this.setStyleProp('va', 'top'), vMiddle: () => this.setStyleProp('va', 'center'), vBottom: () => this.setStyleProp('va', null),
      wrap: () => { this.toggleStyle('wrap'); this.autofitRowsIn(range(this.selection().r1, Math.min(this.selection().r2, this.doc.wb.sheet.maxRow)), false); },
      indent: () => this.styleRange((s) => ({ ...s, ind: Math.min(15, (s.ind || 0) + 1), ha: s.ha === 'right' ? 'right' : 'left' })),
      outdent: () => this.styleRange((s) => ({ ...s, ind: Math.max(0, (s.ind || 0) - 1) || null })),
      mergeCenter: () => this.mergeCommand('center'), mergeAcross: () => this.mergeCommand('across'), merge: () => this.mergeCommand('merge'), unmerge: () => this.mergeCommand('unmerge'),
      numFormat: (v) => { if (v === '__more') return D.formatCellsDialog(this, 'num'); if (v !== '__custom') this.setStyleProp('nf', v === 'General' ? null : v); },
      currency: () => this.setStyleProp('nf', '"₩"#,##0'), percent: () => this.setStyleProp('nf', '0%'), comma: () => this.setStyleProp('nf', '#,##0'),
      decInc: () => this.changeDecimals(1), decDec: () => this.changeDecimals(-1),
      borders: (k) => this.setBorders(k), formatCells: () => D.formatCellsDialog(this),
      insertRow: () => this.insert('row'), insertCol: () => this.insert('col'), deleteRow: () => this.remove('row'), deleteCol: () => this.remove('col'),
      insertCells: () => this.insertDialog(false), deleteCells: () => this.insertDialog(true),
      insertSheet: () => this.doc.wb.addSheet(null, this.doc.wb.active + 1), deleteSheet: () => this.deleteSheet(this.doc.wb.active),
      rowHeight: () => this.sizeDialog('row'), colWidth: () => this.sizeDialog('col'),
      autofitCol: () => { const s = this.selection(); for (let c = s.c1; c <= Math.min(s.c2, s.c1 + 200); c++) this.autofitCols(c); },
      autofitRow: () => this.autofitRowsIn(range(this.selection().r1, Math.min(this.selection().r2, this.doc.wb.sheet.maxRow)), false),
      autosum: (fn) => this.autoSum(fn || 'SUM'), fillDown: () => this.fillDir('down'), fillRight: () => this.fillDir('right'),
      clearAll: () => this.clear('all'), clearFormats: () => this.clear('formats'), clearContents: () => this.clear('contents'), clearComments: () => this.clear('comments'),
      sortAsc: () => this.filters.sortSelection(true), sortDesc: () => this.filters.sortSelection(false), sortDialog: () => D.sortDialog(this),
      filter: () => this.filters.toggle(), filterClear: () => this.filters.clearAll(), filterReapply: () => this.filters.reapply(),
      removeDuplicates: () => this.removeDuplicates(),
      find: () => D.findDialog(this, false), replace: () => D.findDialog(this, true),
      chart: (t) => (t ? this.quickChart(t) : this.chartDialog(null)), chartDialog: () => this.chartDialog(null),
      insertFunction: () => D.insertFunctionDialog(this), comment: () => this.editComment(), link: () => this.editLink(),
      freeze: () => this.freeze('cell'), freezeRow: () => this.freeze('row'), freezeCol: () => this.freeze('col'), unfreeze: () => this.freeze('none'),
      gridlines: () => { const wb = this.doc.wb, sh = wb.sheet; wb.begin('눈금선'); wb.touchMeta(sh); sh.showGrid = sh.showGrid === false; wb.commit(); },
      showFormulas: () => { const v = this.view(); v.showFormulas = !v.showFormulas; this.grid.measureCache.clear(); this.render(); },
      zoomIn: () => this.setZoom(this.view().zoom + 0.1), zoomOut: () => this.setZoom(this.view().zoom - 0.1), zoom100: () => this.setZoom(1),
      zoomSel: () => { const s = this.selection(), g = geom(this.doc.wb.sheet), m = this.grid.metrics(); const w = g.cols.pos(Math.min(s.c2, s.c1 + 200) + 1) - g.cols.pos(s.c1), h = g.rows.pos(Math.min(s.r2, s.r1 + 1000) + 1) - g.rows.pos(s.r1); this.setZoom(Math.min((this.grid.scroller.clientWidth - m.HW) / w, (this.grid.scroller.clientHeight - m.HH) / h) * 0.95); this.grid.scrollTo(s.r1, s.c1); },
      newWindow: () => this.newWindow(), recalc: () => { this.doc.wb.recalc(); this.render(); this.setStatus(`다시 계산 완료 (${this.doc.wb.lastCalcMs.toFixed(0)}ms)`); },
      unhideSheet: () => this.unhideSheet(),
      selectAll: () => this.selectAllSmart(),
      help: () => this.shortcutHelp(),
    };
    return C;
  }

  exec(cmd, arg) {
    const fn = this.commands[cmd];
    if (!fn) return;
    if (this.editor.active && !['undo', 'bold', 'italic', 'underline'].includes(cmd)) { if (!this.editor.commit()) return; }
    try { fn(arg); } catch (e) { console.error(e); this.toast(e.message, 'error'); }
    if (!['undo', 'redo', 'find', 'replace', 'formatCells', 'open', 'save', 'saveAs'].includes(cmd)) this.focusGrid();
  }

  selectAllSmart() {
    const v = this.view(), sh = this.doc.wb.sheet;
    const reg = this.filters.currentRegion(sh, v.active.r, v.active.c);
    const cur = v.sel;
    const all = { r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 };
    if ((reg.r1 !== reg.r2 || reg.c1 !== reg.c2) && !(cur.r1 === reg.r1 && cur.r2 === reg.r2 && cur.c1 === reg.c1 && cur.c2 === reg.c2)) this.selectRange(reg, v.active, { scroll: false });
    else this.selectRange(all, v.active, { scroll: false });
  }

  removeDuplicates() {
    const wb = this.doc.wb, sh = wb.sheet;
    let g = this.selection();
    if (g.r1 === g.r2 && g.c1 === g.c2) g = this.filters.currentRegion(sh, g.r1, g.c1);
    g = { ...g, r2: Math.min(g.r2, sh.maxRow), c2: Math.min(g.c2, sh.maxCol) };
    const header = this.filters.guessHeader(sh, g);
    const seen = new Set(); const keep = [];
    for (let r = g.r1 + (header ? 1 : 0); r <= g.r2; r++) {
      const k = range(g.c1, g.c2).map((c) => { const v = wb.value(sh, r, c); return typeof v === 'string' ? v.toLowerCase() : JSON.stringify(v); }).join('\u0001');
      if (!seen.has(k)) { seen.add(k); keep.push(r); }
    }
    const total = g.r2 - g.r1 + 1 - (header ? 1 : 0);
    const removed = total - keep.length;
    if (!removed) return this.toast('중복된 값이 없습니다.');
    wb.begin('중복된 항목 제거');
    wb.markStructural();
    const rows = keep.map((r) => range(g.c1, g.c2).map((c) => sh.get(r, c)));
    for (let r = g.r1 + (header ? 1 : 0); r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) wb.setCell(sh, r, c, null);
    rows.forEach((row, i) => row.forEach((cell, j) => { const r = g.r1 + (header ? 1 : 0) + i; wb.setCell(sh, r, g.c1 + j, cell ? { ...cell, f: cell.f ? translateFormula(cell.f, r - keep[i], 0) : undefined } : null); }));
    wb.commit();
    this.toast(`중복된 값 ${removed}개를 제거했습니다. 고유한 값 ${keep.length}개가 남아 있습니다.`);
  }

  shortcutHelp() {
    const rows = [
      ['Ctrl+N / O / S', '새로 만들기 / 열기 / 저장'], ['F12', '다른 이름으로 저장'], ['Ctrl+Z / Ctrl+Y', '실행 취소 / 다시 실행'], ['Ctrl+C / X / V', '복사 / 잘라내기 / 붙여넣기'],
      ['Ctrl+Alt+V', '선택하여 붙여넣기'], ['Ctrl+B / I / U / 5', '굵게 / 기울임 / 밑줄 / 취소선'], ['Ctrl+1', '셀 서식'], ['F2', '셀 편집 (편집/입력 모드 전환)'],
      ['F4', '참조 유형 전환($A$1) / 마지막 작업 반복'], ['Alt+Enter', '셀 안에서 줄 바꿈'], ['Ctrl+Enter', '선택 영역 모두에 입력'], ['Ctrl+D / Ctrl+R', '아래로 / 오른쪽으로 채우기'],
      ['Ctrl+; / Ctrl+Shift+;', '오늘 날짜 / 현재 시간'], ['Alt+=', '자동 합계'], ['Ctrl+Shift+L', '필터 켜기/끄기'], ['Alt+↓', '필터 메뉴 열기'], ['Ctrl+F / Ctrl+H', '찾기 / 바꾸기'],
      ['Ctrl+화살표', '데이터 영역 끝으로 이동'], ['Ctrl+Shift+화살표', '데이터 영역 끝까지 선택'], ['Ctrl+Home / End', 'A1 / 마지막 셀'], ['Ctrl+A', '현재 영역 / 전체 선택'],
      ['Ctrl+Space / Shift+Space', '열 전체 / 행 전체 선택'], ['Ctrl++ / Ctrl+-', '삽입 / 삭제'], ['Ctrl+9 / Ctrl+0', '행 / 열 숨기기 (Shift: 숨기기 취소)'],
      ['Ctrl+PgUp / PgDn', '이전 / 다음 시트'], ['Shift+F11', '새 시트'], ['Alt+F1', '차트 삽입'], ['Shift+F2', '메모'], ['Ctrl+K', '하이퍼링크'], ['Ctrl+`', '수식 표시'],
      ['Ctrl+Shift+1/2/3/4/5', '숫자/시간/날짜/통화/백분율 서식'], ['Ctrl+Shift+7 / Ctrl+Shift+-', '윤곽선 테두리 / 테두리 제거'], ['F9', '모두 다시 계산'], ['Ctrl+F4', '통합 문서 닫기'],
      ['Ctrl+Tab', '다음 통합 문서'], ['Ctrl+휠', '확대/축소'],
    ];
    D.modal({ title: '바로 가기 키', width: 560, buttons: [['닫기', 'ok', true]], body: `<table class="kbd-table">${rows.map(([k, d]) => `<tr><td><kbd>${k}</kbd></td><td>${d}</td></tr>`).join('')}</table><p class="muted small">일부 키(Ctrl+N, Ctrl+W, Ctrl+Tab, Ctrl+PgUp/PgDn 등)는 브라우저가 먼저 가로챌 수 있습니다. 에이전트 미리보기 창이나 앱 모드 창에서는 대부분 동작합니다.</p>` });
  }

  // ------------------------------------------------------------------ keyboard
  onKeyDown(e) {
    if (D.modalOpen() || document.querySelector('.find-panel')?.contains(e.target) || document.querySelector('.filter-menu')) return;
    const t = e.target;
    const inField = (t.tagName === 'INPUT' || t.tagName === 'SELECT' || t.tagName === 'TEXTAREA') && t !== this.editor.ta && t !== this.editor.fx;
    const ctrl = e.ctrlKey || e.metaKey, sh = e.shiftKey, alt = e.altKey;
    const k = e.key;
    const K = k.length === 1 ? k.toLowerCase() : k;
    const go = (fn) => { e.preventDefault(); e.stopPropagation(); fn(); };
    // global (work even while editing / in fields)
    if (ctrl && !alt && K === 's' && !sh) return go(() => this.exec('save'));
    if (k === 'F12') return go(() => this.exec('saveAs'));
    if (ctrl && K === 'o') return go(() => this.exec('open'));
    if (ctrl && K === 'n' && !sh) return go(() => this.exec('new'));
    if (alt && K === 'n' && !ctrl) return go(() => this.exec('new'));
    if (ctrl && k === 'F4') return go(() => this.exec('close'));
    if (ctrl && k === 'Tab') return go(() => this.activateDoc((this.di + (sh ? -1 : 1) + this.docs.length) % this.docs.length));
    if (ctrl && (k === 'PageDown' || k === 'PageUp')) return go(() => this.cycleSheet(k === 'PageDown' ? 1 : -1));
    if (inField) return;
    if (this.editor.active) return; // editor handles its own keys
    if (this.painter && k === 'Escape') { this.painter = null; document.body.classList.remove('painting'); this.setStatus(''); return; }
    if (this.charts.selected && (k === 'Delete' || k === 'Backspace')) return go(() => this.charts.remove(this.charts.selected));
    if (k === 'Escape') { this.clipboard.clear(); this.charts.deselect(); return; }
    const v = this.view();
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (alt && k === 'ArrowDown') {
      const s2 = this.doc.wb.sheet;
      if (s2.autoFilter && v.active.r === s2.autoFilter.range.r1 && v.active.c >= s2.autoFilter.range.c1 && v.active.c <= s2.autoFilter.range.c2) {
        const rc = this.grid.cellRect(v.active.r, v.active.c), hr = this.gridHost.getBoundingClientRect();
        return go(() => this.filters.openMenu(v.active.c, hr.left + rc.x + rc.w - 16, hr.top + rc.y + rc.h));
      }
    }
    if (arrows[k] && !alt) {
      const [dr, dc] = arrows[k];
      return go(() => {
        if (ctrl) {
          const from = sh ? (v.ext || v.active) : v.active;
          const to = this.edge(from.r, from.c, dr, dc);
          if (sh) { this.extendTo(to.r, to.c); } else this.select(to.r, to.c);
        } else if (sh) {
          const ext = v.ext && inRange(v.sel, v.ext.r, v.ext.c) ? v.ext : { r: v.sel.r1 === v.anchor.r ? v.sel.r2 : v.sel.r1, c: v.sel.c1 === v.anchor.c ? v.sel.c2 : v.sel.c1 };
          this.extendTo(ext.r + dr, ext.c + dc);
        } else this.moveActive(dr, dc);
      });
    }
    if (k === 'Enter') return go(() => this.moveActive(sh ? -1 : 1, 0, { withinSelection: true }));
    if (k === 'Tab') return go(() => this.moveActive(0, sh ? -1 : 1, { withinSelection: true }));
    if (k === 'Home') return go(() => (ctrl ? this.select(this.doc.wb.sheet.freeze.r || 0, this.doc.wb.sheet.freeze.c || 0) : this.select(v.active.r, 0)));
    if (k === 'End' && ctrl) return go(() => { const s2 = this.doc.wb.sheet; s2.updateExtent(); sh ? this.extendTo(s2.maxRow, s2.maxCol) : this.select(s2.maxRow, s2.maxCol); });
    if (k === 'PageDown' || k === 'PageUp') {
      return go(() => {
        const n = this.grid.visibleRowCount() * (k === 'PageDown' ? 1 : -1);
        if (alt) { this.select(v.active.r, Math.max(0, v.active.c + (n > 0 ? 10 : -10))); return; }
        this.grid.scroller.scrollTop += n * geom(this.doc.wb.sheet).rows.def * this.view().zoom;
        sh ? this.extendTo((v.ext || v.active).r + n, (v.ext || v.active).c) : this.select(v.active.r + n, v.active.c, { scroll: true });
      });
    }
    if (k === 'F2' && !sh) return go(() => this.editor.start({ mode: 'edit' }));
    if (k === 'F2' && sh) return go(() => this.editComment());
    if (k === 'F4' && !ctrl) return go(() => this.lastAction?.());
    if (k === 'F9') return go(() => this.exec('recalc'));
    if (k === 'F11' && sh) return go(() => this.exec('insertSheet'));
    if (k === 'F1' && alt) return go(() => this.exec('chart', 'bar'));
    if (k === 'Delete') return go(() => this.clear('contents'));
    if (k === 'Backspace') return go(() => { this.clear('contents'); this.editor.start({ mode: 'enter', text: '' }); });
    if (alt && (k === '=' || e.code === 'Equal')) return go(() => this.exec('autosum'));
    if (ctrl) {
      const code = e.code;
      if (sh && (code === 'Digit1')) return go(() => this.setStyleProp('nf', '#,##0.00'));
      if (sh && (code === 'Digit2')) return go(() => this.setStyleProp('nf', 'h:mm AM/PM'));
      if (sh && (code === 'Digit3')) return go(() => this.setStyleProp('nf', 'yyyy-mm-dd'));
      if (sh && (code === 'Digit4')) return go(() => this.exec('currency'));
      if (sh && (code === 'Digit5')) return go(() => this.exec('percent'));
      if (sh && (code === 'Digit6')) return go(() => this.setStyleProp('nf', '0.00E+00'));
      if (sh && (code === 'Digit7')) return go(() => this.setBorders('outline'));
      if (sh && (code === 'Minus')) return go(() => this.setBorders('none'));
      if (sh && code === 'Digit9') return go(() => this.hide('row', false));
      if (sh && code === 'Digit0') return go(() => this.hide('col', false));
      if (sh && (code === 'Equal' || k === '+')) return go(() => this.insertDialog(false));
      if (sh && code === 'Semicolon') return go(() => this.insertNow('time'));
      if (sh && K === 'l') return go(() => this.exec('filter'));
      if (sh && K === 'f') return go(() => D.formatCellsDialog(this, 'font'));
      if (sh && K === 'z') return go(() => this.exec('redo'));
      if (alt && K === 'v') return go(() => this.exec('pasteSpecial'));
      if (sh) return;
      switch (code === 'Semicolon' ? ';' : code === 'Minus' ? '-' : code === 'Equal' ? '+' : code === 'Backquote' ? '`' : code === 'Digit1' ? '1' : code === 'Digit5' ? '5' : code === 'Digit9' ? '9' : code === 'Digit0' ? '0' : code === 'Space' ? ' ' : K) {
        case 'z': return go(() => this.exec('undo'));
        case 'y': return go(() => this.exec('redo'));
        case 'b': return go(() => this.exec('bold'));
        case 'i': return go(() => this.exec('italic'));
        case 'u': return go(() => this.exec('underline'));
        case '5': return go(() => this.exec('strike'));
        case '1': return go(() => this.exec('formatCells'));
        case 'a': return go(() => this.selectAllSmart());
        case 'd': return go(() => this.exec('fillDown'));
        case 'r': return go(() => this.exec('fillRight'));
        case 'f': return go(() => this.exec('find'));
        case 'h': return go(() => this.exec('replace'));
        case 'k': return go(() => this.exec('link'));
        case ';': return go(() => this.insertNow('date'));
        case '-': return go(() => this.insertDialog(true));
        case '+': return go(() => this.insertDialog(false));
        case '`': return go(() => this.exec('showFormulas'));
        case '9': return go(() => this.hide('row', true));
        case '0': return go(() => this.hide('col', true));
        case ' ': return go(() => this.selectRange({ r1: 0, c1: v.sel.c1, r2: MAX_ROWS - 1, c2: v.sel.c2 }, v.active, { scroll: false }));
        case 'p': return go(() => this.toast('인쇄는 브라우저 인쇄(Ctrl+P는 차단됨) 대신 Excel에서 수행하세요.'));
        default: return; // let Ctrl+C/X/V fire native clipboard events
      }
    }
    if (sh && k === ' ' && !ctrl) return go(() => this.selectRange({ r1: v.sel.r1, c1: 0, r2: v.sel.r2, c2: MAX_COLS - 1 }, v.active, { scroll: false }));
    // printable keys fall through to the editor textarea 'input' event (IME-safe)
  }

  cycleSheet(d) {
    const wb = this.doc.wb;
    let i = wb.active;
    for (let n = 0; n < wb.sheets.length; n++) { i = (i + d + wb.sheets.length) % wb.sheets.length; if (wb.sheets[i].state === 'visible') break; }
    this.activateSheet(i);
  }

  // ------------------------------------------------------------------ UI binding
  bindUI() {
    window.addEventListener('keydown', (e) => this.onKeyDown(e), true);
    // ribbon tabs
    this.$('#ribbon-tabs').addEventListener('click', (e) => {
      const b = e.target.closest('[data-rtab]'); if (!b) return;
      if (b.dataset.rtab === 'file') { this.fileMenu(b); return; }
      document.querySelectorAll('[data-rtab]').forEach((x) => x.classList.toggle('on', x === b));
      document.querySelectorAll('[data-rpane]').forEach((p) => { p.hidden = p.dataset.rpane !== b.dataset.rtab; });
    });
    document.addEventListener('click', (e) => {
      const b = e.target.closest('button[data-cmd]');
      if (b && !b.closest('.ctx-menu')) { this.exec(b.dataset.cmd, b.dataset.arg); return; }
      const dd = e.target.closest('[data-dd]');
      if (dd) this.dropdown(dd.dataset.dd, dd);
    });
    document.addEventListener('change', (e) => {
      const s = e.target.closest('select[data-cmd], input[data-cmd]');
      if (s) this.exec(s.dataset.cmd, s.value);
    });
    for (const id of ['#font-size']) this.$(id).addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); this.exec('fontSize', e.target.value); } });
    // color split buttons
    this.$('#font-color-pick').addEventListener('click', (e) => D.colorPopup(e.currentTarget, { noneLabel: '자동', onPick: (c) => { this.$('#font-color-bar').style.background = c || '#000'; this.exec('fontColor', c); } }));
    this.$('#fill-color-pick').addEventListener('click', (e) => D.colorPopup(e.currentTarget, { onPick: (c) => { this.$('#fill-color-bar').style.background = c || 'transparent'; this.exec('fillColor', c); } }));
    this.$('#font-color-btn').addEventListener('click', () => this.exec('fontColor', this.$('#font-color-bar').style.background ? rgbHex(this.$('#font-color-bar').style.background) : '#FF0000'));
    this.$('#fill-color-btn').addEventListener('click', () => this.exec('fillColor', rgbHex(this.$('#fill-color-bar').style.background) || '#FFFF00'));
    this.$('#painter-btn').addEventListener('dblclick', () => this.startPainter(true));
    // name box
    const nb = this.$('#name-box');
    nb.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') { if (e.key === 'Escape') { this.syncFormulaBar(); this.focusGrid(); } return; }
      e.preventDefault();
      const txt = nb.value.trim();
      const wb = this.doc.wb;
      const target = wb.definedNames[txt] ? wb.definedNames[txt] : txt;
      const p = parseSheetRef(target);
      if (!p) { this.toast('잘못된 참조입니다.', 'error'); return; }
      if (p.sheet) { const i = wb.sheets.findIndex((s) => s.name.toLowerCase() === p.sheet.toLowerCase()); if (i < 0) return this.toast('시트가 없습니다.', 'error'); this.activateSheet(i); }
      this.selectRange(p.range, { r: p.range.r1, c: p.range.c1 });
      this.focusGrid();
    });
    nb.addEventListener('focus', () => nb.select());
    this.$('#fx-btn').addEventListener('click', () => this.exec('insertFunction'));
    this.$('#fx-cancel').addEventListener('click', () => this.editor.cancel());
    this.$('#fx-ok').addEventListener('click', () => this.editor.commit());
    // doc tabs
    this.$('#doc-tabs').addEventListener('click', (e) => {
      const x = e.target.closest('[data-close]'); if (x) { e.stopPropagation(); this.closeDoc(+x.dataset.close); return; }
      const t = e.target.closest('.doc-tab'); if (t) this.activateDoc(+t.dataset.i);
    });
    this.$('#doc-tabs').addEventListener('auxclick', (e) => { const t = e.target.closest('.doc-tab'); if (t && e.button === 1) this.closeDoc(+t.dataset.i); });
    this.$('#doc-tabs').addEventListener('contextmenu', (e) => {
      const t = e.target.closest('.doc-tab'); if (!t) return; e.preventDefault();
      const i = +t.dataset.i;
      this.menu(e.clientX, e.clientY, [['새 창에서 열기 (동기화)', () => this.newWindow(i)], ['닫기', () => this.closeDoc(i)], ['다른 통합 문서 모두 닫기', async () => { for (let j = this.docs.length - 1; j >= 0; j--) if (this.docs[j] !== this.docs[i]) await this.closeDoc(j); }], null, ['경로 복사', () => navigator.clipboard?.writeText(this.docs[i].wb.path || '')]]);
    });
    // sheet tabs
    const tabs = this.$('#sheet-tabs');
    tabs.addEventListener('click', (e) => { const t = e.target.closest('.sheet-tab'); if (t) this.activateSheet(+t.dataset.i, { keepEditor: this.editor.isFormula() }); });
    tabs.addEventListener('dblclick', (e) => { const t = e.target.closest('.sheet-tab'); if (t) this.renameSheetPrompt(+t.dataset.i); });
    tabs.addEventListener('contextmenu', (e) => {
      const t = e.target.closest('.sheet-tab'); if (!t) return; e.preventDefault();
      const i = +t.dataset.i; this.activateSheet(i);
      const wb = this.doc.wb;
      this.menu(e.clientX, e.clientY, [['삽입', () => wb.addSheet(null, i)], ['삭제', () => this.deleteSheet(i)], ['이름 바꾸기', () => this.renameSheetPrompt(i)],
        ['복사본 만들기', () => wb.duplicateSheet(i)], ['왼쪽으로 이동', () => wb.moveSheet(i, i - 1)], ['오른쪽으로 이동', () => wb.moveSheet(i, i + 1)], null,
        ['탭 색', () => this.tabColor(i, t)], ['숨기기', () => this.hideSheet(i)], ['숨기기 취소…', () => this.unhideSheet()]]);
    });
    let dragFrom = -1;
    tabs.addEventListener('dragstart', (e) => { const t = e.target.closest('.sheet-tab'); dragFrom = t ? +t.dataset.i : -1; e.dataTransfer.effectAllowed = 'move'; });
    tabs.addEventListener('dragover', (e) => { if (dragFrom >= 0) e.preventDefault(); });
    tabs.addEventListener('drop', (e) => { e.preventDefault(); const t = e.target.closest('.sheet-tab'); if (t && dragFrom >= 0) this.doc.wb.moveSheet(dragFrom, +t.dataset.i); dragFrom = -1; });
    this.$('#add-sheet').addEventListener('click', () => this.exec('insertSheet'));
    this.$('#sheet-prev').addEventListener('click', () => this.cycleSheet(-1));
    this.$('#sheet-next').addEventListener('click', () => this.cycleSheet(1));
    // zoom
    this.$('#zoom-range').addEventListener('input', (e) => this.setZoom(+e.target.value / 100));
    this.$('#zoom-label').addEventListener('click', () => this.exec('zoom100'));
    this.$('#win-btn').addEventListener('click', (e) => this.windowsMenu(e.currentTarget));
    // painter completes on mouseup in grid
    this.grid.scroller.addEventListener('pointerup', () => setTimeout(() => this.onSelectionDone(), 0));
    // dropping files would make the *browser* read bytes (DRM!) — refuse and explain
    window.addEventListener('dragover', (e) => { if (e.dataTransfer?.types?.includes('Files')) e.preventDefault(); });
    window.addEventListener('drop', (e) => { if (e.dataTransfer?.files?.length) { e.preventDefault(); this.toast('DRM 보호를 위해 파일은 [열기] (Ctrl+O)로 서버(Python)를 통해 엽니다.', 'warn', 6000); } });
    // ribbon state follows focus changes
    this.$('#font-name').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); this.exec('fontName', e.target.value); } });
  }

  fileMenu(anchor) {
    const r = anchor.getBoundingClientRect();
    this.menu(r.left, r.bottom, [
      ['새로 만들기', () => this.exec('new'), { key: 'Ctrl+N' }], ['열기…', () => this.exec('open'), { key: 'Ctrl+O' }],
      ['저장', () => this.exec('save'), { key: 'Ctrl+S' }], ['다른 이름으로 저장…', () => this.exec('saveAs'), { key: 'F12' }], null,
      ['새 창', () => this.exec('newWindow')], ['닫기', () => this.exec('close'), { key: 'Ctrl+F4' }], null,
      ['바로 가기 키', () => this.exec('help')],
    ]);
  }

  dropdown(name, anchor) {
    const r = anchor.getBoundingClientRect();
    const at = (items) => this.menu(r.left, r.bottom + 2, items);
    const c = (cmd, arg) => () => this.exec(cmd, arg);
    switch (name) {
      case 'paste': return at([['붙여넣기', c('paste'), { key: 'Ctrl+V' }], ['값 붙여넣기', () => this.clipboard.pasteCommand('values')], ['서식 붙여넣기', () => this.clipboard.pasteCommand('formats')], ['수식 붙여넣기', () => this.clipboard.pasteCommand('formulas')], ['행/열 바꿈', () => this.clipboard.pasteCommand('all', true)], ['선택하여 붙여넣기…', c('pasteSpecial'), { key: 'Ctrl+Alt+V' }]]);
      case 'borders': return at([['아래쪽 테두리', c('borders', 'bottom')], ['위쪽 테두리', c('borders', 'top')], ['왼쪽 테두리', c('borders', 'left')], ['오른쪽 테두리', c('borders', 'right')], null, ['테두리 없음', c('borders', 'none')], ['모든 테두리', c('borders', 'all')], ['바깥쪽 테두리', c('borders', 'outline')], ['굵은 바깥쪽 테두리', c('borders', 'thickOutline')], ['안쪽 테두리', c('borders', 'inside')], null, ['아래쪽 이중 테두리', c('borders', 'doubleBottom')], ['굵은 아래쪽 테두리', c('borders', 'thickBottom')], ['위쪽/아래쪽 테두리', c('borders', 'topBottom')], null, ['다른 테두리…', () => D.formatCellsDialog(this, 'border')]]);
      case 'merge': return at([['병합하고 가운데 맞춤', c('mergeCenter')], ['전체 병합', c('mergeAcross')], ['셀 병합', c('merge')], ['셀 분할', c('unmerge')]]);
      case 'insert': return at([['셀 삽입…', c('insertCells')], ['시트 행 삽입', c('insertRow')], ['시트 열 삽입', c('insertCol')], ['시트 삽입', c('insertSheet'), { key: 'Shift+F11' }]]);
      case 'delete': return at([['셀 삭제…', c('deleteCells')], ['시트 행 삭제', c('deleteRow')], ['시트 열 삭제', c('deleteCol')], ['시트 삭제', c('deleteSheet')]]);
      case 'format': return at([['행 높이…', c('rowHeight')], ['행 높이 자동 맞춤', c('autofitRow')], ['열 너비…', c('colWidth')], ['열 너비 자동 맞춤', c('autofitCol')], null, ['행 숨기기', () => this.hide('row', true)], ['열 숨기기', () => this.hide('col', true)], ['행 숨기기 취소', () => this.hide('row', false)], ['열 숨기기 취소', () => this.hide('col', false)], ['시트 숨기기', () => this.hideSheet(this.doc.wb.active)], ['시트 숨기기 취소…', c('unhideSheet')], null, ['시트 이름 바꾸기', () => this.renameSheetPrompt(this.doc.wb.active)], ['탭 색', () => this.tabColor(this.doc.wb.active, anchor)], null, ['셀 서식…', c('formatCells'), { key: 'Ctrl+1' }]]);
      case 'autosum': return at([['합계', c('autosum', 'SUM')], ['평균', c('autosum', 'AVERAGE')], ['숫자 개수', c('autosum', 'COUNT')], ['최대값', c('autosum', 'MAX')], ['최소값', c('autosum', 'MIN')], null, ['기타 함수…', c('insertFunction')]]);
      case 'clear': return at([['모두 지우기', c('clearAll')], ['서식 지우기', c('clearFormats')], ['내용 지우기', c('clearContents'), { key: 'Delete' }], ['메모 지우기', c('clearComments')]]);
      case 'fill': return at([['아래로', c('fillDown'), { key: 'Ctrl+D' }], ['오른쪽으로', c('fillRight'), { key: 'Ctrl+R' }]]);
      case 'sortfilter': return at([['텍스트 오름차순 정렬', c('sortAsc')], ['텍스트 내림차순 정렬', c('sortDesc')], ['사용자 지정 정렬…', c('sortDialog')], null, ['필터', c('filter'), { key: 'Ctrl+Shift+L' }], ['지우기', c('filterClear')], ['다시 적용', c('filterReapply')]]);
      case 'find': return at([['찾기…', c('find'), { key: 'Ctrl+F' }], ['바꾸기…', c('replace'), { key: 'Ctrl+H' }], ['이동…', () => this.$('#name-box').focus(), { key: 'Ctrl+G' }]]);
      case 'freeze': return at([['틀 고정', c('freeze')], ['첫 행 고정', c('freezeRow')], ['첫 열 고정', c('freezeCol')], ['틀 고정 취소', c('unfreeze')]]);
      case 'zoom': return at([['25%', () => this.setZoom(0.25)], ['50%', () => this.setZoom(0.5)], ['75%', () => this.setZoom(0.75)], ['100%', () => this.setZoom(1)], ['150%', () => this.setZoom(1.5)], ['200%', () => this.setZoom(2)], ['선택 영역 확대/축소', c('zoomSel')]]);
      case 'windows': return this.windowsMenu(anchor);
      default: return null;
    }
  }

  async start() {
    const params = new URLSearchParams(location.search);
    const want = params.get('doc'), open = params.get('open');
    let docs = [];
    try { docs = (await api.docs()).docs; } catch (e) {
      this.toast(`서버에 연결할 수 없습니다: ${e.message}. 페이지를 새로 고치세요.`, 'error', 15000);
      return;
    }
    for (const d of docs) await this.attachDoc(d.docId, { activate: d.docId === want });
    if (!this.docs.length) await this.newDoc();
    if (want && !this.docs.some((d) => d.wb.docId === want)) this.toast('요청한 문서가 서버에 없습니다.', 'warn');
    if (want) history.replaceState(null, '', location.pathname);
    if (open) await this.openFile(open);
    this.started = true;
  }

  /** Documents opened by the agent (or another window) appear here automatically. */
  onDocOpened(d) {
    if (!this.started || this.docs.some((x) => x.wb.docId === d.docId)) return;
    this.attachDoc(d.docId, { activate: true }).then((doc) => doc && this.setStatus(`'${d.name}' 문서가 열렸습니다.`));
  }
  onDocsList(docs) { this.onWindowsChanged(docs); }
  /** the agent asked to show a sheet/range in the preview */
  onGoto(d) {
    const i = this.docs.findIndex((x) => x.wb.docId === d.docId);
    if (i < 0 || this.editor.active) return;
    if (i !== this.di) this.activateDoc(i);
    const si = this.doc.wb.sheets.findIndex((s) => s.id === d.sheet);
    if (si >= 0 && si !== this.doc.wb.active) this.activateSheet(si);
    const g = parseRange(d.range);
    if (g) this.selectRange(g, { r: g.r1, c: g.c1 });
  }
  onView(d) {
    const doc = this.docs.find((x) => x.wb.docId === d.docId);
    if (!doc || doc !== this.doc) return;
    if (typeof d.showFormulas === 'boolean' && d.showFormulas !== this.view().showFormulas) this.exec('showFormulas');
    if (d.zoom) this.setZoom(+d.zoom);
  }
}

// ------------------------------------------------------------------ helpers
function range(a, b) { const out = []; for (let i = a; i <= b; i++) out.push(i); return out; }
function rgbHex(c) { if (!c) return null; const m = /rgb\((\d+),\s*(\d+),\s*(\d+)/.exec(c); return m ? '#' + [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase() : c; }
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
function nextSize(cur, d) { if (d > 0) return SIZES.find((s) => s > cur) || cur + 10; return [...SIZES].reverse().find((s) => s < cur) || Math.max(1, cur - 1); }

/** AutoFill series generator for one line of source cells. */
function makeSeries(cells, wb) {
  const vals = cells.map((c) => (c && !c.f ? c.v : undefined));
  const n = cells.length;
  const allNum = vals.every((v) => typeof v === 'number');
  const copy = (k, base) => (base ? { ...base } : null);
  if (allNum) {
    const isDate = cells[0]?.s?.nf && isDateFormat(cells[0].s.nf);
    if (n === 1) return (k, base) => (isDate ? { ...base, v: vals[0] + k } : copy(k, base));
    // least-squares linear trend (Excel AutoFill for numeric series)
    const xs = vals.map((_, i) => i), mx = (n - 1) / 2, my = vals.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0;
    xs.forEach((x, i) => { sxy += (x - mx) * (vals[i] - my); sxx += (x - mx) ** 2; });
    const b = sxx ? sxy / sxx : 0, a = my - b * mx;
    return (k, base) => ({ ...base, v: parseFloat((a + b * k).toPrecision(15)) });
  }
  // known lists
  for (const L of LISTS) {
    const idx = vals.map((v) => (typeof v === 'string' ? L.findIndex((x) => x.toLowerCase() === v.toLowerCase()) : -1));
    if (idx.every((i) => i >= 0)) {
      const step = n > 1 ? (idx[n - 1] - idx[0]) / (n - 1) || 1 : 1;
      return (k, base) => ({ ...base, v: L[(((idx[0] + Math.round(step * k)) % L.length) + L.length) % L.length] });
    }
  }
  // text with trailing number: "항목 1", "Q1"
  const m = vals.map((v) => (typeof v === 'string' ? /^(.*?)(\d+)(\D*)$/.exec(v) : null));
  if (m.every(Boolean) && m.every((x) => x[1] === m[0][1] && x[3] === m[0][3])) {
    const nums = m.map((x) => parseInt(x[2], 10));
    const step = n > 1 ? (nums[n - 1] - nums[0]) / (n - 1) || 1 : 1;
    const width = m[0][2].length;
    return (k, base) => { const v = Math.round(nums[0] + step * k); return { ...base, v: m[0][1] + (m[0][2].startsWith('0') ? String(Math.abs(v)).padStart(width, '0') : v) + m[0][3] }; };
  }
  return copy;
}

const app = new App();
window.webExcel = app; // for debugging & e2e tests
app.start();

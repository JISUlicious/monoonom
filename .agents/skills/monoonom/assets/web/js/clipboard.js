// Clipboard: internal rich copy (formulas/styles) + system clipboard as TSV/HTML for Excel interop.
import { key, MAX_ROWS, MAX_COLS } from './address.js';
import { formatValue, parseInput } from './numfmt.js';
import { internStyle } from './model.js';
import { t } from './i18n.js';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export class Clipboard {
  constructor(app) {
    this.app = app;
    this.state = null; // {wb, sheet, range, cut, block, text}
    this.timer = 0;
    document.addEventListener('copy', (e) => this.onCopy(e, false));
    document.addEventListener('cut', (e) => this.onCopy(e, true));
    document.addEventListener('paste', (e) => this.onPaste(e));
  }

  gridHasFocus() {
    const a = document.activeElement;
    return this.app.editor && !this.app.editor.active && (a === this.app.editor.ta || a === document.body || a?.closest?.('#grid-host'));
  }

  clampRange(sh, g) {
    return { ...g, r2: Math.min(g.r2, Math.max(sh.maxRow, g.r1)), c2: Math.min(g.c2, Math.max(sh.maxCol, g.c1)) };
  }

  capture(cut) {
    const wb = this.app.doc.wb, sh = wb.sheet;
    const range = this.clampRange(sh, this.app.view().sel);
    const block = [], textRows = [], htmlRows = [];
    for (let r = range.r1; r <= range.r2; r++) {
      if (sh.filterHidden.has(r) && !cut) continue; // like Excel: copy visible rows of filtered ranges
      const row = [], trow = [], hrow = [];
      for (let c = range.c1; c <= range.c2; c++) {
        const cell = sh.cells.get(key(r, c));
        const v = wb.value(sh, r, c);
        row.push(cell ? { ...cell, _value: v, _r: r } : null);
        const text = v === null || v === undefined ? '' : formatValue(v, cell?.s?.nf).text;
        trow.push(/[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text);
        const st = cell?.s || {};
        const css = [st.b && 'font-weight:bold', st.i && 'font-style:italic', st.u && 'text-decoration:underline', st.fc && `color:${st.fc}`, st.bg && `background:${st.bg}`, st.ha && `text-align:${st.ha}`, st.fs && `font-size:${st.fs}pt`].filter(Boolean).join(';');
        hrow.push(`<td${css ? ` style="${css}"` : ''}>${esc(text)}</td>`);
      }
      block.push(row); textRows.push(trow.join('\t')); htmlRows.push(`<tr>${hrow.join('')}</tr>`);
    }
    const text = textRows.join('\r\n');
    const html = `<table>${htmlRows.join('')}</table>`;
    this.state = { wb, sheet: sh, range, cut, block, text: normalizeText(text) };
    this.startMarch();
    return { text, html };
  }

  onCopy(e, cut) {
    if (!this.gridHasFocus() || this.app.charts?.selected) {
      if (this.app.charts?.selected && !cut) { /* chart copy unsupported */ }
      return;
    }
    e.preventDefault();
    const { text, html } = this.capture(cut);
    e.clipboardData.setData('text/plain', text);
    e.clipboardData.setData('text/html', html);
    this.app.setStatus(cut ? t('Select where to paste the cut range and press Enter or Paste.') : t('Copied. Select where to paste.'));
  }

  /** copy via button (no user-gesture clipboard event) */
  async copyCommand(cut) {
    const { text, html } = this.capture(cut);
    try {
      if (navigator.clipboard?.write && window.ClipboardItem) {
        await navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([text], { type: 'text/plain' }), 'text/html': new Blob([html], { type: 'text/html' }) })]);
      } else await navigator.clipboard?.writeText(text);
    } catch { /* internal clipboard still works */ }
  }

  async pasteCommand(only = 'all', transpose = false) {
    let text = null;
    try { text = await navigator.clipboard.readText(); } catch { /* permission denied */ }
    if (only === 'all' && !text) { // a picture alone (screenshot): the async clipboard API, where allowed
      try {
        for (const item of await navigator.clipboard.read()) {
          const type = item.types.find((ty) => ty.startsWith('image/'));
          if (type) return this.app.images.paste(await item.getType(type));
        }
      } catch { /* not allowed: fall through */ }
    }
    if (this.state && (text === null || normalizeText(text) === this.state.text)) return this.pasteInternal(only, transpose);
    if (text !== null) return this.pasteText(text, null);
    this.app.toast(t('The browser does not allow reading the clipboard. Use Ctrl+V.'), 'warn');
  }

  onPaste(e) {
    if (!this.gridHasFocus()) return;
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    const html = e.clipboardData.getData('text/html');
    if (this.state && normalizeText(text) === this.state.text) return this.pasteInternal('all');
    // a picture: alone (a screenshot), or an image copied from a web page (HTML that is just one <img>);
    // Excel's cell copies also carry a picture, but with text: those paste as cells
    const pic = [...(e.clipboardData.files || [])].find((f) => f.type.startsWith('image/'));
    const onlyImg = html && /^\s*(<meta[^>]*>\s*)?<img\b[^>]*>\s*$/i.test(html.replace(/<!--.*?-->|<\/?(html|body)[^>]*>/gis, ''));
    if (pic && (!text.trim() || onlyImg)) return this.app.images.paste(pic);
    this.pasteText(text, html);
  }

  pasteInternal(only = 'all', transpose = false) {
    const st = this.state;
    const app = this.app, wb = app.doc.wb, sh = wb.sheet;
    if (st.wb !== wb && only === 'all') { /* cross-workbook: formulas keep text */ }
    let block = st.block;
    if (transpose) block = block[0].map((_, c) => block.map((row) => row[c]));
    const sel = app.view().sel;
    const bh = block.length, bw = block[0]?.length || 0;
    if (!bh || !bw) return;
    // tile when the selection is an exact multiple of the copied block
    const selH = sel.r2 - sel.r1 + 1, selW = sel.c2 - sel.c1 + 1;
    const tilesR = selH % bh === 0 && selH < MAX_ROWS ? selH / bh : 1;
    const tilesC = selW % bw === 0 && selW < MAX_COLS ? selW / bw : 1;
    wb.begin(st.cut ? t('Cut') : t('Paste'));
    if (st.cut && st.sheet === sh || st.cut) {
      // clear source first (move semantics)
      for (let r = st.range.r1; r <= st.range.r2; r++) for (let c = st.range.c1; c <= st.range.c2; c++) wb.setCell(st.sheet, r, c, null);
    }
    for (let tr = 0; tr < tilesR; tr++) for (let tc = 0; tc < tilesC; tc++) {
      const r0 = sel.r1 + tr * bh, c0 = sel.c1 + tc * bw;
      const srcR = transpose ? st.range.r1 : st.range.r1, srcC = st.range.c1;
      if (transpose) {
        for (let i = 0; i < bh; i++) for (let j = 0; j < bw; j++) {
          const src = block[i][j];
          wb.pasteBlock(sh, r0 + i, c0 + j, [[src]], { srcR: src ? src._r ?? r0 + i : r0 + i, srcC: srcC + i, translate: !st.cut, only });
        }
      } else {
        // compute per-row source offsets (filtered copies skip rows)
        for (let i = 0; i < bh; i++) {
          const rowSrc = block[i].find(Boolean)?._r ?? srcR + i;
          wb.pasteBlock(sh, r0 + i, c0, [block[i]], { srcR: st.cut ? r0 + i : rowSrc, srcC: st.cut ? c0 : srcC, translate: !st.cut, only });
        }
      }
    }
    // merges inside the copied block
    if (only === 'all' && !transpose) {
      const merges = st.sheet.merges.filter((m) => m.r1 >= st.range.r1 && m.r2 <= st.range.r2 && m.c1 >= st.range.c1 && m.c2 <= st.range.c2);
      if (merges.length) {
        wb.touchMeta(sh);
        for (const m of merges) {
          const dr = sel.r1 - st.range.r1, dc = sel.c1 - st.range.c1;
          sh.merges.push({ r1: m.r1 + dr, c1: m.c1 + dc, r2: m.r2 + dr, c2: m.c2 + dc });
        }
      }
    }
    wb.commit();
    app.selectRange({ r1: sel.r1, c1: sel.c1, r2: sel.r1 + bh * tilesR - 1, c2: sel.c1 + bw * tilesC - 1 }, { r: sel.r1, c: sel.c1 });
    if (st.cut) { this.clear(); }
  }

  pasteText(text, html) {
    const app = this.app, wb = app.doc.wb, sh = wb.sheet;
    let rows = null, styles = null;
    if (html && /<table/i.test(html)) {
      const parsed = parseHtmlTable(html);
      if (parsed) { rows = parsed.rows; styles = parsed.styles; }
    }
    if (!rows) rows = parseTSV(text);
    if (!rows.length) return;
    const sel = app.view().sel;
    wb.begin(t('Paste'));
    rows.forEach((row, i) => row.forEach((t, j) => {
      const r = sel.r1 + i, c = sel.c1 + j;
      if (r >= MAX_ROWS || c >= MAX_COLS) return;
      const cur = sh.get(r, c);
      let cell;
      if (typeof t === 'string' && t.startsWith('=') && t.length > 1) cell = { f: t, s: cur?.s };
      else {
        const p = parseInput(t ?? '');
        let s = cur?.s || null;
        if (p.nf && (!s || !s.nf)) s = internStyle({ ...(s || {}), nf: p.nf });
        const st2 = styles?.[i]?.[j];
        if (st2) s = internStyle({ ...(s || {}), ...st2 });
        cell = { v: p.v, t: p.t, s };
      }
      wb.setCell(sh, r, c, cell);
    }));
    wb.commit();
    const w = Math.max(...rows.map((r) => r.length));
    app.selectRange({ r1: sel.r1, c1: sel.c1, r2: sel.r1 + rows.length - 1, c2: sel.c1 + w - 1 }, { r: sel.r1, c: sel.c1 });
  }

  startMarch() {
    clearInterval(this.timer);
    this.timer = setInterval(() => { this.app.grid.marchOffset = (this.app.grid.marchOffset + 1) % 14; this.app.grid.render(); }, 90);
  }
  clear() { this.state = null; clearInterval(this.timer); this.app.grid?.render(); }
}

function normalizeText(t) { return String(t ?? '').replace(/\r\n/g, '\n').replace(/\n$/, ''); }

export function parseTSV(text) {
  text = String(text ?? '').replace(/\r\n/g, '\n');
  if (text.endsWith('\n')) text = text.slice(0, -1);
  const rows = [];
  let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; }
      else if (ch === '"') q = false;
      else cur += ch;
      continue;
    }
    if (ch === '"' && cur === '') { q = true; continue; }
    if (ch === '\t') { row.push(cur); cur = ''; continue; }
    if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; continue; }
    cur += ch;
  }
  row.push(cur); rows.push(row);
  return rows;
}

function parseHtmlTable(html) {
  try {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const table = doc.querySelector('table');
    if (!table) return null;
    const rows = [], styles = [];
    for (const tr of table.querySelectorAll('tr')) {
      const row = [], srow = [];
      for (const td of tr.querySelectorAll('td,th')) {
        const span = parseInt(td.getAttribute('colspan') || '1', 10);
        const text = td.innerText ?? td.textContent;
        row.push(text.replace(/ /g, ' ').replace(/\n$/, ''));
        const css = td.getAttribute('style') || '';
        const st = {};
        if (/font-weight:\s*(bold|[6-9]00)/i.test(css) || td.querySelector('b,strong')) st.b = true;
        if (/font-style:\s*italic/i.test(css)) st.i = true;
        const bg = /background(?:-color)?:\s*(#[0-9a-f]{6}|#[0-9a-f]{3})/i.exec(css); if (bg) st.bg = bg[1];
        const fc = /(?:^|;)\s*color:\s*(#[0-9a-f]{6}|#[0-9a-f]{3})/i.exec(css); if (fc) st.fc = fc[1];
        srow.push(Object.keys(st).length ? st : null);
        for (let k = 1; k < span; k++) { row.push(''); srow.push(null); }
      }
      rows.push(row); styles.push(srow);
    }
    return rows.length ? { rows, styles } : null;
  } catch { return null; }
}

// Modal dialogs: file open/save (server-side browsing), format cells, find/replace, chart, sort, etc.
import { api } from './api.js';
import { FORMAT_PRESETS, formatValue } from './numfmt.js';
import { CHART_TYPES, ChartLayer } from './charts.js';
import { parseSheetRef, parseRange, rangeA1, colName, key, a1 } from './address.js';
import { quoteSheetName } from './formula.js';
import { combo } from './combo.js';
import { FN_HELP, fnHelp } from './fnhelp.js';
import { FUNCTION_NAMES } from './engine.js';
import { t, LANG } from './i18n.js';

export const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export const THEME_COLORS = [
  ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'],
  ['#F2F2F2', '#7F7F7F', '#D0CECE', '#D6DCE4', '#D9E1F2', '#FCE4D6', '#EDEDED', '#FFF2CC', '#DDEBF7', '#E2EFDA'],
  ['#D9D9D9', '#595959', '#AEAAAA', '#ACB9CA', '#B4C6E7', '#F8CBAD', '#DBDBDB', '#FFE699', '#BDD7EE', '#C6E0B4'],
  ['#BFBFBF', '#404040', '#757171', '#8497B0', '#8EA9DB', '#F4B084', '#C9C9C9', '#FFD966', '#9BC2E6', '#A9D08E'],
  ['#A6A6A6', '#262626', '#3A3838', '#333F4F', '#2F5597', '#C65911', '#7B7B7B', '#BF8F00', '#2F75B5', '#548235'],
];
export const STANDARD_COLORS = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];

let openModal = null;
export function modalOpen() { return !!openModal; }

/** Generic modal. body: HTML string; returns {el, close, done(promise)} */
export function modal({ title, body, buttons = [[t('OK'), 'ok', true], [t('Cancel'), 'cancel']], width = 420, onOk, onOpen, className = '' }) {
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.innerHTML = `<div class="modal ${className}" role="dialog" aria-modal="true" style="width:${width}px">
    <div class="modal-title"><span>${esc(title)}</span><button class="x" data-act="cancel" aria-label="${t('Close')}">×</button></div>
    <div class="modal-body">${body}</div>
    <div class="modal-btns">${buttons.map(([l, a, p]) => `<button class="btn${p ? ' primary' : ''}" data-act="${a}">${l}</button>`).join('')}</div></div>`;
  document.body.append(overlay);
  const el = overlay.querySelector('.modal');
  let resolve;
  const done = new Promise((r) => { resolve = r; });
  const prevFocus = document.activeElement;
  const close = (val) => { overlay.remove(); openModal = null; resolve(val); prevFocus?.focus?.(); };
  const handler = async (act) => {
    if (act === 'cancel') return close(null);
    if (act === 'ok' && onOk) { const r = await onOk(el, close); if (r === false) return; if (r !== undefined) return close(r); return close(true); }
    if (act === 'ok') return close(true);
    return close(act);
  };
  overlay.addEventListener('click', (e) => { const a = e.target.closest('[data-act]')?.dataset.act; if (a && e.target.closest('.modal-btns, .modal-title')) handler(a); });
  overlay.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') { e.preventDefault(); close(null); }
    if (e.key === 'Enter' && !e.isComposing && e.target.tagName !== 'TEXTAREA' && !e.target.closest('.no-enter')) { e.preventDefault(); handler('ok'); }
  });
  // draggable title
  const tb = el.querySelector('.modal-title');
  tb.addEventListener('pointerdown', (e) => {
    if (e.target.closest('button')) return;
    const r = el.getBoundingClientRect(); const sx = e.clientX - r.left, sy = e.clientY - r.top;
    el.style.position = 'fixed'; el.style.margin = '0';
    const mv = (ev) => { el.style.left = ev.clientX - sx + 'px'; el.style.top = ev.clientY - sy + 'px'; };
    const up = () => { document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); };
    document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up);
  });
  openModal = { el, close };
  onOpen?.(el, close);
  setTimeout(() => (el.querySelector('[autofocus]') || el.querySelector('input,select,textarea,button.primary'))?.focus(), 0);
  return { el, close, done };
}

export function alertBox(msg, title = t('Notice')) { return modal({ title, body: `<p class="msg">${esc(msg).replace(/\n/g, '<br>')}</p>`, buttons: [[t('OK'), 'ok', true]] }).done; }
export function confirmBox(msg, { title = t('Confirm'), yes = t('Yes'), no = t('No'), cancel = false } = {}) {
  const buttons = [[yes, 'ok', true], [no, 'no']];
  if (cancel) buttons.push([t('Cancel'), 'cancel']);
  return modal({ title, body: `<p class="msg">${esc(msg).replace(/\n/g, '<br>')}</p>`, buttons }).done.then((v) => (v === true ? 'yes' : v === 'no' ? 'no' : null));
}
export function promptBox(msg, value = '', title = t('Input')) {
  return modal({ title, body: `<label class="fld">${esc(msg)}<input class="inp" id="pb" value="${esc(value)}" autofocus></label>`,
    onOk: (el) => el.querySelector('#pb').value, onOpen: (el) => setTimeout(() => el.querySelector('#pb').select(), 0) }).done;
}

// ------------------------------------------------------------------ file dialog (server-side)
const RECENT_KEY = 'monoonom.recent';
export function recentFiles() { try { return JSON.parse(localStorage.getItem(RECENT_KEY) || '[]'); } catch { return []; } }
export function addRecent(path) {
  try { const l = [path, ...recentFiles().filter((p) => p !== path)].slice(0, 12); localStorage.setItem(RECENT_KEY, JSON.stringify(l)); } catch { /* ignore */ }
}
const LASTDIR_KEY = 'monoonom.lastdir';

export async function fileDialog({ mode = 'open', name = '', format = 'xlsx', startDir = null } = {}) {
  const info = await api.info().catch(() => ({ start: '', sep: '/' }));
  let dir = startDir;
  if (!dir) { try { dir = localStorage.getItem(LASTDIR_KEY); } catch { /* ignore */ } }
  dir = dir || info.start;
  const sep = info.sep || '/';
  const saving = mode === 'save';
  const body = `
    <div class="fd-path"><button class="btn sm" data-fd="up" title="${t('Up one level')}">↑</button>
      ${info.drives?.length ? `<select class="inp sm" data-fd="drive">${info.drives.map((d) => `<option>${esc(d)}</option>`).join('')}</select>` : ''}
      <input class="inp" data-fd="dir" value="${esc(dir)}" spellcheck="false"><button class="btn sm" data-fd="go">${t('Go')}</button></div>
    <div class="fd-main"><div class="fd-side"><div class="fd-side-title">${t('Recent files')}</div>${recentFiles().map((p) => `<div class="fd-recent" data-path="${esc(p)}" title="${esc(p)}">${esc(p.split(/[\\/]/).pop())}</div>`).join('') || `<div class="muted">${t('None')}</div>`}</div>
      <div class="fd-list" tabindex="0"></div></div>
    ${saving ? `<div class="fd-save"><label>${t('File name')} <input class="inp" data-fd="name" value="${esc(name)}"></label>
      <label>${t('Save as type')} <select class="inp" data-fd="fmt"><option value="xlsx">${t('Excel Workbook (*.xlsx)')}</option><option value="xlsm">${t('Excel Macro-Enabled Workbook (*.xlsm)')}</option></select></label></div>`
      : `<div class="fd-save"><label>${t('File path')} <input class="inp" data-fd="name" placeholder="${t('Select a file or type a full path')}"></label></div>`}
    <p class="muted small">${t('Files are read and written by the local Python server (the browser never reads file contents directly).')}</p>`;
  let current = dir;
  const dlg = modal({
    title: saving ? t('Save As') : t('Open'), body, width: 720, className: 'file-dialog',
    buttons: [[saving ? t('Save') : t('Open'), 'ok', true], [t('Cancel'), 'cancel']],
    onOk: (el) => {
      let n = el.querySelector('[data-fd=name]').value.trim();
      if (!n) { const s = el.querySelector('.fd-item.sel'); if (s && !s.dataset.dir) n = s.dataset.path; }
      if (!n) return false;
      const isAbs = /^([A-Za-z]:[\\/]|\/|\\\\)/.test(n);
      let path = isAbs ? n : current.replace(/[\\/]$/, '') + sep + n;
      if (saving) {
        const fmt = el.querySelector('[data-fd=fmt]').value;
        if (!/\.(xlsx|xlsm|xls)$/i.test(path)) path += '.' + fmt;
        else path = path.replace(/\.(xlsx|xlsm|xls)$/i, '.' + fmt);
        try { localStorage.setItem(LASTDIR_KEY, current); } catch { /* ignore */ }
        return { path, format: fmt };
      }
      try { localStorage.setItem(LASTDIR_KEY, current); } catch { /* ignore */ }
      return { path };
    },
  });
  const el = dlg.el;
  if (saving) el.querySelector('[data-fd=fmt]').value = format === 'xlsm' ? 'xlsm' : 'xlsx';
  const list = el.querySelector('.fd-list');
  const load = async (path) => {
    list.innerHTML = `<div class="muted pad">${t('Loading…')}</div>`;
    try {
      const res = await api.fs(path);
      current = res.path;
      el.querySelector('[data-fd=dir]').value = res.path;
      el.dataset.parent = res.parent || '';
      list.innerHTML = res.entries.map((e) => `<div class="fd-item" data-path="${esc(e.path)}" ${e.dir ? 'data-dir="1"' : ''}>
        <span class="ic">${e.dir ? '📁' : /\.xls$/i.test(e.name) ? '📗' : '📊'}</span><span class="nm">${esc(e.name)}</span>
        <span class="sz">${e.dir ? '' : fmtSize(e.size)}</span><span class="dt">${new Date(e.mtime * 1000).toLocaleString(LANG === 'ko' ? 'ko-KR' : 'en-US', { dateStyle: 'short', timeStyle: 'short' })}</span></div>`).join('') || `<div class="muted pad">${t('No Excel files.')}</div>`;
    } catch (err) { list.innerHTML = `<div class="err pad">${esc(err.message)}</div>`; }
  };
  list.addEventListener('click', (e) => {
    const it = e.target.closest('.fd-item'); if (!it) return;
    list.querySelectorAll('.sel').forEach((x) => x.classList.remove('sel'));
    it.classList.add('sel');
    if (!it.dataset.dir) el.querySelector('[data-fd=name]').value = it.querySelector('.nm').textContent;
  });
  list.addEventListener('dblclick', (e) => {
    const it = e.target.closest('.fd-item'); if (!it) return;
    if (it.dataset.dir) load(it.dataset.path);
    else { el.querySelector('[data-fd=name]').value = it.querySelector('.nm').textContent; el.querySelector('[data-act=ok]').click(); }
  });
  el.querySelector('.fd-side').addEventListener('click', (e) => {
    const r = e.target.closest('.fd-recent'); if (!r) return;
    el.querySelector('[data-fd=name]').value = r.dataset.path;
    if (!saving) el.querySelector('[data-act=ok]').click();
  });
  el.querySelector('[data-fd=up]').addEventListener('click', () => el.dataset.parent && load(el.dataset.parent));
  el.querySelector('[data-fd=go]').addEventListener('click', () => load(el.querySelector('[data-fd=dir]').value));
  el.querySelector('[data-fd=dir]').classList.add('no-enter');
  el.querySelector('[data-fd=dir]').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); load(e.target.value); } });
  el.querySelector('[data-fd=drive]')?.addEventListener('change', (e) => load(e.target.value));
  load(dir);
  return dlg.done;
}
function fmtSize(n) { return n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1048576).toFixed(1)} MB`; }

// ------------------------------------------------------------------ color palette popup
export function colorPopup(anchor, { allowNone = true, noneLabel = t('No Fill'), onPick }) {
  document.querySelector('.color-popup')?.remove();
  const p = document.createElement('div');
  p.className = 'color-popup';
  p.innerHTML = `${allowNone ? `<button class="cp-none" data-c="">${noneLabel}</button>` : ''}
    <div class="cp-label">${t('Theme Colors')}</div><div class="cp-grid">${THEME_COLORS.map((row) => row.map((c) => `<button style="background:${c}" data-c="${c}" title="${c}"></button>`).join('')).join('')}</div>
    <div class="cp-label">${t('Standard Colors')}</div><div class="cp-grid">${STANDARD_COLORS.map((c) => `<button style="background:${c}" data-c="${c}" title="${c}"></button>`).join('')}</div>
    <label class="cp-more">${t('More Colors…')} <input type="color"></label>`;
  document.body.append(p);
  const r = anchor.getBoundingClientRect();
  p.style.left = Math.min(r.left, window.innerWidth - p.offsetWidth - 6) + 'px';
  p.style.top = r.bottom + 2 + 'px';
  const close = () => { p.remove(); document.removeEventListener('pointerdown', outside, true); };
  const outside = (e) => { if (!p.contains(e.target) && e.target !== anchor) close(); };
  setTimeout(() => document.addEventListener('pointerdown', outside, true), 0);
  p.addEventListener('click', (e) => { const b = e.target.closest('[data-c]'); if (b) { onPick(b.dataset.c || null); close(); } });
  p.querySelector('input[type=color]').addEventListener('change', (e) => { onPick(e.target.value.toUpperCase()); close(); });
}

// ------------------------------------------------------------------ format cells (Ctrl+1)
export function formatCellsDialog(app, tab = 'num') {
  const sh = app.doc.wb.sheet, v = app.view();
  const cell = sh.get(v.active.r, v.active.c);
  const st = { ...(cell?.s || {}) };
  const sample = app.doc.wb.value(sh, v.active.r, v.active.c) ?? 1234.5678;
  const borderStyles = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair'];
  const body = `
  <div class="tabs"><button data-tab="num">${t('Number format')}</button><button data-tab="align">${t('Alignment')}</button><button data-tab="font">${t('Font')}</button><button data-tab="border">${t('Border')}</button><button data-tab="fill">${t('Fill')}</button></div>
  <section data-pane="num">
    <div class="row2"><select class="inp" size="12" data-k="preset">${FORMAT_PRESETS.map((p) => `<option value="${esc(p.code)}">${esc(t(p.label))}</option>`).join('')}<option value="__custom">${t('Custom')}</option></select>
    <div class="col"><label class="fld">${t('Format code')}<input class="inp" data-k="nf" value="${esc(st.nf || 'General')}"></label>
    <div class="fld">${t('Preview:')} <span class="preview" data-k="nfprev"></span></div>
    <p class="muted small">${esc(t('Examples: #,##0.00 / 0% / yyyy-mm-dd / "$"#,##0.00;[Red]-"$"#,##0.00 / [h]:mm'))}</p></div></div>
  </section>
  <section data-pane="align">
    <div class="grid2">
      <label class="fld">${t('Horizontal')}<select class="inp" data-k="ha"><option value="">${t('General')}</option><option value="left">${t('Left')}</option><option value="center">${t('Centered')}</option><option value="right">${t('Right')}</option><option value="fill">${t('Fill')}</option><option value="justify">${t('Justify')}</option><option value="centerContinuous">${t('Center Across Selection')}</option></select></label>
      <label class="fld">${t('Vertical')}<select class="inp" data-k="va"><option value="top">${t('Top')}</option><option value="center">${t('Middle')}</option><option value="">${t('Bottom')}</option><option value="justify">${t('Justify')}</option></select></label>
      <label class="fld">${t('Indent')}<input class="inp" type="number" min="0" max="15" data-k="ind" value="${st.ind || 0}"></label>
      <label class="chk"><input type="checkbox" data-k="wrap" ${st.wrap ? 'checked' : ''}> ${t('Wrap text')}</label>
      <label class="chk"><input type="checkbox" data-k="merge"> ${t('Merge cells')}</label>
    </div>
  </section>
  <section data-pane="font">
    <div class="grid2">
      <label class="fld">${t('Font')}<input class="inp" data-k="ff" value="${esc(st.ff || app.doc.wb.defaultFont.name)}"></label>
      <label class="fld">${t('Size')}<input class="inp" type="number" min="1" max="409" step="0.5" data-k="fs" value="${st.fs || app.doc.wb.defaultFont.size}"></label>
      <label class="chk"><input type="checkbox" data-k="b" ${st.b ? 'checked' : ''}> ${t('Bold')}</label>
      <label class="chk"><input type="checkbox" data-k="i" ${st.i ? 'checked' : ''}> ${t('Italic')}</label>
      <label class="fld">${t('Underline')}<select class="inp" data-k="u"><option value="">${t('None')}</option><option value="1">${t('Single')}</option><option value="double">${t('Double')}</option></select></label>
      <label class="chk"><input type="checkbox" data-k="st" ${st.st ? 'checked' : ''}> ${t('Strikethrough')}</label>
      <label class="fld">${t('Color')}<input class="inp" type="color" data-k="fc" value="${st.fc || '#000000'}"></label>
    </div>
    <div class="font-preview" data-k="fprev">${t('AaBbCc 123')}</div>
  </section>
  <section data-pane="border">
    <div class="grid2"><label class="fld">${t('Line style')}<select class="inp" data-k="bs">${borderStyles.map((s) => `<option>${s}</option>`).join('')}</select></label>
    <label class="fld">${t('Color')}<input class="inp" type="color" data-k="bc" value="#000000"></label></div>
    <div class="border-btns"><button class="btn" data-b="none">${t('None')}</button><button class="btn" data-b="outline">${t('Outline')}</button><button class="btn" data-b="inside">${t('Inside')}</button><button class="btn" data-b="all">${t('All Borders')}</button>
    <button class="btn" data-b="top">${t('Top edge')}</button><button class="btn" data-b="bottom">${t('Bottom edge')}</button><button class="btn" data-b="left">${t('Left edge')}</button><button class="btn" data-b="right">${t('Right edge')}</button></div>
    <p class="muted small">${t('Selected border action:')} <b data-k="bact">${t('No change')}</b></p>
  </section>
  <section data-pane="fill">
    <div class="cp-grid big">${THEME_COLORS.flat().concat(STANDARD_COLORS).map((c) => `<button style="background:${c}" data-fill="${c}"></button>`).join('')}</div>
    <div class="row2"><button class="btn" data-fill="">${t('No Fill')}</button><label>${t('More Colors')} <input type="color" data-k="bgc" value="${st.bg || '#ffffff'}"></label> <span class="swatch" data-k="bgprev" style="background:${st.bg || '#fff'}"></span></div>
  </section>`;
  let fillChoice = undefined, borderAct = null;
  const dlg = modal({
    title: t('Format Cells'), body, width: 560, className: 'fmt-dialog',
    onOpen: (el) => {
      const ffIn = el.querySelector('[data-k=ff]');
      combo(ffIn, { items: () => app.fontItems(), onPick: () => ffIn.dispatchEvent(new Event('input', { bubbles: true })), width: 240 });
      const show = (t) => { el.querySelectorAll('[data-tab]').forEach((b) => b.classList.toggle('on', b.dataset.tab === t)); el.querySelectorAll('[data-pane]').forEach((p) => { p.hidden = p.dataset.pane !== t; }); };
      el.querySelector('.tabs').addEventListener('click', (e) => { const t = e.target.dataset.tab; if (t) show(t); });
      show(tab);
      const q = (k) => el.querySelector(`[data-k="${k}"]`);
      q('ha').value = st.ha || ''; q('va').value = st.va || ''; q('u').value = st.u === 'double' ? 'double' : st.u ? '1' : '';
      const preset = q('preset');
      preset.value = FORMAT_PRESETS.some((p) => p.code === (st.nf || 'General')) ? (st.nf || 'General') : '__custom';
      const upd = () => { q('nfprev').textContent = formatValue(typeof sample === 'number' ? sample : sample, q('nf').value).text; };
      preset.addEventListener('change', () => { if (preset.value !== '__custom') q('nf').value = preset.value; upd(); });
      q('nf').addEventListener('input', upd); upd();
      const fprev = () => { const p = q('fprev'); p.style.fontFamily = q('ff').value; p.style.fontSize = q('fs').value + 'pt'; p.style.fontWeight = q('b').checked ? 'bold' : ''; p.style.fontStyle = q('i').checked ? 'italic' : ''; p.style.color = q('fc').value; p.style.textDecoration = [q('u').value && 'underline', q('st').checked && 'line-through'].filter(Boolean).join(' '); };
      el.querySelector('[data-pane=font]').addEventListener('input', fprev); fprev();
      const sel = v.sel; q('merge').checked = !!sh.mergeAt(sel.r1, sel.c1) && sh.mergeAt(sel.r1, sel.c1).r2 === sel.r2;
      el.querySelector('[data-pane=fill]').addEventListener('click', (e) => { const b = e.target.closest('[data-fill]'); if (b) { fillChoice = b.dataset.fill || null; q('bgprev').style.background = fillChoice || '#fff'; } });
      q('bgc').addEventListener('input', (e) => { fillChoice = e.target.value; q('bgprev').style.background = fillChoice; });
      el.querySelector('.border-btns').addEventListener('click', (e) => { const b = e.target.dataset.b; if (b) { borderAct = b; q('bact').textContent = e.target.textContent; } });
    },
    onOk: (el) => {
      const q = (k) => el.querySelector(`[data-k="${k}"]`);
      const nf = q('nf').value.trim() || 'General';
      const patch = {
        nf: nf === 'General' ? null : nf, ha: q('ha').value || null, va: q('va').value || null, ind: +q('ind').value || null,
        wrap: q('wrap').checked || null, ff: q('ff').value === app.doc.wb.defaultFont.name ? null : q('ff').value,
        fs: +q('fs').value === app.doc.wb.defaultFont.size ? null : +q('fs').value, b: q('b').checked || null, i: q('i').checked || null,
        u: q('u').value === 'double' ? 'double' : q('u').value ? true : null, st: q('st').checked || null,
        fc: q('fc').value.toUpperCase() === '#000000' ? null : q('fc').value.toUpperCase(),
      };
      if (fillChoice !== undefined) patch.bg = fillChoice;
      app.applyStylePatch(patch, { merge: q('merge').checked, border: borderAct ? { kind: borderAct, s: q('bs').value, c: q('bc').value.toUpperCase() } : null });
    },
  });
  return dlg.done;
}

// ------------------------------------------------------------------ find & replace
export function findDialog(app, replace = false) {
  document.querySelector('.find-panel')?.remove();
  const p = document.createElement('div');
  p.className = 'find-panel';
  p.innerHTML = `<div class="fp-head"><b>${replace ? t('Find and Replace') : t('Find')}</b><button class="x" data-a="close">×</button></div>
    <label class="fld">${t('Find what')}<input class="inp" data-k="q" autofocus></label>
    ${replace ? `<label class="fld">${t('Replace with')}<input class="inp" data-k="r"></label>` : ''}
    <div class="fp-opts"><label><input type="checkbox" data-k="case"> ${t('Match case')}</label><label><input type="checkbox" data-k="whole"> ${t('Match entire cell contents')}</label>
    <label>${t('Within')} <select data-k="scope"><option value="sheet">${t('Sheet')}</option><option value="wb">${t('Workbook')}</option></select></label>
    <label>${t('Look in')} <select data-k="look"><option value="values">${t('Values')}</option><option value="formulas">${t('Formulas')}</option></select></label></div>
    <div class="fp-btns"><button class="btn" data-a="all">${t('Find All')}</button><button class="btn primary" data-a="next">${t('Find Next')}</button>${replace ? `<button class="btn" data-a="rep">${t('Replace')}</button><button class="btn" data-a="repall">${t('Replace All')}</button>` : ''}</div>
    <div class="fp-results"></div>`;
  document.body.append(p);
  const q = (k) => p.querySelector(`[data-k="${k}"]`);
  const matcher = () => {
    const s = q('q').value; if (!s) return null;
    const cs = q('case').checked, whole = q('whole').checked;
    const re = new RegExp((whole ? '^' : '') + s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/~\./g, '\\.') + (whole ? '$' : ''), cs ? 'g' : 'gi');
    return re;
  };
  const cellText = (sh, r, c, cell) => (q('look').value === 'formulas' && cell.f ? cell.f : formatValue(app.doc.wb.value(sh, r, c), cell.s?.nf).text);
  const all = () => {
    const re = matcher(); if (!re) return [];
    const wb = app.doc.wb;
    const sheets = q('scope').value === 'wb' ? wb.sheets : [wb.sheet];
    const out = [];
    for (const sh of sheets) {
      const keys = [...sh.cells.keys()].sort((a, b) => a - b);
      for (const k of keys) {
        const r = Math.floor(k / 16384), c = k % 16384, cell = sh.cells.get(k);
        const t = cellText(sh, r, c, cell);
        re.lastIndex = 0;
        if (t && re.test(t)) out.push({ sh, r, c, t });
      }
    }
    return out;
  };
  let lastIdx = -1;
  const goto = (m) => { const wb = app.doc.wb; const i = wb.sheets.indexOf(m.sh); if (i !== wb.active) app.activateSheet(i); app.select(m.r, m.c); };
  p.addEventListener('click', (e) => {
    const a = e.target.dataset.a;
    if (a === 'close') { p.remove(); app.focusGrid(); return; }
    const res = e.target.closest('.fp-res');
    if (res) { const m = all()[+res.dataset.i]; if (m) goto(m); return; }
    if (!a) return;
    const list = all();
    const box = p.querySelector('.fp-results');
    if (a === 'all') { box.innerHTML = `<div class="muted small">${t('{n} cell(s) found.', { n: list.length })}</div>` + list.slice(0, 500).map((m, i) => `<div class="fp-res" data-i="${i}">${esc(m.sh.name)}!${colName(m.c)}${m.r + 1} — ${esc(m.t.slice(0, 60))}</div>`).join(''); return; }
    if (!list.length) { box.innerHTML = `<div class="muted small">${t('No matches found.')}</div>`; return; }
    if (a === 'next') {
      const v = app.view(), wb = app.doc.wb;
      const curPos = wb.sheets.indexOf(wb.sheet) * 1e12 + key(v.active.r, v.active.c);
      let i = list.findIndex((m) => wb.sheets.indexOf(m.sh) * 1e12 + key(m.r, m.c) > curPos);
      if (i < 0) i = 0;
      lastIdx = i; goto(list[i]);
      box.innerHTML = `<div class="muted small">${i + 1} / ${list.length}</div>`;
    }
    if (a === 'rep' || a === 'repall') {
      const re = matcher(); const rep = q('r').value; const wb = app.doc.wb;
      const targets = a === 'repall' ? list : list.filter((m) => m.sh === wb.sheet && m.r === app.view().active.r && m.c === app.view().active.c);
      if (!targets.length && a === 'rep') { p.querySelector('[data-a=next]').click(); return; }
      wb.begin(t('Replace'));
      let n = 0;
      for (const m of targets) {
        const cell = m.sh.get(m.r, m.c);
        if (cell.f && q('look').value === 'formulas') { wb.setCell(m.sh, m.r, m.c, { ...cell, f: cell.f.replace(re, rep) }); n++; }
        else if (!cell.f && typeof cell.v === 'string') { wb.setCell(m.sh, m.r, m.c, { ...cell, v: cell.v.replace(re, rep) }); n++; }
        else if (!cell.f) { app.setCellInput(m.sh, m.r, m.c, String(m.t).replace(re, rep), { inTxn: true }); n++; }
      }
      wb.commit();
      box.innerHTML = `<div class="muted small">${t('Replaced {n} item(s).', { n })}</div>`;
      if (a === 'rep') p.querySelector('[data-a=next]').click();
    }
  });
  p.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); p.querySelector('[data-a=next]').click(); } if (e.key === 'Escape') { p.remove(); app.focusGrid(); } });
  q('q').focus();
}

// ------------------------------------------------------------------ chart dialog
export function chartDialog(app, model = null, presetType = null) {
  const wb = app.doc.wb, sh = wb.sheet, v = app.view();
  const isNew = !model;
  let range = model?._range;
  if (!range) {
    if (!isNew && model.series[0]?.val) {
      const p = parseSheetRef(model.series[0].val);
      range = p ? rangeA1(p.range) : '';
    } else {
      let g = v.sel;
      if (g.r1 === g.r2 && g.c1 === g.c2) g = app.filters.tableAt(sh, g.r1, g.c1);
      range = rangeA1({ ...g, r2: Math.min(g.r2, sh.maxRow), c2: Math.min(g.c2, sh.maxCol) });
    }
  }
  const curType = model ? (model.type === 'bar' && model.barDir === 'bar' ? 'barH' : model.type === 'bar' && model.grouping === 'stacked' ? 'barStacked' : model.type) : presetType || 'bar';
  const body = `
    <div class="chart-types">${CHART_TYPES.map(([k, l]) => `<button class="ct${k === curType ? ' on' : ''}" data-t="${k}">${chartIcon(k)}<span>${t(l)}</span></button>`).join('')}</div>
    <div class="grid2">
      <label class="fld">${t('Data range')}<input class="inp" data-k="range" value="${esc(range)}" placeholder="${t('e.g. A1:D7')}"></label>
      <label class="fld">${t('Series in')}<select class="inp" data-k="dir"><option value="auto">${t('Auto')}</option><option value="cols">${t('Columns')}</option><option value="rows">${t('Rows')}</option></select></label>
      <label class="fld">${t('Chart title')}<input class="inp" data-k="title" value="${esc(model?.title || '')}"></label>
      <label class="fld">${t('Legend')}<select class="inp" data-k="legend"><option value="r">${t('Right')}</option><option value="b">${t('Bottom')}</option><option value="t">${t('Top')}</option><option value="l">${t('Left')}</option><option value="">${t('None')}</option></select></label>
      <label class="fld">${t('Horizontal (category) axis title')}<input class="inp" data-k="xt" value="${esc(model?.xTitle || '')}"></label>
      <label class="fld">${t('Vertical (value) axis title')}<input class="inp" data-k="yt" value="${esc(model?.yTitle || '')}"></label>
    </div>
    ${model?.pivot ? `<p class="muted small">${t("Pivot chart of '{name}': its data comes from the pivot table, so the data range cannot be changed here.", { name: esc(model.pivot.table) })}</p>`
    : model && !isNew ? `<p class="muted small">${t('If you leave the data range unchanged, the original series definitions (formula references) are kept.')}</p>` : ''}`;
  let type = curType;
  return modal({
    title: isNew ? t('Insert Chart') : t('Edit Chart'), body, width: 600, className: 'chart-dialog',
    onOpen: (el) => {
      el.querySelector('[data-k=legend]').value = model ? (model.legend || '') : 'r';
      if (model?.pivot) {  // Excel: a pivot chart takes its data from the pivot table and cannot be XY scatter
        el.querySelector('[data-k=range]').disabled = true; el.querySelector('[data-k=dir]').disabled = true;
        const sc = el.querySelector('[data-t=scatter]'); if (sc) sc.disabled = true;
      }
      el.querySelector('.chart-types').addEventListener('click', (e) => { const b = e.target.closest('[data-t]'); if (!b) return; type = b.dataset.t; el.querySelectorAll('.ct').forEach((x) => x.classList.toggle('on', x === b)); });
    },
    onOk: (el) => {
      const q = (k) => el.querySelector(`[data-k="${k}"]`).value;
      const preset = Object.fromEntries(CHART_TYPES.map(([k, , o]) => [k, o || {}]))[type];
      const baseType = preset.type || type;
      let m;
      const rangeText = q('range').trim();
      if (isNew || rangeText !== range) {
        const p = parseSheetRef(rangeText);
        if (!p) { app.toast(t('The data range is not valid.'), 'error'); return false; }
        const dsh = p.sheet ? wb.sheetByName(p.sheet) : sh;
        if (!dsh) { app.toast(t('Sheet not found.'), 'error'); return false; }
        const dir = q('dir');
        m = ChartLayer.fromRange(wb, dsh, p.range, baseType, dir === 'auto' ? null : dir === 'rows');
        if (!isNew) { m.id = model.id; m.anchor = model.anchor; }
      } else m = { ...structuredClone(model), type: baseType };
      m.type = baseType;
      m.barDir = preset.barDir || 'col';
      m.grouping = preset.grouping || (baseType === 'line' || baseType === 'area' ? 'standard' : 'clustered');
      m.title = q('title') || null; m.legend = q('legend') || null; m.xTitle = q('xt') || null; m.yTitle = q('yt') || null;
      m.dirty = true;
      if (m.series.some((s) => s.type)) m.series.forEach((s) => { if (s.type === baseType) delete s.type; });
      app.saveChart(m, isNew);
    },
  }).done;
}

function chartIcon(k) {
  const svg = (inner) => `<svg viewBox="0 0 32 24" width="40" height="30" aria-hidden="true">${inner}</svg>`;
  switch (k) {
    case 'bar': return svg('<rect x="4" y="10" width="5" height="12" fill="#4472C4"/><rect x="12" y="4" width="5" height="18" fill="#ED7D31"/><rect x="20" y="13" width="5" height="9" fill="#A5A5A5"/>');
    case 'barH': return svg('<rect x="3" y="3" width="18" height="5" fill="#4472C4"/><rect x="3" y="10" width="24" height="5" fill="#ED7D31"/><rect x="3" y="17" width="12" height="5" fill="#A5A5A5"/>');
    case 'barStacked': return svg('<rect x="5" y="12" width="6" height="10" fill="#4472C4"/><rect x="5" y="6" width="6" height="6" fill="#ED7D31"/><rect x="17" y="9" width="6" height="13" fill="#4472C4"/><rect x="17" y="3" width="6" height="6" fill="#ED7D31"/>');
    case 'line': return svg('<polyline points="3,18 10,10 17,14 28,4" fill="none" stroke="#4472C4" stroke-width="2"/><polyline points="3,21 10,17 17,19 28,12" fill="none" stroke="#ED7D31" stroke-width="2"/>');
    case 'area': return svg('<polygon points="3,22 3,14 11,8 19,12 29,5 29,22" fill="#4472C4"/>');
    case 'pie': return svg('<circle cx="16" cy="12" r="10" fill="#4472C4"/><path d="M16 12 L16 2 A10 10 0 0 1 25 16 Z" fill="#ED7D31"/>');
    case 'doughnut': return svg('<circle cx="16" cy="12" r="10" fill="none" stroke="#4472C4" stroke-width="6"/><path d="M16 2 A10 10 0 0 1 26 12" fill="none" stroke="#ED7D31" stroke-width="6"/>');
    case 'scatter': return svg('<circle cx="7" cy="17" r="2" fill="#4472C4"/><circle cx="12" cy="11" r="2" fill="#4472C4"/><circle cx="19" cy="14" r="2" fill="#4472C4"/><circle cx="25" cy="6" r="2" fill="#4472C4"/>');
    case 'radar': return svg('<polygon points="16,2 27,10 23,22 9,22 5,10" fill="none" stroke="#999"/><polygon points="16,6 23,11 20,18 12,19 9,11" fill="#4472C455" stroke="#4472C4"/>');
  }
  return '';
}

// ------------------------------------------------------------------ sort dialog
export function sortDialog(app) {
  const sh = app.doc.wb.sheet, v = app.view();
  let range = v.sel.r1 === v.sel.r2 && v.sel.c1 === v.sel.c2 ? app.filters.tableAt(sh, v.sel.r1, v.sel.c1) : { ...v.sel, r2: Math.min(v.sel.r2, sh.maxRow), c2: Math.min(v.sel.c2, sh.maxCol) };
  if (sh.autoFilter && v.active.r >= sh.autoFilter.range.r1 && v.active.r <= sh.autoFilter.range.r2) range = app.filters.tableBody(sh, sh.autoFilter.range);
  const header = app.filters.guessHeader(sh, range);
  const colOpts = (hdr) => { const o = []; for (let c = range.c1; c <= range.c2; c++) { const h = hdr ? formatValue(app.doc.wb.value(sh, range.r1, c), null).text : ''; o.push(`<option value="${c}">${esc(h || t('Column {name}', { name: colName(c) }))}</option>`); } return o.join(''); };
  const levelRow = (i) => `<div class="sort-level"><span>${i === 0 ? t('Sort by') : t('Then by')}</span><select class="inp" data-col="${i}">${i ? `<option value="">${t('(none)')}</option>` : ''}${colOpts(header)}</select><select class="inp" data-ord="${i}"><option value="asc">${t('Ascending')}</option><option value="desc">${t('Descending')}</option></select></div>`;
  return modal({
    title: t('Sort'), width: 480,
    body: `<p class="muted small">${t('Range: {range}', { range: rangeA1(range) })}</p><label class="chk"><input type="checkbox" data-k="hdr" ${header ? 'checked' : ''}> ${t('My data has headers')}</label>${[0, 1, 2].map(levelRow).join('')}`,
    onOk: (el) => {
      const keys = [];
      for (let i = 0; i < 3; i++) { const c = el.querySelector(`[data-col="${i}"]`).value; if (c !== '') keys.push({ c: +c, asc: el.querySelector(`[data-ord="${i}"]`).value === 'asc' }); }
      app.filters.sort(sh, range, keys, el.querySelector('[data-k=hdr]').checked);
    },
  }).done;
}

// ------------------------------------------------------------------ paste special
export function pasteSpecialDialog(app) {
  return modal({
    title: t('Paste Special'), width: 360,
    body: `<div class="radio-col"><label><input type="radio" name="ps" value="all" checked> ${t('All')}</label><label><input type="radio" name="ps" value="formulas"> ${t('Formulas')}</label><label><input type="radio" name="ps" value="values"> ${t('Values')}</label><label><input type="radio" name="ps" value="formats"> ${t('Formats')}</label></div><label class="chk"><input type="checkbox" data-k="tr"> ${t('Transpose')}</label>`,
    onOk: (el) => { app.clipboard.pasteCommand(el.querySelector('input[name=ps]:checked').value, el.querySelector('[data-k=tr]').checked); },
  }).done;
}

// ------------------------------------------------------------------ insert function
export function insertFunctionDialog(app) {
  const cats = ['', ...new Set(Object.keys(FN_HELP).map((n) => fnHelp(n)[0]))]; // '' = all categories
  const body = `<label class="fld">${t('Search for a function')}<input class="inp" data-k="q" placeholder="${t('e.g. SUM, VLOOKUP')}" autofocus></label>
    <label class="fld">${t('Category')}<select class="inp" data-k="cat">${cats.map((c) => `<option value="${esc(c)}">${esc(c ? t(c) : t('All functions'))}</option>`).join('')}</select></label>
    <select class="inp fn-list" size="10" data-k="list"></select><div class="fn-desc" data-k="desc"></div>`;
  return modal({
    title: t('Insert Function'), body, width: 480,
    onOpen: (el) => {
      const q = (k) => el.querySelector(`[data-k="${k}"]`);
      const fill = () => {
        const s = q('q').value.toUpperCase(), cat = q('cat').value;
        const names = FUNCTION_NAMES.filter((n) => (cat === '' || fnHelp(n)?.[0] === cat) && (!s || n.includes(s) || (fnHelp(n)?.[2] || '').includes(q('q').value)));
        q('list').innerHTML = names.map((n) => `<option>${n}</option>`).join('');
        q('list').selectedIndex = 0; desc();
      };
      const desc = () => { const n = q('list').value; const h = fnHelp(n); q('desc').innerHTML = n ? `<b>${esc(h?.[1] || n + '(...)')}</b><br>${esc(h?.[2] || '')}` : ''; };
      q('q').addEventListener('input', fill); q('cat').addEventListener('change', fill); q('list').addEventListener('change', desc);
      q('list').addEventListener('dblclick', () => el.querySelector('[data-act=ok]').click());
      fill();
    },
    onOk: (el) => { const n = el.querySelector('[data-k=list]').value; if (n) app.insertFunction(n); },
  }).done;
}

// ------------------------------------------------------------------ pivot table dialog
const PIVOT_FNS = [['sum', 'Sum'], ['count', 'Count'], ['average', 'Average'], ['max', 'Max'], ['min', 'Min'], ['product', 'Product'], ['countNums', 'Count Numbers'],
  ['stdDev', 'StdDev'], ['stdDevp', 'StdDevp'], ['var', 'Var'], ['varp', 'Varp']];

/** Insert → PivotTable: source range, a role per field (rows / columns / values / filter), where to put it.
 * The server builds it (POST /pivots, same as `xl.py pivot add`), so the agent and the preview make the
 * same pivot; saving writes a real pivot table that Excel rebuilds when it opens the file. */
export function pivotDialog(app) {
  const wb = app.doc.wb, sh = wb.sheet, v = app.view();
  let g = v.sel;
  if (g.r1 === g.r2 && g.c1 === g.c2) g = app.filters.tableAt(sh, g.r1, g.c1);
  g = { ...g, r2: Math.min(g.r2, sh.maxRow), c2: Math.min(g.c2, sh.maxCol) };
  const source0 = `${quoteSheetName(sh.name)}!${rangeA1(g)}`;
  const n = wb.sheets.length + 1;
  const body = `
    <label class="fld">${t('Source range')}<input class="inp" data-k="source" value="${esc(source0)}"></label>
    <div class="pv-fields" data-k="fields"></div>
    <div class="grid2">
      <label class="chk"><input type="radio" name="pv-dest" value="new" checked> ${t('New worksheet')} <input class="inp" data-k="sheet" value="${esc(t('Pivot{n}', { n }))}" style="width:110px"></label>
      <label class="chk"><input type="radio" name="pv-dest" value="here"> ${t('Existing location')} <input class="inp" data-k="at" value="${esc(sh.name)}!${a1(v.active.r, v.active.c)}" style="width:140px"></label>
    </div>
    <p class="muted small">${t('Put each field in Rows, Columns, Values or Filters. When you save, Excel rebuilds the pivot from the source when it opens the file.')}</p>
    <p class="err small" data-k="err" style="color:#c00000"></p>`;
  const headers = (src) => {
    const p = parseSheetRef(src);
    const s = p && (p.sheet ? wb.sheetByName(p.sheet) : sh);
    if (!s) return [];
    const out = [];
    for (let c = p.range.c1; c <= Math.min(p.range.c2, p.range.c1 + 60); c++) {
      const h = wb.value(s, p.range.r1, c);
      if (h !== null && h !== undefined && h !== '') out.push(String(h));
    }
    return out;
  };
  const fieldRows = (names) => names.map((f, i) => `<div class="pv-row" data-f="${esc(f)}"><span class="pv-name">${esc(f)}</span>
      <select class="inp" data-role><option value="">—</option><option value="rows"${i === 0 ? ' selected' : ''}>${t('Rows')}</option><option value="cols">${t('Columns')}</option><option value="values"${i === names.length - 1 ? ' selected' : ''}>${t('Values')}</option><option value="filters">${t('Filters')}</option></select>
      <select class="inp" data-fn>${PIVOT_FNS.map(([k, l]) => `<option value="${k}">${t(l)}</option>`).join('')}</select></div>`).join('');
  return modal({
    title: t('Create PivotTable'), body, width: 520, className: 'pivot-dialog',
    onOpen: (el) => {
      const box = el.querySelector('[data-k=fields]');
      const fill = () => { box.innerHTML = fieldRows(headers(el.querySelector('[data-k=source]').value)); };
      fill();
      el.querySelector('[data-k=source]').addEventListener('change', fill);
    },
    onOk: async (el) => {
      const q = (k) => el.querySelector(`[data-k="${k}"]`);
      const spec = { source: q('source').value.trim(), rows: [], cols: [], values: [], filters: [] };
      for (const row of el.querySelectorAll('.pv-row')) {
        const role = row.querySelector('[data-role]').value, f = row.dataset.f;
        if (!role) continue;
        if (role === 'values') spec.values.push({ field: f, fn: row.querySelector('[data-fn]').value });
        else spec[role].push(f);
      }
      const dest = el.querySelector('input[name=pv-dest]:checked').value;
      spec.at = dest === 'new' ? `${q('sheet').value.trim() || t('Pivot{n}', { n })}!A3` : q('at').value.trim();
      try {
        const res = await app.createPivot(spec);
        app.toast(t("Created pivot table '{name}' at {sheet}!{range}. Excel refreshes it when it opens the saved file.", { name: res.name, sheet: res.sheet, range: res.range }), 'info');
        return true;
      } catch (e) { q('err').textContent = e.message; return false; }
    },
  }).done;
}

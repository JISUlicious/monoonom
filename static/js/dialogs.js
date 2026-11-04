// Modal dialogs: file open/save (server-side browsing), format cells, find/replace, chart, sort, etc.
import { api } from './api.js';
import { FORMAT_PRESETS, formatValue } from './numfmt.js';
import { CHART_TYPES, ChartLayer } from './charts.js';
import { parseSheetRef, parseRange, rangeA1, colName, key } from './address.js';
import { FN_HELP } from './fnhelp.js';
import { FUNCTION_NAMES } from './engine.js';

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
export function modal({ title, body, buttons = [['확인', 'ok', true], ['취소', 'cancel']], width = 420, onOk, onOpen, className = '' }) {
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.innerHTML = `<div class="modal ${className}" role="dialog" aria-modal="true" style="width:${width}px">
    <div class="modal-title"><span>${esc(title)}</span><button class="x" data-act="cancel" aria-label="닫기">×</button></div>
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

export function alertBox(msg, title = '알림') { return modal({ title, body: `<p class="msg">${esc(msg).replace(/\n/g, '<br>')}</p>`, buttons: [['확인', 'ok', true]] }).done; }
export function confirmBox(msg, { title = '확인', yes = '예', no = '아니요', cancel = false } = {}) {
  const buttons = [[yes, 'ok', true], [no, 'no']];
  if (cancel) buttons.push(['취소', 'cancel']);
  return modal({ title, body: `<p class="msg">${esc(msg).replace(/\n/g, '<br>')}</p>`, buttons }).done.then((v) => (v === true ? 'yes' : v === 'no' ? 'no' : null));
}
export function promptBox(msg, value = '', title = '입력') {
  return modal({ title, body: `<label class="fld">${esc(msg)}<input class="inp" id="pb" value="${esc(value)}" autofocus></label>`,
    onOk: (el) => el.querySelector('#pb').value, onOpen: (el) => setTimeout(() => el.querySelector('#pb').select(), 0) }).done;
}

// ------------------------------------------------------------------ file dialog (server-side)
const RECENT_KEY = 'webexcel.recent';
export function recentFiles() { try { return JSON.parse(localStorage.getItem(RECENT_KEY) || '[]'); } catch { return []; } }
export function addRecent(path) {
  try { const l = [path, ...recentFiles().filter((p) => p !== path)].slice(0, 12); localStorage.setItem(RECENT_KEY, JSON.stringify(l)); } catch { /* ignore */ }
}
const LASTDIR_KEY = 'webexcel.lastdir';

export async function fileDialog({ mode = 'open', name = '', format = 'xlsx', startDir = null } = {}) {
  const info = await api.info().catch(() => ({ start: '', sep: '/' }));
  let dir = startDir;
  if (!dir) { try { dir = localStorage.getItem(LASTDIR_KEY); } catch { /* ignore */ } }
  dir = dir || info.start;
  const sep = info.sep || '/';
  const saving = mode === 'save';
  const body = `
    <div class="fd-path"><button class="btn sm" data-fd="up" title="상위 폴더">↑</button>
      ${info.drives?.length ? `<select class="inp sm" data-fd="drive">${info.drives.map((d) => `<option>${esc(d)}</option>`).join('')}</select>` : ''}
      <input class="inp" data-fd="dir" value="${esc(dir)}" spellcheck="false"><button class="btn sm" data-fd="go">이동</button></div>
    <div class="fd-main"><div class="fd-side"><div class="fd-side-title">최근 파일</div>${recentFiles().map((p) => `<div class="fd-recent" data-path="${esc(p)}" title="${esc(p)}">${esc(p.split(/[\\/]/).pop())}</div>`).join('') || '<div class="muted">없음</div>'}</div>
      <div class="fd-list" tabindex="0"></div></div>
    ${saving ? `<div class="fd-save"><label>파일 이름 <input class="inp" data-fd="name" value="${esc(name)}"></label>
      <label>형식 <select class="inp" data-fd="fmt"><option value="xlsx">Excel 통합 문서 (*.xlsx)</option><option value="xlsm">Excel 매크로 사용 통합 문서 (*.xlsm)</option><option value="xls">Excel 97-2003 통합 문서 (*.xls)</option></select></label></div>`
      : `<div class="fd-save"><label>파일 경로 <input class="inp" data-fd="name" placeholder="파일을 선택하거나 전체 경로를 입력"></label></div>`}
    <p class="muted small">파일 읽기/쓰기는 로컬 Python 서버가 수행합니다 (브라우저는 파일 내용에 직접 접근하지 않음).</p>`;
  let current = dir;
  const dlg = modal({
    title: saving ? '다른 이름으로 저장' : '열기', body, width: 720, className: 'file-dialog',
    buttons: [[saving ? '저장' : '열기', 'ok', true], ['취소', 'cancel']],
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
  if (saving) el.querySelector('[data-fd=fmt]').value = format === 'xls' ? 'xls' : format === 'xlsm' ? 'xlsm' : 'xlsx';
  const list = el.querySelector('.fd-list');
  const load = async (path) => {
    list.innerHTML = '<div class="muted pad">불러오는 중…</div>';
    try {
      const res = await api.fs(path);
      current = res.path;
      el.querySelector('[data-fd=dir]').value = res.path;
      el.dataset.parent = res.parent || '';
      list.innerHTML = res.entries.map((e) => `<div class="fd-item" data-path="${esc(e.path)}" ${e.dir ? 'data-dir="1"' : ''}>
        <span class="ic">${e.dir ? '📁' : /\.xls$/i.test(e.name) ? '📗' : '📊'}</span><span class="nm">${esc(e.name)}</span>
        <span class="sz">${e.dir ? '' : fmtSize(e.size)}</span><span class="dt">${new Date(e.mtime * 1000).toLocaleString('ko-KR', { dateStyle: 'short', timeStyle: 'short' })}</span></div>`).join('') || '<div class="muted pad">Excel 파일이 없습니다.</div>';
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
export function colorPopup(anchor, { allowNone = true, noneLabel = '채우기 없음', onPick }) {
  document.querySelector('.color-popup')?.remove();
  const p = document.createElement('div');
  p.className = 'color-popup';
  p.innerHTML = `${allowNone ? `<button class="cp-none" data-c="">${noneLabel}</button>` : ''}
    <div class="cp-label">테마 색</div><div class="cp-grid">${THEME_COLORS.map((row) => row.map((c) => `<button style="background:${c}" data-c="${c}" title="${c}"></button>`).join('')).join('')}</div>
    <div class="cp-label">표준 색</div><div class="cp-grid">${STANDARD_COLORS.map((c) => `<button style="background:${c}" data-c="${c}" title="${c}"></button>`).join('')}</div>
    <label class="cp-more">다른 색… <input type="color"></label>`;
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
  const fonts = ['맑은 고딕', 'Calibri', 'Arial', '굴림', '돋움', '바탕', '궁서', 'Times New Roman', 'Consolas', 'Segoe UI', 'Noto Sans KR'];
  const borderStyles = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair'];
  const body = `
  <div class="tabs"><button data-tab="num">표시 형식</button><button data-tab="align">맞춤</button><button data-tab="font">글꼴</button><button data-tab="border">테두리</button><button data-tab="fill">채우기</button></div>
  <section data-pane="num">
    <div class="row2"><select class="inp" size="12" data-k="preset">${FORMAT_PRESETS.map((p) => `<option value="${esc(p.code)}">${p.label}</option>`).join('')}<option value="__custom">사용자 지정</option></select>
    <div class="col"><label class="fld">형식 코드<input class="inp" data-k="nf" value="${esc(st.nf || 'General')}"></label>
    <div class="fld">보기: <span class="preview" data-k="nfprev"></span></div>
    <p class="muted small">예: #,##0.00 / 0% / yyyy-mm-dd / "₩"#,##0;[빨강]-"₩"#,##0 / [h]:mm</p></div></div>
  </section>
  <section data-pane="align">
    <div class="grid2">
      <label class="fld">가로<select class="inp" data-k="ha"><option value="">일반</option><option value="left">왼쪽</option><option value="center">가운데</option><option value="right">오른쪽</option><option value="fill">채우기</option><option value="justify">양쪽 맞춤</option><option value="centerContinuous">선택 영역의 가운데로</option></select></label>
      <label class="fld">세로<select class="inp" data-k="va"><option value="top">위쪽</option><option value="center">가운데</option><option value="">아래쪽</option><option value="justify">양쪽 맞춤</option></select></label>
      <label class="fld">들여쓰기<input class="inp" type="number" min="0" max="15" data-k="ind" value="${st.ind || 0}"></label>
      <label class="chk"><input type="checkbox" data-k="wrap" ${st.wrap ? 'checked' : ''}> 텍스트 줄 바꿈</label>
      <label class="chk"><input type="checkbox" data-k="merge"> 셀 병합</label>
    </div>
  </section>
  <section data-pane="font">
    <div class="grid2">
      <label class="fld">글꼴<input class="inp" list="fontlist" data-k="ff" value="${esc(st.ff || app.doc.wb.defaultFont.name)}"><datalist id="fontlist">${fonts.map((f) => `<option>${f}</option>`).join('')}</datalist></label>
      <label class="fld">크기<input class="inp" type="number" min="1" max="409" step="0.5" data-k="fs" value="${st.fs || app.doc.wb.defaultFont.size}"></label>
      <label class="chk"><input type="checkbox" data-k="b" ${st.b ? 'checked' : ''}> 굵게</label>
      <label class="chk"><input type="checkbox" data-k="i" ${st.i ? 'checked' : ''}> 기울임꼴</label>
      <label class="fld">밑줄<select class="inp" data-k="u"><option value="">없음</option><option value="1">실선</option><option value="double">이중 실선</option></select></label>
      <label class="chk"><input type="checkbox" data-k="st" ${st.st ? 'checked' : ''}> 취소선</label>
      <label class="fld">색<input class="inp" type="color" data-k="fc" value="${st.fc || '#000000'}"></label>
    </div>
    <div class="font-preview" data-k="fprev">가나다 AaBbCc 123</div>
  </section>
  <section data-pane="border">
    <div class="grid2"><label class="fld">선 스타일<select class="inp" data-k="bs">${borderStyles.map((s) => `<option>${s}</option>`).join('')}</select></label>
    <label class="fld">색<input class="inp" type="color" data-k="bc" value="#000000"></label></div>
    <div class="border-btns"><button class="btn" data-b="none">없음</button><button class="btn" data-b="outline">윤곽선</button><button class="btn" data-b="inside">안쪽</button><button class="btn" data-b="all">모든 테두리</button>
    <button class="btn" data-b="top">위</button><button class="btn" data-b="bottom">아래</button><button class="btn" data-b="left">왼쪽</button><button class="btn" data-b="right">오른쪽</button></div>
    <p class="muted small">선택한 테두리 동작: <b data-k="bact">변경 없음</b></p>
  </section>
  <section data-pane="fill">
    <div class="cp-grid big">${THEME_COLORS.flat().concat(STANDARD_COLORS).map((c) => `<button style="background:${c}" data-fill="${c}"></button>`).join('')}</div>
    <div class="row2"><button class="btn" data-fill="">채우기 없음</button><label>다른 색 <input type="color" data-k="bgc" value="${st.bg || '#ffffff'}"></label> <span class="swatch" data-k="bgprev" style="background:${st.bg || '#fff'}"></span></div>
  </section>`;
  let fillChoice = undefined, borderAct = null;
  const dlg = modal({
    title: '셀 서식', body, width: 560, className: 'fmt-dialog',
    onOpen: (el) => {
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
  p.innerHTML = `<div class="fp-head"><b>${replace ? '찾기 및 바꾸기' : '찾기'}</b><button class="x" data-a="close">×</button></div>
    <label class="fld">찾을 내용<input class="inp" data-k="q" autofocus></label>
    ${replace ? '<label class="fld">바꿀 내용<input class="inp" data-k="r"></label>' : ''}
    <div class="fp-opts"><label><input type="checkbox" data-k="case"> 대/소문자 구분</label><label><input type="checkbox" data-k="whole"> 전체 셀 내용 일치</label>
    <label>범위 <select data-k="scope"><option value="sheet">시트</option><option value="wb">통합 문서</option></select></label>
    <label>찾는 위치 <select data-k="look"><option value="values">값</option><option value="formulas">수식</option></select></label></div>
    <div class="fp-btns"><button class="btn" data-a="all">모두 찾기</button><button class="btn primary" data-a="next">다음 찾기</button>${replace ? '<button class="btn" data-a="rep">바꾸기</button><button class="btn" data-a="repall">모두 바꾸기</button>' : ''}</div>
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
    if (a === 'all') { box.innerHTML = `<div class="muted small">${list.length}개 셀을 찾았습니다.</div>` + list.slice(0, 500).map((m, i) => `<div class="fp-res" data-i="${i}">${esc(m.sh.name)}!${colName(m.c)}${m.r + 1} — ${esc(m.t.slice(0, 60))}</div>`).join(''); return; }
    if (!list.length) { box.innerHTML = '<div class="muted small">찾는 항목이 없습니다.</div>'; return; }
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
      wb.begin('바꾸기');
      let n = 0;
      for (const m of targets) {
        const cell = m.sh.get(m.r, m.c);
        if (cell.f && q('look').value === 'formulas') { wb.setCell(m.sh, m.r, m.c, { ...cell, f: cell.f.replace(re, rep) }); n++; }
        else if (!cell.f && typeof cell.v === 'string') { wb.setCell(m.sh, m.r, m.c, { ...cell, v: cell.v.replace(re, rep) }); n++; }
        else if (!cell.f) { app.setCellInput(m.sh, m.r, m.c, String(m.t).replace(re, rep), { inTxn: true }); n++; }
      }
      wb.commit();
      box.innerHTML = `<div class="muted small">${n}개 항목을 바꿨습니다.</div>`;
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
      if (g.r1 === g.r2 && g.c1 === g.c2) g = app.filters.currentRegion(sh, g.r1, g.c1);
      range = rangeA1({ ...g, r2: Math.min(g.r2, sh.maxRow), c2: Math.min(g.c2, sh.maxCol) });
    }
  }
  const curType = model ? (model.type === 'bar' && model.barDir === 'bar' ? 'barH' : model.type === 'bar' && model.grouping === 'stacked' ? 'barStacked' : model.type) : presetType || 'bar';
  const body = `
    <div class="chart-types">${CHART_TYPES.map(([k, l]) => `<button class="ct${k === curType ? ' on' : ''}" data-t="${k}">${chartIcon(k)}<span>${l}</span></button>`).join('')}</div>
    <div class="grid2">
      <label class="fld">데이터 범위<input class="inp" data-k="range" value="${esc(range)}" placeholder="예: A1:D7"></label>
      <label class="fld">계열 방향<select class="inp" data-k="dir"><option value="auto">자동</option><option value="cols">열</option><option value="rows">행</option></select></label>
      <label class="fld">차트 제목<input class="inp" data-k="title" value="${esc(model?.title || '')}"></label>
      <label class="fld">범례<select class="inp" data-k="legend"><option value="r">오른쪽</option><option value="b">아래쪽</option><option value="t">위쪽</option><option value="l">왼쪽</option><option value="">없음</option></select></label>
      <label class="fld">가로(항목) 축 제목<input class="inp" data-k="xt" value="${esc(model?.xTitle || '')}"></label>
      <label class="fld">세로(값) 축 제목<input class="inp" data-k="yt" value="${esc(model?.yTitle || '')}"></label>
    </div>
    ${model && !isNew ? '<p class="muted small">데이터 범위를 바꾸지 않으면 원본 계열 정의(수식 참조)를 그대로 유지합니다.</p>' : ''}`;
  let type = curType;
  return modal({
    title: isNew ? '차트 삽입' : '차트 편집', body, width: 600, className: 'chart-dialog',
    onOpen: (el) => {
      el.querySelector('[data-k=legend]').value = model ? (model.legend || '') : 'r';
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
        if (!p) { app.toast('데이터 범위가 올바르지 않습니다.', 'error'); return false; }
        const dsh = p.sheet ? wb.sheetByName(p.sheet) : sh;
        if (!dsh) { app.toast('시트를 찾을 수 없습니다.', 'error'); return false; }
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
  let range = v.sel.r1 === v.sel.r2 && v.sel.c1 === v.sel.c2 ? app.filters.currentRegion(sh, v.sel.r1, v.sel.c1) : { ...v.sel, r2: Math.min(v.sel.r2, sh.maxRow), c2: Math.min(v.sel.c2, sh.maxCol) };
  if (sh.autoFilter && v.active.r >= sh.autoFilter.range.r1 && v.active.r <= sh.autoFilter.range.r2) range = sh.autoFilter.range;
  const header = app.filters.guessHeader(sh, range);
  const colOpts = (hdr) => { const o = []; for (let c = range.c1; c <= range.c2; c++) { const t = hdr ? formatValue(app.doc.wb.value(sh, range.r1, c), null).text : ''; o.push(`<option value="${c}">${esc(t || '열 ' + colName(c))}</option>`); } return o.join(''); };
  const levelRow = (i) => `<div class="sort-level"><span>${i === 0 ? '정렬 기준' : '다음 기준'}</span><select class="inp" data-col="${i}">${i ? '<option value="">(없음)</option>' : ''}${colOpts(header)}</select><select class="inp" data-ord="${i}"><option value="asc">오름차순</option><option value="desc">내림차순</option></select></div>`;
  return modal({
    title: '정렬', width: 480,
    body: `<p class="muted small">범위: ${rangeA1(range)}</p><label class="chk"><input type="checkbox" data-k="hdr" ${header ? 'checked' : ''}> 내 데이터에 머리글 표시</label>${[0, 1, 2].map(levelRow).join('')}`,
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
    title: '선택하여 붙여넣기', width: 360,
    body: `<div class="radio-col"><label><input type="radio" name="ps" value="all" checked> 모두</label><label><input type="radio" name="ps" value="formulas"> 수식</label><label><input type="radio" name="ps" value="values"> 값</label><label><input type="radio" name="ps" value="formats"> 서식</label></div><label class="chk"><input type="checkbox" data-k="tr"> 행/열 바꿈</label>`,
    onOk: (el) => { app.clipboard.pasteCommand(el.querySelector('input[name=ps]:checked').value, el.querySelector('[data-k=tr]').checked); },
  }).done;
}

// ------------------------------------------------------------------ insert function
export function insertFunctionDialog(app) {
  const cats = ['전체', ...new Set(Object.values(FN_HELP).map((h) => h[0]))];
  const body = `<label class="fld">함수 검색<input class="inp" data-k="q" placeholder="예: 합계, VLOOKUP" autofocus></label>
    <label class="fld">범주<select class="inp" data-k="cat">${cats.map((c) => `<option>${c}</option>`).join('')}</select></label>
    <select class="inp fn-list" size="10" data-k="list"></select><div class="fn-desc" data-k="desc"></div>`;
  return modal({
    title: '함수 삽입', body, width: 480,
    onOpen: (el) => {
      const q = (k) => el.querySelector(`[data-k="${k}"]`);
      const fill = () => {
        const s = q('q').value.toUpperCase(), cat = q('cat').value;
        const names = FUNCTION_NAMES.filter((n) => (cat === '전체' || FN_HELP[n]?.[0] === cat) && (!s || n.includes(s) || (FN_HELP[n]?.[2] || '').includes(q('q').value)));
        q('list').innerHTML = names.map((n) => `<option>${n}</option>`).join('');
        q('list').selectedIndex = 0; desc();
      };
      const desc = () => { const n = q('list').value; const h = FN_HELP[n]; q('desc').innerHTML = n ? `<b>${esc(h?.[1] || n + '(...)')}</b><br>${esc(h?.[2] || '')}` : ''; };
      q('q').addEventListener('input', fill); q('cat').addEventListener('change', fill); q('list').addEventListener('change', desc);
      q('list').addEventListener('dblclick', () => el.querySelector('[data-act=ok]').click());
      fill();
    },
    onOk: (el) => { const n = el.querySelector('[data-k=list]').value; if (n) app.insertFunction(n); },
  }).done;
}

// A text box with a real dropdown list (font name, font size). <datalist> was used before, but a
// browser only lists the suggestions that match what is already in the box, so with "Calibri" or
// "11" in it the list looked empty. Here the button / ArrowDown / Alt+ArrowDown open the whole
// list with the current value marked; typing narrows it; Enter / click picks; Esc closes.
import { t } from './i18n.js';

/**
 * @param {HTMLInputElement} input
 * @param {{ items: () => Promise<Array<{value: string, label?: string, group?: string, font?: string}>> | Array,
 *           onPick: (value: string) => void, width?: number }} opts
 */
export function combo(input, { items, onPick, width = null }) {
  const wrap = document.createElement('span');
  wrap.className = 'combo';
  input.parentNode.insertBefore(wrap, input);
  wrap.append(input);
  input.removeAttribute('list');
  input.setAttribute('autocomplete', 'off');
  input.setAttribute('role', 'combobox');
  input.setAttribute('aria-expanded', 'false');
  const btn = document.createElement('button');
  btn.type = 'button'; btn.className = 'combo-btn'; btn.tabIndex = -1; btn.setAttribute('aria-label', t('Open list'));
  btn.textContent = '▾';
  wrap.append(btn);
  let pop = null, list = [], shown = [], active = -1, filterText = null;

  const close = () => {
    pop?.remove(); pop = null; active = -1; filterText = null;
    input.setAttribute('aria-expanded', 'false');
    document.removeEventListener('pointerdown', outside, true);
  };
  const outside = (e) => { if (!wrap.contains(e.target) && !pop?.contains(e.target)) close(); };
  const esc = (t) => String(t ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

  function render() {
    const q = (filterText || '').trim().toLowerCase();
    if (q) { // one entry per value when narrowing (a font can be in several groups)
      const seen = new Set();
      shown = list.filter((it) => ((it.label || it.value).toLowerCase().includes(q) || it.value.toLowerCase().includes(q)) && !seen.has(it.value.toLowerCase()) && seen.add(it.value.toLowerCase()));
    } else shown = list;
    const cur = input.value.trim().toLowerCase();
    let html = '', last = null;
    shown.forEach((it, i) => {
      if (!q && it.group && it.group !== last) { html += `<div class="combo-group">${esc(it.group)}</div>`; last = it.group; }
      const style = it.font ? ` style="font-family:&quot;${esc(it.font)}&quot;, sans-serif"` : '';
      html += `<div class="combo-item${i === active ? ' active' : ''}${it.value.toLowerCase() === cur ? ' current' : ''}" data-i="${i}" role="option"${style}>${esc(it.label || it.value)}</div>`;
    });
    pop.innerHTML = html || `<div class="combo-empty">${t('No matches')}</div>`;
    const el = pop.querySelector('.combo-item.active') || (!q && pop.querySelector('.combo-item.current'));
    if (el) el.scrollIntoView({ block: active >= 0 ? 'nearest' : 'center' });
  }

  async function open(filter = null) {
    if (!pop) {
      pop = document.createElement('div');
      pop.className = 'combo-pop'; pop.setAttribute('role', 'listbox');
      document.body.append(pop);
      const r = wrap.getBoundingClientRect();
      Object.assign(pop.style, { left: `${r.left}px`, top: `${r.bottom + 2}px`, minWidth: `${width || r.width}px` });
      pop.addEventListener('pointerdown', (e) => e.preventDefault()); // keep the focus in the box
      pop.addEventListener('click', (e) => { const it = e.target.closest('.combo-item'); if (it) pick(shown[+it.dataset.i]); });
      input.setAttribute('aria-expanded', 'true');
      document.addEventListener('pointerdown', outside, true);
      pop.innerHTML = `<div class="combo-empty">${t('Loading…')}</div>`;
      list = await items();
      if (!pop) return;
    }
    filterText = filter;
    active = -1;
    render();
  }

  function pick(it) {
    if (!it) return;
    input.value = it.value;
    close();
    onPick(it.value);
  }

  btn.addEventListener('pointerdown', (e) => e.preventDefault());
  btn.addEventListener('click', () => { if (pop) close(); else { input.focus(); open(null); } });
  input.addEventListener('input', () => open(input.value));
  input.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      if (!pop) { open(null); return; }
      const n = shown.length;
      if (!n) return;
      const down = e.key === 'ArrowDown';
      if (active < 0) { // from the current value: the next / previous item (for numbers, the next larger / smaller)
        const cur = input.value.trim().toLowerCase(), i = shown.findIndex((it) => it.value.toLowerCase() === cur);
        const num = parseFloat(cur);
        if (i >= 0) active = Math.min(n - 1, Math.max(0, i + (down ? 1 : -1)));
        else if (!Number.isNaN(num) && shown.every((it) => !Number.isNaN(parseFloat(it.value)))) {
          const j = down ? shown.findIndex((it) => parseFloat(it.value) > num) : shown.map((it) => parseFloat(it.value) < num).lastIndexOf(true);
          active = j >= 0 ? j : down ? n - 1 : 0;
        } else active = 0;
      } else active = (active + (down ? 1 : -1) + n) % n;
      render();
    } else if (e.key === 'Enter' && pop && active >= 0) {
      e.preventDefault(); e.stopImmediatePropagation(); pick(shown[active]);
    } else if (e.key === 'Escape' && pop) {
      e.preventDefault(); e.stopPropagation(); close();
    } else if (e.key === 'Tab') close();
  });
  input.addEventListener('blur', () => setTimeout(() => { if (document.activeElement !== input) close(); }, 0));
  return { open, close, get isOpen() { return !!pop; } };
}

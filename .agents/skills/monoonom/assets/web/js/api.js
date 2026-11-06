// Thin client for the local Python server. The browser only ever sends *paths*; Python reads/writes files.
//
// Auth: the server prints the page URL as http://localhost:PORT/#t=TOKEN. The token is read from the
// fragment (never sent to the server as part of the URL), kept in sessionStorage and sent as the
// X-Monoonom-Token header. This works in previews that embed the page in a cross-site iframe
// (VS Code Simple Browser, Tauri apps), where cookies are not sent. Top-level tabs opened without the
// token still work through the HttpOnly cookie the page gets when it loads.
import { LANG, serverText } from './i18n.js';

const KEY = 'mn_token';
let token = null;
let fromUrl = false; // the token came with this page's URL (not from storage left by an earlier server)
(function readToken() {
  if (typeof location === 'undefined') return; // imported outside a browser
  const m = /(?:^#|&)t=([^&]*)/.exec(location.hash);
  if (m) {
    try { token = decodeURIComponent(m[1]) || null; } catch { token = null; } // mangled URL: no token
    fromUrl = !!token;
  }
  try {
    if (token) sessionStorage.setItem(KEY, token); else token = sessionStorage.getItem(KEY);
    // stored: drop t= from the address bar (only when it survives a reload in storage), keep the rest
    if (m && token && sessionStorage.getItem(KEY) === token) {
      const rest = location.hash.slice(1).split('&').filter((p) => !p.startsWith('t=')).join('&');
      history.replaceState(null, '', location.pathname + location.search + (rest ? '#' + rest : ''));
    }
  } catch { /* storage blocked: keep the token in memory and in the URL */ }
})();

/** A stored token the server rejects belongs to another server that used this port: forget it, so
 * the page falls back to its cookie and the 401 hint says the token is missing. */
function forgetStaleToken() {
  if (!token || fromUrl) return;
  token = null;
  try { sessionStorage.removeItem(KEY); } catch { /* ignore */ }
}

/** '#t=TOKEN' for URLs of new windows of this page ('' without a token). */
export const tokenFragment = () => (token ? `#t=${encodeURIComponent(token)}` : '');
export const hasToken = () => !!token;

async function call(method, url, body) {
  const headers = { 'X-Monoonom': '1' };
  if (body !== undefined) headers['Content-Type'] = 'application/json';
  if (token) headers['X-Monoonom-Token'] = token;
  const res = await fetch(url, {
    method,
    credentials: 'same-origin',
    headers,
    body: body === undefined ? undefined : typeof body === 'string' ? body : JSON.stringify(body),
  });
  let data;
  try { data = await res.json(); } catch { data = {}; }
  if (!res.ok) {
    if (res.status === 401 && data.reason === 'token-mismatch') forgetStaleToken();
    // server messages are English: e.message is in the UI language, e.raw the server's text
    const e = new Error(serverText(data.error) || `HTTP ${res.status}`);
    e.status = res.status; e.data = data; e.raw = data.error || '';
    throw e;
  }
  return data;
}
const D = (id) => `/api/docs/${encodeURIComponent(id)}`;

export const api = {
  info: () => call('GET', '/api/info'),
  fs: (path) => call('GET', '/api/fs' + (path ? `?path=${encodeURIComponent(path)}` : '')),
  docs: () => call('GET', '/api/docs'),
  openDoc: (path, clientId) => call('POST', '/api/docs', { path, actor: 'user', clientId }),
  newDoc: (name, clientId) => call('POST', '/api/docs', { new: true, name, lang: LANG, actor: 'user', clientId }),
  snapshot: (id) => call('GET', `${D(id)}/snapshot`),
  ops: (id, payload) => call('POST', `${D(id)}/ops`, payload),
  presence: (id, payload) => call('POST', `${D(id)}/presence`, payload),
  calc: (id, payload) => call('POST', `${D(id)}/calc`, payload),
  saveDoc: (id, opts) => call('POST', `${D(id)}/save`, opts),
  pivotRefresh: (id, payload) => call('POST', `${D(id)}/pivots/refresh`, payload),
  fonts: () => call('GET', `/api/fonts?lang=${LANG}`),
  pivotCreate: (id, payload) => call('POST', `${D(id)}/pivots`, { lang: LANG, ...payload }),
  closeDoc: (id, force = false) => call('DELETE', D(id), { force }),
  streamTicket: () => call('POST', '/api/stream/ticket', {}),
  subscribe: (payload) => call('POST', '/api/stream/subscribe', payload),
  unsubscribe: (payload) => call('POST', '/api/stream/unsubscribe', payload),
};

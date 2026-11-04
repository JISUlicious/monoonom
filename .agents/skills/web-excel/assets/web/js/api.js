// Thin client for the local Python server. The browser only ever sends *paths*; Python reads/writes files.
// Auth: the page receives an HttpOnly SameSite=Strict cookie when it loads; POSTs also carry X-WebExcel.
async function call(method, url, body) {
  const res = await fetch(url, {
    method,
    credentials: 'same-origin',
    headers: body !== undefined ? { 'Content-Type': 'application/json', 'X-WebExcel': '1' } : { 'X-WebExcel': '1' },
    body: body === undefined ? undefined : typeof body === 'string' ? body : JSON.stringify(body),
  });
  let data;
  try { data = await res.json(); } catch { data = {}; }
  if (!res.ok) {
    const e = new Error(data.error || `HTTP ${res.status}`);
    e.status = res.status; e.data = data;
    throw e;
  }
  return data;
}
const D = (id) => `/api/docs/${encodeURIComponent(id)}`;

export const api = {
  info: () => call('GET', '/api/info'),
  fs: (path) => call('GET', '/api/fs' + (path ? `?path=${encodeURIComponent(path)}` : '')),
  docs: () => call('GET', '/api/docs'),
  openDoc: (path) => call('POST', '/api/docs', { path }),
  newDoc: (name) => call('POST', '/api/docs', { new: true, name }),
  snapshot: (id) => call('GET', `${D(id)}/snapshot`),
  ops: (id, payload) => call('POST', `${D(id)}/ops`, payload),
  presence: (id, payload) => call('POST', `${D(id)}/presence`, payload),
  calc: (id, payload) => call('POST', `${D(id)}/calc`, payload),
  saveDoc: (id, opts) => call('POST', `${D(id)}/save`, opts),
  closeDoc: (id, force = false) => call('DELETE', D(id), { force }),
  subscribe: (payload) => call('POST', '/api/stream/subscribe', payload),
  unsubscribe: (payload) => call('POST', '/api/stream/unsubscribe', payload),
};

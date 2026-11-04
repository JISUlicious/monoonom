// Thin client for the local Python server. The browser only ever sends *paths*; Python reads/writes files.
async function call(method, url, body) {
  const res = await fetch(url, {
    method,
    headers: body !== undefined ? { 'Content-Type': 'application/json', 'X-WebExcel': '1' } : { 'X-WebExcel': '1' },
    body: body === undefined ? undefined : typeof body === 'string' ? body : JSON.stringify(body),
  });
  let data;
  try { data = await res.json(); } catch { data = {}; }
  if (!res.ok) { const e = new Error(data.error || `HTTP ${res.status}`); e.status = res.status; throw e; }
  return data;
}

export const api = {
  info: () => call('GET', '/api/info'),
  fs: (path) => call('GET', '/api/fs' + (path ? `?path=${encodeURIComponent(path)}` : '')),
  open: (path) => call('POST', '/api/open', { path }),
  save: (path, workbook, { overwrite = false, base = null } = {}) => call('POST', '/api/save', { path, workbook, overwrite, base }),
  putSession: (id, json) => call('POST', `/api/session/${encodeURIComponent(id)}`, JSON.stringify(json)),
  getSession: (id) => call('GET', `/api/session/${encodeURIComponent(id)}`),
};

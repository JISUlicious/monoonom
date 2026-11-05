// MoNooNoM icons: a 3x3 grid with two outlined pens writing on it, point-symmetric about the centre
// (one in the middle row's right cell, its 180° copy in the left cell).
//   node tools/make_icons.cjs     -> .agents/skills/monoonom/assets/web/img/{monoonom,favicon,monoonom-white}.svg
// monoonom.svg: olive on a transparent background, the pens' bodies filled white (they lie on the grid);
// favicon.svg: the same with heavier lines for 16 px; monoonom-white.svg: white lines for coloured
// backgrounds: the grid is masked out under the pens instead of filling them, so any background shows.
const fs = require('fs');
const path = require('path');

const OUT = path.join(__dirname, '..', '.agents', 'skills', 'monoonom', 'assets', 'web', 'img');
const OLIVE = '#4B7A55';
const X0 = 12, W = 40, N = 3, cell = W / N;
const ANG = -70, L = 27, PW = 9;
const tip1 = [X0 + cell * 2.5, X0 + cell * 1.55], tip2 = [64 - tip1[0], 64 - tip1[1]];
const f = (n) => +n.toFixed(2);

function penOutline(Wd) {
  const nib = Wd * 1.25;
  return `M0 0 L${f(nib)} ${f(-Wd / 2)} L${f(L - Wd * 0.4)} ${f(-Wd / 2)} Q${L} ${f(-Wd / 2)} ${L} 0 Q${L} ${f(Wd / 2)} ${f(L - Wd * 0.4)} ${f(Wd / 2)} L${f(nib)} ${f(Wd / 2)} Z`;
}
const at = (tip, ang, inner) => `<g transform="translate(${f(tip[0])} ${f(tip[1])}) rotate(${ang})">${inner}</g>`;

function grid(c, sw, fill) {
  let g = `<rect x="${X0}" y="${X0}" width="${W}" height="${W}" rx="2" fill="${fill}" stroke="${c}" stroke-width="${f(sw)}"/>`;
  for (let i = 1; i < N; i++) {
    const t = f(X0 + cell * i);
    g += `<line x1="${t}" y1="${X0}" x2="${t}" y2="${X0 + W}" stroke="${c}" stroke-width="${f(sw * 0.8)}"/>`
       + `<line x1="${X0}" y1="${t}" x2="${X0 + W}" y2="${t}" stroke="${c}" stroke-width="${f(sw * 0.8)}"/>`;
  }
  return g;
}
function pen(c, sw, fill) {
  const nib = PW * 1.25, cap = PW * 0.85;
  return `<path d="${penOutline(PW)}" fill="${fill}" stroke="${c}" stroke-width="${f(sw)}" stroke-linejoin="round"/>`
    + `<line x1="${f(nib)}" y1="${f(-PW / 2)}" x2="${f(nib)}" y2="${f(PW / 2)}" stroke="${c}" stroke-width="${f(sw * 0.85)}"/>`
    + `<line x1="${f(L - cap)}" y1="${f(-PW / 2)}" x2="${f(L - cap)}" y2="${f(PW / 2)}" stroke="${c}" stroke-width="${f(sw * 0.85)}"/>`
    + `<circle cx="${f(nib * 0.45)}" cy="0" r="${f(sw * 0.5)}" fill="${c}"/>`;
}
const svg = (inner) => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">${inner}</svg>\n`;

function solid(s) { // olive, pens filled white over a white grid
  return svg(grid(OLIVE, 2.6 * s, '#fff') + at(tip1, ANG, pen(OLIVE, 2.4 * s, '#fff')) + at(tip2, ANG + 180, pen(OLIVE, 2.4 * s, '#fff')));
}
function knockout(s, c) { // c lines, nothing filled; the grid hidden under the pens (mask)
  const cut = (tip, ang) => at(tip, ang, `<path d="${penOutline(PW)}" fill="#000" stroke="#000" stroke-width="${f(2.4 * s)}" stroke-linejoin="round"/>`);
  const mask = `<mask id="m" maskUnits="userSpaceOnUse" x="0" y="0" width="64" height="64"><rect width="64" height="64" fill="#fff"/>${cut(tip1, ANG)}${cut(tip2, ANG + 180)}</mask>`;
  return svg(`<defs>${mask}</defs><g mask="url(#m)">${grid(c, 2.6 * s, 'none')}</g>` + at(tip1, ANG, pen(c, 2.4 * s, 'none')) + at(tip2, ANG + 180, pen(c, 2.4 * s, 'none')));
}
fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(path.join(OUT, 'monoonom.svg'), solid(1));
fs.writeFileSync(path.join(OUT, 'favicon.svg'), solid(1.6));
fs.writeFileSync(path.join(OUT, 'monoonom-white.svg'), knockout(1.15, '#fff'));
console.log('wrote', ['monoonom.svg', 'favicon.svg', 'monoonom-white.svg'].map((n) => `${n} (${fs.statSync(path.join(OUT, n)).size} B)`).join(', '));

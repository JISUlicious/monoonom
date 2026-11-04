// Reads JSON lines {f, op, args} and prints the JS result for each (used by refs_parity.test.py).
import { translateFormula, adjustFormula, renameSheetInFormula } from '../.agents/skills/web-excel/assets/web/js/formula.js';
import { createInterface } from 'readline';
const rl = createInterface({ input: process.stdin });
const out = [];
rl.on('line', (line) => {
  const { f, op, args } = JSON.parse(line);
  let r;
  try {
    r = op === 'translate' ? translateFormula(f, ...args) : op === 'adjust' ? adjustFormula(f, ...args) : renameSheetInFormula(f, ...args);
  } catch (e) { r = 'ERR:' + e.message; }
  out.push(r);
});
rl.on('close', () => process.stdout.write(JSON.stringify(out)));

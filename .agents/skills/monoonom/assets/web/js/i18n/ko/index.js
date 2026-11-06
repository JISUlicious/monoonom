// Korean UI text, keyed by the English text in the code (see ../../i18n.js). One file per module.
import page from './page.js';
import app from './app.js';
import dialogs from './dialogs.js';
import fnhelp from './fnhelp.js';
import misc from './misc.js';
import server from './server.js';

// Korean Excel's currency and long date formats and new-workbook font
const formats = { currency: '"₩"#,##0', currencySymbol: '₩', currencyDecimals: 0, longDate: 'yyyy"년" m"월" d"일" dddd', monthDay: 'm"월" d"일"', defaultFont: '맑은 고딕' };

export default { text: { ...page, ...misc, ...dialogs, ...app, ...fnhelp }, server, formats };

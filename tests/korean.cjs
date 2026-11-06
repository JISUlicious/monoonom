// The browser tests check the Korean UI text: every page and context they open gets the ko-KR
// locale (navigator.language), which makes the preview pick Korean. tests/i18n.e2e.cjs checks English.
module.exports = function korean(browser, locale = 'ko-KR') {
  for (const m of ['newPage', 'newContext']) {
    const f = browser[m].bind(browser);
    browser[m] = (opts = {}) => f({ locale, ...opts });
  }
  return browser;
};

/* The Google Sheets flow in the real app, headless: Google's sign-in and Picker scripts are replaced by stubs and
 * sheets.googleapis.com by tests/fake-sheets.js. Covers save as a new sheet, a collaborator's edits merged on save,
 * an overlapping edit reported, opening through the Picker, Recent, and a ?gsheet= link.
 * Run: node tests/google-e2e.js  (Chromium; PW_EXE=/path/to/chrome to use a preinstalled browser) */
const http = require('http');
const fs = require('fs');
const path = require('path');
const playwright = require('playwright');
const FakeSheets = require('./fake-sheets.js');

const root = path.resolve(__dirname, '..');
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css', '.svg': 'image/svg+xml' };
const server = http.createServer((req, res) => {
  fs.readFile(path.join(root, decodeURIComponent(req.url.split('?')[0])), (err, data) => {
    if (err) { res.writeHead(404); res.end(); return; }
    res.writeHead(200, { 'Content-Type': types[path.extname(req.url.split('?')[0])] || 'application/octet-stream' }); res.end(data);
  });
});
const GSI_STUB = `window.google = window.google || {};
google.accounts = { oauth2: {
  initTokenClient(cfg) { window.__clientId = cfg.client_id; window.__scope = cfg.scope; return { requestAccessToken() { window.__tokenRequests = (window.__tokenRequests || 0) + 1; setTimeout(() => cfg.callback({ access_token: 'fake-token', expires_in: 3600 }), 5); } }; },
  revoke(t, cb) { window.__revoked = t; cb && cb(); } } };`;
const GAPI_STUB = `window.gapi = { load(name, o) { window.google = window.google || {};
  google.picker = { ViewId: { SPREADSHEETS: 'spreadsheets' }, DocsViewMode: { LIST: 'list' }, Response: { ACTION: 'action', DOCUMENTS: 'docs' }, Action: { PICKED: 'picked', CANCEL: 'cancel' }, Document: { ID: 'id' },
    DocsView: class { setMode() { return this; } setFileIds(id) { window.__pickerFileIds = id; return this; } },
    PickerBuilder: class { setAppId(a) { window.__appId = a; return this; } setOAuthToken() { return this; } setDeveloperKey(k) { window.__apiKey = k; return this; } addView() { return this; } setTitle() { return this; } setCallback(cb) { this.cb = cb; return this; }
      build() { const cb = this.cb; return { setVisible() { setTimeout(() => cb(window.__pick ? { action: 'picked', docs: [{ id: window.__pick }] } : { action: 'cancel' }), 5); } }; } } };
  (o.callback || o)(); } };`;

let failed = 0;
const check = (cond, msg) => { console.log((cond ? '  ok   ' : ' FAIL  ') + msg); if (!cond) failed++; };

(async () => {
  await new Promise(r => server.listen(0, '127.0.0.1', r));
  const base = `http://127.0.0.1:${server.address().port}/` + (process.argv[2] || 'src/index.html');
  const browser = await playwright.chromium.launch(process.env.PW_EXE ? { executablePath: process.env.PW_EXE } : {});
  const ctx = await browser.newContext({ viewport: { width: 1300, height: 900 } });
  const store = new Map();
  let badAuth = 0;
  await ctx.route('https://accounts.google.com/gsi/client', r => r.fulfill({ contentType: 'text/javascript', body: GSI_STUB }));
  await ctx.route('https://apis.google.com/js/api.js', r => r.fulfill({ contentType: 'text/javascript', body: GAPI_STUB }));
  await ctx.route('https://sheets.googleapis.com/**', r => {
    const q = r.request();
    if (q.headers().authorization !== 'Bearer fake-token') badAuth++;
    const res = FakeSheets.handle(store, q.method(), q.url(), q.postData() ? JSON.parse(q.postData()) : null);
    r.fulfill({ status: res.status, contentType: 'application/json', body: JSON.stringify(res.json) });
  });
  await ctx.addInitScript(() => { window.SHEETWRITER_GOOGLE_TEST = true; });
  const p = await ctx.newPage();
  p.on('pageerror', e => { console.log('page error:', e.message); failed++; });
  p.on('dialog', d => { console.log('unexpected alert:', d.message()); failed++; d.dismiss(); });
  const status = () => p.textContent('#status');
  const idle = () => p.waitForFunction(() => !/Saving to Google|Opening from Google|Creating a Google/.test(document.querySelector('#status').textContent), null, { timeout: 10000 });
  const menu = async label => { await p.click('#btn-google'); await p.click(`#popup-menu button:has-text("${label}")`); };
  const cellText = async i => p.$eval(`.card[data-i="${i}"] textarea.main`, t => t.value);
  const typeInto = async (i, text) => { await p.click(`.card[data-i="${i}"] .mainwrap`); await p.keyboard.press('Control+a'); await p.keyboard.type(text); };

  await p.goto(base); await p.waitForTimeout(500);
  // 1. A new workbook saved as a new Google Sheet
  await typeInto(0, 'First paragraph.'); await p.keyboard.press('Enter'); await p.keyboard.type('Second paragraph.');
  await menu('Save as a new Google Sheet');
  await p.fill('#ask-input', 'Shared essay'); await p.click('#ask-ok'); await idle();
  const ss = [...store.values()][0];
  check(ss && ss.title === 'Shared essay', 'a spreadsheet named "Shared essay" was created');
  const tab = () => ss.sheets.find(s => !s.title.startsWith('.'));
  check(FakeSheets.shown(tab()).some(r => r.includes('Second paragraph.')), 'its chapter tab holds the text');
  check(/Google Sheets/.test(await status()) && !(await p.evaluate(() => window.SheetWriter.state.dirty)), 'the status bar says Google Sheets, nothing unsaved');
  check(await p.evaluate(() => window.__scope) === 'https://www.googleapis.com/auth/drive.file', 'sign-in asks only for drive.file');

  // 2. A collaborator edits in Google Sheets: a new row, and a note on row 1
  const head = FakeSheets.shown(tab())[0], col = n => head.indexOf(n) + 1;
  const t = tab(); t.cells.push(head.map(c => (c === 'text' ? 'Anna wrote a third paragraph.' : c === '.kind' ? 'p' : '')));
  FakeSheets.setCell(ss, t.title, 2, col('notes'), 'Anna: source?');
  await typeInto(1, 'Second paragraph, revised.');
  await p.keyboard.press('Control+s'); await idle();
  const shown = FakeSheets.shown(tab()).map(r => r[col('text') - 1]);
  check(shown.includes('Second paragraph, revised.') && shown.includes('Anna wrote a third paragraph.'), 'save kept both: my revision and Anna\'s new row');
  check(FakeSheets.shown(tab())[1][col('notes') - 1] === 'Anna: source?', 'and her note');
  check(await p.$$eval('.card', cs => cs.length) === 3 && (await cellText(2)) === 'Anna wrote a third paragraph.', 'her row appears in the editor after saving');

  // 3. The same cell changed on both sides: mine is kept, hers is shown
  FakeSheets.setCell(ss, t.title, 2, col('text'), 'Anna changed the first one.');
  await typeInto(0, 'I changed the first one too.');
  await p.keyboard.press('Control+s'); await idle();
  check(await p.$eval('#dlg-conflicts', d => d.open), 'the overlap dialog opens');
  check((await p.$eval('#cf-list', t => t.value)).includes('Anna changed the first one.'), 'listing her text');
  check(FakeSheets.shown(tab())[1][col('text') - 1] === 'I changed the first one too.', 'my version was written');
  await p.click('#dlg-conflicts button[data-close]');

  // 4. Opening through the Picker in a fresh page, and from Recent
  await p.evaluate(id => { window.__pick = id; }, ss.spreadsheetId);
  await p.goto(base); await p.waitForTimeout(400);
  await p.evaluate(id => { window.__pick = id; }, ss.spreadsheetId);
  await menu('Open a Google Sheet');
  await p.waitForFunction(() => window.SheetWriter.state.google, null, { timeout: 10000 }); await idle();
  check(await p.evaluate(() => window.__appId) === '993934460159', 'the Picker gets the project number');
  check((await cellText(0)) === 'I changed the first one too.' && await p.$$eval('.card', cs => cs.length) === 3, 'opened from the Picker');
  await p.click('#btn-recent'); await p.waitForSelector('#popup-menu button');
  check(await p.$$eval('#popup-menu button', bs => bs.some(b => b.textContent === 'Shared essay (Google Sheets)')), 'listed in Recent');
  await p.keyboard.press('Escape');

  // 5. A ?gsheet= link copied from the menu, followed in a fresh page, lands on the row
  await p.click('.card[data-i="2"] .mainwrap'); await p.keyboard.press('End');
  await ctx.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: new URL(base).origin });
  await p.click('#btn-recent'); await p.click('#popup-menu button:has-text("Copy Markdown link to row")');
  const md = await p.evaluate(() => navigator.clipboard.readText());
  const link = (md.match(/\((http[^)]+)\)$/) || [])[1] || '';
  check(link.includes('gsheet=' + ss.spreadsheetId) && link.includes('row='), 'link carries the sheet id and the row');
  await p.goto('about:blank'); await p.goto(link); await p.waitForTimeout(400);
  check(await p.$eval('#dlg-link', d => d.open), 'the link asks before signing in');
  await p.click('#link-open');
  await p.waitForFunction(() => window.SheetWriter.state.google, null, { timeout: 10000 }); await idle();
  check(await p.evaluate(() => document.activeElement.value) === 'Anna wrote a third paragraph.', 'and opens at the linked row');

  // 6. Sign out
  await menu('Sign out of Google');
  check(await p.evaluate(() => window.__revoked) === 'fake-token', 'signing out revokes the token');
  check(badAuth === 0, 'every Sheets request carried the token');

  console.log(failed ? `\n${failed} failed` : '\nall passed');
  await browser.close(); server.close();
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(2); });

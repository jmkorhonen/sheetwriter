/* A small in-memory stand-in for the Google Sheets API v4, for tests: the requests gsheets.js sends (spreadsheet
 * metadata, values:batchGet, spreadsheets:batchUpdate with addSheet / deleteSheet / updateSheetProperties /
 * repeatCell / updateDimensionProperties, values:batchUpdate RAW, create). It refuses what Google refuses and the
 * writer must avoid: two tabs with one name, removing the last tab, freezing every row or column, writing outside
 * the grid. Used by tests.html directly and by tests/google-e2e.js behind a routed sheets.googleapis.com. */
const FakeSheets = (() => {
  let seq = 0;
  function create(title, tabs = [{ title: 'Sheet1' }]) {
    const ss = { spreadsheetId: 'fake' + (++seq), title, sheets: [] };
    for (const t of tabs) ss.sheets.push({ id: t.id ?? (ss.sheets.length ? 1000 + seq * 10 + ss.sheets.length : 0), title: t.title, rowCount: t.rowCount || 1000, columnCount: t.columnCount || 26, frozenRows: 0, frozenCols: 0, hidden: [], cells: (t.values || []).map(r => r.slice()) });
    return ss;
  }
  const fail = (msg, status = 400) => { const e = new Error(msg); e.status = status; throw e; };
  const byId = (ss, id) => ss.sheets.find(s => s.id === id) || fail('No grid with id: ' + id);
  const titleTaken = (ss, title, except) => ss.sheets.some(s => s !== except && s.title.toLowerCase() === String(title).toLowerCase());
  function batchUpdate(ss, requests) {
    const replies = [];
    for (const rq of requests) {
      if (rq.addSheet) {
        const p = rq.addSheet.properties || {};
        const id = p.sheetId ?? (1 + Math.floor(Math.random() * 1e9));
        if (ss.sheets.some(s => s.id === id)) fail('A sheet with id ' + id + ' already exists.');
        const title = p.title || 'Sheet' + (ss.sheets.length + 1);
        if (titleTaken(ss, title)) fail(`A sheet with the name "${title}" already exists. Please enter another name.`);
        const g = p.gridProperties || {};
        ss.sheets.splice(p.index ?? ss.sheets.length, 0, { id, title, rowCount: g.rowCount || 1000, columnCount: g.columnCount || 26, frozenRows: 0, frozenCols: 0, hidden: [], cells: [] });
        replies.push({ addSheet: { properties: { sheetId: id, title } } });
      } else if (rq.deleteSheet) {
        const s = byId(ss, rq.deleteSheet.sheetId);
        if (ss.sheets.length === 1) fail("You can't remove all the sheets in a document.");
        ss.sheets.splice(ss.sheets.indexOf(s), 1); replies.push({});
      } else if (rq.updateSheetProperties) {
        const p = rq.updateSheetProperties.properties, fields = rq.updateSheetProperties.fields.split(',');
        const s = byId(ss, p.sheetId);
        if (fields.includes('title')) { if (titleTaken(ss, p.title, s)) fail(`A sheet with the name "${p.title}" already exists. Please enter another name.`); s.title = p.title; }
        if (fields.includes('index')) { // "before the move" indexes, as Google counts them
          const from = ss.sheets.indexOf(s); let to = p.index;
          if (to > from) to -= 1;
          ss.sheets.splice(from, 1); ss.sheets.splice(Math.min(to, ss.sheets.length), 0, s);
        }
        const g = p.gridProperties || {};
        if (fields.includes('gridProperties.rowCount')) s.rowCount = g.rowCount;
        if (fields.includes('gridProperties.columnCount')) s.columnCount = g.columnCount;
        if (fields.includes('gridProperties.frozenRowCount')) s.frozenRows = g.frozenRowCount || 0;
        if (fields.includes('gridProperties.frozenColumnCount')) s.frozenCols = g.frozenColumnCount || 0;
        if (s.frozenRows && s.frozenRows >= s.rowCount) fail("You can't freeze all visible rows on the sheet.");
        if (s.frozenCols && s.frozenCols >= s.columnCount) fail("You can't freeze all visible columns on the sheet.");
        s.cells.length = Math.min(s.cells.length, s.rowCount);
        replies.push({});
      } else if (rq.updateDimensionProperties) {
        const r = rq.updateDimensionProperties.range, s = byId(ss, r.sheetId);
        for (let c = r.startIndex; c < r.endIndex; c++) if (!s.hidden.includes(c)) s.hidden.push(c);
        replies.push({});
      } else if (rq.repeatCell) { byId(ss, rq.repeatCell.range.sheetId); replies.push({}); }
      else fail('Unsupported request ' + Object.keys(rq)[0]);
    }
    return { spreadsheetId: ss.spreadsheetId, replies };
  }
  function parseRange(ss, range) {
    const m = /^(?:'((?:[^']|'')*)'|([^!]+))(?:!A1)?$/.exec(range) || fail('Unable to parse range: ' + range);
    const title = m[1] != null ? m[1].replace(/''/g, "'") : m[2];
    return ss.sheets.find(s => s.title === title) || fail('Unable to parse range: ' + range);
  }
  function valuesUpdate(ss, body) {
    if (body.valueInputOption !== 'RAW') fail('Expected RAW');
    for (const d of body.data) {
      const s = parseRange(ss, d.range);
      if (d.values.length > s.rowCount) fail(`Range (${d.range}) exceeds grid limits. Max rows: ${s.rowCount}`);
      d.values.forEach((row, i) => {
        if (row.length > s.columnCount) fail(`Range (${d.range}) exceeds grid limits. Max columns: ${s.columnCount}`);
        s.cells[i] = s.cells[i] || [];
        row.forEach((v, j) => { s.cells[i][j] = v; });
      });
    }
    return { totalUpdatedCells: body.data.reduce((n, d) => n + d.values.reduce((m, r) => m + r.length, 0), 0) };
  }
  /** Displayed values, trimmed of empty trailing cells and rows, as values:batchGet returns them. */
  function shown(s) {
    const out = s.cells.map(r => {
      const row = (r || []).map(v => (v == null ? '' : String(v)));
      while (row.length && row[row.length - 1] === '') row.pop();
      return row;
    });
    while (out.length && !out[out.length - 1].length) out.pop();
    return out;
  }
  /** Someone editing in Google Sheets: set a cell (row and column from 1). */
  function setCell(ss, title, r, c, v) { const s = ss.sheets.find(x => x.title === title); s.cells[r - 1] = s.cells[r - 1] || []; s.cells[r - 1][c - 1] = v; }
  /** Answer one HTTP request as sheets.googleapis.com would. → {status, json} */
  function handle(store, method, url, body) {
    try {
      const u = new URL(url), path = decodeURIComponent(u.pathname.replace(/^\/v4\/spreadsheets\/?/, ''));
      if (method === 'POST' && path === '') { const ss = create(body.properties.title, (body.sheets || []).map(t => ({ title: t.properties.title }))); store.set(ss.spreadsheetId, ss); return { status: 200, json: { spreadsheetId: ss.spreadsheetId, properties: { title: ss.title }, sheets: ss.sheets.map((s, index) => ({ properties: { sheetId: s.id, title: s.title, index, gridProperties: { rowCount: s.rowCount, columnCount: s.columnCount } } })) } }; }
      const m = /^([^/:]+)(?::batchUpdate|\/values:batchGet|\/values:batchUpdate)?$/.exec(path);
      const ss = m && store.get(m[1]);
      if (!ss) return { status: 404, json: { error: { code: 404, message: 'Requested entity was not found.' } } };
      if (method === 'GET' && path === ss.spreadsheetId) return { status: 200, json: { properties: { title: ss.title }, sheets: ss.sheets.map((s, index) => ({ properties: { sheetId: s.id, title: s.title, index, gridProperties: { rowCount: s.rowCount, columnCount: s.columnCount } } })) } };
      if (method === 'GET' && path.endsWith('/values:batchGet')) return { status: 200, json: { valueRanges: u.searchParams.getAll('ranges').map(r => ({ range: r, values: shown(parseRange(ss, r)) })) } };
      if (method === 'POST' && path.endsWith(':batchUpdate') && !path.includes('/values')) return { status: 200, json: batchUpdate(ss, body.requests) };
      if (method === 'POST' && path.endsWith('/values:batchUpdate')) return { status: 200, json: valuesUpdate(ss, body) };
      return { status: 400, json: { error: { message: 'Unsupported ' + method + ' ' + url } } };
    } catch (e) { return { status: e.status || 400, json: { error: { code: e.status || 400, message: e.message } } }; }
  }
  /** What fetchRemote would make of the spreadsheet. */
  function remote(ss) { return { title: ss.title, sheets: ss.sheets.map((s, index) => ({ id: s.id, title: s.title, index, rowCount: s.rowCount, columnCount: s.columnCount, values: shown(s) })) }; }
  /** Apply what GSheets.plan() returns, as the two write requests would. */
  function apply(ss, p) { batchUpdate(ss, p.requests); valuesUpdate(ss, { valueInputOption: 'RAW', data: p.data }); }
  return { create, batchUpdate, valuesUpdate, handle, remote, apply, setCell, shown };
})();
if (typeof module !== 'undefined') module.exports = FakeSheets;

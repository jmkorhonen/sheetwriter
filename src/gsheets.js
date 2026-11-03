/* SheetWriter — Google Sheets as a place to keep the workbook (online version only).
 *
 * Sign-in is Google Identity Services' token flow with the drive.file scope: SheetWriter sees only the spreadsheets a
 * person opens with it (through the Google Picker) or creates with it. The browser talks to the Sheets API directly;
 * there is no server. A spreadsheet holds the same structure as the .xlsx: a tab per chapter, the .sheetwriter tab,
 * the hidden .id column.
 *
 * Saving merges by row id, so several people can work on one spreadsheet: SheetWriter reads the spreadsheet again,
 * keeps what others changed since this copy was opened (rows, cells, added and deleted rows, new tabs), and writes only
 * on top of that. The same cell changed on both sides keeps this side's text and is reported.
 *
 *   GSheets.available() / unavailableReason()
 *   GSheets.pick(fileId?) → spreadsheet id or null        GSheets.open(id) → {doc, title, base, warnings}
 *   GSheets.save(id, mine, base) → {doc, conflicts, base}  GSheets.saveNew(title, doc) → {id, doc, base}
 *   pure, for tests: docFromRemote, merge, plan, writtenCopy, prepare
 */
const GSheets = (() => {
  const CLIENT_ID = '993934460159-uglvpp10gmlgmj625k64ejfbk496iulk.apps.googleusercontent.com';
  const API_KEY = 'AIzaSyBDI_gszEvxrNofJV3Sjred7Lf_t5vXfZk'; // for the Picker; restricted to the site in the Cloud console
  const APP_ID = '993934460159'; // the Cloud project number: files chosen in the Picker are opened to this app
  const SCOPE = 'https://www.googleapis.com/auth/drive.file';
  const ORIGINS = ['https://jmkorhonen.github.io']; // the sign-in client accepts only these
  const GSI = 'https://accounts.google.com/gsi/client', GAPI = 'https://apis.google.com/js/api.js';
  const SHEETS = 'https://sheets.googleapis.com/v4/spreadsheets';
  const SETTINGS = '.sheetwriter';
  const IDENT = Model.IDENT;

  const available = () => ORIGINS.includes(location.origin) || !!window.SHEETWRITER_GOOGLE_TEST;
  const unavailableReason = () => location.protocol === 'file:'
    ? 'Google Sheets works in the online version (jmkorhonen.github.io/sheetwriter), not in a downloaded copy.'
    : 'Google Sheets works only in the online version at jmkorhonen.github.io/sheetwriter.';
  const url = id => 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(id) + '/edit';

  // ---------- Google's scripts, sign-in, requests ----------
  const scripts = {};
  function loadScript(src) {
    return scripts[src] || (scripts[src] = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = src; s.async = true;
      s.onload = () => res();
      s.onerror = () => { delete scripts[src]; rej(new Error('Could not reach Google (offline?).')); };
      document.head.appendChild(s);
    }));
  }
  /** Load the sign-in script ahead of a click, so the sign-in pop-up opens straight from that click. */
  function preload() { if (available()) loadScript(GSI).catch(() => {}); }

  let token = null, tokenUntil = 0;
  const signedIn = () => !!token && Date.now() < tokenUntil;
  async function getToken() {
    if (signedIn()) return token;
    await loadScript(GSI);
    return new Promise((resolve, reject) => {
      const client = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID, scope: SCOPE,
        callback: r => {
          if (!r || r.error) { reject(new Error('Google sign-in failed: ' + ((r && (r.error_description || r.error)) || 'no answer'))); return; }
          token = r.access_token; tokenUntil = Date.now() + (Number(r.expires_in) || 3600) * 1000 - 60000;
          resolve(token);
        },
        error_callback: e => {
          const type = e && e.type;
          const err = new Error(type === 'popup_closed' ? 'Google sign-in was closed.'
            : type === 'popup_failed_to_open' ? 'The Google sign-in window was blocked. Allow pop-ups for this site and try again.'
              : 'Google sign-in failed' + (type ? ' (' + type + ')' : '') + '.');
          err.cancelled = type === 'popup_closed';
          reject(err);
        },
      });
      client.requestAccessToken({ prompt: '' });
    });
  }
  function signOut() {
    if (token && window.google && google.accounts && google.accounts.oauth2) google.accounts.oauth2.revoke(token, () => {});
    token = null; tokenUntil = 0;
  }
  async function api(path, { method = 'GET', body } = {}, attempt = 0) {
    const t = await getToken();
    const r = await fetch(SHEETS + path, { method, headers: Object.assign({ Authorization: 'Bearer ' + t }, body ? { 'Content-Type': 'application/json' } : {}), body: body ? JSON.stringify(body) : undefined });
    if (r.status === 401 && attempt < 1) { token = null; return api(path, { method, body }, attempt + 1); }
    if ((r.status === 429 || r.status >= 500) && attempt < 3) { // over the per-minute quota, or Google busy: wait and retry
      await new Promise(res => setTimeout(res, 1000 * 2 ** attempt + Math.random() * 400));
      return api(path, { method, body }, attempt + 1);
    }
    const data = await r.json().catch(() => ({}));
    if (!r.ok) { const e = new Error((data.error && data.error.message) || 'Google Sheets answered ' + r.status); e.status = r.status; throw e; }
    return data;
  }

  /** The Google Picker, showing spreadsheets; with fileId, that one is offered (a link to a sheet not yet opened here). */
  async function pick(fileId) {
    const t = await getToken();
    await loadScript(GAPI);
    await new Promise((res, rej) => gapi.load('picker', { callback: res, onerror: () => rej(new Error('Could not load the Google file picker.')) }));
    const P = google.picker;
    return new Promise(resolve => {
      const view = new P.DocsView(P.ViewId.SPREADSHEETS).setMode(P.DocsViewMode.LIST);
      if (fileId) view.setFileIds(fileId);
      new P.PickerBuilder().setAppId(APP_ID).setOAuthToken(t).setDeveloperKey(API_KEY).addView(view)
        .setTitle(fileId ? 'Confirm the Google Sheet to open in SheetWriter' : 'Open a Google Sheet in SheetWriter')
        .setCallback(d => {
          const a = d[P.Response.ACTION];
          if (a === P.Action.PICKED) resolve(d[P.Response.DOCUMENTS][0][P.Document.ID]);
          else if (a === P.Action.CANCEL) resolve(null);
        })
        .build().setVisible(true);
    });
  }

  // ---------- reading ----------
  const quote = t => "'" + String(t).replace(/'/g, "''") + "'";
  /** Every tab's properties and values (as displayed). Two requests. */
  async function fetchRemote(id) {
    const enc = encodeURIComponent(id);
    const meta = await api(`/${enc}?fields=properties.title,sheets.properties(sheetId,title,index,gridProperties(rowCount,columnCount))`);
    const props = (meta.sheets || []).map(s => s.properties).sort((a, b) => a.index - b.index);
    let ranges = [];
    if (props.length) {
      const qs = props.map(p => 'ranges=' + encodeURIComponent(quote(p.title))).join('&');
      ranges = (await api(`/${enc}/values:batchGet?${qs}&valueRenderOption=FORMATTED_VALUE&majorDimension=ROWS`)).valueRanges || [];
    }
    return {
      title: (meta.properties && meta.properties.title) || 'Google Sheet',
      sheets: props.map((p, k) => ({
        id: p.sheetId, title: p.title, index: p.index,
        rowCount: (p.gridProperties && p.gridProperties.rowCount) || 0, columnCount: (p.gridProperties && p.gridProperties.columnCount) || 0,
        values: (ranges[k] && ranges[k].values) || [],
      })),
    };
  }
  /** FNV-1a, as six base-36 characters. */
  function hash36(str) {
    let h = 0x811c9dc5;
    for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193); }
    return ((h >>> 0).toString(36) + '000000').slice(0, 6);
  }
  /** Rows without an .id (added in Google Sheets, or never saved by SheetWriter) get one made from their content, so
   *  every read of the same spreadsheet gives them the same id and a later save can match them. */
  function assignRowIds(doc) {
    const taken = new Set();
    for (const s of doc.sheets) if (s.kind === 'chapter') for (const r of s.rows) if (r[IDENT]) taken.add(r[IDENT]);
    for (const s of doc.sheets) if (s.kind === 'chapter') {
      const cols = Model.userColumns(doc, s).concat('.kind'), seen = new Map();
      for (const r of s.rows) {
        if (String(r[IDENT] || '').trim()) continue;
        const key = s.name + '\u001f' + cols.map(c => r[c] || '').join('\u001f');
        const n = (seen.get(key) || 0) + 1; seen.set(key, n);
        let id = hash36(key + '\u001e' + n);
        for (let k = 1; taken.has(id); k++) id = hash36(key + '\u001e' + n + '\u001e' + k);
        taken.add(id); r[IDENT] = id;
      }
    }
    Model.ensureRowIds(doc); // adds the column where it was missing; keeps the ids above
  }
  /** The workbook model of a spreadsheet read with fetchRemote. Tabs keep their Google id in sheet._gid. */
  function docFromRemote(remote) {
    const adapters = remote.sheets.map(s => {
      const cells = s.values.map(r => r.map(v => (v == null ? '' : String(v))));
      return {
        name: s.title, rowCount: cells.length, columnCount: cells.reduce((n, r) => Math.max(n, r.length), 0),
        getRow(r) { const arr = cells[r - 1] || []; return { cellCount: arr.length, getCell(c) { return { value: arr[c - 1] == null ? '' : arr[c - 1] }; } }; },
      };
    });
    const byName = new Map(adapters.map(a => [a.name, a]));
    const res = XlsxIO.loadFromSheets(adapters, n => byName.get(n));
    const gids = new Map(remote.sheets.map(s => [s.title, s.id]));
    for (const s of res.doc.sheets) if (gids.has(s.name)) s._gid = gids.get(s.name);
    assignRowIds(res.doc);
    Model.ensureIds(res.doc);
    return res;
  }

  // ---------- merging ----------
  const J = x => JSON.stringify(x);
  const RECOMPUTED = new Set(['.no', '.words', '.chars']); // written fresh on every save
  const QUIET = new Set(['.updated', '.author', IDENT]); // follow the text; never reported as conflicts
  const content = (s, r) => s.columns.filter(c => !RECOMPUTED.has(c) && !QUIET.has(c)).map(c => r[c] ?? '');
  const sheetSig = s => (s.kind === 'chapter' ? J([s.name, s.rows.map(r => [r[IDENT], content(s, r)])]) : J([s.name, s.cells]));
  /** Where an item that only `theirs` has goes: after the nearest earlier item (in theirs' order) that `list` has. */
  function insertAfterPredecessor(list, item, theirsList, k, key) {
    let at = 0;
    for (let j = k - 1; j >= 0; j--) { const p = list.findIndex(x => key(x) === key(theirsList[j])); if (p >= 0) { at = p + 1; break; } }
    list.splice(at, 0, item);
  }
  function mergeRows(b, m, t, conflicts) {
    const id = r => r[IDENT];
    const bm = new Map(b.rows.map(r => [id(r), r])), tm = new Map(t.rows.map(r => [id(r), r]));
    t.columns.forEach((c, k) => { if (!m.columns.includes(c) && !b.columns.includes(c)) insertAfterPredecessor(m.columns, c, t.columns, k, x => x); });
    const rows = [];
    for (const r of m.rows) {
      const k = id(r), br = bm.get(k), tr = tm.get(k);
      if (!br) { rows.push(r); continue; } // added here
      if (!tr) { // deleted by someone else
        if (J(content(m, br)) === J(content(m, r))) continue;
        conflicts.push({ sheet: m.name, id: k, text: 'was deleted by someone else, but kept because you changed it' });
        rows.push(r); continue;
      }
      for (const c of m.columns) {
        if (RECOMPUTED.has(c)) continue;
        const bv = br[c] ?? '', mv = r[c] ?? '', tv = tr[c] ?? '';
        if (mv === bv) { if (tv !== bv) r[c] = tv; } // only they changed it
        else if (tv !== bv && tv !== mv && !QUIET.has(c)) conflicts.push({ sheet: m.name, id: k, column: c, theirs: tv });
      }
      rows.push(r);
    }
    const have = new Set(rows.map(id));
    t.rows.forEach((tr, k) => {
      const key = id(tr);
      if (have.has(key)) return;
      const br = bm.get(key);
      if (br) { // deleted here
        if (J(content(t, br)) === J(content(t, tr))) return;
        conflicts.push({ sheet: m.name, id: key, text: 'was deleted here, but kept because someone else changed it' });
      }
      const copy = Object.assign({}, tr); delete copy._id;
      insertAfterPredecessor(rows, copy, t.rows, k, id);
      have.add(key);
    });
    m.rows = rows;
  }
  /** Three-way merge of whole workbooks: base is what this copy last read or wrote, mine the editor, theirs the
   *  spreadsheet now. → {doc, conflicts: [{sheet, id?, column?, theirs?, text?}]} */
  function merge(base, mine, theirs) {
    const conflicts = [];
    const out = JSON.parse(J(mine));
    const sb = base.settings || {}, sm = mine.settings || {}, st = theirs.settings || {};
    for (const k of new Set([...Object.keys(sb), ...Object.keys(sm), ...Object.keys(st)])) {
      if (k === 'view') continue; // where each person is in the workbook is their own
      if (J(sm[k]) === J(sb[k]) && J(st[k]) !== J(sb[k]) && st[k] !== undefined) out.settings[k] = JSON.parse(J(st[k]));
    }
    const bS = new Map((base.sheets || []).filter(s => s._gid != null).map(s => [s._gid, s]));
    const tS = new Map(theirs.sheets.filter(s => s._gid != null).map(s => [s._gid, s]));
    const result = [];
    for (const m of out.sheets) {
      const g = m._gid, b = g != null ? bS.get(g) : null;
      if (!b) { result.push(m); continue; } // new here
      const t = tS.get(g);
      if (!t) { // tab deleted by someone else
        if (sheetSig(b) === sheetSig(m)) continue;
        conflicts.push({ sheet: m.name, text: 'was deleted by someone else, but kept because you changed it' });
        delete m._gid; result.push(m); continue;
      }
      if (m.kind !== 'chapter' || t.kind !== 'chapter') { result.push(m.kind === 'chapter' ? m : JSON.parse(J(t))); continue; } // data tabs are read-only here
      if (m.name === b.name && t.name !== b.name) m.name = t.name;
      mergeRows(b, m, t, conflicts);
      result.push(m);
    }
    const have = new Set(result.map(s => s._gid).filter(g => g != null));
    theirs.sheets.forEach((t, k) => {
      const g = t._gid;
      if (g == null || have.has(g)) return;
      const b = bS.get(g);
      if (b) { // deleted here
        if (sheetSig(b) === sheetSig(t)) return;
        conflicts.push({ sheet: t.name, text: 'was deleted here, but kept because someone else changed it' });
      }
      insertAfterPredecessor(result, JSON.parse(J(t)), theirs.sheets, k, s => s._gid);
      have.add(g);
    });
    const names = new Set();
    for (const s of result) { let n = s.name, k = 2; while (names.has(n.toLowerCase())) n = s.name + ' ' + k++; s.name = n; names.add(n.toLowerCase()); }
    out.sheets = result;
    Model.ensureMetaColumns(out);
    return { doc: Model.ensureIds(out), conflicts };
  }

  // ---------- writing ----------
  /** Header row and rows as the .xlsx and .ods writers write them: numbers, counts and ids filled in. */
  function sheetValues(doc, s) {
    const columns = XlsxIO.fileColumns(doc, s);
    const num = Model.numbering(doc, doc.sheets.indexOf(s));
    const out = [columns.slice()];
    const written = s.rows.filter(r => !Model.rowIsEmpty(doc, s, r));
    for (const r of written.length ? written : s.rows.slice(0, 1)) {
      const i = s.rows.indexOf(r);
      const rc = doc.settings.trackCounts ? Model.rowCounts(doc, s, r) : null;
      out.push(columns.map(c => {
        if (c === '.no') return num.numbers[i];
        if (c === '.indent') return Model.indentOf(r) ? String(Model.indentOf(r)) : '';
        if (rc && c === '.words') return rc.words;
        if (rc && c === '.chars') return rc.chars;
        return String(r[c] ?? '');
      }));
    }
    return out;
  }
  function settingsValues(doc) {
    const rows = XlsxIO.settingsNotes(doc).map(n => ['', n.text, '']);
    rows.push(['', '', ''], ['key', 'value', 'description']);
    for (const r of XlsxIO.settingsRows(doc)) rows.push([r.key, r.value, r.desc]);
    return rows;
  }
  function freshGid(used) { let g; do g = 1 + Math.floor(Math.random() * 2e9); while (used.has(g)); used.add(g); return g; }
  /** Tabs without a Google id get one now, so copies of the workbook taken during a save agree on it. */
  function prepare(doc) {
    const used = new Set(doc.sheets.map(s => s._gid).filter(g => g != null));
    for (const s of doc.sheets) if (s._gid == null) s._gid = freshGid(used);
    return doc;
  }
  /** The requests that make the spreadsheet `remote` hold `doc`: tab changes for spreadsheets:batchUpdate (adds,
   *  renames and deletes ordered so names never collide, then order, size and frozen panes), and the values for
   *  values:batchUpdate, padded so that leftovers of longer old content are cleared. Data tabs are left alone. */
  function plan(doc, remote) {
    Model.ensureRowIds(doc);
    prepare(doc);
    const byId = new Map(remote.sheets.map(r => [r.id, r]));
    const want = doc.sheets.map(s => ({ gid: s._gid, title: s.name, sheet: s, values: s.kind === 'chapter' ? sheetValues(doc, s) : (byId.has(s._gid) ? null : (s.cells || [])) }));
    const sw = remote.sheets.find(r => r.title === SETTINGS);
    const used = new Set(remote.sheets.map(r => r.id).concat(doc.sheets.map(s => s._gid)));
    want.push({ gid: sw ? sw.id : freshGid(used), title: SETTINGS, sheet: null, values: settingsValues(doc) });
    const requests = [], temp = g => '~sheetwriter ' + g;
    const size = w => ({ rows: Math.max(1, w.values ? w.values.length : 0), cols: Math.max(1, w.values ? w.values.reduce((n, r) => Math.max(n, r.length), 0) : 0) });
    for (const w of want) if (!byId.has(w.gid)) requests.push({ addSheet: { properties: { sheetId: w.gid, title: temp(w.gid), gridProperties: { rowCount: size(w).rows, columnCount: size(w).cols } } } });
    for (const w of want) { const r = byId.get(w.gid); if (r && r.title !== w.title) requests.push({ updateSheetProperties: { properties: { sheetId: w.gid, title: temp(w.gid) }, fields: 'title' } }); }
    const keep = new Set(want.map(w => w.gid));
    for (const r of remote.sheets) if (!keep.has(r.id) && !r.title.startsWith('.')) requests.push({ deleteSheet: { sheetId: r.id } });
    // Order: simulate the tab order after the steps above, then move each tab back into place in turn.
    const order = remote.sheets.filter(r => keep.has(r.id) || r.title.startsWith('.')).map(r => r.id).concat(want.filter(w => !byId.has(w.gid)).map(w => w.gid));
    const data = [];
    want.forEach((w, k) => {
      const r = byId.get(w.gid), sz = size(w), chapter = w.sheet && w.sheet.kind === 'chapter';
      const props = { sheetId: w.gid, title: w.title }, fields = ['title'];
      const p = order.indexOf(w.gid);
      if (p !== k) { props.index = k; fields.push('index'); order.splice(p, 1); order.splice(k, 0, w.gid); }
      if (w.values) {
        const usedRows = r ? r.values.length : 0, usedCols = r ? r.values.reduce((n, x) => Math.max(n, x.length), 0) : 0;
        const rows = Math.max(sz.rows, usedRows), cols = Math.max(sz.cols, usedCols);
        const freeze = chapter ? Math.max(0, Math.min(doc.settings.freezeColumns ?? 1, cols - 1)) : 0;
        // a frozen header row needs at least one row below it
        props.gridProperties = { rowCount: Math.max(rows + (chapter ? 1 : 0), r ? r.rowCount : 0), columnCount: Math.max(cols, r ? r.columnCount : 0), frozenRowCount: chapter ? 1 : 0, frozenColumnCount: freeze };
        fields.push('gridProperties.rowCount', 'gridProperties.columnCount', 'gridProperties.frozenRowCount', 'gridProperties.frozenColumnCount');
        data.push({ range: quote(w.title) + '!A1', values: Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => { const v = w.values[i] && w.values[i][j]; return v == null ? '' : v; })) });
      }
      requests.push({ updateSheetProperties: { properties: props, fields: fields.join(',') } });
      if (chapter) {
        requests.push({ repeatCell: { range: { sheetId: w.gid, startRowIndex: 0, endRowIndex: 1 }, cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat.textFormat.bold' } });
        const idCol = w.values[0].indexOf(IDENT);
        if (idCol >= 0) requests.push({ updateDimensionProperties: { range: { sheetId: w.gid, dimension: 'COLUMNS', startIndex: idCol, endIndex: idCol + 1 }, properties: { hiddenByUser: true }, fields: 'hiddenByUser' } });
      }
    });
    return { requests, data };
  }
  async function write(id, doc, remote) {
    const p = plan(doc, remote), enc = encodeURIComponent(id);
    await api(`/${enc}:batchUpdate`, { method: 'POST', body: { requests: p.requests } });
    await api(`/${enc}/values:batchUpdate`, { method: 'POST', body: { valueInputOption: 'RAW', data: p.data } });
  }
  /** The workbook as the spreadsheet holds it after a save (empty rows are not written): the base of the next merge. */
  function writtenCopy(doc) {
    const d = JSON.parse(J(doc));
    for (const s of d.sheets) if (s.kind === 'chapter') { const w = s.rows.filter(r => !Model.rowIsEmpty(d, s, r)); s.rows = w.length ? w : s.rows.slice(0, 1); }
    return J(d);
  }

  // ---------- the operations the app calls ----------
  async function open(id) {
    const remote = await fetchRemote(id);
    const r = docFromRemote(remote);
    return { doc: r.doc, title: remote.title, warnings: r.warnings, hasSettings: r.hasSettings, base: writtenCopy(r.doc) };
  }
  async function save(id, mine, base) {
    const remote = await fetchRemote(id);
    const theirs = docFromRemote(remote).doc;
    const { doc, conflicts } = merge(base ? JSON.parse(base) : { settings: {}, sheets: [] }, mine, theirs);
    await write(id, doc, remote);
    return { doc, conflicts, base: writtenCopy(doc) };
  }
  async function saveNew(title, mine) {
    const created = await api('', { method: 'POST', body: { properties: { title }, sheets: [{ properties: { title: SETTINGS } }] } });
    const doc = JSON.parse(J(mine)); // tabs keep the ids the caller gave them (prepare): none exists yet, so all are added
    const remote = { title, sheets: (created.sheets || []).map(s => ({ id: s.properties.sheetId, title: s.properties.title, index: s.properties.index || 0, rowCount: 1000, columnCount: 26, values: [] })) };
    await write(created.spreadsheetId, doc, remote);
    return { id: created.spreadsheetId, doc, base: writtenCopy(doc) };
  }

  return { available, unavailableReason, url, preload, signedIn, signOut, pick, open, save, saveNew, docFromRemote, merge, plan, prepare, writtenCopy, assignRowIds };
})();

/* SheetWriter — document model and mutators. No DOM, no I/O.
 *
 * doc = {
 *   version, mainColumn,
 *   settings: {numbering: 'continuous'|'per-sheet', title, author, description, created, trackUpdated, trackAuthor,
 *              freezeColumns, countColumns: [...], extra:{key:[value, description]}},
 *   sheets: [
 *     {name, kind:'chapter', columns:[...], rows:[{col: text, _id}]},
 *     {name, kind:'data', cells:[[...]]}
 *   ]
 * }
 * System columns start with a dot: .no (computed numbering, written on save), .kind, .indent,
 * .words/.chars (computed) and .updated/.author (meta, maintained when tracking is on). Everything else is the user's.
 */
const Model = (() => {
  const RESERVED = ['.no', '.kind', '.indent'];
  const META = ['.updated', '.author'];
  const COMPUTED = ['.words', '.chars']; // per-row counts, written on save like "no", recomputed on load
  const IDENT = '.id'; // short random row id, written to the file (hidden column) so formatting added in Excel can follow the row
  const KINDS = ['h1', 'h2', 'h3', 'h4', 'p', 's', 'x'];
  const KIND_ORDER = ['h1', 'h2', 'h3', 'h4', 'p', 's']; // promote/demote ladder
  const DEFAULT_COLUMNS = ['.no', '.kind', '.indent', 'text', 'notes', 'sources'];

  function normKind(k) {
    k = String(k == null ? '' : k).trim().toLowerCase();
    return KINDS.includes(k) ? k : 'p';
  }
  function isHeading(kind) { return /^h[1-6]$/.test(kind || ''); }
  const indentOf = row => SheetNumbering.indentOf(row);

  // Rows carry a private _id (never saved) so UI state such as "collapsed" survives moves and undo.
  let nextId = 1;
  function newId() { return nextId++; }
  function ensureIds(doc) {
    let max = 0;
    for (const s of doc.sheets) if (s.kind === 'chapter') for (const r of s.rows) if (typeof r._id === 'number' && r._id > max) max = r._id;
    nextId = Math.max(nextId, max + 1);
    for (const s of doc.sheets) if (s.kind === 'chapter') for (const r of s.rows) if (typeof r._id !== 'number') r._id = newId();
    return doc;
  }
  /** Six base-36 characters, unique within the document. */
  function newRowId(taken) {
    for (;;) {
      let id = '';
      while (id.length < 6) id += Math.floor(Math.random() * 36).toString(36);
      if (!taken || !taken.has(id)) { if (taken) taken.add(id); return id; }
    }
  }
  /** Give every chapter row a persistent .id (and every chapter sheet the column). Duplicates, which Excel copy-paste
   *  produces, are replaced from the second occurrence on. Called by the writers, so the ids also sit in the model. */
  function ensureRowIds(doc) {
    const taken = new Set();
    const valid = id => /^[a-z0-9]{1,16}$/.test(id);
    for (const s of doc.sheets) if (s.kind === 'chapter') {
      addColumnTo(s, IDENT);
      for (const r of s.rows) { const id = String(r[IDENT] || '').trim(); if (valid(id) && !taken.has(id)) { taken.add(id); r[IDENT] = id; } else r[IDENT] = ''; }
    }
    for (const s of doc.sheets) if (s.kind === 'chapter') for (const r of s.rows) if (!r[IDENT]) r[IDENT] = newRowId(taken);
    return doc;
  }
  function emptyRow(columns, kind = 'p', indent = 0) {
    const r = {};
    for (const c of columns) r[c] = '';
    r['.kind'] = kind;
    r['.indent'] = indent ? String(indent) : '';
    r._id = newId();
    return r;
  }

  /** "# text" → h1 … "#### text" → h4. Returns {kind, text} or null. */
  function detectKindPrefix(text) {
    const m = /^(#{1,4})[ \t]+([\s\S]*)$/.exec(text);
    if (m) return { kind: 'h' + m[1].length, text: m[2] };
    return null;
  }
  /** Leading indent trigger (default three spaces) → {text} or null. */
  function detectIndentPrefix(text, trigger) {
    if (!trigger || !text.startsWith(trigger)) return null;
    return { text: text.slice(trigger.length) };
  }

  /** Timestamp for the "updated" meta column: local "YYYY-MM-DD HH:MM". */
  function stamp(d = new Date()) {
    const p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }
  function isMeta(doc, col) {
    return (col === '.updated' && !!doc.settings.trackUpdated) || (col === '.author' && !!doc.settings.trackAuthor);
  }
  function isComputed(doc, col) { return !!doc.settings.trackCounts && COMPUTED.includes(col); }
  const isSystem = (doc, col) => RESERVED.includes(col) || col === IDENT || isMeta(doc, col) || isComputed(doc, col);
  /** Mark a row as edited: maintains updated/author when tracking is on. */
  function touch(doc, sheet, row) {
    if (doc.settings.trackUpdated) { if (!sheet.columns.includes('.updated')) addColumnTo(sheet, '.updated'); row['.updated'] = stamp(); }
    if (doc.settings.trackAuthor) { if (!sheet.columns.includes('.author')) addColumnTo(sheet, '.author'); row['.author'] = doc.settings.author || ''; }
  }
  function addColumnTo(sheet, name, at) {
    if (sheet.columns.includes(name)) return;
    if (at == null) at = sheet.columns.length;
    sheet.columns.splice(at, 0, name);
    sheet.rows.forEach(r => { if (r[name] == null) r[name] = ''; });
  }
  /** Ensure meta and computed columns exist on every chapter sheet when the settings ask for them. */
  function ensureMetaColumns(doc) {
    for (const s of doc.sheets) {
      if (s.kind !== 'chapter') continue;
      if (doc.settings.trackCounts) { addColumnTo(s, '.words'); addColumnTo(s, '.chars'); }
      if (doc.settings.trackUpdated) addColumnTo(s, '.updated');
      if (doc.settings.trackAuthor) addColumnTo(s, '.author');
    }
  }

  /** End (exclusive) of the block a row owns: a heading owns rows up to the next heading of the
   *  same or higher level; a body row owns deeper-indented rows, and a paragraph also owns the
   *  "s" rows (same indent) that continue it. */
  function sectionEnd(rows, i) {
    const r = rows[i];
    if (!r) return i + 1;
    const level = SheetNumbering.headingLevel(r['.kind']);
    const n = rows.length;
    let j = i + 1;
    if (level) {
      while (j < n) { const l = SheetNumbering.headingLevel(rows[j]['.kind']); if (l && l <= level) break; j++; }
      return j;
    }
    const k = indentOf(r), kind = normKind(r['.kind']);
    while (j < n) {
      const rj = rows[j];
      if (SheetNumbering.headingLevel(rj['.kind'])) break;
      const kj = indentOf(rj);
      if (kj > k) { j++; continue; }
      if (kj === k && normKind(rj['.kind']) === 's' && kind !== 's') { j++; continue; }
      break;
    }
    return j;
  }
  /** Section ends for every row at once, near-linear: if i owns j it owns everything j owns, so jump over j's block. */
  function sectionEnds(rows) {
    const n = rows.length;
    const end = new Array(n);
    const lvl = rows.map(r => SheetNumbering.headingLevel(r['.kind']));
    const ind = rows.map(r => indentOf(r));
    const knd = rows.map(r => normKind(r['.kind']));
    const owns = (i, j) => {
      if (lvl[i]) return !lvl[j] || lvl[j] > lvl[i];
      if (lvl[j]) return false;
      if (ind[j] > ind[i]) return true;
      return ind[j] === ind[i] && knd[j] === 's' && knd[i] !== 's';
    };
    for (let i = n - 1; i >= 0; i--) {
      let j = i + 1;
      while (j < n && owns(i, j)) j = end[j];
      end[i] = j;
    }
    return end;
  }
  function isCollapsible(rows, i) { return sectionEnd(rows, i) > i + 1; }
  /** [start, end) of the unit that moves with row i: always its whole block. */
  function blockOf(rows, i) { return [i, sectionEnd(rows, i)]; }
  /** Move `len` rows starting at `from` so they end up before the row originally at index `to`. Returns new start. */
  function moveBlock(doc, si, from, len, to) {
    const rows = doc.sheets[si].rows;
    if (to >= from && to <= from + len) return from;
    const block = rows.splice(from, len);
    const t = to > from ? to - len : to;
    rows.splice(t, 0, ...block);
    return t;
  }
  /** Move the section owned by row i (a heading with its rows) before row `to` of sheet ti, possibly another sheet.
   *  Returns {si, i} of the moved heading, or null when nothing moved. */
  function moveSection(doc, si, i, ti, to) {
    const rows = doc.sheets[si].rows;
    const [start, end] = blockOf(rows, i);
    if (ti === si) {
      if (to >= start && to <= end) return null;
      return { si, i: moveBlock(doc, si, start, end - start, to) };
    }
    const idxs = []; for (let k = start; k < end; k++) idxs.push(k);
    const at = moveRowsToSheet(doc, si, idxs, ti, to);
    return at == null ? null : { si: ti, i: at };
  }
  /** Promote (dir -1) or demote (dir +1) every heading in the section owned by heading i, keeping their relative levels.
   *  Refused (false) when a heading would leave h1..h4. */
  function shiftSectionLevels(doc, si, i, dir) {
    const s = doc.sheets[si], rows = s.rows;
    const [start, end] = blockOf(rows, i);
    const heads = [];
    for (let k = start; k < end; k++) { const hl = SheetNumbering.headingLevel(rows[k]['.kind']); if (hl) heads.push([k, hl]); }
    if (!heads.length || heads.some(([, hl]) => hl + dir < 1 || hl + dir > 4)) return false;
    for (const [k, hl] of heads) { rows[k]['.kind'] = 'h' + (hl + dir); touch(doc, s, rows[k]); }
    return true;
  }
  /** Sibling-aware targets. Returns the `to` index for moveBlock, or null. */
  function siblingMoveTarget(doc, si, i, dir) {
    const rows = doc.sheets[si].rows;
    const levels = numbering(doc, si).levels;
    const [start, end] = blockOf(rows, i);
    const my = levels[start];
    if (dir < 0) {
      for (let j = start - 1; j >= 0; j--) if (levels[j] <= my) return j; // previous sibling or parent
      return null;
    }
    for (let j = end; j < rows.length; j++) {
      if (levels[j] < my) return j + 1;                 // leaving the parent: become first child of the next parent
      if (levels[j] === my) return sectionEnd(rows, j); // past the next sibling
    }
    return null;
  }

  /** Grid column width in pixels: the user's setting, else a default by column type. */
  function columnWidth(doc, col) {
    const w = doc.settings.widths && doc.settings.widths[col];
    if (w > 0) return w;
    if (col === '.no') return 64;
    if (col === '.kind') return 56;
    if (col === '.indent') return 48;
    if (col === doc.mainColumn) return 480;
    if (isComputed(doc, col)) return 60;
    if (isMeta(doc, col)) return 130;
    if (col === IDENT) return 80;
    return 200;
  }
  /** Display state saved with the workbook: which view, sheet and row were open, what the Draft view shows. */
  function defaultView() {
    return { mode: 'draft', sheet: '', row: 0, toc: 'off', side: true, hidden: [], counts: true, gridHidden: ['.words', '.chars', '.updated', '.author', '.id'] };
  }
  function newChapter(name, columns = DEFAULT_COLUMNS) {
    const cols = ensureReserved(columns.slice());
    return { name, kind: 'chapter', columns: cols, rows: [emptyRow(cols)] };
  }
  function ensureReserved(columns) {
    if (!columns.includes('.indent')) { const k = columns.indexOf('.kind'); columns.splice(k >= 0 ? k + 1 : 0, 0, '.indent'); }
    if (!columns.includes('.kind')) columns.unshift('.kind');
    if (!columns.includes('.no')) columns.unshift('.no');
    return columns;
  }
  function newDoc() {
    return {
      version: 1,
      mainColumn: 'text',
      settings: { numbering: 'continuous', title: '', author: '', description: '', created: new Date().toISOString(), trackUpdated: false, trackAuthor: false, trackCounts: true, freezeColumns: 1, countColumns: [], widths: {}, wordTarget: 0, roles: { status: '', target: '' }, protectHeaders: true, contentsSheet: true, exportPresets: {}, rowDefaults: {}, view: defaultView(), extra: {} },
      sheets: [newChapter('Chapter 1')],
    };
  }

  // ---- queries ----
  function chapterSheets(doc) { return doc.sheets.filter(s => s.kind === 'chapter'); }
  function chapterIndex(doc, si) {
    let n = 0;
    for (let i = 0; i <= si; i++) if (doc.sheets[i].kind === 'chapter') n++;
    return n;
  }
  /** Numbering of sheet si. With settings.numbering === 'continuous' (default) counters carry on
   *  from the previous chapter sheets; with 'per-sheet' every sheet starts at 1. */
  function numbering(doc, si) {
    const s = doc.sheets[si];
    if (!s || s.kind !== 'chapter') return { numbers: [], warnings: [], levels: [], indents: [] };
    let state = null;
    if (doc.settings.numbering !== 'per-sheet') {
      for (let k = 0; k < si; k++) {
        const sk = doc.sheets[k];
        if (sk.kind === 'chapter') state = SheetNumbering.compute(sk.rows, { state }).state;
      }
    }
    return SheetNumbering.compute(s.rows, { state });
  }
  /** Columns counted in the status bar: the chosen ones that exist in this sheet, else the main column. */
  function countColumns(doc, sheet) {
    const chosen = (doc.settings.countColumns || []).filter(c => sheet.columns.includes(c));
    return chosen.length ? chosen : [doc.mainColumn];
  }
  function rowCounts(doc, sheet, row) {
    let words = 0, chars = 0;
    for (const c of countColumns(doc, sheet)) { words += wordCount(row[c]); chars += charCount(row[c]); }
    return { words, chars };
  }
  /** Totals of the block each row owns (heading section, paragraph group, indented group), including
   *  the row itself and excluding "x" rows. null for rows that own nothing. */
  function sectionCounts(doc, sheet) {
    const rows = sheet.rows, n = rows.length;
    const pw = [0], pc = [0], pr = [0];
    for (let i = 0; i < n; i++) {
      const x = normKind(rows[i]['.kind']) === 'x';
      const rc = x ? { words: 0, chars: 0 } : rowCounts(doc, sheet, rows[i]);
      pw.push(pw[i] + rc.words); pc.push(pc[i] + rc.chars); pr.push(pr[i] + (x ? 0 : 1));
    }
    const ends = sectionEnds(rows);
    return rows.map((r, i) => {
      const end = ends[i];
      if (end <= i + 1) return null;
      return { words: pw[end] - pw[i], chars: pc[end] - pc[i], rows: pr[end] - pr[i] };
    });
  }
  // Column roles are explicit settings (roles.status, roles.target); names only serve as suggestions.
  const ROLE_NAMES = { status: ['status', 'tag', 'tags', 'state', 'tila'], target: ['target', 'words_target', 'word_target', 'tavoite'] };
  function roleColumn(doc, sheet, role) {
    const c = doc.settings.roles && doc.settings.roles[role];
    return c && sheet.columns.includes(c) ? c : null;
  }
  /** A column whose name suggests the role, for Settings to propose. */
  function suggestRole(doc, sheet, role) {
    const names = ROLE_NAMES[role] || [];
    return userColumns(doc, sheet).find(c => names.includes(c.toLowerCase())) || null;
  }
  /** Column used for status chips and colouring. */
  function statusColumn(doc, sheet) { return roleColumn(doc, sheet, 'status'); }
  /** Column holding per-section word targets (a number on a heading row). */
  function targetColumn(doc, sheet) { return roleColumn(doc, sheet, 'target'); }
  function rowTarget(doc, sheet, row) {
    const c = targetColumn(doc, sheet); if (!c) return 0;
    const n = parseInt(String(row[c] || '').replace(/\s/g, ''), 10);
    return n > 0 ? n : 0;
  }
  /** Stable pastel colour index (0..7) for a status value. */
  function colorIndex(value) {
    let h = 0; for (const ch of String(value || '').toLowerCase()) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
    return h % 8;
  }
  function userColumns(doc, sheet) { return sheet.columns.filter(c => !isSystem(doc, c)); }
  /** Table of contents: one group per chapter sheet (or only sheet `only`), with its heading rows. */
  function tocEntries(doc, only) {
    const out = [];
    doc.sheets.forEach((s, si) => {
      if (s.kind !== 'chapter' || (only != null && si !== only)) return;
      const num = numbering(doc, si);
      const sec = sectionCounts(doc, s);
      const entries = [];
      s.rows.forEach((r, i) => {
        if (!SheetNumbering.headingLevel(r['.kind'])) return;
        const text = String(r[doc.mainColumn] || '').replace(/^#{1,6}[ \t]+/, '').trim();
        entries.push({ si, i, level: num.levels[i], number: num.numbers[i], text: text || '(untitled)', words: sec[i] ? sec[i].words : rowCounts(doc, s, r).words, target: rowTarget(doc, s, r) });
      });
      out.push({ si, name: s.name, entries });
    });
    return out;
  }
  /** Index of the heading whose section contains row i (i itself if it is a heading), or null. */
  function headingFor(rows, i) {
    for (let j = i; j >= 0; j--) {
      if (SheetNumbering.headingLevel(rows[j]['.kind']) && (j === i || sectionEnd(rows, j) > i)) return j;
    }
    return null;
  }
  function sideColumns(doc, sheet) { return userColumns(doc, sheet).filter(c => c !== doc.mainColumn); }
  function rowIsEmpty(doc, sheet, row) { return userColumns(doc, sheet).every(c => !String(row[c] || '').trim()); }
  function wordCount(text) {
    const t = String(text || '').trim();
    return t ? t.split(/\s+/).length : 0;
  }
  function charCount(text) { return String(text || '').length; }
  /** Counts over the counted columns, excluding "x" rows: {rows, words, chars} */
  function sheetCounts(doc, sheet) {
    const c = { rows: 0, words: 0, chars: 0 };
    if (sheet.kind !== 'chapter') return c;
    for (const r of sheet.rows) {
      if (normKind(r['.kind']) === 'x') continue;
      const rc = rowCounts(doc, sheet, r);
      c.rows++; c.words += rc.words; c.chars += rc.chars;
    }
    return c;
  }
  function docCounts(doc) {
    const t = { rows: 0, words: 0, chars: 0 };
    for (const s of doc.sheets) { const c = sheetCounts(doc, s); t.rows += c.rows; t.words += c.words; t.chars += c.chars; }
    return t;
  }
  function sheetWords(doc, sheet) { return sheetCounts(doc, sheet).words; }
  function docWords(doc) { return docCounts(doc).words; }
  /** File-name-safe version of a title. */
  function safeFileName(title) {
    return String(title || '').normalize('NFKD').replace(/[̀-ͯ]/g, '').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 80);
  }

  // ---- row mutators (si = sheet index) ----
  /** "column=value; column=value" ↔ object. Used for row defaults in Settings and in the .sheetwriter sheet. */
  function parsePairs(text) {
    const out = {};
    for (const part of String(text || '').split(';')) { const m = /^\s*([^=]+?)\s*=\s*(.*?)\s*$/.exec(part); if (m && m[1]) out[m[1]] = m[2]; }
    return out;
  }
  function pairsText(obj) { return Object.entries(obj || {}).map(([k, v]) => `${k}=${v}`).join('; '); }
  /** Fill the side columns of a row the user just created from settings.rowDefaults (never the main or system columns, never over existing text). */
  function applyRowDefaults(doc, sheet, row) {
    for (const [c, v] of Object.entries(doc.settings.rowDefaults || {})) {
      if (c === doc.mainColumn || isSystem(doc, c) || !sheet.columns.includes(c) || String(row[c] || '').trim()) continue;
      row[c] = v;
    }
    return row;
  }
  function addRow(doc, si, at, kind = 'p', indent = 0) {
    const s = doc.sheets[si];
    at = Math.max(0, Math.min(at, s.rows.length));
    const r = applyRowDefaults(doc, s, emptyRow(s.columns, kind, indent));
    s.rows.splice(at, 0, r);
    touch(doc, s, r);
    return at;
  }
  function deleteRow(doc, si, i) {
    const s = doc.sheets[si];
    s.rows.splice(i, 1);
    if (!s.rows.length) s.rows.push(emptyRow(s.columns));
  }
  function moveRow(doc, si, from, to) {
    const s = doc.sheets[si];
    if (from === to || from < 0 || from >= s.rows.length) return from;
    to = Math.max(0, Math.min(to, s.rows.length - 1));
    const [r] = s.rows.splice(from, 1);
    s.rows.splice(to, 0, r);
    return to;
  }
  function duplicateRow(doc, si, i) {
    const s = doc.sheets[si];
    const r = { ...s.rows[i], _id: newId(), [IDENT]: '' };
    s.rows.splice(i + 1, 0, r);
    touch(doc, s, r);
    return i + 1;
  }
  function splitRow(doc, si, i, caret) {
    const s = doc.sheets[si], main = doc.mainColumn;
    const row = s.rows[i];
    const text = row[main] || '';
    const a = text.slice(0, caret), b = text.slice(caret);
    row[main] = a.replace(/\s+$/, '');
    const nr = applyRowDefaults(doc, s, emptyRow(s.columns, isHeading(row['.kind']) ? 'p' : row['.kind'], isHeading(row['.kind']) ? 0 : indentOf(row)));
    nr[main] = b.replace(/^\s+/, '');
    s.rows.splice(i + 1, 0, nr);
    touch(doc, s, row); touch(doc, s, nr);
    return i + 1;
  }
  function mergeRow(doc, si, i) {
    const s = doc.sheets[si], main = doc.mainColumn;
    if (i + 1 >= s.rows.length) return null;
    const a = s.rows[i], b = s.rows[i + 1];
    const caret = (a[main] || '').length;
    for (const c of s.columns) {
      if (isSystem(doc, c)) continue;
      const av = a[c] || '', bv = b[c] || '';
      if (!bv) continue;
      a[c] = av ? (c === main ? av + ' ' + bv : av + '\n' + bv) : bv;
    }
    s.rows.splice(i + 1, 1);
    touch(doc, s, a);
    return caret;
  }
  function setCell(doc, si, i, col, val) {
    const s = doc.sheets[si];
    if (!s.rows[i]) return;
    const row = s.rows[i];
    const v = col === '.kind' ? normKind(val) : String(val);
    if (row[col] === v) return;
    row[col] = v;
    if (col === '.kind' && isHeading(v)) row['.indent'] = '';
    touch(doc, s, row);
  }
  function setIndent(doc, si, i, indent) {
    const s = doc.sheets[si]; const row = s.rows[i];
    if (!row || isHeading(row['.kind'])) return false;
    indent = Math.max(0, Math.min(8, indent | 0));
    const v = indent ? String(indent) : '';
    if (row['.indent'] === v) return false;
    row['.indent'] = v;
    touch(doc, s, row);
    return true;
  }
  /** Indent/outdent a row together with its block (children move with it). */
  function shiftIndent(doc, si, i, dir) {
    const s = doc.sheets[si]; const rows = s.rows;
    const row = rows[i];
    if (!row || isHeading(row['.kind'])) return false;
    const cur = indentOf(row);
    if (dir < 0 && cur === 0) return false;
    const [start, end] = blockOf(rows, i);
    for (let k = start; k < end; k++) {
      if (isHeading(rows[k]['.kind'])) continue;
      const v = Math.max(0, indentOf(rows[k]) + dir);
      rows[k]['.indent'] = v ? String(v) : '';
    }
    touch(doc, s, row);
    return true;
  }
  function cycleKind(doc, si, i, dir = 1) {
    const s = doc.sheets[si]; const row = s.rows[i];
    const order = ['p', 'h1', 'h2', 'h3', 'h4', 's', 'x'];
    const k = order.indexOf(normKind(row['.kind']));
    row['.kind'] = order[(k + dir + order.length) % order.length];
    if (isHeading(row['.kind'])) row['.indent'] = '';
    touch(doc, s, row);
  }
  /** dir -1 = promote (towards h1), +1 = demote (towards s) */
  function shiftKind(doc, si, i, dir) {
    const s = doc.sheets[si]; const row = s.rows[i];
    const cur = normKind(row['.kind']);
    let k = KIND_ORDER.indexOf(cur);
    if (k < 0) k = KIND_ORDER.indexOf('p') - dir; // x → p either way
    k = Math.max(0, Math.min(KIND_ORDER.length - 1, k + dir));
    row['.kind'] = KIND_ORDER[k];
    if (isHeading(row['.kind'])) row['.indent'] = '';
    touch(doc, s, row);
  }

  // ---- column mutators ----
  function validColumnName(sheet, name, allowExisting) {
    name = String(name || '').trim();
    if (!name) return { ok: false, reason: 'Column name is empty.' };
    const low = name.toLowerCase();
    if (name.startsWith('.') || name.startsWith('_')) return { ok: false, reason: 'Names starting with a dot are SheetWriter\'s own columns, and names starting with an underscore are not allowed.' };
    if (!allowExisting && sheet.columns.some(c => c.toLowerCase() === name.toLowerCase())) return { ok: false, reason: `Column "${name}" already exists.` };
    return { ok: true, name };
  }
  // Column operations are workbook-wide: sheet si is the reference whose order the other chapter sheets follow.
  function syncColumnOrder(doc, ref) {
    for (const s of doc.sheets) {
      if (s.kind !== 'chapter' || s === ref) continue;
      const shared = ref.columns.filter(c => s.columns.includes(c));
      const rest = s.columns.filter(c => !ref.columns.includes(c));
      s.columns = shared.concat(rest);
    }
  }
  function addColumn(doc, si, name, at) {
    const ref = doc.sheets[si];
    addColumnTo(ref, name, at);
    for (const s of chapterSheets(doc)) if (s !== ref) addColumnTo(s, name);
    syncColumnOrder(doc, ref);
  }
  function renameColumn(doc, si, oldName, newName) {
    if (oldName === newName) return;
    for (const s of chapterSheets(doc)) {
      const k = s.columns.indexOf(oldName);
      if (k < 0) continue;
      s.columns[k] = newName;
      s.rows.forEach(r => { r[newName] = r[oldName] || ''; delete r[oldName]; });
    }
    if (oldName === doc.mainColumn) setMainColumn(doc, newName, oldName);
  }
  function deleteColumn(doc, si, name) {
    if (RESERVED.includes(name) || name === IDENT || name === doc.mainColumn) return;
    for (const s of chapterSheets(doc)) {
      s.columns = s.columns.filter(c => c !== name);
      s.rows.forEach(r => { delete r[name]; });
    }
  }
  function moveColumn(doc, si, name, dir) {
    const s = doc.sheets[si];
    const k = s.columns.indexOf(name);
    const j = k + dir;
    if (k < 0 || j < 0 || j >= s.columns.length) return;
    [s.columns[k], s.columns[j]] = [s.columns[j], s.columns[k]];
    syncColumnOrder(doc, s);
  }
  /** Move column `name` before column `before` (null = to the end). */
  function moveColumnBefore(doc, si, name, before) {
    const s = doc.sheets[si];
    if (!s.columns.includes(name) || name === before) return;
    s.columns = s.columns.filter(c => c !== name);
    const k = before ? s.columns.indexOf(before) : -1;
    if (k < 0) s.columns.push(name); else s.columns.splice(k, 0, name);
    syncColumnOrder(doc, s);
  }
  /** Where a column holds data: [{sheet, count}] over all chapter sheets. */
  function columnData(doc, name) {
    const out = [];
    for (const s of chapterSheets(doc)) {
      const n = s.rows.filter(r => String(r[name] || '').trim()).length;
      if (n) out.push({ sheet: s.name, count: n });
    }
    return out;
  }
  /** Change the main column name across all chapter sheets. */
  function setMainColumn(doc, newName, oldName = doc.mainColumn) {
    doc.mainColumn = newName;
    for (const s of doc.sheets) {
      if (s.kind !== 'chapter') continue;
      if (s.columns.includes(newName)) continue;
      const k = s.columns.indexOf(oldName);
      if (k >= 0) {
        s.columns[k] = newName;
        s.rows.forEach(r => { r[newName] = r[oldName] || ''; delete r[oldName]; });
      } else {
        s.columns.push(newName);
        s.rows.forEach(r => { r[newName] = ''; });
      }
    }
  }

  // ---- sheet mutators ----
  function addSheet(doc, name, at, columns) {
    if (at == null) at = doc.sheets.length;
    const template = columns || (doc.sheets.find(s => s.kind === 'chapter') || {}).columns || DEFAULT_COLUMNS;
    const cols = template.slice();
    if (!cols.includes(doc.mainColumn)) cols.push(doc.mainColumn);
    const s = newChapter(uniqueSheetName(doc, name), cols);
    doc.sheets.splice(at, 0, s);
    return at;
  }
  function uniqueSheetName(doc, name) {
    name = String(name || 'Chapter').trim() || 'Chapter';
    let n = name, k = 2;
    while (doc.sheets.some(s => s.name.toLowerCase() === n.toLowerCase())) n = `${name} ${k++}`;
    return n;
  }
  function renameSheet(doc, si, name) {
    name = String(name || '').trim();
    if (!name) return false;
    if (doc.sheets.some((s, i) => i !== si && s.name.toLowerCase() === name.toLowerCase())) return false;
    doc.sheets[si].name = name;
    return true;
  }
  function deleteSheet(doc, si) {
    doc.sheets.splice(si, 1);
    if (!doc.sheets.length) doc.sheets.push(newChapter('Chapter 1'));
  }
  function moveSheet(doc, from, to) {
    if (from === to) return from;
    to = Math.max(0, Math.min(to, doc.sheets.length - 1));
    const [s] = doc.sheets.splice(from, 1);
    doc.sheets.splice(to, 0, s);
    return to;
  }
  // ---- find and replace ----
  /** All occurrences of q: [{si, i, col, index, len}] in document order. columns: null = every user column. */
  function findMatches(doc, q, { matchCase = false, scope = 'all', si = 0, columns = null } = {}) {
    if (!q) return [];
    const needle = matchCase ? q : q.toLowerCase();
    const out = [];
    doc.sheets.forEach((s, k) => {
      if (s.kind !== 'chapter' || (scope !== 'all' && k !== si)) return;
      const cols = columns ? columns.filter(c => s.columns.includes(c)) : userColumns(doc, s);
      s.rows.forEach((r, i) => {
        for (const c of cols) {
          const v = String(r[c] || '');
          const hay = matchCase ? v : v.toLowerCase();
          let from = 0, idx;
          while ((idx = hay.indexOf(needle, from)) >= 0) { out.push({ si: k, i, col: c, index: idx, len: q.length }); from = idx + Math.max(1, q.length); }
        }
      });
    });
    return out;
  }
  /** Replace the given matches (from findMatches on the same document state). Returns the count. */
  function replaceMatches(doc, matches, replacement) {
    const byCell = new Map();
    for (const m of matches) { const key = `${m.si} ${m.i} ${m.col}`; if (!byCell.has(key)) byCell.set(key, []); byCell.get(key).push(m); }
    let n = 0;
    for (const [key, ms] of byCell) {
      const [si, i, col] = key.split(' ');
      const s = doc.sheets[+si]; const r = s && s.rows[+i];
      if (!r) continue;
      let v = String(r[col] || '');
      ms.sort((a, b) => b.index - a.index);
      for (const m of ms) { v = v.slice(0, m.index) + replacement + v.slice(m.index + m.len); n++; }
      setCell(doc, +si, +i, col, v);
    }
    return n;
  }

  // ---- multi-row operations (selection clipboard) ----
  /** Plain copies of rows without private fields, safe to keep in a clipboard. */
  function cloneRows(rows) {
    return rows.map(r => { const c = {}; for (const [k, v] of Object.entries(r)) if (!k.startsWith('_') && k !== IDENT) c[k] = v; return c; });
  }
  function deleteRows(doc, si, idxs) {
    const s = doc.sheets[si];
    const set = new Set(idxs);
    s.rows = s.rows.filter((r, i) => !set.has(i));
    if (!s.rows.length) s.rows.push(emptyRow(s.columns));
  }
  /** Insert plain row objects at `at` (default: end); missing columns with data are added workbook-wide. Returns at. */
  function insertRows(doc, si, plain, at) {
    const s = doc.sheets[si];
    if (!s || s.kind !== 'chapter') return null;
    const needed = new Set();
    plain.forEach(r => Object.keys(r).forEach(c => { if (!c.startsWith('_') && !RESERVED.includes(c) && !s.columns.includes(c) && String(r[c] || '').trim()) needed.add(c); }));
    for (const c of needed) addColumn(doc, si, c);
    const made = plain.map(r => {
      const row = emptyRow(s.columns, normKind(r['.kind']), parseInt(r['.indent'], 10) || 0);
      for (const c of s.columns) if (!RESERVED.includes(c) && c !== IDENT && r[c] != null) row[c] = String(r[c]);
      touch(doc, s, row);
      return row;
    });
    if (s.rows.length === 1 && rowIsEmpty(doc, s, s.rows[0])) { s.rows.length = 0; at = 0; }
    if (at == null || at > s.rows.length) at = s.rows.length;
    s.rows.splice(at, 0, ...made);
    return at;
  }
  /** Move the rows at `idxs` (any order, possibly non-contiguous) so they sit, in order, before the row originally at `to`. Returns new start. */
  function moveRows(doc, si, idxs, to) {
    const s = doc.sheets[si];
    const set = new Set(idxs);
    const moving = [], rest = [];
    let t = to;
    s.rows.forEach((r, i) => { if (set.has(i)) { moving.push(r); if (i < to) t--; } else rest.push(r); });
    t = Math.max(0, Math.min(t, rest.length));
    rest.splice(t, 0, ...moving);
    s.rows = rest;
    return t;
  }
  /** Insert imported rows ({kind, indent, text, side}) into sheet si at `at` (default: end). Creates side columns as needed.
   *  With replace=true the sheet's rows are replaced. Returns the index of the first inserted row. */
  function importRows(doc, si, imported, { at, replace } = {}) {
    const s = doc.sheets[si];
    if (!s || s.kind !== 'chapter') return null;
    const cols = new Set();
    imported.forEach(r => Object.keys(r.side || {}).forEach(c => cols.add(c)));
    for (const c of cols) if (!s.columns.some(x => x.toLowerCase() === c.toLowerCase())) addColumn(doc, si, c);
    const colFor = c => s.columns.find(x => x.toLowerCase() === c.toLowerCase());
    const rows = imported.map(r => {
      const row = emptyRow(s.columns, normKind(r.kind), isHeading(r.kind) ? 0 : (parseInt(r.indent, 10) || 0));
      row[doc.mainColumn] = r.text || '';
      for (const [c, v] of Object.entries(r.side || {})) row[colFor(c)] = v;
      touch(doc, s, row);
      return row;
    });
    if (replace) { s.rows = rows.length ? rows : [emptyRow(s.columns)]; return 0; }
    if (s.rows.length === 1 && rowIsEmpty(doc, s, s.rows[0])) { s.rows.length = 0; at = 0; }
    if (at == null || at > s.rows.length) at = s.rows.length;
    s.rows.splice(at, 0, ...rows);
    return at;
  }
  /** The section owned by heading i becomes a new sheet right after si, named after the heading. Returns the new sheet's index. */
  function sectionToNewSheet(doc, si, i) {
    const s = doc.sheets[si];
    const [start, end] = blockOf(s.rows, i);
    const name = String(s.rows[i][doc.mainColumn] || '').replace(/^#{1,6}[ \t]+/, '').trim().slice(0, 31) || 'Section';
    const at = addSheet(doc, name, si + 1);
    const idxs = []; for (let k = start; k < end; k++) idxs.push(k);
    moveRowsToSheet(doc, si, idxs, at);
    return at;
  }
  /** Append every row of chapter si to chapter ti and delete si. Returns ti's index afterwards, or null. */
  function mergeSheetInto(doc, si, ti) {
    const src = doc.sheets[si], dst = doc.sheets[ti];
    if (!src || !dst || si === ti || src.kind !== 'chapter' || dst.kind !== 'chapter') return null;
    const idxs = src.rows.map((r, k) => k).filter(k => !rowIsEmpty(doc, src, src.rows[k]));
    if (idxs.length) moveRowsToSheet(doc, si, idxs, ti);
    deleteSheet(doc, si);
    return ti > si ? ti - 1 : ti;
  }
  /** Move rows (by index) from sheet si to the end (or `at`) of sheet ti. Returns new index of first moved row. */
  function moveRowsToSheet(doc, si, idxs, ti, at) {
    const src = doc.sheets[si], dst = doc.sheets[ti];
    if (!src || !dst || src === dst || dst.kind !== 'chapter') return null;
    idxs = [...new Set(idxs)].sort((a, b) => a - b);
    const moved = idxs.map(i => src.rows[i]);
    for (let k = idxs.length - 1; k >= 0; k--) src.rows.splice(idxs[k], 1);
    if (!src.rows.length) src.rows.push(emptyRow(src.columns));
    for (const r of moved) for (const c of Object.keys(r)) if (!c.startsWith('_') && !dst.columns.includes(c) && r[c]) addColumn(doc, ti, c);
    for (const r of moved) for (const c of dst.columns) if (r[c] == null) r[c] = '';
    if (at == null) at = dst.rows.length;
    if (dst.rows.length === 1 && rowIsEmpty(doc, dst, dst.rows[0])) { dst.rows.length = 0; at = 0; }
    dst.rows.splice(at, 0, ...moved);
    return at;
  }

  // ---- comparing two documents ----
  const words = t => String(t || '').toLowerCase().split(/\W+/).filter(Boolean);
  function similarity(a, b) {
    const A = new Set(words(a)), B = new Set(words(b));
    if (!A.size && !B.size) return 1;
    let n = 0; for (const w of A) if (B.has(w)) n++;
    return n / (A.size + B.size - n);
  }
  /** Word-level diff of two strings: [{op: '=', '-', '+', text}], LCS on words. */
  function wordDiff(x, y) {
    const a = String(x || '').split(/\s+/).filter(Boolean), b = String(y || '').split(/\s+/).filter(Boolean);
    const n = a.length, m = b.length;
    if (n * m > 250000) return [{ op: '-', text: a.join(' ') }, { op: '+', text: b.join(' ') }];
    const L = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = a[i] === b[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    const out = []; const push = (op, t) => { const last = out[out.length - 1]; if (last && last.op === op) last.text += ' ' + t; else out.push({ op, text: t }); };
    let i = 0, j = 0;
    while (i < n && j < m) { if (a[i] === b[j]) { push('=', a[i]); i++; j++; } else if (L[i + 1][j] >= L[i][j + 1]) push('-', a[i++]); else push('+', b[j++]); }
    while (i < n) push('-', a[i++]);
    while (j < m) push('+', b[j++]);
    return out;
  }
  /** What changed from document a (older) to b (newer). Rows are matched by .id, then by identical text in the same sheet,
   *  then by word overlap. Returns {entries: [{type: 'added'|'removed'|'changed', sheet, si, i, number, before, after, cols, kindBefore, kindAfter}],
   *  same, sheetsAdded, sheetsRemoved}. si/i point into b for added and changed rows, into a for removed rows. */
  function diffDocs(a, b) {
    const main = b.mainColumn, mainA = a.mainColumn || main;
    const textOf = (doc, r) => String(r[doc.mainColumn] || '').trim();
    const rowsOf = doc => doc.sheets.flatMap((s, si) => s.kind === 'chapter' ? s.rows.map((r, i) => ({ doc, s, si, i, r, empty: rowIsEmpty(doc, s, r) })).filter(x => !x.empty) : []);
    const A = rowsOf(a), B = rowsOf(b);
    const pairs = [], usedA = new Set();
    const byId = new Map(); A.forEach(x => { const id = String(x.r[IDENT] || '').trim(); if (id && !byId.has(id)) byId.set(id, x); });
    const restB = [];
    for (const y of B) { const id = String(y.r[IDENT] || '').trim(); const x = id && byId.get(id); if (x && !usedA.has(x)) { pairs.push([x, y]); usedA.add(x); } else restB.push(y); }
    const byText = new Map(); A.forEach(x => { if (usedA.has(x)) return; const k = x.s.name + '\u0000' + textOf(a, x.r); if (!byText.has(k)) byText.set(k, []); byText.get(k).push(x); });
    const restB2 = [];
    for (const y of restB) { const l = byText.get(y.s.name + '\u0000' + textOf(b, y.r)); const x = l && l.find(z => !usedA.has(z)); if (x) { pairs.push([x, y]); usedA.add(x); } else restB2.push(y); }
    const restB3 = [];
    for (const y of restB2) {
      let best = null, bs = 0.5;
      for (const x of A) { if (usedA.has(x) || x.s.name !== y.s.name) continue; const sc = similarity(textOf(a, x.r), textOf(b, y.r)); if (sc > bs) { bs = sc; best = x; } }
      if (best) { pairs.push([best, y]); usedA.add(best); } else restB3.push(y);
    }
    const entries = []; let same = 0;
    const numB = new Map(), numA = new Map();
    const numberOf = (doc, cache, si, i) => { if (!cache.has(si)) cache.set(si, numbering(doc, si).numbers); return cache.get(si)[i]; };
    for (const [x, y] of pairs) {
      const cols = new Set([...userColumns(a, x.s), ...userColumns(b, y.s)]);
      const changed = [];
      for (const c of cols) { const ca = c === mainA ? main : c; if (String(x.r[c] ?? '').trim() !== String(y.r[ca] ?? '').trim()) changed.push(ca); }
      const kindA = normKind(x.r['.kind']), kindB = normKind(y.r['.kind']);
      if (kindA !== kindB || indentOf(x.r) !== indentOf(y.r)) changed.push('.kind');
      if (!changed.length) { same++; continue; }
      entries.push({ type: 'changed', sheet: y.s.name, si: y.si, i: y.i, number: numberOf(b, numB, y.si, y.i), before: textOf(a, x.r), after: textOf(b, y.r), cols: changed, kindBefore: kindA, kindAfter: kindB, rowBefore: x.r, rowAfter: y.r });
    }
    for (const y of restB3) entries.push({ type: 'added', sheet: y.s.name, si: y.si, i: y.i, number: numberOf(b, numB, y.si, y.i), before: '', after: textOf(b, y.r), cols: [], kindAfter: normKind(y.r['.kind']), rowAfter: y.r });
    for (const x of A) if (!usedA.has(x)) entries.push({ type: 'removed', sheet: x.s.name, si: x.si, i: x.i, number: numberOf(a, numA, x.si, x.i), before: textOf(a, x.r), after: '', cols: [], kindBefore: normKind(x.r['.kind']), rowBefore: x.r });
    const order = new Map(b.sheets.map((s, k) => [s.name, k]));
    entries.sort((p, q) => ((order.has(p.sheet) ? order.get(p.sheet) : 1e9) - (order.has(q.sheet) ? order.get(q.sheet) : 1e9)) || (p.type === 'removed') - (q.type === 'removed') || p.i - q.i);
    const namesA = new Set(a.sheets.map(s => s.name)), namesB = new Set(b.sheets.map(s => s.name));
    return { entries, same, sheetsAdded: [...namesB].filter(n => !namesA.has(n)), sheetsRemoved: [...namesA].filter(n => !namesB.has(n)) };
  }

  return {
    RESERVED, META, COMPUTED, IDENT, KINDS, DEFAULT_COLUMNS, normKind, isHeading, indentOf, emptyRow, newChapter, newDoc, ensureIds, newId, ensureRowIds, newRowId,
    detectKindPrefix, detectIndentPrefix, stamp, isMeta, isComputed, isSystem, touch, ensureMetaColumns,
    sectionEnd, sectionEnds, isCollapsible, blockOf, moveBlock, moveSection, shiftSectionLevels, siblingMoveTarget, statusColumn, targetColumn, rowTarget, roleColumn, suggestRole, colorIndex,
    chapterSheets, chapterIndex, numbering, countColumns, rowCounts, sectionCounts, userColumns, sideColumns, rowIsEmpty, tocEntries, headingFor,
    wordCount, charCount, sheetCounts, docCounts, sheetWords, docWords, safeFileName,
    addRow, deleteRow, moveRow, duplicateRow, splitRow, mergeRow, setCell, setIndent, shiftIndent, cycleKind, shiftKind,
    validColumnName, addColumn, renameColumn, deleteColumn, moveColumn, moveColumnBefore, columnData, syncColumnOrder, setMainColumn, defaultView, columnWidth,
    addSheet, renameSheet, deleteSheet, moveSheet, moveRowsToSheet, uniqueSheetName, importRows, sectionToNewSheet, mergeSheetInto, parsePairs, pairsText, applyRowDefaults,
    cloneRows, deleteRows, insertRows, moveRows, findMatches, replaceMatches, diffDocs, wordDiff, similarity,
  };
})();
if (typeof module !== 'undefined') module.exports = Model;

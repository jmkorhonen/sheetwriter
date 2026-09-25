/* SheetWriter — XLSX load/save with ExcelJS.
 *
 * Chapter sheets are rewritten from the model with a house style.
 * Data sheets (no main-text column) are copied through from the loaded workbook.
 * Settings and metadata live in a visible, protected sheet named ".sheetwriter" (same dot as the system columns).
 */
const XlsxIO = (() => {
  const SETTINGS_SHEET = '.sheetwriter';
  const LEGACY_SETTINGS_SHEET = '_sheetwriter'; // written before 0.10.1; read, never written
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const KNOWN = {
    title: 'Document title',
    author: 'Author',
    description: 'Short description',
    main_column: 'Column that holds the text you are writing. A sheet with this column is a chapter.',
    numbering: 'continuous: numbering carries on from sheet to sheet (h1 is always one number, 1, 2, 3 …); per-sheet: every sheet starts at 1',
    freeze_columns: 'How many leading columns stay frozen in the Grid view and in Excel',
    count_columns: 'Columns whose words and characters the status bar counts (empty = the main column)',
    column_widths: 'Grid column widths in pixels, name:px pairs; also used for the Excel column widths',
    word_target: 'Target for the whole workbook, in words or characters as target_unit says (0 = none); per-section targets go in the target column on heading rows',
    target_unit: 'words or chars: what the workbook target and the section targets count',
    row_defaults: 'Values given to the side columns of rows added in SheetWriter, as column=value pairs separated by semicolons (e.g. status=todo)',
    protect_headers: 'yes/no: protect the header row of chapter sheets in Excel (data cells stay editable; Review → Unprotect Sheet to rename columns)',
    contents_sheet: 'yes/no: write a .contents sheet listing every heading with a link to it (rewritten on every save, not shown in SheetWriter)',
    status_column: 'Column whose values show as coloured chips (any name; set in Settings → Columns)',
    target_column: 'Column holding per-section targets (words or characters, see target_unit) on heading rows (any name; set in Settings → Counts and targets)',
    track_updated: 'yes/no: keep an ".updated" column with the time each row was last edited in SheetWriter',
    track_author: 'yes/no: keep an ".author" column with the author who last edited each row in SheetWriter',
    track_counts: 'yes/no: keep ".words" and ".chars" columns with per-row counts over the counted columns (written on save, recomputed on load)',
    view_mode: 'Display state when last saved: draft, grid or read',
    view_sheet: 'Sheet that was open when last saved',
    view_row: 'Row (1-based, within that sheet) that was being edited when last saved',
    view_toc: 'Table of contents: off, sheet or all',
    view_side: 'yes/no: show side columns in Draft view',
    view_hidden_columns: 'Side columns hidden in Draft view (comma-separated)',
    view_counts: 'yes/no: show word and character counts under each card in Draft view',
    view_grid_hidden_columns: 'Columns hidden in Grid view (comma-separated)',
    created: 'First saved (ISO date)',
    modified: 'Last saved by SheetWriter (ISO date)',
    app: 'Editor that wrote this workbook',
    app_url: 'Open the editor online',
    download: 'Download the latest editor as a single HTML file (right-click, Save link as…)',
    repo: 'Source code',
    readme: 'Documentation',
  };
  const LINK_KEYS = ['app_url', 'download', 'repo', 'readme'];
  const PRESET_DESC = 'Export preset (Export dialog → Preset). column, scope (all/sheet), numbering (none/headings/all), indented (paragraphs/lists), side (column or empty), side_as (quote/comment/footnote), titles (yes/no)';
  const PRESET_KEYS = { column: 'column', scope: 'scope', numbering: 'numbering', indented: 'indented', side: 'side', side_as: 'sideMode', titles: 'titles' };
  function presetToText(p) {
    return Object.entries(PRESET_KEYS).map(([k, f]) => `${k}=${f === 'titles' ? (p.titles ? 'yes' : 'no') : (p[f] == null ? '' : String(p[f]))}`).join('; ');
  }
  function presetFromText(text) {
    const p = { column: '', scope: 'all', numbering: '', indented: 'paragraphs', side: '', sideMode: 'quote', titles: false };
    for (const part of String(text).split(';')) {
      const m = /^\s*([a-z_]+)\s*=\s*(.*?)\s*$/.exec(part);
      if (!m || !PRESET_KEYS[m[1]]) continue;
      const f = PRESET_KEYS[m[1]];
      if (f === 'titles') p.titles = yes(m[2]); else p[f] = m[2];
    }
    if (!['all', 'sheet'].includes(p.scope)) p.scope = 'all';
    if (!['', 'none', 'headings', 'all'].includes(p.numbering)) p.numbering = ''; if (p.numbering === 'none') p.numbering = '';
    if (!['paragraphs', 'lists'].includes(p.indented)) p.indented = 'paragraphs';
    if (!['quote', 'comment', 'footnote'].includes(p.sideMode)) p.sideMode = 'quote';
    return p;
  }

  /* Instructions written into the .sheetwriter sheet for people who open the workbook in Excel.
   * RELEASE CHECKLIST: revise these lines whenever the file format changes, a system column is added,
   * or what is safe to edit in Excel changes. tests.html checks that every system column is mentioned. */
  function excelNotes(doc) {
    const main = doc.mainColumn;
    const sys = [...Model.RESERVED, ...Model.COMPUTED, ...Model.META, Model.IDENT].join(', ');
    return [
      `HOW TO WORK WITH THIS WORKBOOK IN EXCEL (written by ${APP.name} ${APP.version}; this block is rewritten on every save)`,
      'SAFE TO DO IN EXCEL:',
      `• Edit any text in your own columns (${main}, notes, sources, …). Cells are Markdown: **bold**, *italic*, [link](https://…).`,
      '• Add rows anywhere. Leave .no empty and they go to the end of the chapter as paragraphs, or type a .no such as 2.1 to place them.',
      '• Sort or filter rows with the drop-downs in the header row: the .no column restores the order when the file is opened. Reorder rows by editing .no. (Data → Sort with the header row selected is blocked while the header is protected; sort the data rows only, or use the drop-downs.)',
      '• Change .kind (h1, h2, h3, h4, p, s, x) and .indent (0, 1, 2 …).',
      '• Add columns with any name that does not start with a dot or an underscore. Rename your own columns (then reassign roles in SheetWriter Settings if needed).',
      '• Add key/value rows to this sheet: they are kept. Edit the values of the settings rows above.',
      `• Add sheets without a ${main} column (data sheets): they are copied through unchanged, formatting included.`,
      `• Colour cells, change fonts, add borders and comments: on chapter sheets they follow the row through the hidden ${Model.IDENT} column (unhide it if you are curious; leave its values alone, copied rows get a fresh id).`,
      'LOST ON THE NEXT SAVE FROM SHEETWRITER:',
      '• Formulas, number formats, row heights and merged cells in chapter sheets (data sheets are kept as they are). Bold, italic and font size follow the row kind. Hyperlinks survive as Markdown links.',
      `• ${Model.COMPUTED.join(', ')}, ${Model.META.join(', ')}: rewritten from SheetWriter’s own data.`,
      'BREAKS THE FILE OR ITS STRUCTURE:',
      `• Renaming or deleting the dotted columns (${sys}) or the ${main} column, or giving two columns the same name. The header row is protected for this reason (Review → Unprotect Sheet lifts it).`,
      '• Renaming this sheet, or changing the key column of the settings rows above.',
      '• Merged cells in chapter sheets.',
    ];
  }

  function cellText(v) {
    if (v == null) return '';
    if (v instanceof Date) return isNaN(v) ? '' : v.toISOString().slice(0, 10);
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
      if (v.hyperlink !== undefined) {
        // Excel hyperlinks become Markdown links so the URL survives the round trip (Excel then shows the Markdown).
        const t = cellText(v.text), url = String(v.hyperlink).replace(/^mailto:/i, m => m);
        return t && t !== url ? `[${t}](${url})` : url;
      }
      if (v.formula !== undefined || v.sharedFormula !== undefined) return cellText(v.result);
      if (v.error) return String(v.error);
      return String(v);
    }
    return String(v);
  }
  const yes = v => !/^(no|false|0|off|)$/i.test(String(v || '').trim());

  function safeSheetName(name) {
    let n = String(name || 'Sheet').replace(/[\[\]:*?\/\\]/g, ' ').trim();
    if (n.startsWith("'")) n = n.slice(1);
    if (n.endsWith("'")) n = n.slice(0, -1);
    if (n.startsWith('.')) n = n.slice(1); // dotted sheet names are SheetWriter's own
    return (n || 'Sheet').slice(0, 31);
  }

  function readSettings(ws, doc) {
    const extra = {};
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const key = cellText(row.getCell(1).value).trim();
      if (!key) continue;
      const value = cellText(row.getCell(2).value);
      const desc = cellText(row.getCell(3).value);
      switch (key) {
        case 'title': doc.settings.title = value; break;
        case 'author': doc.settings.author = value; break;
        case 'description': doc.settings.description = value; break;
        case 'main_column': if (value.trim()) doc.mainColumn = value.trim(); break;
        case 'chapter_prefix': doc.settings.numbering = yes(value) ? 'continuous' : 'per-sheet'; break; // files from 0.3.0
        case 'numbering': doc.settings.numbering = value.trim() === 'per-sheet' ? 'per-sheet' : 'continuous'; break;
        case 'freeze_columns': { const n = parseInt(value, 10); doc.settings.freezeColumns = n >= 0 ? Math.min(n, 10) : 1; break; }
        case 'count_columns': doc.settings.countColumns = value.split(/[,;]/).map(s => s.trim()).filter(Boolean); break;
        case 'word_target': { const n = parseInt(value.replace(/\s/g, ''), 10); doc.settings.wordTarget = n > 0 ? n : 0; break; }
        case 'row_defaults': doc.settings.rowDefaults = Model.parsePairs(value); break;
        case 'status_column': doc.settings.roles.status = value.trim(); break;
        case 'protect_headers': doc.settings.protectHeaders = yes(value); break;
        case 'contents_sheet': doc.settings.contentsSheet = yes(value); break;
        case 'target_unit': doc.settings.targetUnit = /^c/i.test(value.trim()) ? 'chars' : 'words'; break;
        case 'target_column': doc.settings.roles.target = value.trim(); break;
        case 'column_widths': {
          const w = {};
          value.split(/[,;]/).forEach(p => { const m = /^\s*(.+?)\s*:\s*(\d+)\s*$/.exec(p); if (m && +m[2] > 0) w[m[1]] = +m[2]; });
          doc.settings.widths = w; break;
        }
        case 'track_updated': doc.settings.trackUpdated = yes(value); break;
        case 'track_author': doc.settings.trackAuthor = yes(value); break;
        case 'track_counts': doc.settings.trackCounts = yes(value); break;
        case 'view_mode': if (['draft', 'grid', 'read'].includes(value.trim())) doc.settings.view.mode = value.trim(); break;
        case 'view_sheet': doc.settings.view.sheet = value.trim(); break;
        case 'view_row': { const n = parseInt(value, 10); doc.settings.view.row = n > 0 ? n - 1 : 0; break; }
        case 'view_toc': doc.settings.view.toc = ['sheet', 'all'].includes(value.trim()) ? value.trim() : 'off'; break;
        case 'view_side': doc.settings.view.side = yes(value); break;
        case 'view_hidden_columns': doc.settings.view.hidden = value.split(/[,;]/).map(s => s.trim()).filter(Boolean); break;
        case 'view_counts': doc.settings.view.counts = yes(value); break;
        case 'view_grid_hidden_columns': doc.settings.view.gridHidden = value.split(/[,;]/).map(s => s.trim()).filter(Boolean); break;
        case 'created': if (value) doc.settings.created = value; break;
        case 'modified': case 'app': break;
        default:
          if (LINK_KEYS.includes(key)) break;
          if (key.startsWith('export:') && key.length > 7) { doc.settings.exportPresets[key.slice(7).trim()] = presetFromText(value); break; }
          extra[key] = [value, desc];
      }
    }
    doc.settings.extra = extra;
  }

  /** Rows of the settings sheet, format-neutral: [{key, value, desc, link, bold, note}]. Shared by the XLSX and ODS writers. */
  function settingsRows(doc) {
    const now = new Date().toISOString();
    const view = Object.assign(Model.defaultView(), doc.settings.view || {});
    const rows = [
      ['title', doc.settings.title || ''],
      ['author', doc.settings.author || ''],
      ['description', doc.settings.description || ''],
      ['main_column', doc.mainColumn],
      ['numbering', doc.settings.numbering === 'per-sheet' ? 'per-sheet' : 'continuous'],
      ['freeze_columns', String(doc.settings.freezeColumns ?? 1)],
      ['count_columns', (doc.settings.countColumns || []).join(', ')],
      ['column_widths', Object.entries(doc.settings.widths || {}).map(([k, v]) => `${k}:${v}`).join(', ')],
      ['word_target', String(doc.settings.wordTarget || 0)],
      ['target_unit', doc.settings.targetUnit === 'chars' ? 'chars' : 'words'],
      ['row_defaults', Model.pairsText(doc.settings.rowDefaults)],
      ['protect_headers', doc.settings.protectHeaders === false ? 'no' : 'yes'],
      ['contents_sheet', doc.settings.contentsSheet === false ? 'no' : 'yes'],
      ['status_column', (doc.settings.roles && doc.settings.roles.status) || ''],
      ['target_column', (doc.settings.roles && doc.settings.roles.target) || ''],
      ['track_updated', doc.settings.trackUpdated ? 'yes' : 'no'],
      ['track_author', doc.settings.trackAuthor ? 'yes' : 'no'],
      ['track_counts', doc.settings.trackCounts ? 'yes' : 'no'],
      ['view_mode', view.mode || 'draft'],
      ['view_sheet', view.sheet || ''],
      ['view_row', String((view.row || 0) + 1)],
      ['view_toc', view.toc || 'off'],
      ['view_side', view.side === false ? 'no' : 'yes'],
      ['view_hidden_columns', (view.hidden || []).join(', ')],
      ['view_counts', view.counts === false ? 'no' : 'yes'],
      ['view_grid_hidden_columns', (view.gridHidden || []).join(', ')],
      ['created', doc.settings.created || now],
      ['modified', now],
      ['app', `${APP.name} ${APP.version}`],
      ['app_url', APP.site],
      ['download', APP.download],
      ['repo', APP.repo],
      ['readme', APP.readme],
    ];
    for (const [name, p] of Object.entries(doc.settings.exportPresets || {})) rows.push(['export:' + name, presetToText(p)]);
    const out = rows.map(r => ({ key: r[0], value: r[1], desc: KNOWN[r[0]] || (r[0].startsWith('export:') ? PRESET_DESC : ''), link: LINK_KEYS.includes(r[0]) }));
    for (const [k, v] of Object.entries(doc.settings.extra || {})) out.push({ key: k, value: v[0] || '', desc: v[1] || '' });
    out.push({ key: '', value: '', desc: '', note: true });
    excelNotes(doc).forEach((line, k) => out.push({ key: '', value: line, desc: '', note: true, bold: k === 0 || /^[A-Z ]+:$/.test(line) }));
    return out;
  }

  async function writeSettings(wb, doc) {
    const ws = wb.addWorksheet(SETTINGS_SHEET);
    ws.columns = [
      { header: 'key', key: 'key', width: 16 },
      { header: 'value', key: 'value', width: 56 },
      { header: 'description', key: 'description', width: 70 },
    ];
    for (const r of settingsRows(doc)) {
      const row = ws.addRow([r.key, r.value, r.desc]);
      if (r.link) { row.getCell(2).value = { text: r.value, hyperlink: r.value }; row.getCell(2).font = { color: { argb: 'FF2F6FDB' }, underline: true }; }
      if (r.bold) row.getCell(2).font = { bold: true };
    }
    ws.getRow(1).font = { bold: true };
    ws.eachRow(row => row.eachCell(c => { c.alignment = { wrapText: true, vertical: 'top' }; }));
    // Protect against accidental edits in Excel. No password: "Unprotect Sheet" in Excel is one click.
    try { await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true }); } catch (e) { /* optional */ }
  }

  function dedupeColumns(headers) {
    const seen = new Map();
    return headers.map(h => {
      const key = h.toLowerCase();
      const n = seen.get(key) || 0;
      seen.set(key, n + 1);
      return n ? `${h}_${n + 1}` : h;
    });
  }

  function orderedSheets(wb) {
    return wb.worksheets.slice().sort((a, b) => (a.orderNo ?? 0) - (b.orderNo ?? 0));
  }

  /** Rows edited or sorted in Excel: restore the order given by the "no" column; rows without one go last. */
  function sortByNo(rows) {
    const parse = v => { const s = String(v || '').trim(); return /^\d+(\.\d+)*$/.test(s) ? s.split('.').map(Number) : null; };
    const withNo = [], without = [];
    rows.forEach((r, k) => { const p = parse(r['.no']); (p ? withNo : without).push({ r, k, p }); });
    if (!withNo.length) return rows;
    withNo.sort((a, b) => {
      const n = Math.max(a.p.length, b.p.length);
      for (let i = 0; i < n; i++) { const x = a.p[i] ?? -1, y = b.p[i] ?? -1; if (x !== y) return x - y; }
      return a.k - b.k;
    });
    return withNo.map(x => x.r).concat(without.map(x => x.r));
  }

  function readHeaders(ws) {
    const headers = [];
    const hr = ws.getRow(1);
    const colCount = Math.max(ws.columnCount || 0, hr.cellCount || 0);
    for (let c = 1; c <= colCount; c++) headers[c - 1] = cellText(hr.getCell(c).value).trim();
    while (headers.length && !headers[headers.length - 1]) headers.pop();
    return headers;
  }

  function chapterFromSheet(ws, headers, doc) {
    const mainLower = doc.mainColumn.toLowerCase();
    for (let i = 0; i < headers.length; i++) {
      if (!headers[i]) headers[i] = 'col' + (i + 1);
      const low = headers[i].toLowerCase();
      const sys = [...Model.RESERVED, ...Model.META, ...Model.COMPUTED, Model.IDENT];
      const legacy = [...Model.RESERVED, ...Model.META, ...Model.COMPUTED]; // bare names written before 0.10; a plain "id" column stays the user's
      if (sys.includes(low)) headers[i] = low; // .kind etc.
      else if (legacy.includes('.' + low) && !headers.some(h => h.toLowerCase() === '.' + low)) headers[i] = '.' + low; // files written before 0.10: kind -> .kind
      if (low === mainLower) headers[i] = doc.mainColumn;
    }
    const columns = dedupeColumns(headers);
    let rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const obj = {};
      let any = false;
      columns.forEach((c, i) => {
        const t = cellText(row.getCell(i + 1).value);
        obj[c] = t;
        if (c !== '.no' && t.trim()) any = true;
      });
      if (any) rows.push(obj);
    }
    if (columns.includes('.no')) rows = sortByNo(rows);
    if (!columns.includes('.kind')) columns.unshift('.kind');
    if (!columns.includes('.no')) columns.unshift('.no');
    if (!columns.includes('.indent')) columns.splice(columns.indexOf('.kind') + 1, 0, '.indent');
    rows.forEach(r => {
      r['.kind'] = Model.normKind(r['.kind']); r['.no'] = '';
      const ind = parseInt(r['.indent'], 10); r['.indent'] = ind > 0 && !Model.isHeading(r['.kind']) ? String(ind) : '';
      for (const c of columns) if (r[c] == null) r[c] = '';
    });
    if (!rows.length) rows.push(Model.emptyRow(columns));
    return { name: ws.name, kind: 'chapter', columns, rows };
  }

  /** buffer: ArrayBuffer → {doc, sources: Map<name, Worksheet>, warnings: string[]} */
  async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const res = loadFromSheets(orderedSheets(wb), name => wb.getWorksheet(name));
    res.formats = readFormats(wb, res.doc);
    return res;
  }
  const APP_GREY = 'FF8A8A8A';
  /** Formatting people add in Excel on chapter sheets, keyed by row id then column name: {fill, border, font: {name, color, strike, underline}, note}.
   *  Bold, italic and size are the app's (they follow the row kind) and are not kept. */
  function readFormats(wb, doc) {
    const formats = new Map();
    for (const s of doc.sheets) {
      if (s.kind !== 'chapter' || !s.columns.includes(Model.IDENT)) continue;
      const ws = wb.getWorksheet(s.name);
      if (!ws) continue;
      const names = readHeaders(ws);
      const idCol = names.findIndex(h => h.toLowerCase() === Model.IDENT) + 1;
      if (!idCol) continue;
      const lowerCols = new Map(s.columns.map(c => [c.toLowerCase(), c]));
      ws.eachRow((row, rn) => {
        if (rn === 1) return;
        const id = cellText(row.getCell(idCol).value).trim();
        if (!id || formats.has(id)) return;
        const cells = {};
        row.eachCell({ includeEmpty: false }, (cell, cn) => {
          const col = lowerCols.get((names[cn - 1] || '').toLowerCase());
          if (!col || col === Model.IDENT) return;
          const f = {};
          const fill = cell.fill;
          if (fill && fill.type === 'pattern' && fill.pattern && fill.pattern !== 'none' && fill.fgColor) f.fill = fill;
          const b = cell.border;
          if (b && ['top', 'left', 'bottom', 'right'].some(k => b[k] && b[k].style)) f.border = b;
          const fo = cell.font || {};
          const font = {};
          if (fo.name && fo.name !== 'Calibri') font.name = fo.name; // Calibri: Excel's default, written explicitly by Excel on every styled cell
          if (fo.color && fo.color.argb !== APP_GREY && fo.color.theme !== 1) font.color = fo.color; // theme 1: default text colour
          if (fo.strike) font.strike = true;
          if (fo.underline) font.underline = fo.underline;
          if (Object.keys(font).length) f.font = font;
          if (cell.note) f.note = cell.note;
          if (Object.keys(f).length) cells[col] = f;
        });
        if (Object.keys(cells).length) formats.set(id, { cells });
      });
    }
    return formats;
  }
  /** Build a document from worksheet-like objects ({name, rowCount, columnCount, getRow(r).getCell(c).value}),
   *  in workbook order. Used for XLSX (ExcelJS worksheets) and ODS (adapters). Only ExcelJS sheets can be
   *  copied through with formatting; others fall back to their cell values. */
  function loadFromSheets(allSheets, getWorksheet) {
    const doc = Model.newDoc();
    doc.sheets = [];
    const sources = new Map();
    const warnings = [];

    const sws = getWorksheet(SETTINGS_SHEET) || getWorksheet(LEGACY_SETTINGS_SHEET);
    if (sws) readSettings(sws, doc);
    let mainCased = !!sws; // without a settings sheet, adopt the casing used in the file
    // Sheets whose names start with a dot are SheetWriter's own (.sheetwriter, .contents) and are regenerated on save.
    const sheets = allSheets.filter(ws => !ws.name.startsWith('.') && ws.name !== LEGACY_SETTINGS_SHEET);

    for (const ws of sheets) {
      const headers = readHeaders(ws);
      const mainIdx = headers.findIndex(h => h.toLowerCase() === doc.mainColumn.toLowerCase());
      if (mainIdx >= 0) {
        if (!mainCased) { doc.mainColumn = headers[mainIdx]; mainCased = true; }
        doc.sheets.push(chapterFromSheet(ws, headers, doc));
      } else {
        const cells = [];
        const cc = ws.columnCount || 0;
        for (let r = 1; r <= ws.rowCount; r++) {
          const row = ws.getRow(r);
          const arr = [];
          for (let c = 1; c <= cc; c++) arr.push(cellText(row.getCell(c).value));
          cells.push(arr);
        }
        doc.sheets.push({ name: ws.name, kind: 'data', cells });
        if (typeof ws.eachRow === 'function') sources.set(ws.name, ws);
      }
    }

    if (!doc.sheets.some(s => s.kind === 'chapter') && !sws) {
      // Plain spreadsheet without a "text" column: take the first sheet with a header row and
      // treat its longest text column as the main column.
      const first = sheets.find(ws => readHeaders(ws).some(Boolean) && ws.rowCount > 1);
      if (first) {
        const headers = readHeaders(first);
        const avg = headers.map((h, i) => {
          if (!h || Model.RESERVED.includes(h.toLowerCase())) return -1;
          let n = 0, len = 0;
          for (let r = 2; r <= first.rowCount; r++) { const t = cellText(first.getRow(r).getCell(i + 1).value); if (t.trim()) { n++; len += t.length; } }
          return n ? len / n : 0;
        });
        const best = avg.indexOf(Math.max(...avg));
        if (best >= 0 && avg[best] > 0) {
          doc.mainColumn = headers[best];
          const k = doc.sheets.findIndex(s => s.name === first.name);
          doc.sheets[k] = chapterFromSheet(first, headers, doc);
          sources.delete(first.name);
          warnings.push(`This workbook has no "text" column. Sheet "${first.name}" was imported as a chapter using column "${headers[best]}" as the text. Change the main column in Settings if that is wrong.`);
        }
      }
    }
    if (!doc.sheets.length) doc.sheets.push(Model.newChapter('Chapter 1'));
    if (!doc.sheets.some(s => s.kind === 'chapter')) {
      warnings.push(`No sheet has a "${doc.mainColumn}" column, so there is nothing to draft in. A new chapter was added.`);
      doc.sheets.push(Model.newChapter('Chapter 1'));
    }
    Model.ensureMetaColumns(doc);
    // Duplicate row ids (a row copied in Excel): the first in file order is the original and keeps the id; the rest get a fresh one on save.
    const seen = new Set();
    for (const s of doc.sheets) if (s.kind === 'chapter' && s.columns.includes(Model.IDENT)) for (const r of s.rows) {
      const id = String(r[Model.IDENT] || '').trim();
      if (id && seen.has(id)) r[Model.IDENT] = ''; else if (id) seen.add(id);
    }
    return { doc, sources, warnings, hasSettings: !!sws };
  }

  function widthFor(col, doc) {
    const px = doc.settings.widths && doc.settings.widths[col];
    if (px > 0) return Math.max(4, Math.round(px / 7)); // Excel width unit ≈ 7 px at the default font
    if (col === '.no') return 8;
    if (col === '.kind') return 6;
    if (col === '.indent') return 7;
    if (col === '.updated') return 17;
    if (col === '.author' && doc.settings.trackAuthor) return 16;
    if (Model.isComputed(doc, col)) return 7;
    if (col === doc.mainColumn) return 80;
    return 32;
  }

  function copySheet(src, s, ws) {
    if (src) {
      try { (src.columns || []).forEach((col, i) => { if (col && col.width) ws.getColumn(i + 1).width = col.width; }); } catch (e) { /* no column defs */ }
      src.eachRow({ includeEmpty: true }, (row, rn) => {
        const dst = ws.getRow(rn);
        if (row.height) dst.height = row.height;
        row.eachCell({ includeEmpty: true }, (cell, cn) => {
          const d = dst.getCell(cn);
          d.value = cell.value;
          try { if (cell.style && Object.keys(cell.style).length) d.style = cell.style; } catch (e) { /* ignore */ }
        });
      });
      try { ((src.model && src.model.merges) || []).forEach(m => ws.mergeCells(m)); } catch (e) { /* ignore */ }
      if (src.state && src.state !== 'visible') ws.state = src.state;
      if (src.views && src.views.length) ws.views = src.views;
    } else if (s.cells) {
      s.cells.forEach(r => ws.addRow(r));
    }
  }

  /** Column order in the file: as in the sheet, then computed counts (words, chars), then meta (updated, author). */
  function fileColumns(doc, s) {
    const computed = s.columns.filter(c => Model.isComputed(doc, c));
    const meta = s.columns.filter(c => Model.isMeta(doc, c));
    const ident = s.columns.filter(c => c === Model.IDENT);
    return s.columns.filter(c => !Model.isMeta(doc, c) && !Model.isComputed(doc, c) && c !== Model.IDENT).concat(computed, meta, ident);
  }

  /** → ArrayBuffer-like (Uint8Array/Buffer) suitable for new Blob([...]). formats: from load(), applied back by row id. */
  async function save(doc, sources = new Map(), formats = new Map()) {
    Model.ensureRowIds(doc);
    const wb = new ExcelJS.Workbook();
    wb.creator = doc.settings.author || APP.name;
    wb.lastModifiedBy = APP.name;
    wb.created = doc.settings.created ? new Date(doc.settings.created) : new Date();
    wb.modified = new Date();
    if (doc.settings.title) wb.title = doc.settings.title;

    const usedNames = new Set([SETTINGS_SHEET.toLowerCase(), LEGACY_SETTINGS_SHEET.toLowerCase(), '.contents']);
    const sheetNames = new Map(); // model sheet → name written to the file
    for (const s of doc.sheets) {
      let name = safeSheetName(s.name), base = name, k = 2;
      while (usedNames.has(name.toLowerCase())) name = (base.slice(0, 28) + ' ' + k++);
      usedNames.add(name.toLowerCase());
      sheetNames.set(s, name);
      const ws = wb.addWorksheet(name);
      if (s.kind !== 'chapter') { copySheet(sources.get(s.name), s, ws); continue; }

      const si = doc.sheets.indexOf(s);
      const num = Model.numbering(doc, si);
      const columns = fileColumns(doc, s);
      ws.columns = columns.map(c => ({ header: c, key: c, width: widthFor(c, doc) }));
      const rows = s.rows.filter(r => !Model.rowIsEmpty(doc, s, r));
      if (!rows.length) rows.push(s.rows[0]);
      rows.forEach(r => {
        const i = s.rows.indexOf(r);
        const vals = {};
        const rc = doc.settings.trackCounts ? Model.rowCounts(doc, s, r) : null;
        for (const c of columns) {
          if (c === '.no') vals[c] = num.numbers[i];
          else if (c === '.indent') { const n = Model.indentOf(r); vals[c] = n ? n : ''; }
          else if (rc && c === '.words') vals[c] = rc.words;
          else if (rc && c === '.chars') vals[c] = rc.chars;
          else vals[c] = String(r[c] ?? '');
        }
        const row = ws.addRow(vals);
        const hl = SheetNumbering.headingLevel(r['.kind']);
        row.eachCell({ includeEmpty: true }, cell => { cell.alignment = { wrapText: true, vertical: 'top' }; });
        if (hl) row.font = { bold: true, size: Math.max(11, 15 - hl), italic: hl >= 3 };
        else if (r['.kind'] === 'x') row.font = { italic: true, color: { argb: 'FF8A8A8A' } };
        const ind = num.indents[i] + (r['.kind'] === 's' ? 1 : 0);
        if (!hl && ind) row.getCell(doc.mainColumn).alignment = { wrapText: true, vertical: 'top', indent: ind };
        for (const c of columns) if (Model.isMeta(doc, c) || Model.isComputed(doc, c) || c === Model.IDENT) row.getCell(c).font = { color: { argb: APP_GREY }, size: 9 };
        const kept = formats.get(r[Model.IDENT]);
        if (kept) for (const [c, f] of Object.entries(kept.cells)) {
          if (!columns.includes(c)) continue;
          const cell = row.getCell(c);
          if (f.fill) cell.fill = f.fill;
          if (f.border) cell.border = f.border;
          if (f.font) cell.font = { ...(cell.font || {}), ...f.font };
          if (f.note) cell.note = f.note;
        }
      });
      if (columns.includes(Model.IDENT)) ws.getColumn(columns.indexOf(Model.IDENT) + 1).hidden = true;
      const hdr = ws.getRow(1);
      hdr.font = { bold: true };
      hdr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9E9E9' } };
      columns.forEach((c, k) => { if (c.startsWith('.')) ws.getCell(1, k + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }; });
      const xSplit = Math.max(0, Math.min(doc.settings.freezeColumns ?? 1, columns.length - 1));
      ws.views = [{ state: 'frozen', xSplit, ySplit: 1 }];
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };
      if (doc.settings.protectHeaders !== false) {
        // Lock only the header row: every column (including empty cells below the data) stays editable,
        // rows and columns can be inserted, deleted, sorted and filtered. Renaming a header needs Unprotect Sheet.
        columns.forEach((c, k) => { ws.getColumn(k + 1).protection = { locked: false }; });
        ws.eachRow(row => row.eachCell({ includeEmpty: true }, cell => { cell.protection = { locked: row.number === 1 }; }));
        try {
          await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatCells: true, formatColumns: true, formatRows: true,
            insertColumns: true, insertRows: true, insertHyperlinks: true, deleteColumns: true, deleteRows: true, sort: true, autoFilter: true, pivotTables: true });
        } catch (e) { /* optional */ }
      }
    }
    if (doc.settings.contentsSheet !== false) await writeContents(wb, doc, sheetNames);
    await writeSettings(wb, doc);
    return wb.xlsx.writeBuffer();
  }

  const colLetter = n => { let s = ''; for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s; return s; };
  /** Entries of the ".contents" sheet, format-neutral: [{no, heading, sheet, words, level, col (letter), row (1-based file row) | null}]. */
  function contentsEntries(doc, sheetNames) {
    const out = [];
    for (const g of Model.tocEntries(doc, null)) {
      const s = doc.sheets[g.si];
      const cols = fileColumns(doc, s);
      const col = colLetter(cols.indexOf(doc.mainColumn) + 1);
      const fileName = sheetNames.get(s) || s.name;
      // Row numbers in the file: header + rows that are actually written (empty rows are skipped).
      const written = s.rows.filter(r => !Model.rowIsEmpty(doc, s, r));
      for (const e of g.entries) {
        const idx = written.indexOf(s.rows[e.i]);
        out.push({ no: e.number, heading: e.text, sheet: fileName, words: e.words, level: e.level, col, row: idx >= 0 ? idx + 2 : null });
      }
    }
    return out;
  }
  /** A ".contents" sheet: every heading with its number, section words, and an internal link to the row. */
  async function writeContents(wb, doc, sheetNames) {
    const ws = wb.addWorksheet('.contents');
    ws.columns = [
      { header: 'no', key: 'no', width: 10 }, { header: 'heading', key: 'heading', width: 70 },
      { header: 'sheet', key: 'sheet', width: 24 }, { header: 'words', key: 'words', width: 9 },
    ];
    for (const e of contentsEntries(doc, sheetNames)) {
      const row = ws.addRow({ no: e.no, heading: e.heading, sheet: e.sheet, words: e.words });
      if (e.row) row.getCell('heading').value = { text: e.heading, hyperlink: `#'${e.sheet.replace(/'/g, "''")}'!${e.col}${e.row}` };
      row.getCell('heading').alignment = { indent: Math.max(0, e.level - 1) };
      row.getCell('heading').font = { bold: e.level === 1, color: { argb: 'FF2F6FDB' }, underline: true };
    }
    ws.getRow(1).font = { bold: true };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    try { await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true }); } catch (e) { /* optional */ }
  }

  return { load, loadFromSheets, save, cellText, sortByNo, excelNotes, settingsRows, contentsEntries, presetToText, presetFromText, fileColumns, widthFor, safeSheetName, colLetter, SETTINGS_SHEET, MIME };
})();

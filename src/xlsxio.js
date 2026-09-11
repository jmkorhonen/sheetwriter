/* SheetWriter — XLSX load/save with ExcelJS.
 *
 * Chapter sheets are rewritten from the model with a house style.
 * Data sheets (no main-text column) are copied through from the loaded workbook.
 * Settings and metadata live in a visible, protected sheet named "_sheetwriter".
 */
const XlsxIO = (() => {
  const SETTINGS_SHEET = '_sheetwriter';
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const KNOWN = {
    title: 'Document title',
    author: 'Author',
    description: 'Short description',
    main_column: 'Column that holds the text you are writing. A sheet with this column is a chapter.',
    numbering: 'continuous: numbering carries on from sheet to sheet (h1 is always one number, 1, 2, 3 …); per-sheet: every sheet starts at 1',
    freeze_columns: 'How many leading columns stay frozen in the Grid view and in Excel',
    count_columns: 'Columns whose words and characters the status bar counts (empty = the main column)',
    track_updated: 'yes/no: keep an "updated" column with the time each row was last edited in SheetWriter',
    track_author: 'yes/no: keep an "author" column with the author who last edited each row in SheetWriter',
    created: 'First saved (ISO date)',
    modified: 'Last saved by SheetWriter (ISO date)',
    app: 'Editor that wrote this workbook',
    app_url: 'Open the editor online',
    download: 'Download the latest editor as a single HTML file (right-click, Save link as…)',
    repo: 'Source code',
    readme: 'Documentation',
  };
  const LINK_KEYS = ['app_url', 'download', 'repo', 'readme'];

  function cellText(v) {
    if (v == null) return '';
    if (v instanceof Date) return isNaN(v) ? '' : v.toISOString().slice(0, 10);
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
      if (v.hyperlink !== undefined) return cellText(v.text) || String(v.hyperlink);
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
        case 'track_updated': doc.settings.trackUpdated = yes(value); break;
        case 'track_author': doc.settings.trackAuthor = yes(value); break;
        case 'created': if (value) doc.settings.created = value; break;
        case 'modified': case 'app': break;
        default:
          if (LINK_KEYS.includes(key)) break;
          extra[key] = [value, desc];
      }
    }
    doc.settings.extra = extra;
  }

  async function writeSettings(wb, doc) {
    const ws = wb.addWorksheet(SETTINGS_SHEET);
    ws.columns = [
      { header: 'key', key: 'key', width: 16 },
      { header: 'value', key: 'value', width: 56 },
      { header: 'description', key: 'description', width: 70 },
    ];
    const now = new Date().toISOString();
    const rows = [
      ['title', doc.settings.title || ''],
      ['author', doc.settings.author || ''],
      ['description', doc.settings.description || ''],
      ['main_column', doc.mainColumn],
      ['numbering', doc.settings.numbering === 'per-sheet' ? 'per-sheet' : 'continuous'],
      ['freeze_columns', String(doc.settings.freezeColumns ?? 1)],
      ['count_columns', (doc.settings.countColumns || []).join(', ')],
      ['track_updated', doc.settings.trackUpdated ? 'yes' : 'no'],
      ['track_author', doc.settings.trackAuthor ? 'yes' : 'no'],
      ['created', doc.settings.created || now],
      ['modified', now],
      ['app', `${APP.name} ${APP.version}`],
      ['app_url', APP.site],
      ['download', APP.download],
      ['repo', APP.repo],
      ['readme', APP.readme],
    ];
    for (const r of rows) {
      const row = ws.addRow([r[0], r[1], KNOWN[r[0]] || '']);
      if (LINK_KEYS.includes(r[0])) row.getCell(2).value = { text: r[1], hyperlink: r[1] };
    }
    for (const [k, v] of Object.entries(doc.settings.extra || {})) ws.addRow([k, v[0] || '', v[1] || '']);
    ws.getRow(1).font = { bold: true };
    ws.eachRow(row => row.eachCell(c => { c.alignment = { wrapText: true, vertical: 'top' }; }));
    LINK_KEYS.forEach(() => {});
    ws.eachRow((row, n) => { if (n > 1 && LINK_KEYS.includes(cellText(row.getCell(1).value))) row.getCell(2).font = { color: { argb: 'FF2F6FDB' }, underline: true }; });
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
    rows.forEach((r, k) => { const p = parse(r.no); (p ? withNo : without).push({ r, k, p }); });
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
      if (Model.RESERVED.includes(low) || Model.META.includes(low)) headers[i] = low;
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
        if (c !== 'no' && t.trim()) any = true;
      });
      if (any) rows.push(obj);
    }
    if (columns.includes('no')) rows = sortByNo(rows);
    if (!columns.includes('kind')) columns.unshift('kind');
    if (!columns.includes('no')) columns.unshift('no');
    if (!columns.includes('indent')) columns.splice(columns.indexOf('kind') + 1, 0, 'indent');
    rows.forEach(r => {
      r.kind = Model.normKind(r.kind); r.no = '';
      const ind = parseInt(r.indent, 10); r.indent = ind > 0 && !Model.isHeading(r.kind) ? String(ind) : '';
      for (const c of columns) if (r[c] == null) r[c] = '';
    });
    if (!rows.length) rows.push(Model.emptyRow(columns));
    return { name: ws.name, kind: 'chapter', columns, rows };
  }

  /** buffer: ArrayBuffer → {doc, sources: Map<name, Worksheet>, warnings: string[]} */
  async function load(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const doc = Model.newDoc();
    doc.sheets = [];
    const sources = new Map();
    const warnings = [];

    const sws = wb.getWorksheet(SETTINGS_SHEET);
    if (sws) readSettings(sws, doc);
    let mainCased = !!sws; // without a settings sheet, adopt the casing used in the file
    const sheets = orderedSheets(wb).filter(ws => ws.name !== SETTINGS_SHEET);

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
        sources.set(ws.name, ws);
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
    return { doc, sources, warnings, hasSettings: !!sws };
  }

  function widthFor(col, doc) {
    if (col === 'no') return 8;
    if (col === 'kind') return 6;
    if (col === 'indent') return 7;
    if (col === 'updated') return 17;
    if (col === 'author' && doc.settings.trackAuthor) return 16;
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

  /** Column order in the file: as in the sheet, but meta columns (updated, author) always last. */
  function fileColumns(doc, s) {
    const meta = s.columns.filter(c => Model.isMeta(doc, c));
    return s.columns.filter(c => !Model.isMeta(doc, c)).concat(meta);
  }

  /** → ArrayBuffer-like (Uint8Array/Buffer) suitable for new Blob([...]) */
  async function save(doc, sources = new Map()) {
    const wb = new ExcelJS.Workbook();
    wb.creator = doc.settings.author || APP.name;
    wb.lastModifiedBy = APP.name;
    wb.created = doc.settings.created ? new Date(doc.settings.created) : new Date();
    wb.modified = new Date();
    if (doc.settings.title) wb.title = doc.settings.title;

    const usedNames = new Set([SETTINGS_SHEET.toLowerCase()]);
    for (const s of doc.sheets) {
      let name = safeSheetName(s.name), base = name, k = 2;
      while (usedNames.has(name.toLowerCase())) name = (base.slice(0, 28) + ' ' + k++);
      usedNames.add(name.toLowerCase());
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
        for (const c of columns) {
          if (c === 'no') vals[c] = num.numbers[i];
          else if (c === 'indent') { const n = Model.indentOf(r); vals[c] = n ? n : ''; }
          else vals[c] = String(r[c] ?? '');
        }
        const row = ws.addRow(vals);
        const hl = SheetNumbering.headingLevel(r.kind);
        row.eachCell({ includeEmpty: true }, cell => { cell.alignment = { wrapText: true, vertical: 'top' }; });
        if (hl) row.font = { bold: true, size: Math.max(11, 15 - hl), italic: hl >= 3 };
        else if (r.kind === 'x') row.font = { italic: true, color: { argb: 'FF8A8A8A' } };
        const ind = num.indents[i] + (r.kind === 's' ? 1 : 0);
        if (!hl && ind) row.getCell(doc.mainColumn).alignment = { wrapText: true, vertical: 'top', indent: ind };
        for (const c of columns) if (Model.isMeta(doc, c)) row.getCell(c).font = { color: { argb: 'FF8A8A8A' }, size: 9 };
      });
      const hdr = ws.getRow(1);
      hdr.font = { bold: true };
      hdr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9E9E9' } };
      const xSplit = Math.max(0, Math.min(doc.settings.freezeColumns ?? 1, columns.length - 1));
      ws.views = [{ state: 'frozen', xSplit, ySplit: 1 }];
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };
    }
    await writeSettings(wb, doc);
    return wb.xlsx.writeBuffer();
  }

  return { load, save, cellText, sortByNo, SETTINGS_SHEET, MIME };
})();

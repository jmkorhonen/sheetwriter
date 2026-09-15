/* SheetWriter — OpenDocument: ODT export, ODS save and load. Small OpenDocument writers/reader on JSZip.
 *
 *   Odf.buildOdt(doc, opts) → Blob     text document from Exporter.toBlocks (same options as Word)
 *   Odf.saveOds(doc)        → Blob     spreadsheet with the same layout as the XLSX writer
 *   Odf.loadOds(buffer)     → {doc, sources, warnings, hasSettings}   via XlsxIO.loadFromSheets
 */
const Odf = (() => {
  const ODT_MIME = 'application/vnd.oasis.opendocument.text';
  const ODS_MIME = 'application/vnd.oasis.opendocument.spreadsheet';
  const NS = {
    office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0', style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0', table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', xlink: 'http://www.w3.org/1999/xlink',
    dc: 'http://purl.org/dc/elements/1.1/', meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
    svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
  };
  const NSDECL = Object.entries(NS).filter(([k]) => k !== 'manifest').map(([k, v]) => `xmlns:${k}="${v}"`).join(' ') + ' office:version="1.2"';
  const XML = '<?xml version="1.0" encoding="UTF-8"?>\n';
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  /** Text inside a paragraph, with ODF whitespace rules. */
  const odfText = s => esc(s).replace(/ {2,}/g, m => ' ' + `<text:s text:c="${m.length - 1}"/>`).replace(/\t/g, '<text:tab/>').replace(/\n/g, '<text:line-break/>');

  function manifest(mime) {
    const entry = (p, t) => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="${t}"/>`;
    return XML + `<manifest:manifest xmlns:manifest="${NS.manifest}" manifest:version="1.2">${entry('/', mime)}${entry('content.xml', 'text/xml')}${entry('styles.xml', 'text/xml')}${entry('meta.xml', 'text/xml')}</manifest:manifest>`;
  }
  function metaXml(doc) {
    const now = new Date().toISOString().replace(/\.\d+Z$/, '');
    return XML + `<office:document-meta ${NSDECL}><office:meta><meta:generator>${esc(APP.name + ' ' + APP.version)}</meta:generator><dc:title>${esc(doc.settings.title || '')}</dc:title><dc:description>${esc(doc.settings.description || '')}</dc:description><meta:initial-creator>${esc(doc.settings.author || '')}</meta:initial-creator><dc:creator>${esc(doc.settings.author || '')}</dc:creator><meta:creation-date>${esc((doc.settings.created || now).replace(/\.\d+Z$/, '').replace(/Z$/, ''))}</meta:creation-date><dc:date>${now}</dc:date></office:meta></office:document-meta>`;
  }
  /** Zip with the mimetype entry first and stored, as the ODF packaging spec requires. */
  async function pack(mime, files) {
    const zip = new JSZip();
    zip.file('mimetype', mime, { compression: 'STORE' });
    for (const [name, content] of Object.entries(files)) zip.file(name, content, { compression: 'DEFLATE' });
    return zip.generateAsync({ type: 'blob', mimeType: mime });
  }

  // ---------- ODT ----------
  const HEADINGS = [[1, '18pt', 'bold', ''], [2, '15pt', 'bold', ''], [3, '13pt', 'bold', 'italic'], [4, '12pt', 'normal', 'italic'], [5, '11pt', 'normal', 'italic'], [6, '11pt', 'normal', 'italic']];
  function odtStyles() {
    return XML + `<office:document-styles ${NSDECL}><office:font-face-decls><style:font-face style:name="Liberation Serif" svg:font-family="'Liberation Serif'" style:font-family-generic="roman"/><style:font-face style:name="Liberation Sans" svg:font-family="'Liberation Sans'" style:font-family-generic="swiss"/><style:font-face style:name="Liberation Mono" svg:font-family="'Liberation Mono'" style:font-family-generic="modern"/></office:font-face-decls><office:styles>`
      + `<style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.25cm" fo:line-height="120%"/><style:text-properties style:font-name="Liberation Serif" fo:font-size="11pt" fo:language="en" fo:country="GB"/></style:default-style>`
      + `<style:style style:name="Standard" style:family="paragraph" style:class="text"/>`
      + `<style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.25cm"/></style:style>`
      + `<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text"><style:paragraph-properties fo:margin-top="0.42cm" fo:margin-bottom="0.21cm" fo:keep-with-next="always"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="14pt"/></style:style>`
      + HEADINGS.map(([n, sz, w, st]) => `<style:style style:name="Heading_20_${n}" style:display-name="Heading ${n}" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="${n}" style:class="text"><style:text-properties fo:font-size="${sz}" fo:font-weight="${w}"${st ? ` fo:font-style="${st}"` : ''}/></style:style>`).join('')
      + `<style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:paragraph-properties fo:margin-left="0.6cm" fo:margin-bottom="0.1cm" fo:text-indent="-0.6cm"/><style:text-properties fo:font-size="9pt"/></style:style>`
      + `<style:style style:name="Side_20_Note" style:display-name="Side Note" style:family="paragraph" style:parent-style-name="Text_20_body"><style:paragraph-properties fo:margin-left="1.27cm" fo:margin-bottom="0.35cm"/><style:text-properties fo:font-size="9pt" fo:font-style="italic" fo:color="#666666"/></style:style>`
      + `</office:styles></office:document-styles>`;
  }
  function odtContent(doc, opts) {
    const spanStyles = new Map(); // flags → style name
    const styleFor = r => {
      const key = (r.b ? 'b' : '') + (r.i ? 'i' : '') + (r.strike ? 's' : '') + (r.code ? 'c' : '');
      if (!key) return null;
      if (!spanStyles.has(key)) spanStyles.set(key, 'T_' + key);
      return spanStyles.get(key);
    };
    const blocks = Exporter.toBlocks(doc, opts);
    const fnBlock = blocks.find(b => b.type === 'footnotes');
    const fnText = new Map((fnBlock ? fnBlock.notes : []).map(f => [f.n, f.text]));
    const runs = rs => rs.map(r => {
      if (r.br) return '<text:line-break/>';
      if (r.fn) return `<text:note text:id="ftn${r.fn}" text:note-class="footnote"><text:note-citation>${r.fn}</text:note-citation><text:note-body>${Docx.paragraphs(fnText.get(r.fn) || '').map(p => `<text:p text:style-name="Footnote">${odfText(p.prefix || '')}${runs(p.runs)}</text:p>`).join('')}</text:note-body></text:note>`;
      const st = styleFor(r);
      return st ? `<text:span text:style-name="${st}">${odfText(r.text || '')}</text:span>` : odfText(r.text || '');
    }).join('');
    const body = [];
    let maxIndent = 0;
    for (const b of blocks) {
      if (b.type === 'sheetTitle' || b.type === 'heading') {
        const p = Docx.paragraphs(b.text)[0];
        const lv = Math.min(b.level, 6);
        body.push(`<text:h text:style-name="Heading_20_${lv}" text:outline-level="${lv}">${runs(p.runs.map(r => ({ ...r, b: false })))}</text:h>`);
      } else if (b.type === 'para') {
        Docx.paragraphs(b.text).forEach((p, k) => {
          const ind = (b.indent || 0) + (p.prefix ? 1 : 0);
          maxIndent = Math.max(maxIndent, ind);
          const prefix = (b.list && k === 0 ? '• ' : '') + (p.prefix || '');
          body.push(`<text:p text:style-name="${ind ? 'P_ind' + ind : 'Text_20_body'}">${odfText(prefix)}${runs(p.runs)}</text:p>`);
        });
      } else if (b.type === 'side') {
        const parts = b.values.map(v => Docx.paragraphs(v).map(p => runs(p.runs)).join('<text:line-break/>'));
        body.push(`<text:p text:style-name="Side_20_Note"><text:span text:style-name="T_b">${odfText(b.label + ': ')}</text:span>${parts.join('<text:line-break/>')}</text:p>`);
      }
    }
    if (!spanStyles.has('b')) spanStyles.set('b', 'T_b');
    const auto = [...spanStyles.entries()].map(([key, name]) => `<style:style style:name="${name}" style:family="text"><style:text-properties${key.includes('b') ? ' fo:font-weight="bold"' : ''}${key.includes('i') ? ' fo:font-style="italic"' : ''}${key.includes('s') ? ' style:text-line-through-style="solid"' : ''}${key.includes('c') ? ' style:font-name="Liberation Mono"' : ''}/></style:style>`).join('')
      + Array.from({ length: maxIndent }, (_, k) => `<style:style style:name="P_ind${k + 1}" style:family="paragraph" style:parent-style-name="Text_20_body"><style:paragraph-properties fo:margin-left="${((k + 1) * 0.9).toFixed(2)}cm"/></style:style>`).join('');
    return XML + `<office:document-content ${NSDECL}><office:font-face-decls><style:font-face style:name="Liberation Mono" svg:font-family="'Liberation Mono'" style:font-family-generic="modern"/></office:font-face-decls><office:automatic-styles>${auto}</office:automatic-styles><office:body><office:text>${body.join('')}</office:text></office:body></office:document-content>`;
  }
  async function buildOdt(doc, opts = {}) {
    return pack(ODT_MIME, { 'META-INF/manifest.xml': manifest(ODT_MIME), 'content.xml': odtContent(doc, opts), 'styles.xml': odtStyles(), 'meta.xml': metaXml(doc) });
  }

  // ---------- ODS write ----------
  const cmFromExcelWidth = w => (w * 7 / 37.8).toFixed(3) + 'cm';
  function cellXml(v, style, href) {
    const attrs = style ? ` table:style-name="${style}"` : '';
    if (v === '' || v == null) return `<table:table-cell${attrs}/>`;
    if (typeof v === 'number') return `<table:table-cell${attrs} office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>`;
    const ps = String(v).split('\n').map(line => `<text:p>${href ? `<text:a xlink:type="simple" xlink:href="${esc(href)}">${odfText(line)}</text:a>` : odfText(line)}</text:p>`).join('');
    return `<table:table-cell${attrs} office:value-type="string">${ps}</table:table-cell>`;
  }
  const odsLink = (sheet, col, row) => `#${/[^A-Za-z0-9_]/.test(sheet) ? `'${sheet.replace(/'/g, "''")}'` : sheet}.${col}${row}`;
  function odsContent(doc) {
    const widths = new Set();
    const colStyle = w => { widths.add(w); return 'co' + String(w).replace('.', '_'); };
    const tables = [];
    const used = new Set(['.sheetwriter', '.contents']);
    const sheetNames = new Map();
    for (const s of doc.sheets) {
      let name = XlsxIO.safeSheetName(s.name), base = name, k = 2;
      while (used.has(name.toLowerCase())) name = base.slice(0, 28) + ' ' + k++;
      used.add(name.toLowerCase()); sheetNames.set(s, name);
    }
    const table = (name, cols, rows, prot) => `<table:table table:name="${esc(name)}"${prot ? ' table:protected="true"' : ''}>${cols.map(w => typeof w === 'object' ? `<table:table-column table:style-name="${colStyle(w.w)}" table:visibility="collapse"/>` : `<table:table-column table:style-name="${colStyle(w)}"/>`).join('')}${rows.map(r => `<table:table-row table:style-name="ro1">${r.map(c => cellXml(c.v, c.s, c.href)).join('')}</table:table-row>`).join('')}</table:table>`;
    for (const s of doc.sheets) {
      const name = sheetNames.get(s);
      if (s.kind !== 'chapter') {
        const cells = s.cells || [];
        const ncol = Math.max(1, ...cells.map(r => r.length));
        tables.push(table(name, Array.from({ length: ncol }, () => 14), cells.map(r => Array.from({ length: ncol }, (_, i) => ({ v: r[i] || '', s: 'cell' }))), false));
        continue;
      }
      const columns = XlsxIO.fileColumns(doc, s);
      const num = Model.numbering(doc, doc.sheets.indexOf(s));
      const rows = [columns.map(c => ({ v: c, s: c.startsWith('.') ? 'hdrsys' : 'hdr' }))];
      const written = s.rows.filter(r => !Model.rowIsEmpty(doc, s, r));
      (written.length ? written : [s.rows[0]]).forEach(r => {
        const i = s.rows.indexOf(r);
        const hl = SheetNumbering.headingLevel(r['.kind']);
        const rc = doc.settings.trackCounts ? Model.rowCounts(doc, s, r) : null;
        rows.push(columns.map(c => {
          let v, st = 'cell';
          if (c === '.no') v = num.numbers[i];
          else if (c === '.indent') { const n = Model.indentOf(r); v = n ? n : ''; }
          else if (rc && c === '.words') v = rc.words;
          else if (rc && c === '.chars') v = rc.chars;
          else v = String(r[c] ?? '');
          if (Model.isMeta(doc, c) || Model.isComputed(doc, c)) st = 'cellmeta';
          else if (hl) st = 'cellh';
          else if (r['.kind'] === 'x') st = 'cellx';
          return { v, s: st };
        }));
      });
      tables.push(table(name, columns.map(c => c === Model.IDENT ? { w: XlsxIO.widthFor(c, doc) } : XlsxIO.widthFor(c, doc)), rows, doc.settings.protectHeaders !== false));
    }
    if (doc.settings.contentsSheet !== false) {
      const rows = [['no', 'heading', 'sheet', 'words'].map(v => ({ v, s: 'hdr' }))];
      for (const e of XlsxIO.contentsEntries(doc, sheetNames)) {
        rows.push([{ v: e.no, s: 'cell' }, { v: e.heading, s: e.level === 1 ? 'linkb' : 'link', href: e.row ? odsLink(e.sheet, e.col, e.row) : null }, { v: e.sheet, s: 'cell' }, { v: e.words, s: 'cell' }]);
      }
      tables.push(table('.contents', [10, 70, 24, 9], rows, true));
    }
    const srows = [['key', 'value', 'description'].map(v => ({ v, s: 'hdr' }))];
    for (const r of XlsxIO.settingsRows(doc)) srows.push([{ v: r.key, s: 'cell' }, { v: r.value, s: r.link ? 'link' : r.bold ? 'cellh' : 'cell', href: r.link ? r.value : null }, { v: r.desc, s: 'cell' }]);
    tables.push(table('.sheetwriter', [16, 56, 70], srows, true));
    const cellStyle = (name, props, textProps) => `<style:style style:name="${name}" style:family="table-cell"><style:table-cell-properties fo:wrap-option="wrap" style:vertical-align="top" ${props}/>${textProps ? `<style:text-properties ${textProps}/>` : ''}</style:style>`;
    const auto = [...widths].map(w => `<style:style style:name="${colStyle(w)}" style:family="table-column"><style:table-column-properties style:column-width="${cmFromExcelWidth(w)}"/></style:style>`).join('')
      + `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:use-optimal-row-height="true"/></style:style>`
      + cellStyle('hdr', 'fo:background-color="#e9e9e9" style:cell-protect="protected"', 'fo:font-weight="bold"')
      + cellStyle('hdrsys', 'fo:background-color="#d9d9d9" style:cell-protect="protected"', 'fo:font-weight="bold"')
      + cellStyle('cell', 'style:cell-protect="none"', '')
      + cellStyle('cellh', 'style:cell-protect="none"', 'fo:font-weight="bold"')
      + cellStyle('cellx', 'style:cell-protect="none"', 'fo:font-style="italic" fo:color="#8a8a8a"')
      + cellStyle('cellmeta', 'style:cell-protect="none"', 'fo:font-size="9pt" fo:color="#8a8a8a"')
      + cellStyle('link', 'style:cell-protect="none"', 'fo:color="#2f6fdb" style:text-underline-style="solid"')
      + cellStyle('linkb', 'style:cell-protect="none"', 'fo:color="#2f6fdb" style:text-underline-style="solid" fo:font-weight="bold"');
    return XML + `<office:document-content ${NSDECL}><office:automatic-styles>${auto}</office:automatic-styles><office:body><office:spreadsheet>${tables.join('')}</office:spreadsheet></office:body></office:document-content>`;
  }
  function odsStyles() {
    return XML + `<office:document-styles ${NSDECL}><office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt"/></style:default-style></office:styles></office:document-styles>`;
  }
  async function saveOds(doc) {
    Model.ensureRowIds(doc);
    return pack(ODS_MIME, { 'META-INF/manifest.xml': manifest(ODS_MIME), 'content.xml': odsContent(doc), 'styles.xml': odsStyles(), 'meta.xml': metaXml(doc) });
  }

  // ---------- ODS read ----------
  function textOf(node) {
    let s = '';
    for (const c of node.childNodes) {
      if (c.nodeType === 3) s += c.nodeValue;
      else if (c.nodeType === 1) {
        const ln = c.localName;
        if (ln === 's') s += ' '.repeat(+c.getAttributeNS(NS.text, 'c') || 1);
        else if (ln === 'line-break') s += '\n';
        else if (ln === 'tab') s += '\t';
        else if (ln === 'note' || ln === 'annotation') continue;
        else s += textOf(c);
      }
    }
    return s;
  }
  function cellValue(el) {
    const ps = [...el.children].filter(e => e.localName === 'p');
    const text = ps.map(textOf).join('\n');
    if (text) {
      const a = ps.length === 1 && ps[0].children.length === 1 && ps[0].children[0].localName === 'a' ? ps[0].children[0] : null;
      const href = a && a.getAttributeNS(NS.xlink, 'href');
      if (href && !href.startsWith('#') && href !== text) return `[${text}](${href})`;
      return text;
    }
    const type = el.getAttributeNS(NS.office, 'value-type');
    if (type === 'float' || type === 'percentage' || type === 'currency') return el.getAttributeNS(NS.office, 'value') || '';
    if (type === 'date') return el.getAttributeNS(NS.office, 'date-value') || '';
    if (type === 'time') return el.getAttributeNS(NS.office, 'time-value') || '';
    if (type === 'boolean') return el.getAttributeNS(NS.office, 'boolean-value') || '';
    return '';
  }
  /** A worksheet-like adapter over one table:table, expanding repeated rows and cells (capped). */
  function adapter(t) {
    const name = t.getAttributeNS(NS.table, 'name') || 'Sheet';
    const cells = [];
    for (const rowEl of t.getElementsByTagNameNS(NS.table, 'table-row')) {
      const row = [];
      for (const cellEl of rowEl.children) {
        if (cellEl.localName !== 'table-cell' && cellEl.localName !== 'covered-table-cell') continue;
        const rep = Math.min(+cellEl.getAttributeNS(NS.table, 'number-columns-repeated') || 1, 1024);
        const v = cellValue(cellEl);
        for (let k = 0; k < rep; k++) row.push(v);
      }
      while (row.length && row[row.length - 1] === '') row.pop();
      const rep = row.length ? Math.min(+rowEl.getAttributeNS(NS.table, 'number-rows-repeated') || 1, 5000) : 1;
      for (let k = 0; k < rep; k++) cells.push(row);
    }
    while (cells.length && !cells[cells.length - 1].length) cells.pop();
    const columnCount = Math.max(0, ...cells.map(r => r.length));
    return {
      name, rowCount: cells.length, columnCount,
      getRow(r) { const arr = cells[r - 1] || []; return { cellCount: arr.length, getCell(c) { return { value: arr[c - 1] == null ? '' : arr[c - 1] }; } }; },
    };
  }
  async function loadOds(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const f = zip.file('content.xml');
    if (!f) throw new Error('Not an OpenDocument spreadsheet: content.xml is missing.');
    const dom = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
    if (dom.querySelector('parsererror')) throw new Error('content.xml is not well-formed XML.');
    const adapters = [...dom.getElementsByTagNameNS(NS.table, 'table')].map(adapter);
    const byName = new Map(adapters.map(a => [a.name, a]));
    return XlsxIO.loadFromSheets(adapters, name => byName.get(name));
  }
  async function looksLikeOds(buffer) {
    try { const zip = await JSZip.loadAsync(buffer); const m = zip.file('mimetype'); return !!m && (await m.async('string')).trim() === ODS_MIME; } catch (e) { return false; }
  }

  return { buildOdt, saveOds, loadOds, looksLikeOds, odtContent, odsContent, ODT_MIME, ODS_MIME };
})();

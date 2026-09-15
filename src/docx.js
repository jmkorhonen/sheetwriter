/* SheetWriter — Word (.docx) export. A minimal OOXML writer on JSZip; no other dependency.
 *
 * Docx.build(doc, opts) → Blob. opts as for Exporter.toBlocks. Headings use Word's built-in
 * heading styles (so the navigation pane and a table of contents work), paragraphs keep bold,
 * italic, strikethrough, code and links from the cell Markdown, indented rows are indented,
 * side columns become small indented notes, or real footnotes in footnote mode.
 */
const Docx = (() => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const unent = s => String(s).replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");

  /** Markdown text → paragraphs of runs: [{runs: [{text, b, i, strike, code} | {br}], prefix}] */
  function paragraphs(md) {
    let tokens;
    try { tokens = marked.lexer(String(md)); } catch (e) { tokens = [{ type: 'paragraph', text: String(md), tokens: [{ type: 'text', text: String(md) }] }]; }
    const out = [];
    const inline = (toks, fmt, runs) => {
      for (const t of toks || []) {
        switch (t.type) {
          case 'text': case 'escape': case 'html': runs.push({ text: unent(t.text != null ? t.text : t.raw), ...fmt }); break;
          case 'strong': inline(t.tokens, { ...fmt, b: true }, runs); break;
          case 'em': inline(t.tokens, { ...fmt, i: true }, runs); break;
          case 'del': inline(t.tokens, { ...fmt, strike: true }, runs); break;
          case 'codespan': runs.push({ text: unent(t.text), ...fmt, code: true }); break;
          case 'link': inline(t.tokens, fmt, runs); if (t.href && !/^#/.test(t.href)) runs.push({ text: ` (${t.href})`, ...fmt, i: true }); break;
          case 'image': runs.push({ text: t.text || t.href || '', ...fmt }); break;
          case 'br': runs.push({ br: true }); break;
          default: if (t.tokens) inline(t.tokens, fmt, runs); else if (t.text) runs.push({ text: unent(t.text), ...fmt });
        }
      }
    };
    const block = (toks, prefix) => {
      for (const t of toks) {
        switch (t.type) {
          case 'paragraph': case 'text': { const runs = []; inline(t.tokens || [{ type: 'text', text: t.text }], {}, runs); out.push({ runs, prefix }); break; }
          case 'heading': { const runs = []; inline(t.tokens, { b: true }, runs); out.push({ runs, prefix }); break; }
          case 'list': for (const item of t.items) {
            block(item.tokens.filter(x => x.type !== 'list'), prefix + (t.ordered ? '' : '• '));
            for (const n of item.tokens.filter(x => x.type === 'list')) block([n], prefix + ' ');
          } break;
          case 'blockquote': block(t.tokens, prefix + ' '); break;
          case 'code': out.push({ runs: [{ text: t.text, code: true }], prefix }); break;
          case 'space': case 'hr': break;
          default: if (t.raw && t.raw.trim()) out.push({ runs: [{ text: t.raw.trim() }], prefix });
        }
      }
    };
    block(tokens, '');
    if (!out.length) out.push({ runs: [{ text: String(md) }], prefix: '' });
    // [^n] markers written by the exporter's footnote mode become footnote-reference runs {fn: n}.
    for (const p of out) p.runs = p.runs.flatMap(r => {
      if (r.br || !r.text || !/\[\^\d+\]/.test(r.text)) return [r];
      return r.text.split(/(\[\^\d+\])/).filter(Boolean).map(t => { const m = /^\[\^(\d+)\]$/.exec(t); return m ? { fn: +m[1] } : { ...r, text: t }; });
    });
    return out;
  }

  function runXml(r) {
    if (r.br) return '<w:r><w:br/></w:r>';
    if (r.fn) return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${r.fn}"/></w:r>`;
    const p = [];
    if (r.b) p.push('<w:b/>');
    if (r.i) p.push('<w:i/>');
    if (r.strike) p.push('<w:strike/>');
    if (r.code) p.push('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>');
    return `<w:r>${p.length ? `<w:rPr>${p.join('')}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(r.text || '')}</w:t></w:r>`;
  }
  function paraXml(style, runs, indentTwips, prefix) {
    const pPr = [];
    if (style) pPr.push(`<w:pStyle w:val="${style}"/>`);
    if (indentTwips) pPr.push(`<w:ind w:left="${indentTwips}"/>`);
    const all = (prefix ? [{ text: prefix }] : []).concat(runs);
    return `<w:p>${pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : ''}${all.map(runXml).join('')}</w:p>`;
  }

  function bodyXml(blocks) {
    const parts = [];
    for (const b of blocks) {
      if (b.type === 'sheetTitle' || b.type === 'heading') {
        const p = paragraphs(b.text)[0];
        parts.push(paraXml('Heading' + Math.min(b.level, 6), p.runs, 0, ''));
      } else if (b.type === 'para') {
        paragraphs(b.text).forEach((p, k) => {
          const indent = (b.indent || 0) * 360 + (p.prefix ? 360 : 0);
          parts.push(paraXml(b.list ? 'ListParagraph' : null, p.runs, indent, (b.list && k === 0 ? '• ' : '') + (p.prefix || '')));
        });
      } else if (b.type === 'side') {
        const runs = [{ text: b.label + ': ', b: true }];
        b.values.forEach((v, k) => { if (k) runs.push({ br: true }); paragraphs(v).forEach((p, j) => { if (j) runs.push({ br: true }); runs.push(...p.runs); }); });
        parts.push(paraXml('SideNote', runs, 0, ''));
      }
    }
    return parts.join('');
  }

  const heading = (n, size, extra) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="${n === 1 ? 480 : 280}" w:after="120"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr>${extra}<w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
${heading(1, 32, '<w:b/>')}${heading(2, 28, '<w:b/>')}${heading(3, 26, '<w:b/><w:i/>')}${heading(4, 24, '<w:i/>')}${heading(5, 22, '<w:i/>')}${heading(6, 22, '<w:i/>')}
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:uiPriority w:val="99"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="SideNote"><w:name w:val="Side Note"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:spacing w:before="0" w:after="200"/></w:pPr><w:rPr><w:i/><w:color w:val="666666"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>
</w:styles>`;
  const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
  const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>`;
  /** word/footnotes.xml: the two separator entries Word expects, then one footnote per note. */
  function footnotesXml(notes) {
    const sep = `<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>`;
    const body = (notes || []).map(f => {
      const runs = [];
      paragraphs(f.text).forEach((p, k) => { if (k) runs.push({ br: true }); if (p.prefix) runs.push({ text: p.prefix }); runs.push(...p.runs); });
      return `<w:footnote w:id="${f.n}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>${runs.map(runXml).join('')}</w:p></w:footnote>`;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes xmlns:w="${W}">${sep}${body}</w:footnotes>`;
  }
  const APP = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>SheetWriter</Application></Properties>`;

  function coreXml(doc) {
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(doc.settings.title || '')}</dc:title><dc:creator>${esc(doc.settings.author || '')}</dc:creator><dc:description>${esc(doc.settings.description || '')}</dc:description><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  }

  function documentXml(doc, opts, blocks) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body>${bodyXml(blocks || Exporter.toBlocks(doc, opts))}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  }

  async function build(doc, opts = {}) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', CONTENT_TYPES);
    zip.file('_rels/.rels', RELS);
    zip.file('word/_rels/document.xml.rels', DOC_RELS);
    const blocks = Exporter.toBlocks(doc, opts);
    const fn = blocks.find(b => b.type === 'footnotes');
    zip.file('word/document.xml', documentXml(doc, opts, blocks));
    zip.file('word/footnotes.xml', footnotesXml(fn ? fn.notes : []));
    zip.file('word/styles.xml', STYLES);
    zip.file('docProps/core.xml', coreXml(doc));
    zip.file('docProps/app.xml', APP);
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
  }

  return { build, documentXml, footnotesXml, paragraphs, MIME: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
})();

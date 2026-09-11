/* SheetWriter — Markdown export, by column.
 *
 * opts = {
 *   column:      which column becomes the document (default doc.mainColumn)
 *   scope:       'all' | sheet index
 *   sheetTitles: emit "# <sheet name>" per chapter (default false)
 *   numbering:   false | 'headings' | 'all' (true = 'all'): prepend computed numbers
 *   indented:    'paragraphs' (default) | 'lists': how rows with indent > 0 are written
 *   side:        null | {column, mode: 'quote'|'comment'}
 * }
 * Row kinds: hN heading · p paragraph · s continues previous paragraph · x omitted.
 * Cell text is already Markdown and passes through unchanged.
 */
const Exporter = (() => {
  function toMarkdown(doc, opts = {}) {
    const column = opts.column || doc.mainColumn;
    const chapters = doc.sheets.map((s, i) => ({ s, i })).filter(x => x.s.kind === 'chapter');
    const selected = opts.scope === 'all' || opts.scope == null
      ? chapters
      : chapters.filter(x => x.i === +opts.scope);
    const sheetTitles = !!opts.sheetTitles;
    const offset = sheetTitles ? 1 : 0;
    const numAll = opts.numbering === true || opts.numbering === 'all';
    const numHead = numAll || opts.numbering === 'headings';
    const lists = opts.indented === 'lists';
    const parts = [];
    for (const { s, i } of selected) {
      const blocks = []; // {text, list?: true}
      if (sheetTitles) blocks.push({ text: '# ' + s.name });
      const num = Model.numbering(doc, i);
      let para = null;
      const flush = () => {
        if (!para) return;
        const text = para.text.join(' ').trim();
        if (text) {
          const isList = lists && para.indent > 0;
          const prefix = isList ? '  '.repeat(para.indent - 1) + '- ' : '';
          blocks.push({ text: prefix + (numAll ? para.num + ' ' : '') + text, list: isList });
          for (const b of sideBlocks(para.side, opts.side)) blocks.push({ text: b });
        }
        para = null;
      };
      s.rows.forEach((r, k) => {
        const kind = Model.normKind(r.kind);
        const text = String(r[column] || '').trim();
        const sideVal = opts.side && opts.side.column ? String(r[opts.side.column] || '').trim() : '';
        if (kind === 'x') return;
        const hl = SheetNumbering.headingLevel(kind);
        if (hl) {
          flush();
          // A leading "#" typed into the cell is a heading marker, not text.
          const htext = (text || String(r[doc.mainColumn] || '').trim()).replace(/^#{1,6}[ \t]+/, '');
          if (!htext) return;
          const level = Math.min(6, hl + offset);
          blocks.push({ text: '#'.repeat(level) + ' ' + (numHead ? num.numbers[k] + ' ' : '') + htext });
          for (const b of sideBlocks(sideVal ? [sideVal] : [], opts.side)) blocks.push({ text: b });
          return;
        }
        if (kind === 's' && para) {
          if (text) para.text.push(text);
          if (sideVal) para.side.push(sideVal);
          return;
        }
        flush();
        para = { num: num.numbers[k], indent: num.indents[k], text: text ? [text] : [], side: sideVal ? [sideVal] : [] };
      });
      flush();
      let out = '';
      blocks.forEach((b, k) => {
        if (k) out += (b.list && blocks[k - 1].list) ? '\n' : '\n\n';
        out += b.text;
      });
      parts.push(out);
    }
    return parts.join('\n\n') + '\n';
  }

  function sideBlocks(values, side) {
    if (!side || !side.column || !values.length) return [];
    const label = side.column;
    if (side.mode === 'comment') return [`<!-- ${label}: ${values.join(' | ').replace(/-->/g, '--​>')} -->`];
    // blockquote: keep it readable, one value per line
    return ['> **' + label + ':** ' + values.map(v => v.replace(/\n/g, ' ')).join('\n> ')];
  }

  return { toMarkdown };
})();
if (typeof module !== 'undefined') module.exports = Exporter;

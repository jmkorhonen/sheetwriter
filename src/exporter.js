/* SheetWriter — export: blocks, Markdown.
 *
 * toBlocks(doc, opts) → document-order blocks, each addressable by its row:
 *   {type: 'sheetTitle', si, level: 1, text}
 *   {type: 'heading', si, i, level, text, num}
 *   {type: 'para', si, i, text, indent, list, num}          text = paragraph + its "s" rows
 *   {type: 'side', si, i, label, values, mode, text}         a side column shown after a heading/paragraph
 * toMarkdown(doc, opts) joins the blocks. The Read view and the Word export render them one by one.
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
  function toBlocks(doc, opts = {}) {
    const column = opts.column || doc.mainColumn;
    const chapters = doc.sheets.map((s, i) => ({ s, i })).filter(x => x.s.kind === 'chapter');
    const selected = opts.scope === 'all' || opts.scope == null ? chapters : chapters.filter(x => x.i === +opts.scope);
    const sheetTitles = !!opts.sheetTitles;
    const offset = sheetTitles ? 1 : 0;
    const numAll = opts.numbering === true || opts.numbering === 'all';
    const numHead = numAll || opts.numbering === 'headings';
    const lists = opts.indented === 'lists';
    const blocks = [];
    for (const { s, i } of selected) {
      let first = true;
      const push = b => { b.first = first; first = false; blocks.push(b); };
      if (sheetTitles) push({ type: 'sheetTitle', si: i, level: 1, text: s.name });
      const num = Model.numbering(doc, i);
      let para = null;
      const flush = () => {
        if (!para) return;
        const text = para.text.join(' ').trim();
        if (text) {
          push({ type: 'para', si: i, i: para.i, text: (numAll ? para.num + ' ' : '') + text, indent: para.indent, list: lists && para.indent > 0, num: para.num });
          const sb = sideBlock(para.side, opts.side);
          if (sb) push({ type: 'side', si: i, i: para.i, ...sb });
        }
        para = null;
      };
      s.rows.forEach((r, k) => {
        const kind = Model.normKind(r['.kind']);
        const text = String(r[column] || '').trim();
        const sideVal = opts.side && opts.side.column ? String(r[opts.side.column] || '').trim() : '';
        if (kind === 'x') return;
        const hl = SheetNumbering.headingLevel(kind);
        if (hl) {
          flush();
          // A leading "#" typed into the cell is a heading marker, not text.
          const htext = (text || String(r[doc.mainColumn] || '').trim()).replace(/^#{1,6}[ \t]+/, '');
          if (!htext) return;
          push({ type: 'heading', si: i, i: k, level: Math.min(6, hl + offset), text: (numHead ? num.numbers[k] + ' ' : '') + htext, num: num.numbers[k] });
          const sb = sideBlock(sideVal ? [sideVal] : [], opts.side);
          if (sb) push({ type: 'side', si: i, i: k, ...sb });
          return;
        }
        if (kind === 's' && para) {
          if (text) para.text.push(text);
          if (sideVal) para.side.push(sideVal);
          return;
        }
        flush();
        para = { i: k, num: num.numbers[k], indent: num.indents[k], text: text ? [text] : [], side: sideVal ? [sideVal] : [] };
      });
      flush();
    }
    return blocks;
  }

  function sideBlock(values, side) {
    if (!side || !side.column || !values.length) return null;
    const label = side.column;
    const mode = side.mode === 'comment' ? 'comment' : 'quote';
    const text = mode === 'comment'
      ? `<!-- ${label}: ${values.join(' | ').replace(/-->/g, '--​>')} -->`
      : '> **' + label + ':** ' + values.map(v => v.replace(/\n/g, ' ')).join('\n> ');
    return { label, values, mode, text };
  }

  function blockMarkdown(b) {
    if (b.type === 'heading' || b.type === 'sheetTitle') return '#'.repeat(b.level) + ' ' + b.text;
    if (b.type === 'para' && b.list) return '  '.repeat(b.indent - 1) + '- ' + b.text;
    return b.text;
  }

  function toMarkdown(doc, opts = {}) {
    const blocks = toBlocks(doc, opts);
    let out = '';
    blocks.forEach((b, k) => {
      if (k) out += (b.list && blocks[k - 1].list && !b.first) ? '\n' : '\n\n';
      out += blockMarkdown(b);
    });
    return out + '\n';
  }

  /** The headings toMarkdown emits, in order: {sheetTitle: si} or {si, i}. */
  function headingOrder(doc, opts = {}) {
    return toBlocks(doc, opts).filter(b => b.type === 'heading' || b.type === 'sheetTitle')
      .map(b => (b.type === 'sheetTitle' ? { sheetTitle: b.si } : { si: b.si, i: b.i }));
  }

  return { toBlocks, toMarkdown, blockMarkdown, headingOrder };
})();
if (typeof module !== 'undefined') module.exports = Exporter;

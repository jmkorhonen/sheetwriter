/* SheetWriter — Markdown import. Pure: string + options → rows. Uses marked's block lexer.
 *
 * parse(md, opts) → { rows: [{kind, indent, text, side: {column: value}}], columns: [side columns used],
 *                     settings: {title?, author?, description?}, warnings: [], stats: {rows, headings} }
 * opts.granularity: 'paragraphs' (default) | 'sentences' | 'lines'
 * opts.stripNumbers: true | false | 'auto' (default: strip when most headings carry section numbers)
 *
 * analyze(md) → { semanticLines, numbered, title } for pre-selecting dialog options.
 */
const Importer = (() => {
  // Abbreviations (English and Finnish) after which a period does not end a sentence.
  const ABBR = new Set(['e.g', 'i.e', 'vs', 'etc', 'cf', 'ca', 'approx', 'dr', 'mr', 'mrs', 'ms', 'prof', 'jr', 'sr', 'st', 'no', 'vol', 'fig', 'eq', 'al', 'ed', 'eds', 'pp', 'p', 'ibid', 'op',
    'esim', 'mm', 'ks', 'vrt', 'jne', 'ns', 'n', 's', 'k', 'mrd', 'milj', 'tms', 'ym', 'yms', 'em', 'jk', 'huom', 'kts', 'nk', 'os', 'puh', 'klo', 'v', 'vrk', 'ent', 'ts', 'ml']);

  function stripFrontMatter(md) {
    md = md.replace(/^﻿/, '').replace(/\r\n?/g, '\n');
    const m = /^---[ \t]*\n([\s\S]*?)\n---[ \t]*\n?/.exec(md);
    const settings = {};
    if (!m) return { body: md, settings };
    for (const line of m[1].split('\n')) {
      const kv = /^([A-Za-z_][\w-]*)\s*:\s*(.*)$/.exec(line);
      if (!kv) continue;
      const key = kv[1].toLowerCase(), val = kv[2].trim().replace(/^["'](.*)["']$/, '$1');
      if (['title', 'author', 'description'].includes(key)) settings[key] = val;
    }
    return { body: md.slice(m[0].length), settings };
  }

  function analyze(md) {
    const { body, settings } = stripFrontMatter(md);
    const lines = body.split('\n');
    const blocks = []; let cur = [];
    for (const l of lines) {
      const t = l.trim();
      if (!t || /^(#{1,6}\s|[-*+]\s|\d+[.)]\s|>|```|~~~|\||---|===|<!--)/.test(t)) { if (cur.length) { blocks.push(cur); cur = []; } }
      else cur.push(t);
    }
    if (cur.length) blocks.push(cur);
    const multi = blocks.filter(b => b.length >= 2).flat();
    const ended = multi.filter(l => /[.!?:;»”"')\]]$/.test(l)).length;
    const avgLen = multi.reduce((n, l) => n + l.length, 0) / (multi.length || 1);
    const semanticLines = multi.length >= 4 && ended / multi.length >= 0.6 && avgLen < 200;
    const headings = lines.filter(l => /^#{1,6}\s/.test(l));
    const numbered = headings.length > 0 && headings.filter(l => /^#{1,6}\s+\d+(\.\d+)*\.?\s/.test(l)).length / headings.length >= 0.5;
    const first = lines.find(l => /^#\s+/.test(l));
    const title = settings.title || (first ? first.replace(/^#\s+/, '').replace(/^\d+(\.\d+)*\.?\s+/, '').trim() : '');
    return { semanticLines, numbered, title };
  }

  /** Join soft line breaks with a space, keep hard breaks (two trailing spaces or a backslash) as newlines. */
  function joinSoft(raw) {
    const lines = raw.replace(/\n+$/, '').split('\n');
    let out = '';
    lines.forEach((l, k) => {
      if (k === 0) { out = l; return; }
      const prev = lines[k - 1];
      if (/(\s{2,}|\\)$/.test(prev)) out = out.replace(/(\s{2,}|\\)$/, '') + '\n' + l.trimStart();
      else out += ' ' + l.trim();
    });
    return out.trim();
  }
  function splitLines(raw) {
    return raw.replace(/\n+$/, '').split('\n').map(l => l.replace(/(\s{2,}|\\)$/, '').trim()).filter(Boolean);
  }
  function count(s, ch) { let n = 0; for (const c of s) if (c === ch) n++; return n; }
  function splitSentences(text) {
    const out = [];
    let start = 0;
    const re = /([.!?]+)(["'”’)\]]*)\s+(?=["'“‘(\[]?[A-ZÅÄÖØÆÜÉ0-9])/g;
    let m;
    while ((m = re.exec(text))) {
      const end = m.index + m[1].length + m[2].length;
      const before = text.slice(start, end);
      const lastWord = (/([\p{L}\p{N}.]+)[.!?]+["'”’)\]]*$/u.exec(before) || [])[1] || '';
      const w = lastWord.replace(/\.+$/, '').toLowerCase();
      if (m[1] === '.' && (ABBR.has(w) || /^\p{Lu}$/u.test(lastWord.replace(/\.$/, '')) || /^\d+$/.test(w))) continue;
      if (count(before, '(') !== count(before, ')') || count(before, '[') !== count(before, ']')) continue;
      out.push(before.trim());
      start = m.index + m[0].length;
    }
    out.push(text.slice(start).trim());
    return out.filter(Boolean);
  }

  const stripQuote = raw => raw.replace(/\n+$/, '').split('\n').map(l => l.replace(/^>\s?/, '')).join('\n').trim();
  const SIDE_QUOTE = /^\*\*([^*\n]{1,40}):\*\*\s*([\s\S]*)$/;
  const SIDE_COMMENT = /^<!--\s*([\w\- ]{1,40}?):\s*([\s\S]*?)\s*-->$/;

  function parse(md, opts = {}) {
    const granularity = opts.granularity || 'paragraphs';
    const { body, settings } = stripFrontMatter(md);
    const info = analyze(md);
    const stripNumbers = opts.stripNumbers === 'auto' || opts.stripNumbers == null ? info.numbered : !!opts.stripNumbers;
    const rows = [], warnings = [], columns = new Set();
    const last = () => rows[rows.length - 1];
    const push = (kind, indent, text, side) => { const r = { kind, indent: indent > 0 ? String(indent) : '', text: String(text || '').trim(), side: side || {} }; rows.push(r); return r; };
    const stripNum = (text, heading) => stripNumbers ? text.replace(heading ? /^\d+(\.\d+)*\.?\s+/ : /^\d+(\.\d+)+\.?\s+/, '') : text;
    const addSide = (col, val) => {
      col = col.trim(); columns.add(col);
      const r = last();
      if (!r) { push('x', 0, `${col}: ${val}`); return; }
      r.side[col] = r.side[col] ? r.side[col] + '\n' + val : val;
    };

    function bodyRows(raw, indent) {
      // Emit one paragraph's worth of rows according to the granularity.
      let parts;
      if (granularity === 'lines') parts = splitLines(raw);
      else if (granularity === 'sentences') parts = splitSentences(joinSoft(raw));
      else parts = [joinSoft(raw)];
      if (!parts.length) return;
      parts.forEach((p, k) => push(k === 0 ? 'p' : 's', indent, k === 0 ? stripNum(p, false) : p));
    }
    function listRows(token, depth, hang) {
      for (const item of token.items) {
        const own = item.tokens.filter(t => t.type !== 'list').map(t => t.raw).join('').trim();
        const nested = item.tokens.filter(t => t.type === 'list');
        if (own) bodyRows(own, depth + hang);
        for (const n of nested) listRows(n, depth + 1, hang);
      }
    }

    let tokens;
    try { tokens = marked.lexer(body); } catch (e) { warnings.push('Could not parse the Markdown: ' + e.message); tokens = [{ type: 'paragraph', raw: body }]; }
    for (const t of tokens) {
      switch (t.type) {
        case 'space': case 'hr': break;
        case 'heading': {
          let level = t.depth;
          if (level > 4) { warnings.push(`Heading level ${level} ("${t.text.slice(0, 40)}") imported as h4.`); level = 4; }
          push('h' + level, 0, stripNum(t.text, true));
          break;
        }
        case 'paragraph': case 'text': bodyRows(t.raw, 0); break;
        case 'list': {
          const prev = last();
          const hang = prev && !/^h/.test(prev.kind) ? 1 : 0;
          listRows(t, 0, hang);
          break;
        }
        case 'blockquote': {
          const inner = stripQuote(t.raw);
          const m = SIDE_QUOTE.exec(inner);
          if (m) addSide(m[1], m[2].trim());
          else push('p', 0, t.raw.replace(/\n+$/, ''));
          break;
        }
        case 'html': {
          const raw = t.raw.trim();
          const m = SIDE_COMMENT.exec(raw);
          if (m) addSide(m[1], m[2]);
          else if (/^<!--[\s\S]*-->$/.test(raw)) push('x', 0, raw.replace(/^<!--\s*|\s*-->$/g, ''));
          else push('p', 0, raw);
          break;
        }
        case 'code': case 'table': push('p', 0, t.raw.replace(/\n+$/, '')); break;
        default: if (t.raw && t.raw.trim()) push('p', 0, t.raw.trim());
      }
    }
    const stats = { rows: rows.length, headings: rows.filter(r => /^h/.test(r.kind)).length };
    return { rows, columns: [...columns], settings, warnings, stats, info };
  }

  // ---- CSV / TSV ----
  /** RFC 4180-style parse: quoted fields, "" escapes, newlines inside quotes. Blank lines are dropped. */
  function parseCsv(text, delim) {
    delim = delim || detectDelimiter(text);
    const rows = []; let row = [], field = '', q = false;
    text = String(text).replace(/^\ufeff/, '');
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) { if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; } else field += c; }
      else if (c === '"' && field === '') q = true;
      else if (c === delim) { row.push(field); field = ''; }
      else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(field); rows.push(row); row = []; field = ''; }
      else field += c;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows.filter(r => r.some(c => c.trim()));
  }
  /** The delimiter (comma, semicolon or tab) that appears most consistently on the first lines. */
  function detectDelimiter(text) {
    const lines = String(text).split(/\r?\n/).filter(l => l.trim()).slice(0, 8);
    let best = ',', bestScore = -1;
    for (const d of [',', ';', '\t']) {
      const counts = lines.map(l => l.replace(/"[^"]*"/g, '').split(d).length - 1);
      if (!counts.length || counts[0] === 0) continue;
      const same = counts.filter(c => c === counts[0]).length;
      const score = same * 10 + counts[0];
      if (score > bestScore) { bestScore = score; best = d; }
    }
    return best;
  }
  /** Is this text a delimited table rather than Markdown? A .csv/.tsv name settles it; otherwise every line must carry the same number of delimiters. */
  function looksCsv(text, fileName) {
    if (/\.(csv|tsv)$/i.test(fileName || '')) return true;
    if (/\.(md|markdown|txt)$/i.test(fileName || '')) return false;
    const lines = String(text).split(/\r?\n/).filter(l => l.trim()).slice(0, 12);
    if (lines.length < 2) return false;
    const d = detectDelimiter(text);
    const counts = lines.map(l => l.replace(/"[^"]*"/g, '').split(d).length - 1);
    return counts[0] >= 1 && counts.every(c => c === counts[0]) && !lines.some(l => /^#{1,6} /.test(l));
  }
  const looksNumeric = s => /^\s*-?\d+([.,]\d+)?\s*%?\s*$/.test(s);
  /** A delimited table → importer rows. The first line is the header when it looks like one; the text comes from a
   *  column named text (or the main column), else from the longest column; kind/indent columns are honoured;
   *  every other column becomes a side column. */
  function fromCsv(text, opts = {}) {
    const table = parseCsv(text);
    const warnings = [];
    if (!table.length) return { rows: [], columns: [], settings: {}, warnings, stats: { rows: 0, headings: 0 }, main: '' };
    const width = Math.max(...table.map(r => r.length));
    const first = table[0].map(c => c.trim());
    // A header spans every column, is short, non-numeric and unique; a one-column table needs a name-like first cell ("text", not a sentence).
    const hasHeader = table.length > 1 && first.length === width && first.every(c => c && c.length <= 40 && !looksNumeric(c) && !c.includes('\n'))
      && new Set(first.map(c => c.toLowerCase())).size === first.length && (width > 1 || (/^[\w .-]{1,20}$/.test(first[0]) && !/[.!?]$/.test(first[0])));
    const headers = Array.from({ length: width }, (_, i) => (hasHeader && first[i]) ? first[i] : 'col' + (i + 1));
    const data = hasHeader ? table.slice(1) : table;
    const norm = headers.map(h => h.toLowerCase().replace(/^\./, ''));
    const idx = name => norm.indexOf(name);
    const kindCol = idx('kind'), indentCol = idx('indent'), skip = new Set([kindCol, indentCol, idx('no'), idx('id'), idx('words'), idx('chars'), idx('updated'), idx('author')].filter(i => i >= 0));
    let mainCol = idx((opts.mainColumn || 'text').toLowerCase());
    if (mainCol < 0) mainCol = idx('text');
    if (mainCol < 0) {
      let best = -1, bestLen = -1;
      for (let i = 0; i < width; i++) {
        if (skip.has(i)) continue;
        const vals = data.map(r => (r[i] || '').trim()).filter(Boolean);
        const avg = vals.length ? vals.reduce((a, v) => a + v.length, 0) / vals.length : 0;
        if (avg > bestLen) { bestLen = avg; best = i; }
      }
      mainCol = best;
      if (width > 1 && hasHeader) warnings.push(`No "text" column: using "${headers[mainCol]}" as the text.`);
    }
    const rows = []; let headings = 0;
    for (const r of data) {
      const kind = kindCol >= 0 ? Model.normKind(r[kindCol]) : 'p';
      const ind = indentCol >= 0 ? parseInt(r[indentCol], 10) || 0 : 0;
      const side = {};
      headers.forEach((h, i) => { if (i !== mainCol && !skip.has(i) && String(r[i] || '').trim()) side[h] = String(r[i]).trim(); });
      const t = String(r[mainCol] || '').trim();
      if (!t && !Object.keys(side).length) continue;
      if (Model.isHeading(kind)) headings++;
      rows.push({ kind, indent: Model.isHeading(kind) ? 0 : ind, text: t, side });
    }
    return { rows, columns: headers, settings: {}, warnings, stats: { rows: rows.length, headings }, main: headers[mainCol] || '', hasHeader };
  }

  return { parse, analyze, splitSentences, joinSoft, stripFrontMatter, parseCsv, detectDelimiter, looksCsv, fromCsv };
})();
if (typeof module !== 'undefined') module.exports = Importer;

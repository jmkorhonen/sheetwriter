/* SheetWriter — rich text (HTML from the clipboard, Word, a browser) → Markdown. Small and pragmatic. */
const Html2Md = (() => {
  const BLOCK = new Set(['p', 'div', 'section', 'article', 'main', 'header', 'footer', 'li', 'blockquote', 'pre', 'table', 'ul', 'ol', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'hr', 'tr', 'body']);
  const SKIP = new Set(['script', 'style', 'head', 'meta', 'link', 'title', 'noscript', 'template', 'o:p']);

  function isBold(el) { const w = el.style && el.style.fontWeight; return el.tagName === 'B' || el.tagName === 'STRONG' || w === 'bold' || (+w >= 600); }
  function isItalic(el) { return el.tagName === 'I' || el.tagName === 'EM' || (el.style && el.style.fontStyle === 'italic'); }

  function inline(node) {
    let s = '';
    for (const c of node.childNodes) {
      if (c.nodeType === 3) { s += c.nodeValue.replace(/\s+/g, ' '); continue; }
      if (c.nodeType !== 1) continue;
      const tag = c.tagName.toLowerCase();
      if (SKIP.has(tag)) continue;
      if (tag === 'br') { s += '  \n'; continue; }
      if (tag === 'img') { s += `![${c.getAttribute('alt') || ''}](${c.getAttribute('src') || ''})`; continue; }
      let inner = inline(c);
      if (!inner.trim()) { s += inner; continue; }
      if (tag === 'code' || tag === 'kbd' || tag === 'samp') { s += '`' + inner.trim() + '`'; continue; }
      if (tag === 'a' && c.getAttribute('href')) { s += `[${inner.trim()}](${c.getAttribute('href')})`; continue; }
      if (tag === 'del' || tag === 's' || tag === 'strike') inner = wrap(inner, '~~');
      if (isBold(c)) inner = wrap(inner, '**');
      if (isItalic(c)) inner = wrap(inner, '*');
      s += inner;
    }
    return s;
  }
  function wrap(text, m) {
    const lead = /^\s*/.exec(text)[0], trail = /\s*$/.exec(text)[0];
    const core = text.trim();
    return core ? lead + m + core + m + trail : text;
  }

  function list(el, depth, ordered) {
    const out = [];
    let n = 1;
    for (const li of el.children) {
      if (li.tagName.toLowerCase() !== 'li') continue;
      const own = [], nested = [];
      for (const c of li.childNodes) {
        if (c.nodeType === 1 && (c.tagName === 'UL' || c.tagName === 'OL')) nested.push(c);
        else own.push(c);
      }
      const frag = document.createElement('div'); own.forEach(c => frag.appendChild(c.cloneNode(true)));
      const text = blocks(frag).trim().replace(/\n\n+/g, '\n' + '  '.repeat(depth + 1));
      out.push('  '.repeat(depth) + (ordered ? `${n++}. ` : '- ') + text);
      for (const nl of nested) out.push(list(nl, depth + 1, nl.tagName === 'OL'));
    }
    return out.join('\n');
  }

  function table(el) {
    const rows = [...el.querySelectorAll('tr')].map(tr => [...tr.children].map(td => inline(td).trim().replace(/\|/g, '\\|').replace(/\s*\n\s*/g, ' ')));
    if (!rows.length) return '';
    const width = Math.max(...rows.map(r => r.length));
    const line = r => '| ' + Array.from({ length: width }, (_, k) => r[k] || '').join(' | ') + ' |';
    return [line(rows[0]), '| ' + Array(width).fill('---').join(' | ') + ' |', ...rows.slice(1).map(line)].join('\n');
  }

  function blocks(node) {
    let s = '';
    let run = ''; // inline content accumulating between block elements
    const flush = () => { if (run.trim()) s += '\n\n' + run.trim().replace(/[ \t]+\n/g, '  \n') + '\n\n'; run = ''; };
    for (const c of node.childNodes) {
      if (c.nodeType === 3) { run += c.nodeValue.replace(/\s+/g, ' '); continue; }
      if (c.nodeType !== 1) continue;
      const tag = c.tagName.toLowerCase();
      if (SKIP.has(tag)) continue;
      if (!BLOCK.has(tag)) { run += inline(c.parentNode === node ? wrapOne(c) : c); continue; }
      flush();
      const cls = (c.getAttribute('class') || '');
      if (/^h[1-6]$/.test(tag)) s += `\n\n${'#'.repeat(+tag[1])} ${inline(c).trim()}\n\n`;
      else if (tag === 'ul' || tag === 'ol') s += '\n\n' + list(c, 0, tag === 'ol') + '\n\n';
      else if (tag === 'blockquote') s += '\n\n' + blocks(c).trim().split('\n').map(l => '> ' + l).join('\n') + '\n\n';
      else if (tag === 'pre') s += '\n\n```\n' + c.textContent.replace(/\n$/, '') + '\n```\n\n';
      else if (tag === 'table') s += '\n\n' + table(c) + '\n\n';
      else if (tag === 'hr') s += '\n\n---\n\n';
      else if (tag === 'p' && /MsoListParagraph|MsoListBullet/i.test(cls)) s += '\n\n- ' + inline(c).trim().replace(/^[•·\-–o§]\s*/, '') + '\n\n';
      else if (tag === 'p' && /MsoTitle/i.test(cls)) s += `\n\n# ${inline(c).trim()}\n\n`;
      else if (tag === 'p' || tag === 'li') s += '\n\n' + inline(c).trim() + '\n\n';
      else s += blocks(c); // div, section, body…: recurse
    }
    flush();
    return s;
  }
  function wrapOne(c) { const d = document.createElement('span'); d.appendChild(c.cloneNode(true)); return d; }

  function convert(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    return blocks(doc.body).replace(/[ \t]+$/gm, '').replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }
  /** Rough test: does the clipboard HTML carry structure worth converting? */
  function looksRich(html) { return /<(p|h[1-6]|li|ul|ol|table|b|strong|i|em|a)\b/i.test(html || ''); }

  return { convert, looksRich };
})();

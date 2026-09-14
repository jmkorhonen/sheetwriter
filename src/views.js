/* SheetWriter — DOM rendering for the Draft, Grid and Read views, plus sheet tabs.
 * Views only build DOM; event handling lives in app.js (event delegation on #view / #tabs).
 */
const Views = (() => {
  const RESERVED = Model.RESERVED;

  function el(tag, attrs, ...children) {
    const n = document.createElement(tag);
    if (attrs) for (const [k, v] of Object.entries(attrs)) {
      if (v == null || v === false) continue;
      if (k === 'class') n.className = v;
      else if (k === 'html') n.innerHTML = v;
      else n.setAttribute(k, v === true ? '' : v);
    }
    for (const c of children.flat()) {
      if (c == null) continue;
      n.append(c.nodeType ? c : document.createTextNode(String(c)));
    }
    return n;
  }

  function autosize(t) {
    if (!t || t.offsetParent === null) return;
    t.style.height = 'auto';
    t.style.height = (t.scrollHeight + 2) + 'px';
  }
  function autosizeAll(root) { root.querySelectorAll('textarea.cell').forEach(autosize); }

  const KIND_TITLES = { h1: 'Heading 1', h2: 'Heading 2', h3: 'Heading 3', h4: 'Heading 4', p: 'Paragraph', s: 'Sentence, continues previous paragraph', x: 'Excluded from export' };
  const KIND_LABEL = { h1: 'H1', h2: 'H2', h3: 'H3', h4: 'H4', p: '¶', s: '↳', x: '✕' };

  function placeholderFor(kind) {
    if (Model.isHeading(kind)) return 'Heading…';
    if (kind === 's') return 'Next sentence of the paragraph…';
    if (kind === 'x') return 'Excluded note…';
    return 'One argument per row. Markdown works; "# " makes a heading, Tab indents.';
  }

  /** Row indexes shown in Draft view given the set of collapsed row ids. */
  function visibleIndexes(sheet, collapsed) {
    const out = [];
    const rows = sheet.rows;
    for (let i = 0; i < rows.length; i++) {
      out.push(i);
      if (collapsed && collapsed.has(rows[i]._id)) i = Model.sectionEnd(rows, i) - 1;
    }
    return out;
  }

  const fmt = n => Number(n).toLocaleString('en-US').replace(/,/g, ' ');
  /** The small grey line under a card: row counts, section totals, last edit. */
  function metaLine(ctx, sheet, row, i, sections) {
    const doc = ctx.doc;
    const parts = [];
    if (ctx.showCounts) {
      const rc = Model.rowCounts(doc, sheet, row);
      parts.push(`${fmt(rc.words)} w · ${fmt(rc.chars)} c`);
      const sec = sections[i];
      if (sec) parts.push(`section ${fmt(sec.words)} w · ${fmt(sec.chars)} c · ${sec.rows} rows`);
    }
    if (doc.settings.trackUpdated && row.updated) parts.push(row.updated);
    if (doc.settings.trackAuthor && row.author) parts.push(row.author);
    return parts.join(' '); // em spaces: ordinary spaces would collapse to one
  }
  function setMeta(card, text) {
    let m = card.querySelector('.meta');
    if (!text) { if (m) m.remove(); return; }
    if (!m) { m = el('div', { class: 'meta' }); card.querySelector('.body').appendChild(m); }
    m.textContent = text;
  }
  /** Recompute every card's meta line (section totals change when any row changes). */
  function refreshMeta(root, ctx) {
    const sheet = ctx.doc.sheets[ctx.si];
    if (!sheet || sheet.kind !== 'chapter') return;
    const sections = Model.sectionCounts(ctx.doc, sheet);
    root.querySelectorAll('.card').forEach(c => {
      const i = +c.dataset.i; const row = sheet.rows[i];
      if (row) setMeta(c, metaLine(ctx, sheet, row, i, sections));
    });
  }

  // ---------- Draft ----------
  function renderDraft(root, ctx) {
    const { doc, si } = ctx;
    const sheet = doc.sheets[si];
    root.innerHTML = '';
    if (!sheet) return;
    if (sheet.kind !== 'chapter') { root.appendChild(dataSheetTable(sheet, ctx)); return; }
    const num = Model.numbering(doc, si);
    const hidden = ctx.hiddenColumns || new Set();
    const side = ctx.showSide ? Model.sideColumns(doc, sheet).filter(c => !hidden.has(c)) : [];
    const list = el('div', { class: 'cards' + (side.length ? ' has-side' : '') });
    const sections = Model.sectionCounts(doc, sheet);
    for (const i of visibleIndexes(sheet, ctx.collapsed)) list.appendChild(card(ctx, sheet, sheet.rows[i], i, num, side, sections));
    root.appendChild(list);
    autosizeAll(root);
  }

  function card(ctx, sheet, row, i, num, side, sections) {
    const main = ctx.doc.mainColumn;
    const kind = Model.normKind(row.kind);
    const text = row[main] || '';
    const collapsible = Model.isCollapsible(sheet.rows, i);
    const collapsed = collapsible && ctx.collapsed.has(row._id);
    const hidden = collapsed ? Model.sectionEnd(sheet.rows, i) - i - 1 : 0;
    const indent = num.indents[i] || 0;
    const c = el('div', { class: 'card kind-' + kind + (side.length ? ' with-side' : '') + (collapsed ? ' collapsed' : '') + (indent ? ' indented' : ''), 'data-i': i });
    const gutter = el('div', { class: 'gutter' },
      el('span', { class: 'handle', draggable: 'true', title: 'Drag to move with its sub-rows (Alt+↑/↓)' }, '⋮⋮'),
      collapsible
        ? el('button', { class: 'collapse', type: 'button', title: (collapsed ? 'Expand' : 'Collapse') + ' (Ctrl+.)' }, collapsed ? '▸' : '▾')
        : el('span', { class: 'collapse none' }, ''),
      el('button', { class: 'kindbadge', type: 'button', title: KIND_TITLES[kind] + ' — click or Ctrl+Enter to change, Alt+Shift+←/→ to promote/demote' }, KIND_LABEL[kind]),
      el('span', { class: 'num' + (num.warnings[i] ? ' warn' : ''), title: num.warnings[i] ? 'Skipped level: numbered at the next allowed level' : 'Number' }, num.numbers[i]));
    const ta = el('textarea', { class: 'cell main', 'data-col': main, rows: '1', spellcheck: 'true', placeholder: placeholderFor(kind) });
    ta.value = text;
    const rendered = el('div', { class: 'rendered', html: MD.render(text) });
    const mainWrap = el('div', { class: 'mainwrap' + (text ? '' : ' empty') }, rendered, ta);
    const body = el('div', { class: 'body' }, mainWrap);
    if (indent) body.style.paddingLeft = (indent * 26) + 'px';
    if (collapsed) body.appendChild(el('button', { class: 'pill', type: 'button', title: 'Expand' }, `▸ ${hidden} hidden row${hidden === 1 ? '' : 's'}`));
    const meta = metaLine(ctx, sheet, row, i, sections);
    if (meta) body.appendChild(el('div', { class: 'meta' }, meta));
    c.append(gutter, body);
    if (side.length) {
      const sd = el('div', { class: 'side' });
      for (const col of side) {
        const t = el('textarea', { class: 'cell side-cell', 'data-col': col, rows: '1', placeholder: '…' });
        t.value = row[col] || '';
        sd.appendChild(el('label', { class: 'side-field' + (row[col] ? ' filled' : ''), title: col }, el('span', { class: 'lbl' }, col), t));
      }
      c.appendChild(sd);
    }
    return c;
  }

  /** Update one card's rendered view, empty flag and meta line from the model without rebuilding. */
  function refreshCard(root, ctx, i) {
    const doc = ctx.doc, sheet = doc.sheets[ctx.si];
    const c = root.querySelector(`.card[data-i="${i}"]`);
    if (!c || !sheet || !sheet.rows[i]) return;
    const row = sheet.rows[i];
    const text = row[doc.mainColumn] || '';
    const mw = c.querySelector('.mainwrap');
    mw.classList.toggle('empty', !text);
    mw.querySelector('.rendered').innerHTML = MD.render(text);
    c.querySelectorAll('.side-field').forEach(f => {
      const t = f.querySelector('textarea');
      f.classList.toggle('filled', !!t.value);
    });
    // own section only; other cards' totals are refreshed by refreshMeta when the row is left
    const sections = []; sections[i] = Model.sectionCounts(doc, sheet)[i];
    setMeta(c, metaLine(ctx, sheet, row, i, sections));
  }

  function dataSheetTable(sheet, ctx) {
    const wrap = el('div', { class: 'datasheet' });
    wrap.appendChild(el('p', { class: 'notice' },
      `“${sheet.name}” is a data sheet: it has no “${ctx.doc.mainColumn}” column, so it is shown read-only here and written back unchanged. `,
      'Add a column named ', el('code', {}, ctx.doc.mainColumn), ' in Excel to make it a chapter.'));
    const cells = sheet.cells || [];
    const t = el('table', { class: 'grid readonly' });
    const max = Math.min(cells.length, 500);
    for (let r = 0; r < max; r++) {
      const tr = el('tr', {});
      (cells[r] || []).forEach(v => tr.appendChild(el(r === 0 ? 'th' : 'td', {}, v)));
      t.appendChild(tr);
    }
    if (cells.length > max) wrap.appendChild(el('p', { class: 'muted' }, `Showing the first ${max} of ${cells.length} rows.`));
    wrap.appendChild(el('div', { class: 'scroll' }, t));
    return wrap;
  }

  // ---------- Grid ----------
  function renderGrid(root, ctx) {
    const { doc, si } = ctx;
    const sheet = doc.sheets[si];
    root.innerHTML = '';
    if (!sheet) return;
    if (sheet.kind !== 'chapter') { root.appendChild(dataSheetTable(sheet, ctx)); return; }
    const num = Model.numbering(doc, si);
    const table = el('table', { class: 'grid' });
    const thead = el('thead', {});
    const hr = el('tr', {});
    hr.appendChild(el('th', { class: 'handle-col' }, ''));
    for (const col of sheet.columns) {
      const isMain = col === doc.mainColumn, isRes = RESERVED.includes(col), isMeta = Model.isMeta(doc, col) || Model.isComputed(doc, col);
      const draggable = !isRes && !isMeta;
      const th = el('th', { class: (isMain ? 'main' : '') + (isRes ? ' reserved' : '') + (col === 'no' ? ' num' : '') + (isMeta ? ' meta' : '') + (draggable ? ' col-drag' : ''), 'data-col': col, draggable: draggable ? 'true' : null });
      if (!isRes && !isMeta && ctx.editColumn === col) {
        th.appendChild(el('input', { type: 'text', class: 'colname-edit', 'data-col': col, value: col, spellcheck: 'false' }));
      } else {
        const titles = { no: 'Computed numbering, written to the file on save', kind: 'Row kind', indent: 'Indent level (Tab / Shift+Tab)', updated: 'Last edited (maintained by SheetWriter)', author: 'Last editor (maintained by SheetWriter)', words: 'Words in the counted columns (computed, written on save)', chars: 'Characters in the counted columns (computed, written on save)' };
        th.appendChild(el('span', { class: 'colname' + (isRes || isMeta ? '' : ' editable'), 'data-col': col, title: (isRes || isMeta) ? titles[col] : 'Click to rename, drag the header to reorder (all sheets)' }, col, isMain ? ' ★' : ''));
      }
      if (!isRes && !isMeta) {
        const ops = el('span', { class: 'colops' });
        ops.appendChild(el('button', { type: 'button', 'data-action': 'left', 'data-col': col, title: 'Move left' }, '◀'));
        ops.appendChild(el('button', { type: 'button', 'data-action': 'right', 'data-col': col, title: 'Move right' }, '▶'));
        if (!isMain) ops.appendChild(el('button', { type: 'button', 'data-action': 'main', 'data-col': col, title: 'Make this the main text column' }, '★'));
        if (!isMain) ops.appendChild(el('button', { type: 'button', class: 'danger', 'data-action': 'delete', 'data-col': col, title: 'Delete column' }, '✕'));
        th.appendChild(ops);
      }
      hr.appendChild(th);
    }
    hr.appendChild(el('th', { class: 'addcol' }, el('button', { type: 'button', class: 'addcol-btn', 'data-action': 'add', title: 'Add a column' }, '+ column')));
    thead.appendChild(hr);
    table.appendChild(thead);
    const tbody = el('tbody', {});
    sheet.rows.forEach((row, i) => {
      const tr = el('tr', { class: 'kind-' + Model.normKind(row.kind), 'data-i': i });
      tr.appendChild(el('td', { class: 'handle-col' }, el('span', { class: 'handle', draggable: 'true', title: 'Drag to move (Alt+↑/↓)' }, '⋮⋮')));
      for (const col of sheet.columns) {
        if (col === 'no') { tr.appendChild(el('td', { class: 'num' + (num.warnings[i] ? ' warn' : '') }, num.numbers[i])); continue; }
        if (col === 'kind') {
          const sel = el('select', { class: 'kind-select', 'data-col': 'kind', title: KIND_TITLES[Model.normKind(row.kind)] });
          for (const k of Model.KINDS) sel.appendChild(el('option', { value: k, selected: k === Model.normKind(row.kind) }, k));
          tr.appendChild(el('td', { class: 'kind' }, sel));
          continue;
        }
        if (col === 'indent') { tr.appendChild(el('td', { class: 'indent' }, Model.indentOf(row) || '')); continue; }
        if (Model.isMeta(doc, col)) { tr.appendChild(el('td', { class: 'meta' }, row[col] || '')); continue; }
        if (Model.isComputed(doc, col)) { const rc = Model.rowCounts(doc, sheet, row); tr.appendChild(el('td', { class: 'meta num' }, fmt(col === 'words' ? rc.words : rc.chars))); continue; }
        const t = el('textarea', { class: 'cell' + (col === doc.mainColumn ? ' main' : ''), 'data-col': col, rows: '1' });
        t.value = row[col] || '';
        const td = el('td', { class: col === doc.mainColumn ? 'main' : '' }, t);
        if (col === doc.mainColumn && num.indents[i]) td.style.paddingLeft = (4 + num.indents[i] * 18) + 'px';
        tr.appendChild(td);
      }
      tr.appendChild(el('td', { class: 'rowops' },
        el('button', { type: 'button', class: 'row-add', 'data-action': 'row-add', title: 'Insert a row below' }, '+ row'),
        el('button', { type: 'button', class: 'danger', 'data-action': 'row-del', title: 'Delete this row' }, '✕')));
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    root.appendChild(el('div', { class: 'scroll' }, table));
    autosizeAll(root);
    applyFreeze(table, doc.settings.freezeColumns);
    const edit = root.querySelector('input.colname-edit');
    if (edit) { edit.focus(); edit.select(); }
  }

  /** Freeze the handle column plus the first n sheet columns (sticky left offsets measured from the header). */
  function applyFreeze(table, n) {
    if (!table) return;
    n = Math.max(0, Math.min(+n || 0, 10));
    const headCells = [...table.querySelectorAll('thead th')];
    const count = Math.min(n + 1, Math.max(0, headCells.length - 1)); // never the "+ column" header
    let left = 0;
    for (let k = 0; k < count; k++) {
      const w = headCells[k].getBoundingClientRect().width;
      for (const tr of table.rows) {
        const c = tr.children[k];
        if (!c) continue;
        c.classList.add('frozen');
        c.classList.toggle('frozen-last', k === count - 1);
        c.style.left = left + 'px';
      }
      left += w;
    }
  }

  // ---------- Read ----------
  function renderRead(root, ctx) {
    const { doc, si } = ctx;
    root.innerHTML = '';
    const sheet = doc.sheets[si];
    const cols = [...new Set(doc.sheets.filter(s => s.kind === 'chapter').flatMap(s => Model.userColumns(doc, s)))];
    const bar = el('div', { class: 'readbar' });
    const colSel = el('select', { id: 'read-column' });
    cols.forEach(c => colSel.appendChild(el('option', { value: c, selected: c === (ctx.readColumn || doc.mainColumn) }, c)));
    const scope = el('select', { id: 'read-scope' },
      el('option', { value: 'sheet', selected: ctx.readScope !== 'all' }, 'This chapter'),
      el('option', { value: 'all', selected: ctx.readScope === 'all' }, 'Whole workbook'));
    const numSel = el('select', { id: 'read-numbering' },
      el('option', { value: '', selected: !ctx.readNumbering }, 'none'),
      el('option', { value: 'headings', selected: ctx.readNumbering === 'headings' }, 'headings'),
      el('option', { value: 'all', selected: ctx.readNumbering === 'all' }, 'headings and paragraphs'));
    const indSel = el('select', { id: 'read-indented' },
      el('option', { value: 'paragraphs', selected: ctx.readIndented !== 'lists' }, 'paragraphs'),
      el('option', { value: 'lists', selected: ctx.readIndented === 'lists' }, 'nested lists'));
    bar.append(el('label', {}, 'Column ', colSel), el('label', {}, 'Scope ', scope), el('label', {}, 'Numbering ', numSel), el('label', {}, 'Indented rows as ', indSel));
    root.appendChild(bar);
    const opts = {
      column: ctx.readColumn || doc.mainColumn,
      scope: ctx.readScope === 'all' ? 'all' : si,
      numbering: ctx.readNumbering || false,
      indented: ctx.readIndented || 'paragraphs',
      sheetTitles: ctx.readScope === 'all' && doc.settings.numbering === 'per-sheet' && Model.chapterSheets(doc).length > 1,
    };
    const art = el('article', { class: 'read', html: MD.render(Exporter.toMarkdown(doc, opts)) });
    if (ctx.readScope !== 'all' && sheet && sheet.kind !== 'chapter') art.innerHTML = '<p class="muted">Data sheet: nothing to read here.</p>';
    // Tag rendered headings with their rows so the table of contents can scroll to them.
    const order = Exporter.headingOrder(doc, opts);
    const hs = art.querySelectorAll('h1, h2, h3, h4, h5, h6');
    if (hs.length === order.length) hs.forEach((h, k) => { const o = order[k]; if (o.si != null) { h.dataset.si = o.si; h.dataset.i = o.i; } else h.dataset.sheet = o.sheetTitle; });
    root.appendChild(art);
  }

  // ---------- Table of contents ----------
  function renderToc(root, ctx) {
    const { doc, si } = ctx;
    root.innerHTML = '';
    const head = el('div', { class: 'toc-head' },
      el('span', { class: 'toc-title' }, 'Contents'),
      el('select', { id: 'toc-scope', title: 'Which sheets to list' },
        el('option', { value: 'sheet', selected: ctx.tocScope !== 'all' }, 'This sheet'),
        el('option', { value: 'all', selected: ctx.tocScope === 'all' }, 'All sheets')),
      el('button', { type: 'button', id: 'toc-close', title: 'Close the table of contents' }, '×'));
    root.appendChild(head);
    const list = el('div', { class: 'toc-list' });
    const groups = Model.tocEntries(doc, ctx.tocScope === 'all' ? null : si);
    let any = false;
    for (const g of groups) {
      if (ctx.tocScope === 'all') list.appendChild(el('button', { type: 'button', class: 'toc-sheet' + (g.si === si ? ' active' : ''), 'data-si': g.si, title: 'Open this sheet' }, g.name));
      for (const e of g.entries) {
        any = true;
        const cur = ctx.current && ctx.current.si === e.si && ctx.current.i === e.i;
        list.appendChild(el('button', { type: 'button', class: 'toc-item level-' + Math.min(e.level, 6) + (cur ? ' current' : ''), 'data-si': e.si, 'data-i': e.i, title: e.text },
          el('span', { class: 'toc-num' }, e.number),
          el('span', { class: 'toc-text' }, e.text),
          ctx.showCounts ? el('span', { class: 'toc-count' }, fmt(e.words) + ' w') : null));
      }
    }
    if (!any) list.appendChild(el('p', { class: 'muted toc-empty' }, ctx.tocScope === 'all' ? 'No headings yet.' : 'No headings in this sheet yet. Type "# " at the start of a row to make one.'));
    root.appendChild(list);
    const cur = list.querySelector('.current');
    if (cur) cur.scrollIntoView({ block: 'nearest' });
  }
  function updateTocCurrent(root, current) {
    root.querySelectorAll('.toc-item').forEach(b => b.classList.toggle('current', !!current && +b.dataset.si === current.si && +b.dataset.i === current.i));
    const c = root.querySelector('.toc-item.current');
    if (c) c.scrollIntoView({ block: 'nearest' });
  }

  // ---------- Tabs ----------
  function renderTabs(root, ctx) {
    const { doc, si } = ctx;
    root.innerHTML = '';
    let ci = 0;
    doc.sheets.forEach((s, i) => {
      const isCh = s.kind === 'chapter';
      if (isCh) ci++;
      if (ctx.editSheet === i) {
        root.appendChild(el('input', { type: 'text', class: 'tab tab-edit', 'data-i': i, value: s.name, spellcheck: 'false' }));
        return;
      }
      const b = el('button', {
        type: 'button', class: 'tab' + (i === si ? ' active' : '') + (isCh ? '' : ' data'), 'data-i': i, draggable: 'true',
        title: isCh ? `Chapter ${ci}: ${s.name} — double-click to rename, drag to reorder` : `Data sheet: ${s.name} (read-only here, written back unchanged)`,
      }, isCh ? '' : '⊞ ', s.name);
      root.appendChild(b);
    });
    root.appendChild(el('button', { type: 'button', class: 'tab add', id: 'tab-add', title: 'Add a chapter after the current one' }, '+'));
    root.appendChild(el('button', { type: 'button', class: 'tab menu', id: 'tab-menu', title: 'Sheet actions' }, '⋯'));
    const edit = root.querySelector('input.tab-edit');
    if (edit) { edit.focus(); edit.select(); }
  }

  return { el, autosize, autosizeAll, visibleIndexes, renderDraft, refreshCard, refreshMeta, renderGrid, applyFreeze, renderRead, renderToc, updateTocCurrent, renderTabs };
})();

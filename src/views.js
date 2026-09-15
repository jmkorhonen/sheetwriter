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
      const target = Model.rowTarget(doc, sheet, row);
      if (sec) parts.push(`section ${fmt(sec.words)}${target ? ` / ${fmt(target)}` : ''} w${target ? ` (${Math.round(100 * sec.words / target)} %)` : ''} · ${fmt(sec.chars)} c · ${sec.rows} rows`);
      else if (target) parts.push(`target ${fmt(target)} w (${Math.round(100 * rc.words / target)} %)`);
    }
    if (doc.settings.trackUpdated && row['.updated']) parts.push(row['.updated']);
    if (doc.settings.trackAuthor && row['.author']) parts.push(row['.author']);
    return parts.join(' '); // em spaces: ordinary spaces would collapse to one
  }
  function statusChip(doc, sheet, row) {
    const sc = Model.statusColumn(doc, sheet);
    const v = sc ? String(row[sc] || '').trim() : '';
    return v ? el('span', { class: 'chip c' + Model.colorIndex(v), title: sc }, v) : null;
  }
  function setMeta(card, text, chip) {
    let m = card.querySelector('.meta');
    if (!text && !chip) { if (m) m.remove(); return; }
    if (!m) { m = el('div', { class: 'meta' }); card.querySelector('.body').appendChild(m); }
    m.innerHTML = '';
    m.appendChild(chip || el('span', {}));
    m.appendChild(el('span', { class: 'meta-text' }, text));
  }
  /** Recompute every card's meta line (section totals change when any row changes). */
  function refreshMeta(root, ctx) {
    const sheet = ctx.doc.sheets[ctx.si];
    if (!sheet || sheet.kind !== 'chapter') return;
    const sections = Model.sectionCounts(ctx.doc, sheet);
    root.querySelectorAll('.card').forEach(c => {
      const i = +c.dataset.i; const row = sheet.rows[i];
      if (row) setMeta(c, metaLine(ctx, sheet, row, i, sections), statusChip(ctx.doc, sheet, row));
    });
  }
  /** Filter box shared by Draft and Grid. Query: plain text, or column:value (kind:h2, status:todo). */
  function filterBar(ctx) {
    return el('div', { class: 'filterbar' },
      el('input', { type: 'search', id: 'row-filter', placeholder: 'Filter rows… text, or column:value (status:todo, kind:h2)', value: ctx.rowFilter || '', spellcheck: 'false' }),
      el('span', { id: 'row-filter-info', class: 'muted' }, ''));
  }
  /** Hide rows (cards or grid rows) that do not match the filter. */
  function applyRowFilter(root, ctx) {
    const q = String(ctx.rowFilter || '').trim();
    const sheet = ctx.doc.sheets[ctx.si];
    if (!sheet || sheet.kind !== 'chapter') return;
    const holders = root.querySelectorAll('.card[data-i], tbody tr[data-i]');
    const num = q ? Model.numbering(ctx.doc, ctx.si).numbers : null;
    let col = null, needle = q.toLowerCase();
    const m = /^([^\s:]+):(.*)$/.exec(q);
    if (m) {
      const name = m[1].toLowerCase();
      col = name === 'kind' || name === '.kind' ? '.kind' : name === 'no' || name === '.no' ? '.no' : sheet.columns.find(c => c.toLowerCase() === name) || null;
      if (col) needle = m[2].trim().toLowerCase();
    }
    const gridHidden = ctx.gridHidden || new Set();
    const cols = sheet.columns.filter(c => c !== '.no' && !gridHidden.has(c));
    let shown = 0, all = 0;
    holders.forEach(h => {
      all++;
      const i = +h.dataset.i, row = sheet.rows[i];
      let hit = true;
      if (q && col === '.no') hit = num[i].startsWith(needle);
      else if (q && col === '.kind') hit = Model.normKind(row['.kind']) === needle || (needle === 'h' && Model.isHeading(row['.kind']));
      else if (q && col) hit = String(row[col] || '').toLowerCase().includes(needle);
      else if (q) hit = num[i].startsWith(needle) || cols.some(c => String(row[c] || '').toLowerCase().includes(needle));
      h.classList.toggle('filtered', !hit);
      if (hit) shown++;
    });
    const info = root.querySelector('#row-filter-info');
    if (info) info.textContent = q ? `${shown} of ${all} rows` : '';
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
    if (ctx.rowFilter || ctx.showFilter) root.appendChild(filterBar(ctx));
    root.appendChild(list);
    autosizeAll(root);
    applyRowFilter(root, ctx);
  }

  function card(ctx, sheet, row, i, num, side, sections) {
    const main = ctx.doc.mainColumn;
    const kind = Model.normKind(row['.kind']);
    const text = row[main] || '';
    const collapsible = Model.isCollapsible(sheet.rows, i);
    const collapsed = collapsible && ctx.collapsed.has(row._id);
    const hidden = collapsed ? Model.sectionEnd(sheet.rows, i) - i - 1 : 0;
    const indent = num.indents[i] || 0;
    const selected = ctx.selected && ctx.selected.has(row._id);
    const c = el('div', { class: 'card kind-' + kind + (side.length ? ' with-side' : '') + (collapsed ? ' collapsed' : '') + (indent ? ' indented' : '') + (selected ? ' selected' : ''), 'data-i': i });
    const gutter = el('div', { class: 'gutter' },
      el('span', { class: 'handle', draggable: 'true', title: 'Click to select the row (Shift: range, Ctrl: add), drag to move' }, '⋮⋮'),
      collapsible
        ? el('button', { class: 'collapse', type: 'button', title: (collapsed ? 'Expand' : 'Collapse') + ' (Ctrl+.)' }, collapsed ? '▸' : '▾')
        : el('span', { class: 'collapse none' }, ''),
      el('button', { class: 'kindbadge', type: 'button', title: KIND_TITLES[kind] + ' — click or Ctrl+Enter to change, Alt+Shift+←/→ to promote/demote' }, KIND_LABEL[kind]),
      el('span', { class: 'num' + (num.warnings[i] ? ' warn' : ''), title: num.warnings[i] ? 'Skipped level: numbered at the next allowed level' : 'Click to select the row (Shift: range, Ctrl: add)' }, num.numbers[i]));
    const ta = el('textarea', { class: 'cell main', 'data-col': main, rows: '1', spellcheck: 'true', placeholder: placeholderFor(kind) });
    ta.value = text;
    const rendered = el('div', { class: 'rendered', html: MD.render(text) });
    const mainWrap = el('div', { class: 'mainwrap' + (text ? '' : ' empty') }, rendered, ta);
    const body = el('div', { class: 'body' }, mainWrap);
    if (indent) body.style.paddingLeft = (indent * 26) + 'px';
    if (collapsed) body.appendChild(el('button', { class: 'pill', type: 'button', title: 'Expand' }, `▸ ${hidden} hidden row${hidden === 1 ? '' : 's'}`));
    c.append(gutter, body);
    setMeta(c, metaLine(ctx, sheet, row, i, sections), statusChip(ctx.doc, sheet, row));
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
    setMeta(c, metaLine(ctx, sheet, row, i, sections), statusChip(doc, sheet, row));
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
    const gridHidden = ctx.gridHidden || new Set();
    const columns = sheet.columns.filter(c => c === '.no' || c === doc.mainColumn || !gridHidden.has(c));
    const hiddenCount = sheet.columns.length - columns.length;
    if (ctx.rowFilter || ctx.showFilter) root.appendChild(filterBar(ctx));
    const statusCol = Model.statusColumn(doc, sheet), targetCol = Model.targetColumn(doc, sheet);
    const table = el('table', { class: 'grid' });
    // Fixed layout with explicit widths so columns can be resized; the table is as wide as its columns.
    const cg = el('colgroup', {});
    cg.appendChild(el('col', { style: 'width:22px' }));
    let total = 22;
    for (const col of columns) { const w = Model.columnWidth(doc, col); total += w; cg.appendChild(el('col', { 'data-col': col, style: `width:${w}px` })); }
    const addW = hiddenCount ? 200 : 110; // room for "+ column" and the "N hidden" pill side by side
    cg.appendChild(el('col', { style: `width:${addW}px` })); total += addW;
    table.appendChild(cg);
    table.style.width = total + 'px';
    const thead = el('thead', {});
    const hr = el('tr', {});
    hr.appendChild(el('th', { class: 'handle-col' }, ''));
    for (const col of columns) {
      const isMain = col === doc.mainColumn, isRes = RESERVED.includes(col), isMeta = Model.isMeta(doc, col) || Model.isComputed(doc, col) || col === Model.IDENT;
      const draggable = !isRes && !isMeta;
      const th = el('th', { class: (isMain ? 'main' : '') + (isRes ? ' reserved' : '') + (col === '.no' ? ' num' : '') + (isMeta ? ' meta' : '') + (draggable ? ' col-drag' : ''), 'data-col': col, draggable: draggable ? 'true' : null });
      if (!isRes && !isMeta && ctx.editColumn === col) {
        th.appendChild(el('input', { type: 'text', class: 'colname-edit', 'data-col': col, value: col, spellcheck: 'false' }));
      } else {
        const titles = { '.no': 'Computed numbering, written to the file on save', '.kind': 'Row kind', '.indent': 'Indent level (Tab / Shift+Tab)', '.updated': 'Last edited (maintained by SheetWriter)', '.author': 'Last editor (maintained by SheetWriter)', '.words': 'Words in the counted columns (computed, written on save)', '.chars': 'Characters in the counted columns (computed, written on save)' };
        const role = isMain ? ' ★' : col === statusCol ? ' ●' : col === targetCol ? ' ◎' : '';
        const roleTitle = isMain ? ' — main text column' : col === statusCol ? ' — status column (chips)' : col === targetCol ? ' — word targets' : '';
        th.appendChild(el('span', { class: 'colname' + (isRes || isMeta ? '' : ' editable'), 'data-col': col, title: (isRes || isMeta) ? titles[col] : 'Click to rename, drag the header to reorder (all sheets)' + roleTitle }, col, role));
      }
      if (!isRes && !isMeta) {
        const ops = el('span', { class: 'colops' });
        ops.appendChild(el('button', { type: 'button', 'data-action': 'left', 'data-col': col, title: 'Move left' }, '◀'));
        ops.appendChild(el('button', { type: 'button', 'data-action': 'right', 'data-col': col, title: 'Move right' }, '▶'));
        if (!isMain) ops.appendChild(el('button', { type: 'button', 'data-action': 'main', 'data-col': col, title: 'Make this the main text column' }, '★'));
        if (!isMain) ops.appendChild(el('button', { type: 'button', 'data-action': 'hide', 'data-col': col, title: 'Hide this column in Grid view (Columns ▾ shows it again)' }, '–'));
        if (!isMain) ops.appendChild(el('button', { type: 'button', class: 'danger', 'data-action': 'delete', 'data-col': col, title: 'Delete column' }, '✕'));
        th.appendChild(ops);
      } else if (col !== '.no') {
        th.appendChild(el('span', { class: 'colops' }, el('button', { type: 'button', 'data-action': 'hide', 'data-col': col, title: 'Hide this column in Grid view (Columns ▾ shows it again)' }, '–')));
      }
      th.appendChild(el('span', { class: 'col-resize', 'data-col': col, title: 'Drag to resize, double-click to reset' }));
      hr.appendChild(th);
    }
    const addTh = el('th', { class: 'addcol' }, el('button', { type: 'button', class: 'addcol-btn', 'data-action': 'add', title: 'Add a column' }, '+ column'));
    if (hiddenCount) addTh.appendChild(el('button', { type: 'button', class: 'hidden-pill', 'data-action': 'showhidden', title: 'Hidden: ' + sheet.columns.filter(c => !columns.includes(c)).join(', ') + '. Click to choose.' }, `${hiddenCount} hidden ▾`));
    hr.appendChild(addTh);
    thead.appendChild(hr);
    table.appendChild(thead);
    const tbody = el('tbody', {});
    for (const i of visibleIndexes(sheet, ctx.collapsed)) {
      const row = sheet.rows[i];
      const collapsible = Model.isCollapsible(sheet.rows, i);
      const collapsed = collapsible && ctx.collapsed && ctx.collapsed.has(row._id);
      const tr = el('tr', { class: 'kind-' + Model.normKind(row['.kind']) + (ctx.selected && ctx.selected.has(row._id) ? ' selected' : '') + (collapsed ? ' collapsed' : ''), 'data-i': i });
      tr.appendChild(el('td', { class: 'handle-col' }, el('span', { class: 'handle', draggable: 'true', title: 'Click to select the row (Shift: range, Ctrl: add), drag to move' }, '⋮⋮')));
      for (const col of columns) {
        if (col === '.no') {
          const td = el('td', { class: 'num' + (num.warnings[i] ? ' warn' : ''), title: 'Click to select the row (Shift: range, Ctrl: add)' });
          td.appendChild(collapsible
            ? el('button', { class: 'collapse', type: 'button', title: (collapsed ? 'Expand' : 'Collapse') + ' (Ctrl+.)' }, collapsed ? '▸' : '▾')
            : el('span', { class: 'collapse none' }, ''));
          td.appendChild(el('span', {}, num.numbers[i]));
          if (collapsed) td.appendChild(el('span', { class: 'hidden-rows', title: 'Rows hidden under this one' }, ` +${Model.sectionEnd(sheet.rows, i) - i - 1}`));
          tr.appendChild(td); continue;
        }
        if (col === '.kind') {
          const sel = el('select', { class: 'kind-select', 'data-col': '.kind', title: KIND_TITLES[Model.normKind(row['.kind'])] });
          for (const k of Model.KINDS) sel.appendChild(el('option', { value: k, selected: k === Model.normKind(row['.kind']) }, k));
          tr.appendChild(el('td', { class: 'kind' }, sel));
          continue;
        }
        if (col === '.indent') { tr.appendChild(el('td', { class: 'indent' }, Model.indentOf(row) || '')); continue; }
        if (Model.isMeta(doc, col) || col === Model.IDENT) { tr.appendChild(el('td', { class: 'meta' }, row[col] || '')); continue; }
        if (Model.isComputed(doc, col)) { const rc = Model.rowCounts(doc, sheet, row); tr.appendChild(el('td', { class: 'meta num' }, fmt(col === '.words' ? rc.words : rc.chars))); continue; }
        const t = el('textarea', { class: 'cell' + (col === doc.mainColumn ? ' main' : ''), 'data-col': col, rows: '1' });
        t.value = row[col] || '';
        const chipClass = col === statusCol && String(row[col] || '').trim() ? ' status c' + Model.colorIndex(String(row[col]).trim()) : '';
        const td = el('td', { class: (col === doc.mainColumn ? 'main' : '') + chipClass }, t);
        if (col === doc.mainColumn && num.indents[i]) td.style.paddingLeft = (4 + num.indents[i] * 18) + 'px';
        tr.appendChild(td);
      }
      tr.appendChild(el('td', { class: 'rowops' },
        el('button', { type: 'button', class: 'row-add', 'data-action': 'row-add', title: 'Insert a row below' }, '+ row'),
        el('button', { type: 'button', class: 'danger', 'data-action': 'row-del', title: 'Delete this row' }, '✕')));
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    root.appendChild(el('div', { class: 'scroll' }, table));
    autosizeAll(root);
    applyFreeze(table, doc.settings.freezeColumns);
    applyRowFilter(root, ctx);
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
    const art = el('article', { class: 'read' });
    if (ctx.readScope !== 'all' && sheet && sheet.kind !== 'chapter') { art.innerHTML = '<p class="muted">Data sheet: nothing to read here.</p>'; root.appendChild(art); return; }
    // Rendered block by block so every heading and paragraph knows its row (click to edit, scroll spy, contents pane).
    let listBuf = null;
    const flushList = () => {
      if (!listBuf) return;
      art.appendChild(el('div', { class: 'rblock rlist', 'data-si': listBuf.si, 'data-i': listBuf.i, title: 'Click to edit', html: MD.render(listBuf.md.join('\n')) }));
      listBuf = null;
    };
    for (const b of Exporter.toBlocks(doc, opts)) {
      if (b.type === 'para' && b.list) { if (!listBuf) listBuf = { si: b.si, i: b.i, md: [] }; listBuf.md.push(Exporter.blockMarkdown(b)); continue; }
      flushList();
      if (b.type === 'sheetTitle') art.appendChild(el('h1', { class: 'rblock rsheet', 'data-sheet': b.si, html: MD.renderInline(b.text) }));
      else if (b.type === 'heading') art.appendChild(el('h' + b.level, { class: 'rblock', 'data-si': b.si, 'data-i': b.i, title: 'Click to edit', html: MD.renderInline(b.text) }));
      else if (b.type === 'para') art.appendChild(el('div', { class: 'rblock rpara' + (b.indent ? ' indent-' + Math.min(b.indent, 6) : ''), 'data-si': b.si, 'data-i': b.i, title: 'Click to edit', html: MD.render(b.text) }));
      else if (b.type === 'side') art.appendChild(el('div', { class: 'rblock rside', 'data-si': b.si, 'data-i': b.i, html: MD.render('> **' + b.label + ':** ' + b.values.map(v => v.replace(/\n/g, ' ')).join('\n> ')) }));
    }
    flushList();
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
        const over = e.target && e.words > e.target;
        list.appendChild(el('button', { type: 'button', class: 'toc-item level-' + Math.min(e.level, 6) + (cur ? ' current' : ''), 'data-si': e.si, 'data-i': e.i, title: e.text + (e.target ? ` — target ${fmt(e.target)} words` : '') },
          el('span', { class: 'toc-num' }, e.number),
          el('span', { class: 'toc-text' }, e.text),
          ctx.showCounts ? el('span', { class: 'toc-count' + (over ? ' over' : '') }, e.target ? `${fmt(e.words)} / ${fmt(e.target)} w` : fmt(e.words) + ' w') : null));
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

  return { el, autosize, autosizeAll, visibleIndexes, renderDraft, refreshCard, refreshMeta, renderGrid, applyFreeze, applyRowFilter, renderRead, renderToc, updateTocCurrent, renderTabs };
})();

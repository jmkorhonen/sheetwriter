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

  function metaText(doc, row) {
    const parts = [];
    if (doc.settings.trackUpdated && row.updated) parts.push(row.updated);
    if (doc.settings.trackAuthor && row.author) parts.push(row.author);
    return parts.join(' · ');
  }

  // ---------- Draft ----------
  function renderDraft(root, ctx) {
    const { doc, si } = ctx;
    const sheet = doc.sheets[si];
    root.innerHTML = '';
    if (!sheet) return;
    if (sheet.kind !== 'chapter') { root.appendChild(dataSheetTable(sheet, ctx)); return; }
    const num = Model.numbering(doc, si);
    const side = ctx.showSide ? Model.sideColumns(doc, sheet) : [];
    const list = el('div', { class: 'cards' + (side.length ? ' has-side' : '') });
    for (const i of visibleIndexes(sheet, ctx.collapsed)) list.appendChild(card(ctx, sheet, sheet.rows[i], i, num, side));
    root.appendChild(list);
    autosizeAll(root);
  }

  function card(ctx, sheet, row, i, num, side) {
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
    const meta = metaText(ctx.doc, row);
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

  /** Update one card's rendered view + empty flag from the model without rebuilding. */
  function refreshCard(root, doc, sheet, i) {
    const c = root.querySelector(`.card[data-i="${i}"]`);
    if (!c) return;
    const row = sheet.rows[i];
    const text = row[doc.mainColumn] || '';
    const mw = c.querySelector('.mainwrap');
    mw.classList.toggle('empty', !text);
    mw.querySelector('.rendered').innerHTML = MD.render(text);
    c.querySelectorAll('.side-field').forEach(f => {
      const t = f.querySelector('textarea');
      f.classList.toggle('filled', !!t.value);
    });
    const meta = metaText(doc, row);
    let m = c.querySelector('.meta');
    if (meta && !m) { m = el('div', { class: 'meta' }); c.querySelector('.body').appendChild(m); }
    if (m) m.textContent = meta;
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
      const isMain = col === doc.mainColumn, isRes = RESERVED.includes(col), isMeta = Model.isMeta(doc, col);
      const th = el('th', { class: (isMain ? 'main' : '') + (isRes ? ' reserved' : '') + (col === 'no' ? ' num' : '') + (isMeta ? ' meta' : ''), 'data-col': col });
      if (!isRes && !isMeta && ctx.editColumn === col) {
        th.appendChild(el('input', { type: 'text', class: 'colname-edit', 'data-col': col, value: col, spellcheck: 'false' }));
      } else {
        const titles = { no: 'Computed numbering, written to the file on save', kind: 'Row kind', indent: 'Indent level (Tab / Shift+Tab)', updated: 'Last edited (maintained by SheetWriter)', author: 'Last editor (maintained by SheetWriter)' };
        th.appendChild(el('span', { class: 'colname' + (isRes || isMeta ? '' : ' editable'), 'data-col': col, title: (isRes || isMeta) ? titles[col] : 'Click to rename' }, col, isMain ? ' ★' : ''));
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
    const edit = root.querySelector('input.colname-edit');
    if (edit) { edit.focus(); edit.select(); }
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
    const md = Exporter.toMarkdown(doc, {
      column: ctx.readColumn || doc.mainColumn,
      scope: ctx.readScope === 'all' ? 'all' : si,
      numbering: ctx.readNumbering || false,
      indented: ctx.readIndented || 'paragraphs',
      sheetTitles: ctx.readScope === 'all' && Model.chapterSheets(doc).length > 1,
    });
    const art = el('article', { class: 'read', html: MD.render(md) });
    if (ctx.readScope !== 'all' && sheet && sheet.kind !== 'chapter') art.innerHTML = '<p class="muted">Data sheet: nothing to read here.</p>';
    root.appendChild(art);
  }

  // ---------- Tabs ----------
  function renderTabs(root, ctx) {
    const { doc, si } = ctx;
    root.innerHTML = '';
    let ci = 0;
    const multi = doc.settings.chapterPrefix && Model.chapterSheets(doc).length > 1;
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
      }, isCh ? (multi ? `${ci}. ` : '') : '⊞ ', s.name);
      root.appendChild(b);
    });
    root.appendChild(el('button', { type: 'button', class: 'tab add', id: 'tab-add', title: 'Add a chapter after the current one' }, '+'));
    root.appendChild(el('button', { type: 'button', class: 'tab menu', id: 'tab-menu', title: 'Sheet actions' }, '⋯'));
    const edit = root.querySelector('input.tab-edit');
    if (edit) { edit.focus(); edit.select(); }
  }

  return { el, autosize, autosizeAll, visibleIndexes, renderDraft, refreshCard, renderGrid, renderRead, renderTabs };
})();

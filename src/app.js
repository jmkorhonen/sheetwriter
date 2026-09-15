/* SheetWriter — application wiring: state, undo, file I/O, keyboard, events. */
(() => {
  const $ = s => document.querySelector(s);
  const viewRoot = $('#view'), tabsRoot = $('#tabs'), statusEl = $('#status'), tocRoot = $('#toc');
  const NO_COLLAPSE = new Set();

  // Editor preferences live in this browser, not in the workbook.
  const PREF_KEY = 'sheetwriter.prefs';
  const prefs = Object.assign({ enterMode: 'row', indentTrigger: '   ', theme: 'auto', autosaveCopy: { enabled: false, minutes: 2 } }, (() => { try { return JSON.parse(localStorage.getItem(PREF_KEY) || '{}'); } catch (e) { return {}; } })());
  function savePrefs() { try { localStorage.setItem(PREF_KEY, JSON.stringify(prefs)); } catch (e) { /* ignore */ } }

  const state = {
    doc: Model.ensureIds(Model.newDoc()), si: 0, view: 'draft', showSide: true,
    fileHandle: null, fileName: 'untitled.xlsx', sources: new Map(), dirty: false,
    readScope: 'sheet', readNumbering: '', readColumn: null, readIndented: 'paragraphs',
    focus: null, lastFocus: null, readPos: null,
    fileMtime: null, savedAt: null, copyHandle: null, copyName: '', copyAt: null, copyNeedsPermission: false,
    collapsed: new Set(), editColumn: null, editSheet: null,
    sel: new Set(), selAnchor: null, clipboard: null, // row selection (ids, current sheet) and the internal row clipboard
    // Display state; saved into the workbook on save and restored on load.
    ui: { side: true, hidden: new Set(), counts: true, toc: 'off', lastToc: 'sheet', gridHidden: new Set() },
    rowFilter: '', showFilter: false,
  };
  const history = { undo: [], redo: [] };
  let typingKey = null, typingAt = 0;
  const hasFS = !!window.showOpenFilePicker;

  // ---------- undo ----------
  const snapshot = () => JSON.stringify(state.doc);
  function commit(typing) {
    const now = Date.now();
    if (typing && typing === typingKey && now - typingAt < 1500) { typingAt = now; return; }
    typingKey = typing || null; typingAt = now;
    history.undo.push(snapshot());
    if (history.undo.length > 300) history.undo.shift();
    history.redo.length = 0;
  }
  function mutate(fn, opts = {}) {
    commit(opts.typing);
    const r = fn(state.doc);
    markDirty();
    if (opts.noRender) renderStatus(); else render();
    return r;
  }
  function restore(json) {
    state.doc = Model.ensureIds(JSON.parse(json));
    typingKey = null; clampSi(); markDirty(); state.focus = state.lastFocus; render();
  }
  function undo() { if (!history.undo.length) return; history.redo.push(snapshot()); restore(history.undo.pop()); }
  function redo() { if (!history.redo.length) return; history.undo.push(snapshot()); restore(history.redo.pop()); }
  function clampSi() { state.si = Math.max(0, Math.min(state.si, state.doc.sheets.length - 1)); }
  function markDirty() { state.dirty = true; }

  // ---------- render ----------
  function ctx() {
    return { doc: state.doc, si: state.si, showSide: state.ui.side, hiddenColumns: state.ui.hidden, showCounts: state.ui.counts, readScope: state.readScope, readNumbering: state.readNumbering, readColumn: state.readColumn, readIndented: state.readIndented,
      collapsed: state.view === 'read' ? NO_COLLAPSE : state.collapsed, editColumn: state.editColumn, editSheet: state.editSheet, selected: state.sel,
      gridHidden: state.ui.gridHidden, rowFilter: state.rowFilter, showFilter: state.showFilter };
  }
  let rendering = false; // focusout events fired by replacing the DOM must not trigger row cleanup
  let renderedView = null, renderedSi = null;
  const scroller = () => (state.view === 'grid' ? viewRoot.querySelector('.scroll') : viewRoot);
  function render() {
    rendering = true;
    try {
      clampSi();
      // Rebuilding the DOM resets the scroll position; keep it when staying in the same view and sheet.
      const sameplace = renderedView === state.view && renderedSi === state.si;
      if (renderedSi !== null && renderedSi !== state.si) { state.sel.clear(); state.selAnchor = null; } // selection is per sheet
      const sc = sameplace ? scroller() : null;
      const top = sc ? sc.scrollTop : 0;
      Views.renderTabs(tabsRoot, ctx());
      viewRoot.className = 'view-' + state.view;
      if (state.view === 'draft') Views.renderDraft(viewRoot, ctx());
      else if (state.view === 'grid') Views.renderGrid(viewRoot, ctx());
      else Views.renderRead(viewRoot, ctx());
      if (sc) { const sc2 = scroller(); if (sc2) sc2.scrollTop = top; }
      renderedView = state.view; renderedSi = state.si;
      document.querySelectorAll('#toolbar .views button').forEach(b => b.classList.toggle('active', b.dataset.view === state.view));
      $('#btn-undo').disabled = !history.undo.length;
      $('#btn-redo').disabled = !history.redo.length;
      $('#btn-filter').classList.toggle('active', !!state.showFilter);
      document.body.classList.toggle('toc-open', state.ui.toc !== 'off');
      $('#btn-toc').classList.toggle('active', state.ui.toc !== 'off');
      renderToc();
      renderStatus();
      renderSelBar();
      applyFocus();
    } finally { rendering = false; }
  }

  // ---------- row selection ----------
  const selBar = $('#selbar');
  /** Indexes of selected rows in the current sheet, sorted; a collapsed heading brings its hidden rows. */
  function selectedIndexes() {
    if (!isChapter() || !state.sel.size) return [];
    const rows = sheet().rows;
    const idx = new Set();
    rows.forEach((r, i) => {
      if (!state.sel.has(r._id)) return;
      const [a, b] = state.view !== 'read' && state.collapsed.has(r._id) ? Model.blockOf(rows, i) : [i, i + 1];
      for (let k = a; k < b; k++) idx.add(k);
    });
    return [...idx].sort((a, b) => a - b);
  }
  function applySelection() {
    viewRoot.querySelectorAll('.card[data-i], tr[data-i]').forEach(h => {
      const row = sheet().rows[+h.dataset.i];
      h.classList.toggle('selected', !!row && state.sel.has(row._id));
    });
    renderSelBar();
  }
  function clearSelection() { state.sel.clear(); state.selAnchor = null; applySelection(); }
  function selectRow(i, e) {
    const rows = sheet().rows;
    if (e.shiftKey && state.selAnchor != null && rows[state.selAnchor]) {
      const a = Math.min(state.selAnchor, i), b = Math.max(state.selAnchor, i);
      if (!(e.ctrlKey || e.metaKey)) state.sel.clear();
      for (let k = a; k <= b; k++) state.sel.add(rows[k]._id);
    } else if (e.ctrlKey || e.metaKey) {
      if (state.sel.has(rows[i]._id)) state.sel.delete(rows[i]._id); else state.sel.add(rows[i]._id);
      state.selAnchor = i;
    } else {
      const only = state.sel.size === 1 && state.sel.has(rows[i]._id);
      state.sel.clear();
      if (!only) state.sel.add(rows[i]._id);
      state.selAnchor = i;
    }
    const a = document.activeElement;
    if (a && a.matches && a.matches('textarea, input')) a.blur();
    applySelection();
  }
  function selectAll() { state.sel = new Set(sheet().rows.map(r => r._id)); state.selAnchor = 0; applySelection(); }
  function renderSelBar() {
    const n = selectedIndexes().length;
    const clip = state.clipboard ? state.clipboard.rows.length : 0;
    if (!n && !clip) { selBar.hidden = true; return; }
    selBar.hidden = false;
    $('#sel-info').textContent = n ? `${n} row${n === 1 ? '' : 's'} selected` : `${clip} row${clip === 1 ? '' : 's'} in clipboard`;
    selBar.querySelectorAll('[data-needs=sel]').forEach(b => b.disabled = !n);
    selBar.querySelectorAll('[data-needs=clip]').forEach(b => b.disabled = !clip);
    $('#sel-paste').textContent = n ? 'Paste after selection' : 'Paste after current row';
  }
  function copySelection() {
    const idxs = selectedIndexes(); if (!idxs.length) return;
    const rows = sheet().rows;
    state.clipboard = { rows: Model.cloneRows(idxs.map(i => rows[i])), from: sheet().name };
    const text = idxs.map(i => '  '.repeat(Model.indentOf(rows[i])) + (rows[i][state.doc.mainColumn] || '')).join('\n\n');
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).catch(() => {});
    renderSelBar();
  }
  function deleteSelection() {
    const idxs = selectedIndexes(); if (!idxs.length) return;
    const first = idxs[0];
    state.sel.clear(); state.selAnchor = null;
    mutate(d => Model.deleteRows(d, state.si, idxs));
    focusRow(Math.min(first, sheet().rows.length - 1), state.doc.mainColumn, 'end');
  }
  function pasteTarget() {
    const idxs = selectedIndexes();
    if (idxs.length) return idxs[idxs.length - 1] + 1;
    if (state.lastFocus && state.lastFocus.i != null) return state.lastFocus.i + 1;
    return sheet().rows.length;
  }
  function pasteRows(rowsToPaste, at) {
    if (!isChapter() || !rowsToPaste || !rowsToPaste.length) return;
    if (at == null) at = pasteTarget();
    let start = 0;
    mutate(d => { start = Model.insertRows(d, state.si, rowsToPaste, at); });
    const rows = sheet().rows;
    state.sel = new Set(rows.slice(start, start + rowsToPaste.length).map(r => r._id)); state.selAnchor = start;
    state.focus = { i: start, col: state.doc.mainColumn, caret: 0, block: 'center' };
    render();
  }
  function moveSelectionToSheetMenu(anchor) {
    const idxs = selectedIndexes(); if (!idxs.length) return;
    const chapters = state.doc.sheets.map((s, k) => ({ s, k })).filter(x => x.s.kind === 'chapter' && x.k !== state.si);
    if (!chapters.length) { alert('There is no other chapter to move the rows to.'); return; }
    showMenu(anchor, chapters.map(x => ({
      label: `Move ${idxs.length} row${idxs.length === 1 ? '' : 's'} to “${x.s.name}”`,
      action: () => { state.sel.clear(); state.selAnchor = null; mutate(d => Model.moveRowsToSheet(d, state.si, idxs, x.k)); },
    })));
  }
  selBar.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    switch (b.dataset.act) {
      case 'copy': copySelection(); break;
      case 'cut': copySelection(); deleteSelection(); break;
      case 'paste': pasteRows(state.clipboard && state.clipboard.rows); break;
      case 'dup': { const idxs = selectedIndexes(); pasteRows(Model.cloneRows(idxs.map(i => sheet().rows[i])), idxs[idxs.length - 1] + 1); break; }
      case 'delete': deleteSelection(); break;
      case 'move': moveSelectionToSheetMenu(b); break;
      case 'clear': clearSelection(); break;
      case 'clearclip': state.clipboard = null; renderSelBar(); break;
    }
  });
  // Click on a number or handle selects; Shift+click must not select text.
  viewRoot.addEventListener('mousedown', e => {
    if (e.shiftKey && e.target.closest && e.target.closest('.num, .handle, td.num')) e.preventDefault();
  });
  viewRoot.addEventListener('click', e => {
    if (e.target.closest && e.target.closest('button')) return; // collapse toggle inside the number cell
    const t = e.target.closest && e.target.closest('.card .num, .card .handle, td.num, td.handle-col');
    if (!t || !isChapter()) return;
    const i = rowIndexOf(t); if (i == null) return;
    selectRow(i, e);
  });
  // Keys that act on the selection when no text field has focus.
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && state.sel.size) { clearSelection(); return; }
    const t = e.target;
    if (t && t.matches && t.matches('textarea, input, select')) return;
    if (document.querySelector('dialog[open]')) return;
    if (!isChapter()) return;
    const k = e.key.toLowerCase();
    if (state.sel.size && (e.key === 'Delete' || e.key === 'Backspace')) { e.preventDefault(); deleteSelection(); return; }
    if (state.sel.size && e.ctrlKey && k === 'c') { e.preventDefault(); copySelection(); return; }
    if (state.sel.size && e.ctrlKey && k === 'x') { e.preventDefault(); copySelection(); deleteSelection(); return; }
    if (e.ctrlKey && k === 'v' && state.clipboard) { e.preventDefault(); pasteRows(state.clipboard.rows); return; }
    if (e.ctrlKey && k === 'a') { e.preventDefault(); selectAll(); return; }
    if (state.sel.size && e.ctrlKey && k === 'd') { e.preventDefault(); const idxs = selectedIndexes(); pasteRows(Model.cloneRows(idxs.map(i => sheet().rows[i])), idxs[idxs.length - 1] + 1); }
  });
  function currentHeading() {
    if (state.view === 'read' && state.readPos) return state.readPos;
    const f = state.lastFocus;
    if (!isChapter() || !f || f.i == null || !sheet().rows[f.i]) return null;
    const h = Model.headingFor(sheet().rows, f.i);
    return h == null ? null : { si: state.si, i: h };
  }
  function renderToc() {
    if (state.ui.toc === 'off') { tocRoot.innerHTML = ''; return; }
    Views.renderToc(tocRoot, { ...ctx(), tocScope: state.ui.toc, current: currentHeading() });
  }
  /** Jump to a row (table of contents, view switches), in whichever view is active. */
  function goToRow(si, i, opts = {}) {
    const s = state.doc.sheets[si];
    if (!s || s.kind !== 'chapter' || !s.rows[i]) return;
    state.si = si;
    const rows = s.rows;
    const col = opts.col || state.doc.mainColumn;
    if (state.view !== 'read') rows.forEach((r, j) => { if (j < i && state.collapsed.has(r._id) && Model.sectionEnd(rows, j) > i) state.collapsed.delete(r._id); });
    if (state.view === 'read') {
      state.lastFocus = { i, col, caret: 0 };
      state.readPos = { si, i };
      render();
      // every heading and paragraph is tagged; an "s" or "x" row belongs to the nearest block above it
      let el = null;
      for (let j = i; j >= 0 && !el; j--) el = viewRoot.querySelector(`article.read [data-si="${si}"][data-i="${j}"]`);
      if (el) el.scrollIntoView({ block: opts.block || 'start' }); else viewRoot.scrollTop = 0;
      return;
    }
    state.focus = { i, col, caret: opts.caret == null ? 'end' : opts.caret, block: opts.block || 'start', selectLen: opts.selectLen || 0 };
    state.lastFocus = { i, col, caret: opts.caret == null ? 0 : opts.caret };
    render();
  }
  /** Change view and stay at the same place in the text. */
  function switchView(mode) {
    if (mode === state.view) return;
    let si = state.si, i = state.lastFocus ? state.lastFocus.i : null;
    const col = state.lastFocus ? state.lastFocus.col : null;
    const caret = state.lastFocus ? state.lastFocus.caret : null;
    if (state.view === 'read' && state.readPos) { si = state.readPos.si; i = state.readPos.i; }
    state.view = mode; state.editColumn = null;
    const s = state.doc.sheets[si];
    if (i != null && s && s.kind === 'chapter' && s.rows[i]) goToRow(si, i, { block: 'center', col: mode === 'grid' || mode === 'draft' ? col : null, caret });
    else render();
  }
  // Click a paragraph or heading in Read view to edit that row in Draft.
  viewRoot.addEventListener('click', e => {
    if (state.view !== 'read' || e.target.closest('a')) return;
    const b = e.target.closest('.rblock[data-i]');
    if (!b) return;
    const si = +b.dataset.si, i = +b.dataset.i;
    state.readPos = { si, i };
    state.lastFocus = { i, col: state.doc.mainColumn, caret: 0 };
    switchView('draft');
  });
  // Read view scroll spy: remember which block is being read so other views (and the contents pane) can follow.
  viewRoot.addEventListener('scroll', () => {
    if (state.view !== 'read') return;
    const top = viewRoot.getBoundingClientRect().top + 60;
    let cur = null;
    for (const h of viewRoot.querySelectorAll('article.read [data-i]')) { if (h.getBoundingClientRect().top <= top) cur = h; else break; }
    if (!cur) return;
    const pos = { si: +cur.dataset.si, i: +cur.dataset.i };
    if (state.readPos && state.readPos.si === pos.si && state.readPos.i === pos.i) return;
    state.readPos = pos;
    if (state.ui.toc !== 'off') { const rows = state.doc.sheets[pos.si].rows; const h = Model.headingFor(rows, pos.i); Views.updateTocCurrent(tocRoot, h == null ? null : { si: pos.si, i: h }); }
  }, { passive: true });

  // ---------- find and replace ----------
  const findBar = $('#findbar');
  const find = { matches: [], cur: -1, jumped: false };
  function findOpts() {
    return { matchCase: $('#find-case').checked, scope: $('#find-scope').value === 'all' ? 'all' : 'sheet', si: state.si, columns: $('#find-cols').value === 'main' ? [state.doc.mainColumn] : null };
  }
  function openFind(replace) {
    findBar.hidden = false;
    runFind(true);
    const box = replace ? $('#find-r') : $('#find-q');
    box.focus(); box.select();
  }
  function closeFind() { findBar.hidden = true; find.matches = []; find.cur = -1; find.jumped = false; }
  function runFind(keepCur) {
    const q = $('#find-q').value;
    find.matches = Model.findMatches(state.doc, q, findOpts());
    if (!keepCur || find.cur >= find.matches.length) { find.cur = find.matches.length ? 0 : -1; find.jumped = false; }
    $('#find-info').textContent = q ? (find.matches.length ? `${find.cur + 1} of ${find.matches.length}` : 'No matches') : '';
    $('#find-replace').disabled = $('#find-replace-all').disabled = !find.matches.length;
  }
  function gotoMatch(k) {
    if (!find.matches.length) return;
    find.cur = (k + find.matches.length) % find.matches.length;
    find.jumped = true;
    const m = find.matches[find.cur];
    $('#find-info').textContent = `${find.cur + 1} of ${find.matches.length}`;
    if (state.view === 'read') state.view = 'draft';
    goToRow(m.si, m.i, { col: m.col, caret: m.index, block: 'center', selectLen: m.len });
  }
  function replaceCurrent() {
    if (!find.matches.length) return;
    const m = find.matches[Math.max(0, find.cur)];
    const cur = find.cur;
    mutate(d => Model.replaceMatches(d, [m], $('#find-r').value));
    runFind(true);
    find.cur = Math.min(Math.max(cur, 0), find.matches.length - 1);
    if (find.matches.length) gotoMatch(find.cur); else $('#find-info').textContent = 'No more matches';
  }
  function replaceAll() {
    if (!find.matches.length) return;
    const n = find.matches.length;
    mutate(d => Model.replaceMatches(d, find.matches, $('#find-r').value));
    runFind(false);
    $('#find-info').textContent = `Replaced ${n}`;
  }
  findBar.addEventListener('input', e => { if (e.target.id === 'find-q') runFind(false); });
  findBar.addEventListener('change', e => { if (e.target.id !== 'find-q' && e.target.id !== 'find-r') runFind(false); });
  findBar.addEventListener('keydown', e => {
    if (e.key === 'Escape') { e.preventDefault(); closeFind(); return; }
    if (e.key === 'Enter' && e.target.id === 'find-q') { e.preventDefault(); gotoMatch(find.jumped ? find.cur + (e.shiftKey ? -1 : 1) : Math.max(find.cur, 0)); }
    if (e.key === 'Enter' && e.target.id === 'find-r') { e.preventDefault(); replaceCurrent(); }
  });
  findBar.addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    switch (b.id) {
      case 'find-prev': gotoMatch(find.cur - 1); break;
      case 'find-next': gotoMatch(find.jumped ? find.cur + 1 : Math.max(find.cur, 0)); break;
      case 'find-replace': replaceCurrent(); break;
      case 'find-replace-all': replaceAll(); break;
      case 'find-close': closeFind(); break;
    }
  });
  tocRoot.addEventListener('click', e => {
    const item = e.target.closest('.toc-item');
    if (item) { goToRow(+item.dataset.si, +item.dataset.i); return; }
    const sh = e.target.closest('.toc-sheet');
    if (sh) { state.si = +sh.dataset.si; state.lastFocus = null; render(); return; }
    if (e.target.closest('#toc-close')) { state.ui.lastToc = state.ui.toc; state.ui.toc = 'off'; render(); }
  });
  tocRoot.addEventListener('change', e => {
    if (e.target.id === 'toc-scope') { state.ui.toc = e.target.value; state.ui.lastToc = e.target.value; renderToc(); }
  });
  const fmt = n => n.toLocaleString('en-US').replace(/,/g, ' ');
  function docTitle() { return state.doc.settings.title || state.fileName.replace(/\.xlsx$/i, ''); }
  function renderStatus() {
    const doc = state.doc, sheet = doc.sheets[state.si];
    const hhmm = t => { const d = new Date(t); const p = n => String(n).padStart(2, '0'); return `${p(d.getHours())}:${p(d.getMinutes())}`; };
    const parts = [state.fileName, state.dirty ? 'unsaved changes' : (state.savedAt ? `saved ${hhmm(state.savedAt)}` : 'saved')];
    const f = state.lastFocus;
    if (sheet && sheet.kind === 'chapter' && f && f.i != null && sheet.rows[f.i]) {
      const rc = Model.rowCounts(doc, sheet, sheet.rows[f.i]);
      parts.push(`Row ${Model.numbering(doc, state.si).numbers[f.i]}: ${fmt(rc.words)} words, ${fmt(rc.chars)} chars`);
    }
    if (sheet && sheet.kind === 'chapter') {
      const c = Model.sheetCounts(doc, sheet);
      parts.push(`Chapter: ${fmt(c.rows)} rows, ${fmt(c.words)} words, ${fmt(c.chars)} chars`);
    }
    const d = Model.docCounts(doc);
    const tgt = doc.settings.wordTarget || 0;
    parts.push(`Workbook: ${fmt(d.rows)} rows, ${fmt(d.words)}${tgt ? ` / ${fmt(tgt)}` : ''} words${tgt ? ` (${Math.round(100 * d.words / tgt)} %)` : ''}, ${fmt(d.chars)} chars`);
    parts.push(hasFS ? 'direct file access' : 'download mode');
    if (state.copyHandle && prefs.autosaveCopy.enabled) {
      const p = n => String(n).padStart(2, '0');
      const at = state.copyAt ? new Date(state.copyAt) : null;
      parts.push(state.copyNeedsPermission ? `autosave copy ${state.copyName}: permission needed (save once to grant)` : `autosave copy ${state.copyName}${at ? ` ${p(at.getHours())}:${p(at.getMinutes())}` : ''}`);
    }
    statusEl.textContent = parts.join('  ·  ');
    $('#doc-title').textContent = docTitle();
    $('#dirty-dot').classList.toggle('on', state.dirty);
    $('#dirty-dot').title = state.dirty ? 'Unsaved changes' : 'All changes saved';
    document.title = (state.dirty ? '• ' : '') + docTitle() + ' — SheetWriter';
  }
  function applyFocus() {
    const f = state.focus; state.focus = null;
    if (!f) return;
    const holder = viewRoot.querySelector(`[data-i="${f.i}"]`);
    if (!holder) return;
    const col = f.col || state.doc.mainColumn;
    const t = holder.querySelector(`textarea.cell[data-col="${CSS.escape(col)}"]`);
    if (!t) return;
    const mw = t.closest('.mainwrap');
    if (mw) mw.classList.add('editing');
    holder.classList.add('focusing'); // side fields of an unfocused card are hidden; show them so a side cell can take focus
    Views.autosize(t);
    t.focus();
    holder.classList.remove('focusing');
    const pos = f.caret === 'end' || f.caret == null ? t.value.length : Math.min(f.caret, t.value.length);
    try { t.setSelectionRange(pos, f.selectLen ? Math.min(pos + f.selectLen, t.value.length) : pos); } catch (e) { /* ignore */ }
    holder.scrollIntoView({ block: f.block || 'nearest' });
  }
  function focusRow(i, col, caret) { state.focus = { i, col: col || state.doc.mainColumn, caret }; applyFocus(); }

  // ---------- helpers ----------
  const sheet = () => state.doc.sheets[state.si];
  const isChapter = () => sheet() && sheet().kind === 'chapter';
  function nextKind(kind) { return Model.isHeading(kind) || kind === 'x' ? 'p' : kind; }
  function rowIndexOf(elm) { const h = elm && elm.closest('[data-i]'); return h ? +h.dataset.i : null; }
  function visible() { return state.view !== 'read' ? Views.visibleIndexes(sheet(), state.collapsed) : sheet().rows.map((_, i) => i); }
  function prevVisible(i) { const v = visible(); const k = v.indexOf(i); return k > 0 ? v[k - 1] : null; }
  function nextVisibleAfter(end) { const v = visible(); return v.find(x => x >= end) ?? null; }
  function expandRow(i) { state.collapsed.delete(sheet().rows[i]._id); }

  // ---------- generic popup menu (window.prompt is unavailable in some embedded browsers) ----------
  /** items: array or function returning an array of {label, action, disabled?, title?, checked?, keep?} or '-'.
   *  keep: the menu stays open and is rebuilt after the action (for toggles). */
  function showMenu(anchor, items) {
    closeMenu();
    const list = typeof items === 'function' ? items() : items;
    const menu = Views.el('div', { class: 'popup', id: 'popup-menu' });
    for (const it of list) {
      if (it === '-') { menu.appendChild(Views.el('hr', {})); continue; }
      if (it.heading) { menu.appendChild(Views.el('div', { class: 'popup-heading' }, it.heading)); continue; }
      const label = it.checked === undefined ? it.label : (it.checked ? '☑ ' : '☐ ') + it.label;
      const b = Views.el('button', { type: 'button', disabled: it.disabled || null, title: it.title || null }, label);
      b.onclick = () => { if (it.keep) { it.action(); showMenu(anchor, items); } else { closeMenu(); it.action(); } };
      menu.appendChild(b);
    }
    document.body.appendChild(menu);
    const r = anchor.getBoundingClientRect();
    const mh = menu.offsetHeight, mw = menu.offsetWidth;
    const top = r.bottom + mh + 8 < window.innerHeight ? r.bottom + 4 : Math.max(4, r.top - mh - 4);
    menu.style.top = top + 'px';
    menu.style.left = Math.max(4, Math.min(r.left, window.innerWidth - mw - 8)) + 'px';
    setTimeout(() => document.addEventListener('mousedown', onDocDown, true));
    document.addEventListener('keydown', onMenuKey, true);
  }
  function onDocDown(ev) { const m = $('#popup-menu'); if (m && !m.contains(ev.target)) closeMenu(); }
  function onMenuKey(ev) { if (ev.key === 'Escape') { closeMenu(); ev.stopPropagation(); } }
  function closeMenu() {
    const m = $('#popup-menu'); if (m) m.remove();
    document.removeEventListener('mousedown', onDocDown, true);
    document.removeEventListener('keydown', onMenuKey, true);
  }
  function askText(label, value) {
    return new Promise(resolve => {
      const dlg = $('#dlg-ask');
      $('#ask-label').textContent = label;
      const inp = $('#ask-input'); inp.value = value || '';
      let done = false;
      const finish = v => { if (done) return; done = true; dlg.close(); resolve(v); };
      $('#ask-ok').onclick = () => finish(inp.value);
      $('#ask-cancel').onclick = () => finish(null);
      dlg.onclose = () => finish(null);
      inp.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); finish(inp.value); } };
      dlg.showModal(); inp.focus(); inp.select();
    });
  }

  // ---------- cell editing (draft + grid) ----------
  viewRoot.addEventListener('input', e => {
    const t = e.target;
    if (!t.matches('textarea.cell')) return;
    const i = rowIndexOf(t), col = t.dataset.col;
    if (i == null) return;
    const s = sheet(), row = s.rows[i];
    if (col === state.doc.mainColumn) {
      const m = Model.detectKindPrefix(t.value);
      if (m && (m.kind !== Model.normKind(row['.kind']) || m.text !== t.value)) {
        const caret = Math.max(0, t.selectionStart - (t.value.length - m.text.length));
        state.focus = { i, col, caret };
        mutate(d => { Model.setCell(d, state.si, i, col, m.text); Model.setCell(d, state.si, i, '.kind', m.kind); });
        return;
      }
      const ind = !Model.isHeading(row['.kind']) && Model.detectIndentPrefix(t.value, prefs.indentTrigger);
      if (ind) {
        const caret = Math.max(0, t.selectionStart - (t.value.length - ind.text.length));
        state.focus = { i, col, caret };
        mutate(d => { Model.setCell(d, state.si, i, col, ind.text); Model.shiftIndent(d, state.si, i, 1); });
        return;
      }
    }
    mutate(d => Model.setCell(d, state.si, i, col, t.value), { typing: `${state.si}:${i}:${col}`, noRender: true });
    if (state.lastFocus && state.lastFocus.i === i) state.lastFocus.caret = t.selectionStart;
    Views.autosize(t);
    const mw = t.closest('.mainwrap');
    if (mw) mw.classList.toggle('empty', !t.value);
    if (state.view === 'draft') Views.refreshCard(viewRoot, ctx(), i); // live row counts
  });
  viewRoot.addEventListener('change', e => {
    const t = e.target;
    if (t.matches('select.kind-select')) {
      const i = rowIndexOf(t);
      mutate(d => Model.setCell(d, state.si, i, '.kind', t.value));
    }
    if (t.id === 'read-scope') { state.readScope = t.value; render(); }
    if (t.id === 'read-numbering') { state.readNumbering = t.value; render(); }
    if (t.id === 'read-column') { state.readColumn = t.value; render(); }
    if (t.id === 'read-indented') { state.readIndented = t.value; render(); }
  });
  viewRoot.addEventListener('focusin', e => {
    const t = e.target;
    if (!t.matches('textarea.cell')) return;
    const i = rowIndexOf(t);
    state.lastFocus = { i, col: t.dataset.col, caret: t.selectionStart };
    const mw = t.closest('.mainwrap');
    if (mw) { mw.classList.add('editing'); Views.autosize(t); }
    renderStatus();
    if (state.ui.toc !== 'off') Views.updateTocCurrent(tocRoot, currentHeading());
  });
  viewRoot.addEventListener('focusout', e => {
    const t = e.target;
    if (t.matches('input.colname-edit')) { commitColumnRename(t); return; }
    if (!t.matches('textarea.cell')) return;
    if (rendering || !t.isConnected) return; // blur caused by a re-render, not by the user leaving the row
    const mw = t.closest('.mainwrap');
    if (mw) mw.classList.remove('editing');
    const i = rowIndexOf(t);
    if (i == null || !isChapter()) return;
    if (state.view === 'draft') { Views.refreshCard(viewRoot, ctx(), i); Views.refreshMeta(viewRoot, ctx()); }
    // Empty rows are dropped once you leave them (not when the window loses focus, not within the same row).
    const s = sheet(), row = s.rows[i];
    if (row && s.rows.length > 1 && Model.rowIsEmpty(state.doc, s, row)) {
      const rel = e.relatedTarget;
      const relRow = rel && rel.matches && rel.matches('textarea.cell') ? rowIndexOf(rel) : null;
      if (relRow === i) return;
      setTimeout(() => {
        if (!document.hasFocus()) return;
        const s2 = sheet();
        if (s2 !== s || s2.rows[i] !== row || !Model.rowIsEmpty(state.doc, s2, row)) return;
        const active = document.activeElement;
        let f = null;
        if (active && active.matches && active.matches('textarea.cell') && viewRoot.contains(active)) {
          const ai = rowIndexOf(active);
          if (ai != null && ai !== i) f = { i: ai > i ? ai - 1 : ai, col: active.dataset.col, caret: active.selectionStart };
        }
        state.focus = f;
        mutate(d => Model.deleteRow(d, state.si, i));
      }, 0);
    }
  });
  viewRoot.addEventListener('click', e => {
    const r = e.target.closest('.rendered');
    if (r) {
      if (e.target.closest('a')) return; // let links work
      const mw = r.closest('.mainwrap');
      const t = mw.querySelector('textarea');
      mw.classList.add('editing');
      Views.autosize(t);
      t.focus();
      t.setSelectionRange(t.value.length, t.value.length);
      return;
    }
    const kb = e.target.closest('.kindbadge');
    if (kb) {
      const i = rowIndexOf(kb);
      state.focus = state.lastFocus && state.lastFocus.i === i ? state.lastFocus : null;
      mutate(d => Model.cycleKind(d, state.si, i, e.shiftKey ? -1 : 1));
      return;
    }
    const cb = e.target.closest('.collapse, .pill');
    if (cb) { toggleCollapse(rowIndexOf(cb)); return; }
    const cn = e.target.closest('.colname.editable');
    if (cn) { state.editColumn = cn.dataset.col; render(); return; }
    const btn = e.target.closest('button[data-action]');
    if (btn) gridAction(btn.dataset.action, btn);
  });
  viewRoot.addEventListener('keydown', e => {
    const t = e.target;
    if (t.matches('input.colname-edit')) {
      if (e.key === 'Enter') { e.preventDefault(); commitColumnRename(t); }
      if (e.key === 'Escape') { e.preventDefault(); t.dataset.cancel = '1'; state.editColumn = null; render(); }
    }
  });
  function commitColumnRename(inp) {
    if (state.editColumn == null) return;
    const old = inp.dataset.col, name = inp.value.trim();
    state.editColumn = null;
    if (inp.dataset.cancel || !name || name === old) { render(); return; }
    const v = Model.validColumnName(sheet(), name);
    if (!v.ok) { alert(v.reason); render(); return; }
    mutate(d => Model.renameColumn(d, state.si, old, v.name));
  }
  function toggleCollapse(i) {
    const row = sheet().rows[i];
    if (!row || !Model.isCollapsible(sheet().rows, i)) return;
    if (state.collapsed.has(row._id)) state.collapsed.delete(row._id); else state.collapsed.add(row._id);
    state.focus = { i, col: state.doc.mainColumn, caret: state.lastFocus && state.lastFocus.i === i ? state.lastFocus.caret : 'end' };
    render();
  }
  function collapseAll(on) {
    if (!isChapter()) return;
    const rows = sheet().rows;
    if (!on) state.collapsed.clear();
    else rows.forEach((r, i) => { if (Model.isCollapsible(rows, i)) state.collapsed.add(r._id); });
    if (state.view !== 'read') render();
  }

  function gridAction(action, btn) {
    const s = sheet();
    const col = btn.dataset.col;
    const i = rowIndexOf(btn);
    switch (action) {
      case 'add': {
        let n = s.columns.length, name;
        do { name = 'column ' + n++; } while (s.columns.some(c => c.toLowerCase() === name));
        state.editColumn = name;
        mutate(d => Model.addColumn(d, state.si, name));
        break;
      }
      case 'hide': state.ui.gridHidden.add(col); render(); break;
      case 'showhidden': columnsMenu(btn); break;
      case 'left': mutate(d => Model.moveColumn(d, state.si, col, -1)); break;
      case 'right': mutate(d => Model.moveColumn(d, state.si, col, 1)); break;
      case 'main':
        if (confirm(`Make "${col}" the main text column for the whole workbook?\nSheets that lack a "${col}" column will have their current main column renamed.`))
          mutate(d => Model.setMainColumn(d, col));
        break;
      case 'delete': {
        const data = Model.columnData(state.doc, col);
        const where = data.length ? `\n\nIt holds data in: ${data.map(x => `${x.sheet} (${x.count} row${x.count === 1 ? '' : 's'})`).join(', ')}.\nThat data will be lost (Undo is available).` : '\n\nIt is empty in every sheet.';
        if (confirm(`Delete column "${col}" from all sheets?${where}`)) mutate(d => Model.deleteColumn(d, state.si, col));
        break;
      }
      case 'row-add': state.focus = { i: i + 1, col: state.doc.mainColumn }; mutate(d => Model.addRow(d, state.si, i + 1, nextKind(s.rows[i]['.kind']), Model.indentOf(s.rows[i]))); break;
      case 'row-del': mutate(d => Model.deleteRow(d, state.si, i)); break;
    }
  }

  // ---------- row commands ----------
  function moveRowUp(i, col, caret) {
    const rows = sheet().rows;
    const [start, end] = Model.blockOf(rows, i);
    const to = Model.siblingMoveTarget(state.doc, state.si, i, -1);
    if (to == null) return;
    state.focus = { i: to + (i - start), col, caret };
    mutate(d => Model.moveBlock(d, state.si, start, end - start, to));
  }
  function moveRowDown(i, col, caret) {
    const rows = sheet().rows;
    const [start, end] = Model.blockOf(rows, i);
    const to = Model.siblingMoveTarget(state.doc, state.si, i, 1);
    if (to == null) return;
    state.focus = { i: to - (end - start) + (i - start), col, caret };
    mutate(d => Model.moveBlock(d, state.si, start, end - start, to));
  }
  function deleteRowAt(i, col) {
    const wasLast = i >= sheet().rows.length - 1;
    mutate(d => Model.deleteRow(d, state.si, i));
    focusRow(wasLast ? Math.max(0, i - 1) : i, col, 'end');
  }
  function indentRow(i, col, caret, dir) {
    state.focus = { i, col, caret };
    const changed = mutate(d => Model.shiftIndent(d, state.si, i, dir));
    if (!changed) applyFocus();
  }

  // ---------- keyboard inside cells ----------
  viewRoot.addEventListener('keydown', e => {
    const t = e.target;
    if (!t.matches('textarea.cell')) return;
    const i = rowIndexOf(t); if (i == null) return;
    const col = t.dataset.col, main = state.doc.mainColumn, isMain = col === main;
    const s = sheet(), row = s.rows[i];
    const noMods = !e.ctrlKey && !e.altKey && !e.metaKey;
    const caret = t.selectionStart;

    if (e.key === 'Enter' && noMods) {
      const newRow = prefs.enterMode === 'newline' ? e.shiftKey : !e.shiftKey;
      e.preventDefault();
      if (!newRow) { // line break inside the cell, done by hand so it behaves the same everywhere
        t.setRangeText('\n', t.selectionStart, t.selectionEnd, 'end');
        t.dispatchEvent(new Event('input', { bubbles: true }));
        return;
      }
      if (state.collapsed.has(row._id)) expandRow(i);
      const indent = Model.isHeading(row['.kind']) ? 0 : Model.indentOf(row);
      if (!isMain) { mutate(d => Model.addRow(d, state.si, i + 1, nextKind(row['.kind']), indent)); focusRow(i + 1, col, 0); return; }
      const len = t.value.length;
      if (caret >= len) { mutate(d => Model.addRow(d, state.si, i + 1, nextKind(row['.kind']), indent)); focusRow(i + 1, main, 0); }
      else if (caret === 0 && len) { mutate(d => Model.addRow(d, state.si, i, 'p', indent)); focusRow(i + 1, main, 0); }
      else { mutate(d => Model.splitRow(d, state.si, i, caret)); focusRow(i + 1, main, 0); }
      return;
    }
    if (e.key === 'Enter' && e.ctrlKey && !e.altKey) {
      e.preventDefault();
      state.focus = { i, col, caret };
      mutate(d => Model.cycleKind(d, state.si, i, e.shiftKey ? -1 : 1));
      return;
    }
    if (e.key === 'Tab' && noMods && isMain && state.view === 'draft') {
      e.preventDefault();
      indentRow(i, col, caret, e.shiftKey ? -1 : 1);
      return;
    }
    if (e.ctrlKey && !e.shiftKey && !e.altKey && (e.key === ']' || e.key === '[')) {
      e.preventDefault(); indentRow(i, col, caret, e.key === ']' ? 1 : -1); return;
    }
    if (e.key === 'Backspace' && noMods && isMain) {
      if (t.value === '' && s.rows.length > 1) { e.preventDefault(); const pv = prevVisible(i); mutate(d => Model.deleteRow(d, state.si, i)); focusRow(pv != null ? pv : 0, main, 'end'); return; }
      if (caret === 0 && t.selectionEnd === 0 && Model.indentOf(row) > 0) { e.preventDefault(); indentRow(i, col, 0, -1); return; }
      return;
    }
    if (e.altKey && !e.shiftKey && !e.ctrlKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
      e.preventDefault();
      if (e.key === 'ArrowUp') moveRowUp(i, col, caret); else moveRowDown(i, col, caret);
      return;
    }
    if (e.altKey && e.shiftKey && !e.ctrlKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
      e.preventDefault();
      state.focus = { i, col, caret };
      mutate(d => Model.shiftKind(d, state.si, i, e.key === 'ArrowLeft' ? -1 : 1));
      return;
    }
    if (e.ctrlKey && !e.shiftKey && !e.altKey && e.key === '.') { e.preventDefault(); toggleCollapse(i); return; }
    if (e.ctrlKey && !e.shiftKey && !e.altKey && (e.key === 'b' || e.key === 'i' || e.key === 'k')) {
      e.preventDefault();
      if (e.key === 'k') linkSelection(t); else wrapSelection(t, e.key === 'b' ? '**' : '*');
      return;
    }
    if (e.ctrlKey && e.shiftKey && !e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
      e.preventDefault();
      if (!state.sel.size) { state.sel.add(row._id); state.selAnchor = i; }
      const idxs = selectedIndexes();
      const edge = e.key === 'ArrowUp' ? idxs[0] : idxs[idxs.length - 1];
      const next = e.key === 'ArrowUp' ? prevVisible(edge) : nextVisibleAfter(Model.blockOf(s.rows, edge)[1]);
      if (next != null) state.sel.add(s.rows[next]._id);
      applySelection();
      return;
    }
    if (e.ctrlKey && !e.shiftKey && e.key.toLowerCase() === 'd') {
      e.preventDefault();
      mutate(d => Model.duplicateRow(d, state.si, i)); focusRow(i + 1, main, 'end');
      return;
    }
    if (e.ctrlKey && !e.shiftKey && e.key.toLowerCase() === 'j') {
      e.preventDefault();
      if (i + 1 < s.rows.length) { const c = mutate(d => Model.mergeRow(d, state.si, i)); focusRow(i, main, c); }
      return;
    }
    if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'k') { e.preventDefault(); deleteRowAt(i, col); return; }
    if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'm') { e.preventDefault(); moveRowToSheetMenu(i, t); return; }
    if (noMods && !e.shiftKey && isMain && (e.key === 'ArrowUp' || e.key === 'ArrowDown') && caret === t.selectionEnd) {
      const before = t.value.slice(0, caret), after = t.value.slice(t.selectionEnd);
      if (e.key === 'ArrowUp' && !before.includes('\n')) { const pv = prevVisible(i); if (pv != null) { e.preventDefault(); focusRow(pv, main, 'end'); } }
      if (e.key === 'ArrowDown' && !after.includes('\n')) {
        const nv = state.collapsed.has(row._id) ? nextVisibleAfter(Model.sectionEnd(s.rows, i)) : nextVisibleAfter(i + 1);
        if (nv != null) { e.preventDefault(); focusRow(nv, main, 0); }
      }
    }
  });
  viewRoot.addEventListener('keyup', e => {
    const t = e.target;
    if (t.matches && t.matches('textarea.cell') && state.lastFocus && state.lastFocus.i === rowIndexOf(t)) state.lastFocus.caret = t.selectionStart;
  });
  viewRoot.addEventListener('click', e => {
    const t = e.target;
    if (t.matches && t.matches('textarea.cell') && state.lastFocus && state.lastFocus.i === rowIndexOf(t)) state.lastFocus.caret = t.selectionStart;
  });
  // A bare Alt press (as in Alt+arrows) focuses the browser menu in Chromium; keep focus in the editor.
  document.addEventListener('keyup', e => { if (e.key === 'Alt') e.preventDefault(); });
  document.addEventListener('keydown', e => { if (e.key === 'Alt' && !e.ctrlKey) e.preventDefault(); });

  /** Toggle Markdown markers around the selection (or insert them at the caret). */
  function wrapSelection(t, m) {
    const a = t.selectionStart, b = t.selectionEnd, v = t.value, sel = v.slice(a, b);
    const inside = sel.startsWith(m) && sel.endsWith(m) && sel.length >= 2 * m.length;
    const outside = v.slice(a - m.length, a) === m && v.slice(b, b + m.length) === m;
    if (inside) { t.setRangeText(sel.slice(m.length, -m.length), a, b, 'select'); }
    else if (outside) { t.setRangeText(sel, a - m.length, b + m.length, 'select'); }
    else if (sel) { t.setRangeText(m + sel + m, a, b); t.setSelectionRange(a + m.length, b + m.length); }
    else { t.setRangeText(m + m, a, a); t.setSelectionRange(a + m.length, a + m.length); }
    t.dispatchEvent(new Event('input', { bubbles: true }));
  }
  function linkSelection(t) {
    const a = t.selectionStart, b = t.selectionEnd, sel = t.value.slice(a, b);
    if (/^https?:\/\/\S+$/.test(sel)) { t.setRangeText(`[](${sel})`, a, b); t.setSelectionRange(a + 1, a + 1); }
    else { const text = sel || 'text'; t.setRangeText(`[${text}](url)`, a, b); t.setSelectionRange(a + text.length + 3, a + text.length + 6); }
    t.dispatchEvent(new Event('input', { bubbles: true }));
  }
  function moveRowToSheetMenu(i, anchor) {
    const chapters = state.doc.sheets.map((s, k) => ({ s, k })).filter(x => x.s.kind === 'chapter' && x.k !== state.si);
    if (!chapters.length) { alert('There is no other chapter to move the row to.'); return; }
    const rows = sheet().rows;
    const [start, end] = Model.blockOf(rows, i);
    const idxs = []; for (let k = start; k < end; k++) idxs.push(k);
    showMenu(anchor, chapters.map(x => ({
      label: `Move ${idxs.length > 1 ? idxs.length + ' rows' : 'row'} to “${x.s.name}”`,
      action: () => { mutate(d => Model.moveRowsToSheet(d, state.si, idxs, x.k)); focusRow(Math.min(start, sheet().rows.length - 1), state.doc.mainColumn, 'end'); },
    })));
  }

  // ---------- drag and drop (draft cards and grid rows share the logic; grid headers reorder columns) ----------
  let dragRow = null, dragTab = null, dragCol = null, dragSel = false;
  const holderOf = elm => elm.closest && elm.closest('.card, tr[data-i]');
  // Column resizing (Grid): drag the handle at the right edge of a header; double-click resets.
  let resizing = null;
  viewRoot.addEventListener('mousedown', e => {
    const h = e.target.closest && e.target.closest('.col-resize');
    if (!h) return;
    e.preventDefault();
    const col = h.dataset.col;
    const colEl = viewRoot.querySelector(`table.grid col[data-col="${CSS.escape(col)}"]`);
    resizing = { col, colEl, startX: e.clientX, startW: Model.columnWidth(state.doc, col), table: h.closest('table') };
    document.body.classList.add('col-resizing');
  });
  document.addEventListener('mousemove', e => {
    if (!resizing) return;
    const w = Math.max(40, Math.round(resizing.startW + e.clientX - resizing.startX));
    resizing.w = w;
    resizing.colEl.style.width = w + 'px';
    resizing.table.style.width = ([...resizing.table.querySelectorAll('col')].reduce((n, c) => n + parseFloat(c.style.width), 0)) + 'px';
  });
  document.addEventListener('mouseup', () => {
    if (!resizing) return;
    const { col, w } = resizing; resizing = null;
    document.body.classList.remove('col-resizing');
    if (w) { state.doc.settings.widths = state.doc.settings.widths || {}; state.doc.settings.widths[col] = w; markDirty(); }
    Views.autosizeAll(viewRoot);
    Views.applyFreeze(viewRoot.querySelector('table.grid'), state.doc.settings.freezeColumns);
    renderStatus();
  });
  viewRoot.addEventListener('dblclick', e => {
    const h = e.target.closest && e.target.closest('.col-resize');
    if (!h) return;
    e.preventDefault();
    if (state.doc.settings.widths && state.doc.settings.widths[h.dataset.col]) { delete state.doc.settings.widths[h.dataset.col]; markDirty(); render(); }
  });
  document.addEventListener('dragstart', e => {
    if (resizing || (e.target.closest && e.target.closest('.col-resize'))) { e.preventDefault(); return; }
    const th = e.target.closest && e.target.closest('th.col-drag');
    if (th) { dragCol = th.dataset.col; e.dataTransfer.setData('text/sw-col', dragCol); e.dataTransfer.effectAllowed = 'move'; th.classList.add('dragging'); return; }
    const h = e.target.closest && e.target.closest('.handle');
    if (h) {
      dragRow = rowIndexOf(h);
      const row = sheet().rows[dragRow];
      dragSel = !!row && state.sel.has(row._id) && selectedIndexes().length > 1;
      e.dataTransfer.setData('text/sw-row', String(dragRow)); e.dataTransfer.effectAllowed = 'move';
      if (dragSel) viewRoot.querySelectorAll('.selected').forEach(n => n.classList.add('dragging'));
      else { const holder = holderOf(h); if (holder) holder.classList.add('dragging'); }
      return;
    }
    const tab = e.target.closest && e.target.closest('.tab[data-i]');
    if (tab) { dragTab = +tab.dataset.i; e.dataTransfer.setData('text/sw-tab', String(dragTab)); e.dataTransfer.effectAllowed = 'move'; }
  });
  document.addEventListener('dragend', () => {
    dragRow = null; dragTab = null; dragCol = null; dragSel = false;
    document.querySelectorAll('.dragging, .drop-above, .drop-below, .drop-target, .drop-left, .drop-right').forEach(n => n.classList.remove('dragging', 'drop-above', 'drop-below', 'drop-target', 'drop-left', 'drop-right'));
  });
  const colDropTarget = e => {
    const th = e.target.closest && e.target.closest('table.grid thead th');
    if (!th || !th.dataset.col || th.classList.contains('reserved') || th.classList.contains('meta')) return null;
    const r = th.getBoundingClientRect();
    return { th, before: e.clientX < r.left + r.width / 2 };
  };
  viewRoot.addEventListener('dragover', e => {
    if (dragCol != null) {
      const t = colDropTarget(e); if (!t) return;
      e.preventDefault(); e.dataTransfer.dropEffect = 'move';
      viewRoot.querySelectorAll('.drop-left, .drop-right').forEach(n => n.classList.remove('drop-left', 'drop-right'));
      t.th.classList.add(t.before ? 'drop-left' : 'drop-right');
      return;
    }
    if (dragRow == null) return;
    const holder = holderOf(e.target); if (!holder) return;
    e.preventDefault(); e.dataTransfer.dropEffect = 'move';
    const r = holder.getBoundingClientRect();
    const above = e.clientY < r.top + r.height / 2;
    viewRoot.querySelectorAll('.drop-above, .drop-below').forEach(n => n.classList.remove('drop-above', 'drop-below'));
    holder.classList.add(above ? 'drop-above' : 'drop-below');
  });
  viewRoot.addEventListener('drop', e => {
    if (dragCol != null) {
      const t = colDropTarget(e); const name = dragCol; dragCol = null;
      if (!t) return;
      e.preventDefault();
      const cols = sheet().columns;
      const target = t.th.dataset.col;
      let before = t.before ? target : cols[cols.indexOf(target) + 1] || null;
      if (before === name) return;
      mutate(d => Model.moveColumnBefore(d, state.si, name, before));
      return;
    }
    if (dragRow == null) return;
    const holder = holderOf(e.target); if (!holder) return;
    e.preventDefault();
    const rows = sheet().rows;
    const r = holder.getBoundingClientRect();
    const above = e.clientY < r.top + r.height / 2;
    const ti = +holder.dataset.i;
    const [ts, te] = Model.blockOf(rows, ti);
    const to = above ? ts : te;
    if (dragSel) {
      const idxs = selectedIndexes(); dragRow = null; dragSel = false;
      if (idxs.includes(ti)) { render(); return; }
      mutate(d => Model.moveRows(d, state.si, idxs, to));
      return;
    }
    const [start, end] = Model.blockOf(rows, dragRow);
    dragRow = null;
    if (to >= start && to <= end) { render(); return; }
    const newStart = to > start ? to - (end - start) : to;
    state.focus = { i: newStart, col: state.doc.mainColumn, caret: 'end' };
    mutate(d => Model.moveBlock(d, state.si, start, end - start, to));
  });
  tabsRoot.addEventListener('dragover', e => {
    const tab = e.target.closest('.tab[data-i]'); if (!tab) return;
    const ti = +tab.dataset.i;
    if (dragRow != null && ti !== state.si && state.doc.sheets[ti].kind === 'chapter') { e.preventDefault(); tab.classList.add('drop-target'); }
    if (dragTab != null && ti !== dragTab) { e.preventDefault(); tab.classList.add('drop-target'); }
  });
  tabsRoot.addEventListener('dragleave', e => { const tab = e.target.closest('.tab'); if (tab) tab.classList.remove('drop-target'); });
  tabsRoot.addEventListener('drop', e => {
    const tab = e.target.closest('.tab[data-i]'); if (!tab) return;
    const ti = +tab.dataset.i;
    e.preventDefault();
    if (dragRow != null) {
      const from = dragRow; dragRow = null;
      let idxs;
      if (dragSel) { idxs = selectedIndexes(); dragSel = false; state.sel.clear(); state.selAnchor = null; }
      else { const [start, end] = Model.blockOf(sheet().rows, from); idxs = []; for (let k = start; k < end; k++) idxs.push(k); }
      mutate(d => Model.moveRowsToSheet(d, state.si, idxs, ti));
      return;
    }
    if (dragTab != null) { const from = dragTab; dragTab = null; state.si = ti; mutate(d => Model.moveSheet(d, from, ti)); }
  });

  // ---------- tabs ----------
  tabsRoot.addEventListener('click', e => {
    const tab = e.target.closest('.tab');
    if (!tab || tab.matches('input')) return;
    if (tab.id === 'tab-add') {
      const at = mutate(d => Model.addSheet(d, Model.uniqueSheetName(d, 'Chapter'), state.si + 1));
      state.si = at; state.editSheet = at; render();
      return;
    }
    if (tab.id === 'tab-menu') { sheetMenu(tab); return; }
    state.si = +tab.dataset.i; state.lastFocus = null; render();
  });
  tabsRoot.addEventListener('dblclick', e => {
    const tab = e.target.closest('button.tab[data-i]'); if (!tab) return;
    state.si = +tab.dataset.i; state.editSheet = state.si; render();
  });
  tabsRoot.addEventListener('keydown', e => {
    const t = e.target; if (!t.matches('input.tab-edit')) return;
    if (e.key === 'Enter') { e.preventDefault(); commitSheetRename(t); }
    if (e.key === 'Escape') { e.preventDefault(); t.dataset.cancel = '1'; state.editSheet = null; render(); }
  });
  tabsRoot.addEventListener('focusout', e => { const t = e.target; if (t.matches('input.tab-edit')) commitSheetRename(t); });
  function commitSheetRename(inp) {
    if (state.editSheet == null) return;
    const si = +inp.dataset.i, name = inp.value.trim();
    state.editSheet = null;
    if (inp.dataset.cancel || !name || name === state.doc.sheets[si].name) { render(); return; }
    const ok = mutate(d => Model.renameSheet(d, si, name));
    if (!ok) alert('A sheet with that name already exists.');
  }
  function sheetMenu(anchor) {
    const si = state.si, n = state.doc.sheets.length, s = state.doc.sheets[si];
    showMenu(anchor, [
      { label: 'Rename sheet', action: () => { state.editSheet = si; render(); } },
      { label: 'Move left', disabled: si === 0, action: () => { state.si = si - 1; mutate(d => Model.moveSheet(d, si, si - 1)); } },
      { label: 'Move right', disabled: si >= n - 1, action: () => { state.si = si + 1; mutate(d => Model.moveSheet(d, si, si + 1)); } },
      '-',
      { label: 'Delete sheet', action: () => {
        if (confirm(`Delete sheet "${s.name}"${s.kind === 'chapter' ? ` and its ${s.rows.length} rows` : ''}? (Undo is available.)`)) mutate(d => Model.deleteSheet(d, si));
      } },
    ]);
  }

  // ---------- global keys ----------
  document.addEventListener('keydown', e => {
    const k = e.key.toLowerCase();
    if (e.ctrlKey && !e.shiftKey && k === 's') { e.preventDefault(); saveFile(false); return; }
    if (e.ctrlKey && e.shiftKey && k === 's') { e.preventDefault(); saveFile(true); return; }
    if (e.ctrlKey && k === 'o') { e.preventDefault(); openFile(); return; }
    if (e.ctrlKey && !e.shiftKey && k === 'z') { e.preventDefault(); undo(); return; }
    if ((e.ctrlKey && k === 'y') || (e.ctrlKey && e.shiftKey && k === 'z')) { e.preventDefault(); redo(); return; }
    if (e.ctrlKey && k === 'e') { e.preventDefault(); openExport(); return; }
    if (e.ctrlKey && !e.shiftKey && k === 'f') { e.preventDefault(); openFind(false); return; }
    if (e.ctrlKey && !e.shiftKey && k === 'h') { e.preventDefault(); openFind(true); return; }
    if ((e.key === 'F3' || (e.ctrlKey && k === 'g')) && !findBar.hidden) { e.preventDefault(); gotoMatch(find.jumped ? find.cur + (e.shiftKey ? -1 : 1) : Math.max(find.cur, 0)); return; }
    if (e.key === 'Escape' && !findBar.hidden && !state.sel.size) { closeFind(); return; }
    if (e.ctrlKey && (e.key === 'PageUp' || e.key === 'PageDown')) {
      e.preventDefault();
      const n = state.doc.sheets.length;
      state.si = (state.si + (e.key === 'PageUp' ? -1 : 1) + n) % n; state.lastFocus = null; render();
      return;
    }
    if (e.key === 'F1') { e.preventDefault(); $('#dlg-help').showModal(); }
  });

  // ---------- toolbar ----------
  $('#btn-new').onclick = () => { if (!confirmDiscard()) return; loadDoc(Model.newDoc(), 'untitled.xlsx', null, new Map()); };
  $('#btn-open').onclick = () => openFile();
  $('#btn-recent').onclick = e => recentMenu(e.currentTarget);
  $('#btn-save').onclick = () => saveFile(false);
  $('#btn-saveas').onclick = () => saveFile(true);
  $('#btn-undo').onclick = () => undo();
  $('#btn-redo').onclick = () => redo();
  $('#btn-export').onclick = () => openExport();
  $('#btn-import').onclick = () => openImport('', '');
  $('#btn-settings').onclick = () => openSettings();
  $('#btn-help').onclick = () => $('#dlg-help').showModal();
  $('#btn-toc').onclick = () => { if (state.ui.toc === 'off') state.ui.toc = state.ui.lastToc || 'sheet'; else { state.ui.lastToc = state.ui.toc; state.ui.toc = 'off'; } render(); };
  $('#btn-collapse').onclick = () => collapseAll(true);
  $('#btn-expand').onclick = () => collapseAll(false);
  function columnsMenu(anchor) {
    showMenu(anchor, () => {
      const cols = isChapter() ? Model.sideColumns(state.doc, sheet()) : [];
      const items = [
        { heading: 'Draft view shows' },
        { label: 'Side columns', checked: state.ui.side, keep: true, action: () => { state.ui.side = !state.ui.side; render(); } },
      ];
      for (const c of cols) items.push({ label: c, checked: state.ui.side && !state.ui.hidden.has(c), disabled: !state.ui.side, keep: true, action: () => { if (state.ui.hidden.has(c)) state.ui.hidden.delete(c); else state.ui.hidden.add(c); render(); } });
      if (cols.length > 1) items.push({ label: 'All side columns', keep: true, action: () => { state.ui.hidden.clear(); state.ui.side = true; render(); } });
      items.push({ label: 'Row counts (words, characters, section totals)', checked: state.ui.counts, keep: true, action: () => { state.ui.counts = !state.ui.counts; render(); } });
      if (isChapter()) {
        items.push('-', { heading: 'Grid view shows' });
        const gcols = sheet().columns.filter(c => c !== '.no' && c !== state.doc.mainColumn);
        for (const c of gcols) items.push({ label: c, checked: !state.ui.gridHidden.has(c), keep: true, action: () => { if (state.ui.gridHidden.has(c)) state.ui.gridHidden.delete(c); else state.ui.gridHidden.add(c); render(); } });
        if (state.ui.gridHidden.size) items.push({ label: 'All columns', keep: true, action: () => { state.ui.gridHidden.clear(); render(); } });
      }
      return items;
    });
  }
  $('#btn-columns').onclick = e => columnsMenu(e.currentTarget);
  // Row filter (Draft and Grid): applied live, no re-render, so the box keeps focus.
  viewRoot.addEventListener('input', e => {
    if (e.target.id !== 'row-filter') return;
    state.rowFilter = e.target.value;
    Views.applyRowFilter(viewRoot, ctx());
  });
  $('#btn-filter').onclick = () => { state.showFilter = !state.showFilter; if (!state.showFilter) state.rowFilter = ''; if (state.view === 'read') state.view = 'draft'; render(); if (state.showFilter) { const f = $('#row-filter'); if (f) f.focus(); } };
  // Theme
  function applyTheme() { document.documentElement.dataset.theme = prefs.theme === 'auto' ? '' : prefs.theme; }
  applyTheme();
  // Click the title to edit it in place.
  $('#doc-title').onclick = () => {
    if ($('#title-edit')) return;
    const span = $('#doc-title');
    const inp = Views.el('input', { type: 'text', id: 'title-edit', value: state.doc.settings.title || '', placeholder: 'Title', spellcheck: 'false' });
    span.replaceWith(inp); inp.focus(); inp.select();
    let done = false;
    const finish = commit => {
      if (done) return; done = true;
      const v = inp.value.trim();
      inp.replaceWith(span);
      if (commit && v !== (state.doc.settings.title || '')) mutate(d => { d.settings.title = v; }); else renderStatus();
    };
    inp.onkeydown = ev => { if (ev.key === 'Enter') { ev.preventDefault(); finish(true); } if (ev.key === 'Escape') { ev.preventDefault(); finish(false); } };
    inp.onblur = () => finish(true);
  };
  document.querySelectorAll('#toolbar .views button').forEach(b => b.onclick = () => switchView(b.dataset.view));
  document.querySelectorAll('dialog button[data-close]').forEach(b => b.onclick = () => b.closest('dialog').close());
  $('#file-input').onchange = async e => {
    const f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    await loadBuffer(await f.arrayBuffer(), f.name, null);
  };
  document.querySelectorAll('.app-version').forEach(n => n.textContent = APP.version);
  // About block (in the help dialog)
  $('#ab-repo').href = APP.repo; $('#ab-readme').href = APP.readme; $('#ab-site').href = APP.site; $('#ab-download-link').href = APP.download;
  $('#ab-download').onclick = async () => {
    const b = $('#ab-download'); b.disabled = true; b.textContent = 'Downloading…';
    try {
      const r = await fetch(APP.download, { cache: 'no-store' });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      download(await r.blob(), 'sheetwriter.html');
      b.textContent = 'Downloaded ✓';
    } catch (e) {
      window.open(APP.download, '_blank');
      b.textContent = 'Opened in a new tab (use Save as…)';
    }
    setTimeout(() => { b.disabled = false; b.textContent = 'Download latest version'; }, 3000);
  };

  // ---------- display state saved with the workbook ----------
  function currentViewState() {
    return {
      mode: state.view, sheet: sheet() ? sheet().name : '', row: state.lastFocus && state.lastFocus.i != null ? state.lastFocus.i : 0,
      toc: state.ui.toc, side: state.ui.side, hidden: [...state.ui.hidden], counts: state.ui.counts, gridHidden: [...state.ui.gridHidden],
    };
  }
  function applyViewState(v) {
    v = Object.assign(Model.defaultView(), v || {});
    state.view = ['draft', 'grid', 'read'].includes(v.mode) ? v.mode : 'draft';
    const si = state.doc.sheets.findIndex(s => s.name === v.sheet);
    if (si >= 0) state.si = si;
    state.ui.toc = ['sheet', 'all'].includes(v.toc) ? v.toc : 'off';
    state.ui.lastToc = state.ui.toc === 'off' ? 'sheet' : state.ui.toc;
    state.ui.side = v.side !== false;
    state.ui.hidden = new Set(v.hidden || []);
    state.ui.gridHidden = new Set(v.gridHidden || []);
    state.ui.counts = v.counts !== false;
    state.rowFilter = ''; state.showFilter = false;
    const s = state.doc.sheets[state.si];
    if (s && s.kind === 'chapter' && v.row > 0 && v.row < s.rows.length) {
      state.lastFocus = { i: v.row, col: state.doc.mainColumn, caret: 0 };
      if (state.view !== 'read') state.focus = { i: v.row, col: state.doc.mainColumn, caret: 'end', block: 'center' };
    }
  }

  function confirmDiscard() { return !state.dirty || confirm('Discard unsaved changes?'); }

  // ---------- files ----------
  const FILE_TYPES = [{ description: 'Excel workbook', accept: { [XlsxIO.MIME]: ['.xlsx'] } }];
  async function openFile() {
    if (!confirmDiscard()) return;
    if (hasFS) {
      try {
        const [h] = await window.showOpenFilePicker({ types: FILE_TYPES, multiple: false });
        const f = await h.getFile();
        await loadBuffer(await f.arrayBuffer(), f.name, h, f.lastModified);
      } catch (e) { if (e && e.name !== 'AbortError') alert('Could not open file: ' + e.message); }
    } else {
      $('#file-input').click();
    }
  }
  async function loadBuffer(buffer, name, handle, mtime) {
    try {
      const { doc, sources, warnings, hasSettings } = await XlsxIO.load(buffer);
      if (!hasSettings && !doc.settings.title) doc.settings.title = name.replace(/\.xlsx$/i, '').replace(/_/g, ' ').trim();
      loadDoc(doc, name, handle, sources);
      state.fileMtime = handle ? (mtime || null) : null;
      state.savedAt = mtime || null; renderStatus();
      if (handle) { addRecent(name, handle); loadCopyHandle(name); }
      if (warnings.length) alert(warnings.join('\n'));
    } catch (e) {
      console.error(e);
      alert('Could not read this workbook: ' + (e.message || e));
    }
  }
  function loadDoc(doc, name, handle, sources) {
    state.doc = Model.ensureIds(doc); state.fileName = name; state.fileHandle = handle; state.sources = sources || new Map();
    state.si = doc.sheets.findIndex(s => s.kind === 'chapter'); if (state.si < 0) state.si = 0;
    state.dirty = false; state.lastFocus = null; state.readPos = null; state.collapsed.clear(); state.editColumn = null; state.editSheet = null;
    state.sel.clear(); state.selAnchor = null;
    state.fileMtime = null; state.savedAt = null; state.copyHandle = null; state.copyName = ''; state.copyAt = null; state.copyNeedsPermission = false;
    closeFind();
    history.undo.length = 0; history.redo.length = 0; typingKey = null;
    applyViewState(doc.settings.view);
    idb.del('autosave').catch(() => {});
    render();
  }
  function suggestedName() {
    if (state.fileName && state.fileName !== 'untitled.xlsx') return state.fileName;
    const t = Model.safeFileName(state.doc.settings.title);
    return (t || 'untitled') + '.xlsx';
  }
  async function saveFile(as) {
    let buf;
    state.doc.settings.view = currentViewState(); // display state travels with the file; not an undo step
    try { buf = await XlsxIO.save(state.doc, state.sources); }
    catch (e) { console.error(e); alert('Could not build the workbook: ' + (e.message || e)); return; }
    const blob = new Blob([buf], { type: XlsxIO.MIME });
    try {
      // Guard against overwriting a file that changed on disk since it was opened (edited in Excel, say).
      if (hasFS && state.fileHandle && !as && state.fileMtime) {
        let onDisk = null;
        try { onDisk = (await state.fileHandle.getFile()).lastModified; } catch (e) { /* file gone: fall through to the picker */ }
        if (onDisk && onDisk > state.fileMtime + 1500) {
          const choice = await askConflict(onDisk);
          if (choice === 'cancel') return;
          if (choice === 'copy') as = true;
        }
      }
      if (hasFS && (as || !state.fileHandle)) {
        const h = await window.showSaveFilePicker({ suggestedName: suggestedName(), types: FILE_TYPES });
        state.fileHandle = h; state.fileName = h.name;
      }
      if (state.fileHandle) {
        const w = await state.fileHandle.createWritable();
        await w.write(blob); await w.close();
        try { state.fileMtime = (await state.fileHandle.getFile()).lastModified; } catch (e) { /* ignore */ }
        addRecent(state.fileName, state.fileHandle);
        ensureCopyPermission();
      } else {
        if (as || state.fileName === 'untitled.xlsx') {
          const n = await askText('File name', suggestedName());
          if (n == null || !n.trim()) return;
          state.fileName = /\.xlsx$/i.test(n) ? n.trim() : n.trim() + '.xlsx';
        }
        download(blob, state.fileName);
      }
      state.dirty = false; state.savedAt = Date.now();
      idb.del('autosave').catch(() => {});
      renderStatus(); // nothing in the view changes on save; a full render would lose the scroll position
    } catch (e) {
      if (e && (e.name === 'AbortError' || e.name === 'NotAllowedError')) { statusEl.textContent = 'Save cancelled: no permission to write the file. ' + statusEl.textContent; return; }
      console.error(e);
      if (confirm('Could not save in place: ' + (e.message || e) + '\nDownload a copy instead?')) download(blob, state.fileName);
    }
  }
  function askConflict(onDisk) {
    return new Promise(resolve => {
      const dlg = $('#dlg-conflict');
      $('#conflict-text').textContent = `"${state.fileName}" was changed on disk at ${new Date(onDisk).toLocaleString()}, after it was opened here. Saving now would overwrite those changes.`;
      let done = false;
      const finish = v => { if (done) return; done = true; dlg.close(); resolve(v); };
      $('#conflict-overwrite').onclick = () => finish('overwrite');
      $('#conflict-copy').onclick = () => finish('copy');
      $('#conflict-cancel').onclick = () => finish('cancel');
      dlg.onclose = () => finish('cancel');
      dlg.showModal();
    });
  }

  // ---------- autosave copy to a second workbook (Chromium: needs a file handle) ----------
  const COPY_KEY = 'autosaveCopies';
  async function loadCopyHandle(name) {
    state.copyHandle = null; state.copyName = ''; state.copyNeedsPermission = false;
    if (!hasFS) return;
    try {
      const map = (await idb.get(COPY_KEY)) || {};
      const entry = map[name];
      if (!entry || !entry.handle) return;
      state.copyHandle = entry.handle; state.copyName = entry.name;
      const p = await entry.handle.queryPermission({ mode: 'readwrite' });
      state.copyNeedsPermission = p !== 'granted';
      renderStatus();
    } catch (e) { /* ignore */ }
  }
  async function chooseCopyFile() {
    if (!hasFS) { alert('The autosave copy needs Edge or Chrome (direct file access).'); return; }
    try {
      const base = (state.fileName || 'untitled.xlsx').replace(/\.xlsx$/i, '');
      const h = await window.showSaveFilePicker({ suggestedName: base + '_AUTOSAVE.xlsx', types: FILE_TYPES });
      state.copyHandle = h; state.copyName = h.name; state.copyNeedsPermission = false; lastCopySnap = '';
      const map = (await idb.get(COPY_KEY)) || {};
      map[state.fileName] = { handle: h, name: h.name };
      await idb.set(COPY_KEY, map);
      $('#st-copy-name').textContent = h.name;
      renderStatus();
    } catch (e) { if (e && e.name !== 'AbortError') alert('Could not choose the autosave file: ' + e.message); }
  }
  /** Called from a user gesture (a manual save): re-request write permission for a remembered copy handle. */
  async function ensureCopyPermission() {
    if (!state.copyHandle || !state.copyNeedsPermission) return;
    try { const p = await state.copyHandle.requestPermission({ mode: 'readwrite' }); state.copyNeedsPermission = p !== 'granted'; renderStatus(); } catch (e) { /* ignore */ }
  }
  let lastCopyAt = 0, lastCopySnap = '', copyBusy = false;
  async function autosaveCopyTick() {
    if (copyBusy || !prefs.autosaveCopy.enabled || !state.copyHandle || state.copyNeedsPermission || !state.dirty) return;
    if (Date.now() - lastCopyAt < (prefs.autosaveCopy.minutes || 2) * 60000) return;
    const snap = snapshot();
    if (snap === lastCopySnap) return;
    copyBusy = true;
    try {
      const p = await state.copyHandle.queryPermission({ mode: 'readwrite' });
      if (p !== 'granted') { state.copyNeedsPermission = true; renderStatus(); return; }
      const d = JSON.parse(snap); d.settings.view = currentViewState();
      const buf = await XlsxIO.save(d, state.sources);
      const w = await state.copyHandle.createWritable();
      await w.write(new Blob([buf], { type: XlsxIO.MIME })); await w.close();
      lastCopyAt = Date.now(); lastCopySnap = snap; state.copyAt = lastCopyAt;
      renderStatus();
    } catch (e) { console.warn('autosave copy failed', e); }
    finally { copyBusy = false; }
  }
  setInterval(autosaveCopyTick, 15000);

  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  }

  // ---------- recent files (file handles kept in IndexedDB; Chromium only) ----------
  async function addRecent(name, handle) {
    if (!handle) return;
    try {
      let list = (await idb.get('recents')) || [];
      const same = [];
      for (const r of list) { try { if (r.handle && await r.handle.isSameEntry(handle)) same.push(r); } catch (e) { /* ignore */ } }
      list = list.filter(r => !same.includes(r));
      list.unshift({ name, handle, at: Date.now() });
      await idb.set('recents', list.slice(0, 10));
    } catch (e) { /* handles not storable */ }
  }
  async function recentMenu(anchor) {
    let list = [];
    try { list = (await idb.get('recents')) || []; } catch (e) { /* ignore */ }
    const recover = { label: 'Recover an autosave…', action: () => recoverMenu(anchor) };
    if (!hasFS) { showMenu(anchor, [{ label: 'Recent files need Edge or Chrome (direct file access).', disabled: true, action() {} }, '-', recover]); return; }
    if (!list.length) { showMenu(anchor, [{ label: 'No recent files yet. Files you open or save appear here.', disabled: true, action() {} }, '-', recover]); return; }
    showMenu(anchor, list.map(r => ({
      label: r.name, title: new Date(r.at).toLocaleString(),
      action: async () => {
        if (!confirmDiscard()) return;
        try {
          let p = await r.handle.queryPermission({ mode: 'readwrite' });
          if (p !== 'granted') p = await r.handle.requestPermission({ mode: 'readwrite' });
          if (p !== 'granted') { p = await r.handle.queryPermission({ mode: 'read' }); if (p !== 'granted') p = await r.handle.requestPermission({ mode: 'read' }); }
          if (p !== 'granted') return;
          const f = await r.handle.getFile();
          await loadBuffer(await f.arrayBuffer(), f.name, r.handle);
        } catch (e) {
          alert('Could not open "' + r.name + '": ' + (e.message || e) + '\nIt will be removed from the list.');
          try { const l = ((await idb.get('recents')) || []).filter(x => x !== r && x.name !== r.name); await idb.set('recents', l); } catch (e2) { /* ignore */ }
        }
      },
    })).concat(['-', { label: 'Recover an autosave…', action: () => recoverMenu(anchor) }, { label: 'Clear list', action: () => idb.del('recents').catch(() => {}) }]));
  }

  // ---------- export dialog ----------
  const dlgExport = $('#dlg-export');
  function exportColumns(scopeAll) {
    const sheets = scopeAll ? state.doc.sheets : [sheet()];
    return [...new Set(sheets.filter(s => s && s.kind === 'chapter').flatMap(s => Model.userColumns(state.doc, s)))];
  }
  function openExport() {
    const fill = (sel, cols, current, allowNone) => {
      sel.innerHTML = '';
      if (allowNone) sel.appendChild(Views.el('option', { value: '' }, 'none'));
      cols.forEach(c => sel.appendChild(Views.el('option', { value: c, selected: c === current }, c)));
    };
    const scopeSel = $('#ex-scope');
    const multi = Model.chapterSheets(state.doc).length > 1;
    scopeSel.value = multi ? 'all' : 'sheet';
    $('#ex-titles').checked = multi && state.doc.settings.numbering === 'per-sheet';
    const refreshCols = () => {
      const cols = exportColumns(scopeSel.value === 'all');
      fill($('#ex-column'), cols, $('#ex-column').value || state.doc.mainColumn, false);
      const sideCols = cols.filter(c => c !== $('#ex-column').value).concat(Model.META.filter(m => Model.isMeta(state.doc, m)));
      fill($('#ex-side'), sideCols, $('#ex-side').value, true);
    };
    refreshCols();
    const update = () => {
      const md = Exporter.toMarkdown(state.doc, {
        column: $('#ex-column').value,
        scope: scopeSel.value === 'all' ? 'all' : state.si,
        sheetTitles: $('#ex-titles').checked,
        numbering: $('#ex-numbering').value || false,
        indented: $('#ex-indented').value,
        side: $('#ex-side').value ? { column: $('#ex-side').value, mode: $('#ex-sidemode').value } : null,
      });
      $('#ex-preview').value = md;
      $('#ex-info').textContent = `${Model.wordCount(md)} words · ${md.length} characters`;
    };
    dlgExport.oninput = dlgExport.onchange = e => { if (e.target.id === 'ex-scope' || e.target.id === 'ex-column') refreshCols(); update(); };
    $('#ex-copy').onclick = async () => {
      try { await navigator.clipboard.writeText($('#ex-preview').value); $('#ex-copy').textContent = 'Copied ✓'; setTimeout(() => $('#ex-copy').textContent = 'Copy', 1500); }
      catch (e) { $('#ex-preview').select(); document.execCommand('copy'); }
    };
    const exportName = ext => {
      const base = suggestedName().replace(/\.xlsx$/i, '');
      const col = $('#ex-column').value;
      return (col === state.doc.mainColumn ? base : `${base}-${col}`) + (scopeSel.value === 'all' ? '' : '-' + sheet().name.replace(/[^\w\-]+/g, '_')) + ext;
    };
    const exportOpts = () => ({
      column: $('#ex-column').value, scope: scopeSel.value === 'all' ? 'all' : state.si, sheetTitles: $('#ex-titles').checked,
      numbering: $('#ex-numbering').value || false, indented: $('#ex-indented').value,
      side: $('#ex-side').value ? { column: $('#ex-side').value, mode: $('#ex-sidemode').value } : null,
    });
    $('#ex-download').onclick = () => download(new Blob([$('#ex-preview').value], { type: 'text/markdown;charset=utf-8' }), exportName('.md'));
    $('#ex-docx').onclick = async () => {
      const b = $('#ex-docx'); b.disabled = true;
      try { download(await Docx.build(state.doc, exportOpts()), exportName('.docx')); }
      catch (e) { console.error(e); alert('Could not build the Word file: ' + (e.message || e)); }
      finally { b.disabled = false; }
    };
    update();
    dlgExport.showModal();
  }

  // ---------- import dialog ----------
  const dlgImport = $('#dlg-import');
  let importExtra = []; // further files chosen at once, each imported as its own sheet
  const importOpts = () => ({ granularity: $('#im-gran').value, stripNumbers: $('#im-strip').checked });
  function openImport(text, fileName) {
    importExtra = [];
    $('#im-dest').value = 'new';
    loadImportText(text || '', fileName || '');
    dlgImport.showModal();
    if (!text) $('#im-text').focus();
  }
  /** New text arrived (paste, file, drop): detect its shape and preselect the options. */
  function loadImportText(text, fileName) {
    $('#im-text').value = text;
    $('#im-file').textContent = fileName ? fileName : '';
    const info = Importer.analyze(text);
    $('#im-gran').value = info.semanticLines ? 'lines' : 'paragraphs';
    $('#im-strip').checked = info.numbered;
    const notes = [];
    if (info.semanticLines) notes.push('Looks like one sentence per line: importing lines as rows.');
    if (info.numbered) notes.push('Headings carry section numbers: stripping them.');
    $('#im-detect').textContent = notes.join(' ');
    const base = (fileName || '').replace(/\.[^.]+$/, '').replace(/_/g, ' ').trim();
    $('#im-name').value = info.title || base || 'Imported';
    refreshImportPreview();
  }
  function refreshImportPreview() {
    const text = $('#im-text').value;
    const box = $('#im-preview');
    const dest = $('#im-dest').value;
    $('#im-name').parentElement.style.visibility = dest === 'new' || dest === 'split' ? 'visible' : 'hidden';
    if (!text.trim()) { box.innerHTML = '<p class="muted im-empty">Paste Markdown above, choose a file, or drop a .md file on the window.</p>'; $('#im-info').textContent = ''; $('#im-go').disabled = true; return; }
    const res = Importer.parse(text, importOpts());
    const table = Views.el('table', {});
    res.rows.slice(0, 300).forEach(r => {
      const side = Object.entries(r.side || {}).map(([k, v]) => `${k}: ${v}`).join(' · ');
      table.appendChild(Views.el('tr', { class: 'kind-' + r.kind },
        Views.el('td', { class: 'k' }, r.kind), Views.el('td', { class: 'ind' }, r.indent || ''),
        Views.el('td', { class: 't' }, r.text.length > 140 ? r.text.slice(0, 140) + '…' : r.text, side ? Views.el('span', { class: 'side' }, '  ' + side) : null)));
    });
    box.innerHTML = '';
    box.appendChild(table);
    if (res.rows.length > 300) box.appendChild(Views.el('p', { class: 'muted im-empty' }, `… and ${res.rows.length - 300} more rows`));
    const parts = [`${res.stats.rows} rows, ${res.stats.headings} headings`];
    if (res.columns.length) parts.push('side columns: ' + res.columns.join(', '));
    if (importExtra.length) parts.push(`+ ${importExtra.length} more file${importExtra.length > 1 ? 's' : ''} as separate sheets`);
    if (res.warnings.length) parts.push(res.warnings.join(' '));
    $('#im-info').textContent = parts.join(' · ');
    $('#im-go').disabled = !res.rows.length;
  }
  dlgImport.addEventListener('input', e => { if (e.target.id === 'im-text') { const t = e.target.value; if (t.trim()) { const info = Importer.analyze(t); if (!$('#im-file').textContent) { $('#im-name').value = info.title || 'Imported'; } } } refreshImportPreview(); });
  dlgImport.addEventListener('change', refreshImportPreview);
  $('#im-text').addEventListener('paste', e => {
    const html = e.clipboardData && e.clipboardData.getData('text/html');
    if (html && Html2Md.looksRich(html)) {
      e.preventDefault();
      loadImportText(Html2Md.convert(html), '');
      $('#im-detect').textContent = 'Converted from rich text. ' + $('#im-detect').textContent;
      return;
    }
    setTimeout(() => loadImportText($('#im-text').value, ''), 0);
  });
  $('#im-choose').onclick = () => $('#md-input').click();
  $('#md-input').onchange = async e => {
    const files = [...e.target.files]; e.target.value = '';
    if (!files.length) return;
    await importFromFiles(files);
  };
  async function importFromFiles(files) {
    const first = files[0];
    importExtra = await Promise.all(files.slice(1).map(async f => ({ name: f.name.replace(/\.[^.]+$/, '').replace(/_/g, ' '), text: await f.text() })));
    loadImportText(await first.text(), first.name);
    if (!dlgImport.open) dlgImport.showModal();
  }
  $('#im-go').onclick = () => {
    const text = $('#im-text').value;
    if (!text.trim()) return;
    const opts = importOpts();
    const res = Importer.parse(text, opts);
    const dest = $('#im-dest').value;
    const name = $('#im-name').value.trim() || 'Imported';
    const extra = importExtra.slice();
    dlgImport.close();
    let firstIndex = 0;
    mutate(d => {
      for (const k of ['title', 'author', 'description']) if (res.settings[k] && !d.settings[k]) d.settings[k] = res.settings[k];
      if (dest === 'split') {
        const chunks = []; let cur = { name, rows: [] };
        for (const r of res.rows) { if (r.kind === 'h1' && cur.rows.length) { chunks.push(cur); cur = { name: r.text || name, rows: [] }; } else if (r.kind === 'h1' && !cur.rows.length) cur.name = r.text || name; cur.rows.push(r); }
        if (cur.rows.length) chunks.push(cur);
        let at = state.si;
        for (const c of chunks) { at = Model.addSheet(d, c.name.slice(0, 31), at + 1); Model.importRows(d, at, c.rows, { replace: true }); }
        state.si = at;
      } else if (dest === 'new') {
        const at = Model.addSheet(d, name.slice(0, 31), state.si + 1);
        Model.importRows(d, at, res.rows, { replace: true });
        state.si = at;
      } else if (dest === 'append') {
        firstIndex = Model.importRows(d, state.si, res.rows);
      } else {
        Model.importRows(d, state.si, res.rows, { replace: true });
      }
      let at = state.si;
      for (const f of extra) {
        const r2 = Importer.parse(f.text, opts);
        at = Model.addSheet(d, f.name.slice(0, 31) || 'Imported', at + 1);
        Model.importRows(d, at, r2.rows, { replace: true });
      }
      Model.ensureMetaColumns(d);
    });
    state.collapsed.clear();
    focusRow(firstIndex, state.doc.mainColumn, 0);
  };
  // Files dropped on the window: .xlsx opens, .md/.txt imports.
  document.addEventListener('dragover', e => { if ([...e.dataTransfer.types].includes('Files')) { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; } });
  document.addEventListener('drop', async e => {
    if (![...e.dataTransfer.types].includes('Files')) return;
    e.preventDefault();
    const files = [...e.dataTransfer.files];
    const x = files.find(f => /\.xlsx$/i.test(f.name));
    if (x) { if (!confirmDiscard()) return; await loadBuffer(await x.arrayBuffer(), x.name, null); return; }
    const mds = files.filter(f => /\.(md|markdown|txt)$/i.test(f.name));
    if (mds.length) await importFromFiles(mds);
  });

  // ---------- settings dialog ----------
  function openSettings() {
    const d = state.doc, dlg = $('#dlg-settings');
    $('#st-title').value = d.settings.title || '';
    $('#st-author').value = d.settings.author || '';
    $('#st-description').value = d.settings.description || '';
    $('#st-numbering').value = d.settings.numbering === 'per-sheet' ? 'per-sheet' : 'continuous';
    $('#st-freeze').value = d.settings.freezeColumns ?? 1;
    $('#st-track-updated').checked = !!d.settings.trackUpdated;
    $('#st-track-author').checked = !!d.settings.trackAuthor;
    $('#st-track-counts').checked = !!d.settings.trackCounts;
    $('#st-protect').checked = d.settings.protectHeaders !== false;
    $('#st-contents').checked = d.settings.contentsSheet !== false;
    $('#st-enter').value = prefs.enterMode;
    $('#st-theme').value = prefs.theme || 'auto';
    $('#st-target').value = d.settings.wordTarget || '';
    $('#st-copy-enabled').checked = !!prefs.autosaveCopy.enabled;
    $('#st-copy-minutes').value = String(prefs.autosaveCopy.minutes || 2);
    $('#st-copy-name').textContent = state.copyHandle ? state.copyName + (state.copyNeedsPermission ? ' (permission needed)' : '') : (hasFS ? 'no file chosen' : 'needs Edge or Chrome');
    $('#st-copy-choose').disabled = !hasFS;
    $('#st-copy-choose').onclick = () => chooseCopyFile();
    const indSel = $('#st-indent');
    indSel.value = [...indSel.options].some(o => o.value === prefs.indentTrigger) ? prefs.indentTrigger : '   ';
    const sel = $('#st-main'); sel.innerHTML = '';
    const cols = exportColumns(true);
    cols.forEach(c => sel.appendChild(Views.el('option', { value: c, selected: c === d.mainColumn }, c)));
    const roles = d.settings.roles || { status: '', target: '' };
    for (const role of ['status', 'target']) {
      const rs = $('#st-role-' + role); rs.innerHTML = '';
      rs.appendChild(Views.el('option', { value: '' }, '(none)'));
      cols.filter(c => c !== d.mainColumn).forEach(c => rs.appendChild(Views.el('option', { value: c, selected: c === roles[role] }, c)));
      const hint = $('#st-role-' + role + '-hint');
      const sug = !roles[role] && isChapter() ? Model.suggestRole(d, sheet(), role) : null;
      hint.textContent = sug ? `“${sug}” looks like one` : '';
    }
    const countBox = $('#st-count'); countBox.innerHTML = '';
    const counted = new Set((d.settings.countColumns || []).length ? d.settings.countColumns : [d.mainColumn]);
    cols.forEach(c => countBox.appendChild(Views.el('label', { class: 'chk inline' }, Views.el('input', { type: 'checkbox', value: c, checked: counted.has(c) }), ' ', c)));
    const formSnapshot = () => JSON.stringify([...dlg.querySelectorAll('input, select')].map(n => n.type === 'checkbox' ? n.checked : n.value));
    const initial = formSnapshot();
    const refreshSaveButton = () => { $('#st-save').disabled = formSnapshot() === initial; };
    dlg.oninput = dlg.onchange = refreshSaveButton;
    refreshSaveButton();
    $('#st-extra').textContent = Object.keys(d.settings.extra || {}).length
      ? 'Extra keys kept from the file: ' + Object.keys(d.settings.extra).join(', ')
      : '';
    $('#st-save').onclick = () => {
      const countSel = [...countBox.querySelectorAll('input:checked')].map(n => n.value);
      const vals = { title: $('#st-title').value, author: $('#st-author').value, description: $('#st-description').value,
        numbering: $('#st-numbering').value, freezeColumns: Math.max(0, Math.min(10, parseInt($('#st-freeze').value, 10) || 0)),
        countColumns: countSel.length === 1 && countSel[0] === sel.value ? [] : countSel,
        wordTarget: Math.max(0, parseInt(String($('#st-target').value).replace(/\s/g, ''), 10) || 0),
        roles: { status: $('#st-role-status').value, target: $('#st-role-target').value },
        protectHeaders: $('#st-protect').checked, contentsSheet: $('#st-contents').checked,
        trackUpdated: $('#st-track-updated').checked, trackAuthor: $('#st-track-author').checked, trackCounts: $('#st-track-counts').checked };
      const main = sel.value;
      prefs.enterMode = $('#st-enter').value; prefs.indentTrigger = indSel.value; prefs.theme = $('#st-theme').value; applyTheme();
      prefs.autosaveCopy = { enabled: $('#st-copy-enabled').checked, minutes: parseInt($('#st-copy-minutes').value, 10) || 2 }; savePrefs();
      if (prefs.autosaveCopy.enabled && !state.copyHandle && hasFS) chooseCopyFile();
      dlg.close();
      mutate(doc => { Object.assign(doc.settings, vals); if (main && main !== doc.mainColumn) Model.setMainColumn(doc, main); Model.ensureMetaColumns(doc); });
    };
    dlg.showModal();
  }

  // ---------- autosave (IndexedDB) ----------
  const idb = {
    open() {
      return new Promise((res, rej) => {
        const r = indexedDB.open('sheetwriter', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('kv');
        r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
      });
    },
    async get(k) { const db = await this.open(); return new Promise((res, rej) => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); },
    async set(k, v) { const db = await this.open(); return new Promise((res, rej) => { const q = db.transaction('kv', 'readwrite').objectStore('kv').put(v, k); q.onsuccess = () => res(); q.onerror = () => rej(q.error); }); },
    async del(k) { const db = await this.open(); return new Promise((res, rej) => { const q = db.transaction('kv', 'readwrite').objectStore('kv').delete(k); q.onsuccess = () => res(); q.onerror = () => rej(q.error); }); },
  };
  let lastAutosave = '', lastHistoryAt = 0;
  setInterval(() => {
    if (!state.dirty) return;
    const snap = snapshot();
    if (snap === lastAutosave) return;
    lastAutosave = snap;
    idb.set('autosave', { doc: state.doc, fileName: state.fileName, at: Date.now() }).catch(() => {});
    // Keep a short history of snapshots (one every 5 minutes, six kept) for "Recover an autosave…".
    if (Date.now() - lastHistoryAt > 5 * 60000) {
      lastHistoryAt = Date.now();
      idb.get('autosave_history').then(h => idb.set('autosave_history', [{ doc: state.doc, fileName: state.fileName, at: Date.now() }].concat(h || []).slice(0, 6))).catch(() => {});
    }
  }, 3000);
  async function recoverMenu(anchor) {
    let h = [];
    try { h = (await idb.get('autosave_history')) || []; } catch (e) { /* ignore */ }
    if (!h.length) { showMenu(anchor, [{ label: 'No autosave snapshots yet (one is kept every 5 minutes while you edit).', disabled: true, action() {} }]); return; }
    showMenu(anchor, h.map(e => ({
      label: `${new Date(e.at).toLocaleString()} — ${e.fileName}`,
      action: () => { if (!confirmDiscard()) return; loadDoc(e.doc, e.fileName, null, new Map()); state.dirty = true; render(); },
    })));
  }
  async function offerRestore() {
    try {
      const a = await idb.get('autosave');
      if (!a || !a.doc) return;
      const when = new Date(a.at).toLocaleString();
      if (confirm(`Unsaved work from ${when} (${a.fileName}) was found. Restore it?`)) {
        loadDoc(a.doc, a.fileName, null, new Map());
        state.dirty = true; render();
      } else {
        await idb.del('autosave');
      }
    } catch (e) { /* no IndexedDB */ }
  }
  window.addEventListener('beforeunload', e => { if (state.dirty) { e.preventDefault(); e.returnValue = ''; } });

  // Textarea heights depend on width: re-measure when the view is resized (or first laid out).
  let resizeTimer = null, lastWidth = 0;
  const onResize = () => {
    const w = viewRoot.clientWidth;
    if (w === lastWidth) return;
    lastWidth = w;
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => { Views.autosizeAll(viewRoot); if (state.view === 'grid') Views.applyFreeze(viewRoot.querySelector('table.grid'), state.doc.settings.freezeColumns); }, 50);
  };
  if (window.ResizeObserver) new ResizeObserver(onResize).observe(viewRoot);
  window.addEventListener('resize', onResize);

  // ---------- boot ----------
  render();
  offerRestore();
  window.SheetWriter = { state, render, prefs, Model, XlsxIO, Exporter, Importer, APP, openImport, currentViewState, applyViewState, loadDoc };
})();

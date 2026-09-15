# SheetWriter — plan

*Drafting tool where the document is a spreadsheet: one argument per row, columns for notes, sources, examples, status. Saves to XLSX so the file stays a normal Excel workbook. Runs as a single HTML file, no installation.*

Written 2026-09-10. Revised same day: XLSX from the start; hierarchical numbering, Markdown export and sheets-as-chapters moved into the first milestones. Later the same day: cells are Markdown (no rich text), settings live in a visible protected sheet, and M0 plus most of M1 were built (see §10).

---

## 1. Does this already exist?

Short answer: no. Nothing found uses a spreadsheet as its *native* save format for row-per-argument drafting and also runs without installation. The closest things, and why they fall short:

| Tool | What it does | Why it doesn't fit |
|---|---|---|
| **DraftWriter** (Mac App Store) | Exactly the idea: each sentence is a cell, four assignable columns (topic sentence, support, etc.), reorder with keyboard. | Mac only, in-app subscription, own file format. No Windows, no CSV/XLSX round trip found. |
| **Scrivener** / **OmniOutliner** | Outliner with custom columns per item; can *export* outline to CSV. | Desktop install; CSV is one-way export, not the working format. OmniOutliner is Mac only. |
| **Grist** (hosted or self-hosted), **Airtable**, **Notion** databases | Row = record, long-text fields, record-card view, CSV/XLSX import and export, drag row order. | Needs an account and a hosted service (work-computer policy risk). Editing is form-like, not prose-like; no "read it as an article" view; export is a round trip, not the file itself. Grist is the best of these if you ever want a hosted option. |
| **Semantic line-break editors** (Obsidian with strict line breaks, Writemonkey, VS Code + Markdown folding) | One sentence per line in source, normal paragraphs in preview. | Plain text only: no columns, no metadata per line. |
| **Browser CSV editors** (csv-viewer-online, Grist CSV viewer widget, csvtool.io) | Edit CSV in browser, some do row drag. | Generic grid; long text in cells is as painful as in Excel. No prose view, no headings, no compile. |
| **Zapier, "Write faster with spreadsheets"** | Advocates precisely this workflow in Google Sheets: row per idea, side columns for notes and links. | Confirms the workflow is useful and confirms the pain: spreadsheets are not word processors. |

Conclusion: the workflow is recognised, but the tool is missing. The gap is small enough that a single HTML file can fill it.

---

## 2. Product definition

**One sentence:** a text editor whose document is an Excel workbook.

**Core loop**

1. Open a `.xlsx` workbook, or start a new one.
2. Each sheet is a chapter. Each row is an argument: main text large, other columns beside or under it.
3. Type, add rows with Enter, move rows with Alt+Up/Down or drag, split and merge rows. Numbering updates live.
4. Switch to *Read* view to see the rows flowing as prose. Export the text column as one Markdown file.
5. Save. The file is a plain workbook that Excel opens normally. Sheet order = chapter order, row order = argument order.

**Non-goals for v1:** collaboration, sync, formulas, footnote management. Rich text is replaced by Markdown in cells (see §3).

---

## 3. Data model (what the workbook contains)

Keep the file a boring, honest workbook that makes sense when opened in Excel.

### Sheets

- **Workbook = document. Sheet = chapter.** Sheet order is chapter order; sheet name is the chapter title.
- A sheet is a *chapter sheet* if its header row contains the main text column (default name `text`). Any other sheet (`sources`, `README`, a table of data) is a *data sheet*: shown as a tab but not editable in Draft view, and written back untouched.
- A visible sheet `_sheetwriter` with three columns `key`, `value`, `description` holds settings and metadata: title, author, description, `main_column`, `chapter_prefix`, created, modified, app version. Unknown keys you add yourself are kept. The sheet is protected without a password so it is not edited by accident in Excel (Review → Unprotect Sheet lifts it). The app works without the sheet; it exists so settings travel with the file between home and work.

### Cells are Markdown

No rich text. Bold, italic, links, inline code, lists and line breaks are written as Markdown in the cell. The app renders a row as formatted text when it is not being edited and shows the raw Markdown while you type; Excel shows the raw Markdown. Export passes it through unchanged. Raw HTML in cells is escaped, and links are limited to http(s) and mailto.

### Columns (per chapter sheet)

- **Header row** = column names, row 1.
- **Main text column** (default `text`). This is what you are writing.
- **`kind` column**, structural, values:
  - `h1` `h2` `h3` `h4` — headings
  - `p` (or empty) — a paragraph of body text
  - `s` — a sentence that *continues the previous paragraph*. This lets you keep one argument per row while still producing multi-sentence paragraphs on export.
  - `x` — excluded from export (a note to self, a parked argument). Still numbered, shown dimmed.
- **`no` column**, written by the app on save, ignored on load. It holds the computed hierarchical number so Excel shows the same numbering. Never edit it by hand; it is regenerated.
- **All other columns are yours**: `notes`, `sources`, `examples`, `status`, `counter`. Add, rename and reorder freely. All are free text.

No id column needed. Position is identity.

### Hierarchical numbering (computed, never stored as truth)

Rules:

1. A heading of level L gets number `c1.c2…cL` where cL is incremented and all deeper counters reset.
2. A body row (`p`, `s`, `x`) is numbered *one level below the nearest preceding heading*: after `h1` = 1 it gets `1.1`, `1.2`…; after `h2` = 1.1 it gets `1.1.1`, `1.1.2`…
3. Body rows before any heading are numbered at level 1: `1`, `2`, `3`.
4. A heading that skips a level (`h1` then `h3`) is treated as the next allowed level (`h2`) for numbering only; export still emits it as `###`. The gutter shows a small warning marker so you notice.
5. Numbering restarts per sheet. With more than one chapter sheet, the chapter index is prefixed: chapter 2, first heading = `2.1`, first body row under it = `2.1.1`. Prefix can be switched off in settings.

Worked example (matches the request):

```
kind  text                      no
h1    Introduction              1
p     First argument            1.1
p     Second argument           1.2
```
insert an `h2` between the heading and the first argument:
```
h1    Introduction              1
h2    Background                1.1
p     First argument            1.1.1
p     Second argument           1.1.2
```
Numbering is recomputed on every edit, move, kind change, and sheet reorder, so it is always correct in the app; the `no` column in the file is only as fresh as the last save.

### Excel considerations

- **Styles the app applies on save** so the workbook is pleasant in Excel: bold frozen header row, wrap text in the main column, sensible column widths, heading rows bold and slightly larger, `x` rows grey. This is why the XLSX library must write styles (see §5).
- **Formula injection:** a cell starting with `=`, `+`, `-`, `@` is read by Excel as a formula. The app writes all cells as explicit strings, so this is a non-issue with XLSX. (It was the main CSV headache.)
- **Encoding and delimiters:** non-issue with XLSX. Finnish characters are safe.
- **Excel edits between app sessions** are fine: the app reads whatever is in the sheet, recomputes numbering, and ignores the stale `no` column.

---

## 4. UI

Three views, sheet tabs, one toolbar, keyboard-first.

**Sheet tabs** (bottom, like Excel, or left rail): one per sheet, in workbook order. Drag to reorder chapters, double-click to rename, `+` to add. Data sheets shown with a different icon. Row cut/paste and drag-onto-tab move rows between chapters.

**Draft view (default)**
- Vertical list of cards. Each card: number in the gutter, kind badge, main text as an auto-growing textarea, side columns as small labelled fields, collapsible per card or in a right-hand panel for the focused row.
- Heading rows render larger and can collapse the rows beneath them.
- Keyboard: `Enter` new row below (same kind, or `p` after a heading) · `Shift+Enter` newline inside cell · `Alt+Up/Down` move row · `Alt+Left/Right` promote/demote (`h1`→`h2`→…→`p`→`s`) · `Ctrl+D` duplicate · `Ctrl+Backspace` on empty row deletes it · `Ctrl+Shift+S` split at cursor · `Ctrl+J` merge with next · `Tab` next side column · `Ctrl+Z/Y` undo/redo · `Ctrl+PageUp/PageDown` previous/next sheet.
- Status bar: word count for the row, the chapter, and the workbook.

**Grid view**
- Plain table of the current sheet, all columns, inline editing. For bulk metadata work. Column manager (add/rename/hide/reorder) lives here.

**Read view**
- The current chapter, or the whole workbook, flowing as prose: headings as headings, `p` rows as paragraphs, `s` rows joined to the previous paragraph, `x` rows omitted. Toggle numbering on/off. Optional margin column (e.g. show `sources` beside each paragraph).

**Toolbar:** Open · Save · Save As · New · view switch · Export Markdown · column manager · filter (later) · help.

---

## 5. Markdown export

Export is *by column*: you pick which column becomes the document. Default is the main text column, but exporting `notes` or `counter` as its own document is the same operation.

Options in the export dialog:

| Option | Default | Notes |
|---|---|---|
| Column | main text | any column |
| Scope | whole workbook | or current sheet only |
| Sheet titles | `#` heading per sheet, from sheet name | off for single-chapter files |
| Heading levels | `kind` h1..h4 → `##`..`#####` when sheet titles are on, `#`..`####` when off | |
| Numbering | off | prepends the computed number to headings and paragraphs |
| Paragraphs | `p` starts a paragraph, `s` continues it, `x` omitted | |
| Side columns | none | optional: append a chosen column as a blockquote under each paragraph (`> sources: …`), or as an HTML comment `<!-- notes: … -->` so it survives but does not render |
| Output | copy to clipboard | or download `name.md` |

Also planned, same machinery: plain text, and later `.docx` via an inlined docx library (M2).

---

## 6. Technical design

**Form factor:** one file, `sheetwriter.html`, vanilla JS, no framework. Works from `file://` by double-clicking, and works unchanged on any static host (GitHub Pages, intranet share, OneDrive). No server component.

**Saving files, three tiers (auto-selected):**
1. **File System Access API** (Chrome/Edge): real Open and Save-in-place, remembers the file handle for the session. Works from `file://` in Chromium.
2. **Fallback:** `<input type=file>` to open, browser download to save. Works even under group policy that disables the API.
3. **Crash protection:** autosave the working copy to IndexedDB every few seconds; on startup offer to restore an unsaved session.

**XLSX library: ExcelJS** (browser build, about 1 MB minified, inlined). Chosen over SheetJS community edition because ExcelJS reads *and writes* cell styles, column widths, frozen panes and hidden sheets in the free version, and the workbook must look right in Excel after every save. SheetJS CE would strip formatting on write. Risk: ExcelJS is slower and stricter on odd files; documents here are small, and the app only needs to read files it or Excel wrote. If ExcelJS chokes on some real-world file, SheetJS stays as a fallback reader.

**Round-trip policy:** chapter sheets are rewritten from the app's model with the house style. Data sheets are copied through as loaded. Anything ExcelJS cannot represent (charts, macros, pivot tables) will be lost, and the app warns on open if the workbook contains such things. Keep the drafting workbook separate from heavy analysis workbooks.

**Repo layout**
```
sheetwriter/
  src/
    index.html        # markup + script tags for dev
    app.js            # document model, commands, undo
    numbering.js      # hierarchical numbering (pure function: rows -> numbers)
    xlsx-io.js        # ExcelJS load/save, house style, settings sheet
    export.js         # Markdown / text from a column
    views/            # draft.js, grid.js, read.js, tabs.js
    style.css
  vendor/exceljs.min.js
  build.py            # inlines src + vendor into dist/sheetwriter.html
  dist/sheetwriter.html
  tests.html          # numbering and export cases, runs in the browser
  PLAN.md
```
Development happens in `src/`; `build.py` (a 30-line inliner) produces the single distributable file. Copy `dist/sheetwriter.html` to the work computer, or open it from OneDrive.

**State and undo:** in-memory document `{sheets: [{name, kind: 'chapter'|'data', columns, rows}], mainColumn, settings, dirty}`. Every command pushes a snapshot on an undo stack. Numbering is a pure function of the rows and is recomputed on render, never edited.

**Testing:** `tests.html` runs numbering cases (insert heading between, skipped levels, body before heading, multi-sheet prefix), export cases (`s` joining, `x` omission, side-column modes) and an XLSX round trip in the browser. No test runner to install.

---

## 7. Milestones

**M0 — usable in one sitting**
Open/save XLSX with house style · sheet tabs (switch, add, rename, reorder) · Draft view with main text and `kind` · Enter / Alt-arrows / promote-demote / delete · live hierarchical numbering with `no` column written on save · download fallback · autosave.
*Done when: you can draft a two-chapter piece, save, open in Excel, reorder rows in Excel, reopen in the app, and the numbering is right.*

**M1 — the real v1**
Markdown export by column with the options in §5 · side columns on cards · column manager · Read view · undo/redo · split/merge · move rows between sheets · word counts · settings sheet · `tests.html`.

**M2 — polish and reach**
Filter/search by column value · multi-select move · collapse under headings · `.docx` export · plain-text export · dark mode · hosted copy on GitHub Pages.

**M3 — maybe**
`sources` column linked to a `sources` data sheet (pick from list, render as citations on export) · per-row comments via Excel cell comments · CSV import for legacy files.

---

## 8. Decisions taken (change if you disagree)

- **ExcelJS rather than SheetJS.** For style-preserving writes in the free version. See §6.
- **Body text numbered one level below its heading**, and body rows before any heading numbered at level 1. Alternative would be numbering only headings; rejected because arguments are the unit you rearrange, so they need addresses.
- **Skipped heading levels are clamped for numbering** and flagged, not rendered as `1.0.1`.
- **Chapter prefix from sheet position** rather than requiring an `h1` per sheet. Switchable off.
- **`s` kind for sentence continuation** so one-argument-per-row still yields real paragraphs on export. Alternative was a `para` column; a kind value is simpler and visible in Excel.
- **`no` column written on save, ignored on load.** Excel users see numbers; the app never trusts them.
- **One hidden settings sheet allowed** (`_sheetwriter`), replacing the earlier "nothing hidden" rule, so settings travel with the file. The app tolerates its absence.
- **Main column found by name (`text`)**, not by position, so existing sheets with a different layout still work.
- **No framework.** Vanilla JS keeps the single-file promise honest.

---

## 10. Status (2026-09-10)

Built and tested in the browser: `dist/sheetwriter.html`, one file, about 1 MB.

Done from M0: XLSX open/save with house style, sheet tabs (add, rename, reorder by drag, delete, data sheets read-only), Draft view with Markdown rendering, Enter/split, Alt+arrows, promote/demote, kind cycling, live numbering with the `no` column written on save, File System Access API with download fallback, IndexedDB autosave and restore.

Done from M1: Markdown export by column with scope, sheet titles, numbering and side-column options; side columns on cards; column manager in Grid view; Read view; undo/redo; merge, duplicate, delete; move rows between chapters by drag or Ctrl+Shift+M; word counts; settings sheet; `tests.html` with 20 cases.

Revision of 2026-09-11 after first feedback: compact Draft layout (badge and number left, side columns right, empty side fields shown only while the card is focused); collapse/expand of heading sections and paragraph groups with collapsed sections moving as a unit (keyboard and drag); Markdown prefixes set the kind (`# ` … `#### ` → h1–h4, `- ` or `* ` → s); Enter adds a row in Grid view too, Shift+Enter is a line break; drag handles in Grid; inline renaming of columns and sheets, and in-page menus, replacing every `window.prompt` (some embedded browsers block it); row/word/character counters for row, chapter and workbook; Read view and export numbering: none, headings only, or all; a leading `#` typed into a heading cell is no longer doubled on export.

Collapsed-section moves are defined as "past the next visible unit". Moving a collapsed h1 section down past an *expanded* h1 heading therefore lands it between that heading and its children. Sibling-aware moves (skip the whole next section) are a possible refinement.

Revision of 2026-09-11, second round (v0.3.0): indentation with an `indent` column, numbering that follows it, and collapse of indented groups; sibling-aware moves (a row always moves with its block; up = before the previous row at the same or a higher structural level, down = past the next sibling's block, or as first child of the next parent); the `- ` shortcut is gone, three spaces (configurable) or Tab indent instead; Enter/Shift+Enter behaviour configurable, line breaks inserted by the app; empty rows removed on leaving them and never written; Alt+Shift+←/→ for promote/demote (Alt+arrows alone are back/forward in Edge); undo/redo buttons; title with unsaved-changes dot in the toolbar, file name in the status bar; frozen grid header and number column; Excel-edited workbooks re-sorted by `no`, rows without a number appended; plain spreadsheets imported by their longest column; optional `updated`/`author` meta columns, always last in the file; save cancellation no longer falls back to a download; favicon; version and project links in Settings and in the `_sheetwriter` sheet; Recent files (Chromium); GitHub Pages build in `docs/`.

Bug found on the way: the main textarea had the class `main` too, so `closest('.main')` returned the textarea itself and the "editing" class never reached the wrapper. Focus after a move or promote therefore failed on non-empty rows (it worked on empty rows only because those show the textarea anyway). The wrapper is now `mainwrap`.

Not yet done at that point: filter/search, multi-select, `.docx` export, dark mode, mobile layout, releases with a forced-download asset.

## 12. Status (2026-09-15, 0.13.0)

Since §10 and §11 the app has gained, in order: a table of contents pane; Markdown import with sentence and line modes and rich-text paste; workbook-wide columns with a Columns menu, resizing, hiding and drag-reordering; display state saved in the workbook; view switches that keep the position; multi-row selection with copy, cut, paste, move and delete; find and replace; Word export from a built-in OOXML writer; an overwrite guard, autosave snapshots and an autosave copy to a second workbook; paragraph-level anchors in Read view; Ctrl+B/I/K; status chips, word targets and a row filter; dark mode and print styles; headless tests in GitHub Actions; and the 0.10 file-format change: system columns carry a dot (`.no`, `.kind`, `.indent`, `.words`, `.chars`, `.updated`, `.author`), special columns are chosen by role in Settings rather than by name, the settings sheet is `.sheetwriter`, every workbook carries instructions for editing in Excel, and chapter sheets protect their header row. 0.12 added a `.contents` sheet with links, Export and Import toolbar buttons and saved-time in the status bar; 0.13 added OpenDocument: `.ods` as a second save format (same sheets, settings, contents and protection, written by `src/odf.js` on JSZip and read back through the same loader as XLSX) and `.odt` export alongside Word.

The README is the current user documentation; the help dialog in the app mirrors it. Sections 3 to 9 above describe the original design and are kept for the reasoning; where they conflict with the README, the README is right.

0.14 gave every row a persistent `.id` (hidden last column) and keeps Excel fills, borders, comments and font name/colour/underline/strike on chapter sheets by that id; bold, italic and size stay the app's, formulas and row heights are dropped. 0.15 made the contents pane an outliner: drag sections between headings and sheets, promote and demote with sub-headings, Alt+arrow keys. 0.16 added footnotes as a side-column mode (Markdown `[^n]`, real footnotes in Word and OpenDocument) and export presets stored as `export:name` rows in `.sheetwriter`. 0.17: CSV import, Print / PDF, the suite runs in Firefox and WebKit as well, CHANGELOG.md.

Open: formatting of data sheets in `.ods` files (only cell values are carried through), frozen panes in `.ods`, real Firefox and Safari use on a desk (the automated suite passes in both engines), touch editing on phones, releases with a forced-download asset. The `.ods` and `.odt` writers are checked for well-formedness and round trips by the tests but have not yet been opened in LibreOffice.

One technical note worth keeping: ExcelJS's browser build yields through zero-delay `setTimeout` thousands of times per file, which browsers clamp and, in hidden tabs, throttle. `src/timers.js` routes zero-delay timeouts through a MessageChannel and must load before ExcelJS. It took a 7 KB workbook from 6 to 45+ seconds down to 25 ms.

Not exercised by automation: the native Open and Save dialogs, and opening a file that Excel itself has saved (the tests build such a file with ExcelJS instead).

## 9. Sources consulted

- DraftWriter on the App Store: https://apps.apple.com/us/app/draftwriter/id6484404149
- Zapier, "Write faster with spreadsheets": https://zapier.com/blog/spreadsheets-for-writers/
- Outliner Software forum, "Sentence outliner?": https://www.outlinersoftware.com/topics/viewt/10520/
- Scrivener outliner CSV export: https://www.oreilly.com/library/view/scrivener-for-dummies/9781118312469/a13_15_9781118312469-ch09.html
- OmniOutliner export formats: https://support.omnigroup.com/documentation/omnioutliner/mac/5.1.2/en/importing-exporting-and-printing/
- Grist: https://www.getgrist.com/ and its in-browser CSV viewer https://www.getgrist.com/csv-viewer/
- Gingko Writer export options (no CSV): https://docs.gingkowriter.com
- Browser CSV editors: https://csv-viewer-online.github.io/ , https://www.csvtool.io/csv-editor
- TreeSheets (desktop only, has CSV export): https://github.com/TreeSheets-Spreadsheet
- ExcelJS: https://github.com/exceljs/exceljs

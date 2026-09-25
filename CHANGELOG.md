# Changelog

One line per release; the commit under each tag has the details. Dates are the day the tag was pushed.

## 0.19.8 (2026-09-25)
Grid: Right arrow at the end of a cell's text moves to the next cell, Left arrow at its start to the previous one (wrapping across rows like Tab).

## 0.19.7 (2026-09-18)
Draft: the side area spans the whole card height, so any spot right of the divider opens the side columns; the hint naming them only appears on hover.

## 0.19.6 (2026-09-18)
Draft: clicking the empty area right of the text opens the side columns, which now show a hint naming them. Bigger collapse/expand triangles.

## 0.19.5 (2026-09-16)
Heading rows have a background tint that fades with the level, in Grid and Draft, light and dark.

## 0.19.4 (2026-09-16)
Grid: heading levels are told apart as in Draft (h1 and h2 bold and larger, h3 bold italic, h4 medium italic and greyer).

## 0.19.3 (2026-09-16)
Grid: heading rows are bold (and x rows grey) in the rendered cell too, not only while editing.

## 0.19.2 (2026-09-16)
Grid cells render their Markdown (italics, bold, links, lists) and open for editing on click; Tab moves between cells. A click anywhere outside the selection controls clears the row selection.

## 0.19.1 (2026-09-15)
Selected cards and grid rows drag as a whole (not just their handles); the selection bar says so.

## 0.19.0 (2026-09-15)
Long sheets stay responsive (off-screen cards and read blocks are not laid out; Grid autosizing batched, 30× faster). Defaults for new rows. Merge a chapter into the previous one; a section to a new sheet from the Contents pane. Excel behaviour verified with real Excel: kept formatting, drop-down sort and filter on protected sheets. Phone-width layout: floating contents pane, narrower cards. Screenshots and a GIF in the README.

## 0.18.0 (2026-09-15)
Named snapshots kept in the browser; compare the editor with a snapshot, the saved file or another workbook: rows added, removed and changed with word-level diffs.

## 0.17.0 (2026-09-15)
CSV import, Print / PDF from the Export dialog, page breaks before top headings when printing, narrower toolbar on small screens, tests runnable in Firefox and WebKit (`npm test -- --browser=firefox`), this changelog.

## 0.16.0 (2026-09-15)
Footnotes as a side-column mode (Markdown `[^n]`, real footnotes in Word and OpenDocument). Export presets stored as `export:name` rows in `.sheetwriter`.

## 0.15.0 (2026-09-15)
Contents pane as an outliner: drag sections between headings and sheets, promote and demote with sub-headings, Alt+arrow keys.

## 0.14.0 (2026-09-15)
Row identity: hidden `.id` column; Excel fills, borders, comments and font name/colour/underline/strike on chapter sheets follow the row across saves.

## 0.13.1 (2026-09-15)
Filter and Columns buttons greyed out in Read view.

## 0.13.0 (2026-09-15)
OpenDocument: `.ods` as a save format, `.odt` export.

## 0.12.1 (2026-09-15)
Red delete buttons in Grid.

## 0.12.0 (2026-09-15)
`.contents` sheet with links, Export and Import toolbar buttons, saved time in the status bar, Grid header fixes, filter bar only on request.

## 0.11.0 (2026-09-15)
Instructions for editing in Excel written into every workbook; header rows of chapter sheets protected.

## 0.10.1 (2026-09-15)
Settings sheet renamed `.sheetwriter` (`_sheetwriter` still read).

## 0.10.0 (2026-09-15)
Dotted system columns (`.no`, `.kind`, `.indent`, `.words`, `.chars`, `.updated`, `.author`); status and target columns chosen by role.

## 0.9.0 (2026-09-15)
Ctrl+B/I/K, rich-text paste, status chips, word targets, row filter, dark mode, print styles, CI.

## 0.8.0 (2026-09-15)
Find and replace, Word export, overwrite guard, autosave snapshots and autosave copy, paragraph anchors in Read view.

## 0.7.1 (2026-09-14)
Hide columns and rows in Grid view.

## 0.7.0 (2026-09-14)
Multi-row selection with copy, cut, paste, duplicate, delete and move.

## 0.6.3 (2026-09-14)
Scroll position kept on save and re-render; contents jumps land in view.

## 0.6.2 (2026-09-14)
Resizable Grid columns.

## 0.6.1 (2026-09-14)
View switches keep the position in the text.

## 0.6.0 (2026-09-14)
Workbook-wide columns, Columns menu, display state saved in the workbook.

## 0.5.0 (2026-09-11)
Markdown import.

## 0.4.0 (2026-09-11)
Table of contents pane.

## 0.3.2 (2026-09-11)
Per-row word and character counts, section totals.

## 0.3.1 (2026-09-11)
Absolute heading levels with numbering across sheets, frozen columns, counted columns.

## 0.3.0 (2026-09-11)
First public build: single-file editor for spreadsheet-backed writing.

# Changelog

One line per release; the commit under each tag has the details. Dates are the day the tag was pushed.

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

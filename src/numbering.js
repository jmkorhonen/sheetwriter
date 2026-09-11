/* SheetWriter — hierarchical numbering. Pure function, no DOM.
 *
 * Rules (see PLAN.md §3):
 *  - heading hN: counter N increments, deeper counters reset
 *  - body row (p, s, x): numbered one level below the nearest preceding heading,
 *    plus its indent (0, 1, 2 …): after h1 = 1, a body row at indent 0 is 1.1,
 *    a body row at indent 1 below it is 1.1.1
 *  - body before any heading: level 1 (+ indent)
 *  - skipped heading level (h1 then h3) or skipped indent (0 then 2): clamped
 *    to the next allowed level, flagged
 *  - optional chapter prefix, e.g. "2." for the second chapter sheet
 */
const SheetNumbering = (() => {
  function headingLevel(kind) {
    const m = /^h([1-6])$/.exec(String(kind || '').trim().toLowerCase());
    return m ? +m[1] : 0;
  }
  function indentOf(row) {
    const n = parseInt(row && row.indent, 10);
    return n > 0 ? n : 0;
  }

  /** rows: [{kind, indent}], opts: {prefix?} → {numbers, warnings, levels, indents} */
  function compute(rows, opts = {}) {
    const prefix = (opts.prefix !== null && opts.prefix !== undefined && opts.prefix !== '') ? String(opts.prefix) + '.' : '';
    const counters = [];
    let last = 0;          // effective level of the last heading seen
    let lastIndent = -1;   // effective indent of the previous body row (-1 right after a heading)
    const numbers = [], warnings = [], levels = [], indents = [];
    for (const r of rows) {
      const hl = headingLevel(r.kind);
      let level, warn = false, eff = 0;
      if (hl) {
        level = Math.min(hl, last + 1);
        warn = level !== hl;
        last = level;
        lastIndent = -1;
      } else {
        const ind = indentOf(r);
        eff = Math.min(ind, lastIndent + 1);
        warn = eff !== ind;
        lastIndent = eff;
        level = last + 1 + eff;
      }
      counters.length = level;
      for (let i = 0; i < level; i++) if (!counters[i]) counters[i] = 0;
      counters[level - 1]++;
      numbers.push(prefix + counters.join('.'));
      warnings.push(warn);
      levels.push(level);
      indents.push(eff);
    }
    return { numbers, warnings, levels, indents };
  }

  return { compute, headingLevel, indentOf };
})();
if (typeof module !== 'undefined') module.exports = SheetNumbering;

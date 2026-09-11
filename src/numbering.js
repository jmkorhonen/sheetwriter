/* SheetWriter — hierarchical numbering. Pure function, no DOM.
 *
 * Rules (see PLAN.md §3):
 *  - heading hN: counter N increments, deeper counters reset. Levels are absolute:
 *    h1 is always one number (1, 2, 3), h2 two (2.1), h3 three (2.1.1) …
 *  - body row (p, s, x): numbered one level below the nearest preceding heading,
 *    plus its indent (0, 1, 2 …): after h1 = 1, a body row at indent 0 is 1.1,
 *    a body row at indent 1 below it is 1.1.1
 *  - body before any heading: level 1 (+ indent)
 *  - skipped heading level (h1 then h3) or skipped indent (0 then 2): clamped
 *    to the next allowed level, flagged
 *  - numbering can continue across sheets: pass the previous sheet's `state`
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

  /** rows: [{kind, indent}], opts: {state?} → {numbers, warnings, levels, indents, state} */
  function compute(rows, opts = {}) {
    const st = opts.state
      ? { counters: opts.state.counters.slice(), last: opts.state.last, lastIndent: opts.state.lastIndent }
      : { counters: [], last: 0, lastIndent: -1 };
    const numbers = [], warnings = [], levels = [], indents = [];
    for (const r of rows) {
      const hl = headingLevel(r.kind);
      let level, warn = false, eff = 0;
      if (hl) {
        level = Math.min(hl, st.last + 1);
        warn = level !== hl;
        st.last = level;
        st.lastIndent = -1;
      } else {
        const ind = indentOf(r);
        eff = Math.min(ind, st.lastIndent + 1);
        warn = eff !== ind;
        st.lastIndent = eff;
        level = st.last + 1 + eff;
      }
      st.counters.length = level;
      for (let i = 0; i < level; i++) if (!st.counters[i]) st.counters[i] = 0;
      st.counters[level - 1]++;
      numbers.push(st.counters.join('.'));
      warnings.push(warn);
      levels.push(level);
      indents.push(eff);
    }
    return { numbers, warnings, levels, indents, state: st };
  }

  return { compute, headingLevel, indentOf };
})();
if (typeof module !== 'undefined') module.exports = SheetNumbering;

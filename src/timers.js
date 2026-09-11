/* SheetWriter — zero-delay timers via MessageChannel. Must load BEFORE ExcelJS.
 *
 * ExcelJS's browser bundle yields through process.nextTick → setTimeout(fn, 0)
 * thousands of times per workbook. Browsers clamp nested zero-delay timeouts to
 * 4 ms and throttle them to once per second (later once per minute) in hidden
 * tabs, which turns a 10 KB file into a multi-second or multi-minute load.
 * MessageChannel tasks have neither limit. Timeouts with a positive delay are
 * left alone.
 */
(() => {
  if (typeof MessageChannel === 'undefined' || typeof window === 'undefined') return;
  const nativeSet = window.setTimeout.bind(window);
  const nativeClear = window.clearTimeout.bind(window);
  const pending = new Set();
  const queue = [];
  let seq = 0;
  const ch = new MessageChannel();
  ch.port1.onmessage = () => {
    const job = queue.shift();
    if (!job || !pending.delete(job.id)) return;
    try { job.fn(...job.args); }
    catch (e) { if (window.reportError) window.reportError(e); else console.error(e); }
  };
  window.setTimeout = function (fn, ms, ...args) {
    if (typeof fn !== 'function' || (Number(ms) || 0) > 0) return nativeSet(fn, ms, ...args);
    const id = -(++seq);
    pending.add(id);
    queue.push({ id, fn, args });
    ch.port2.postMessage(null);
    return id;
  };
  window.clearTimeout = function (id) {
    if (typeof id === 'number' && id < 0) pending.delete(id);
    else nativeClear(id);
  };
  if (!window.setImmediate) window.setImmediate = (fn, ...args) => window.setTimeout(fn, 0, ...args);
})();

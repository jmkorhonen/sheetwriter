/* SheetWriter — Markdown rendering for cells and the Read view.
 * Raw HTML in cells is disabled by escaping "<" before parsing; links are
 * restricted to http(s)/mailto and open in a new tab.
 */
const MD = (() => {
  function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function attr(s) { return String(s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;'); }

  let ready = false;
  function setup() {
    if (ready || typeof marked === 'undefined') return;
    marked.use({
      breaks: true,
      gfm: true,
      renderer: {
        link(token) {
          const href = /^(https?:|mailto:|#)/i.test(token.href || '') ? token.href : '#';
          const text = this.parser.parseInline(token.tokens);
          const title = token.title ? ` title="${attr(token.title)}"` : '';
          return `<a href="${attr(href)}" target="_blank" rel="noopener"${title}>${text}</a>`;
        },
        html(token) { return escapeHtml(token.text); },
      },
    });
    ready = true;
  }

  /** Block-level render (paragraphs, lists, headings inside a cell all work). */
  function render(src) {
    if (!src) return '';
    setup();
    try {
      return marked.parse(String(src).replace(/</g, '&lt;'));
    } catch (e) {
      return '<p>' + escapeHtml(src) + '</p>';
    }
  }

  /** Inline render: no wrapping <p>, for one-line cells. */
  function renderInline(src) {
    if (!src) return '';
    setup();
    try {
      return marked.parseInline(String(src).replace(/</g, '&lt;'));
    } catch (e) {
      return escapeHtml(src);
    }
  }

  return { render, renderInline, escapeHtml };
})();

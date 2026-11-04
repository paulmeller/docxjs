describe('Annotation render fixtures', function () {
  // These snapshots project content and card metadata, excluding font/layout pixel
  // values and generated identifiers. The existing full upstream snapshots remain unchanged.
  for (const [name, mode, comments] of [
    ['revision-margin', 'margin', false],
    ['revision-floating', 'floating', false],
    ['comments', 'margin', true],
    ['comments', 'floating', true]
  ]) {
    it(`renders ${name} in ${mode} mode`, async () => {
      const blob = await fetch(`/base/tests/render-test/${name}/document.docx`).then(r => r.blob());
      const expected = await fetch(`/base/tests/render-test/${name}/result.html`).then(r => r.text());
      const div = document.createElement('div');
      document.body.appendChild(div);
      try {
        await docx.renderAsync(blob, div, null, {
          renderChanges: !comments, renderComments: comments,
          trackChangesMode: mode, commentsMode: mode
        });
        const projected = document.createElement('div');
        const article = div.querySelector('article').cloneNode(true);
        // Keep upstream content styles and markup, remove only new revision presentation.
        article.querySelectorAll('[data-tc-id]').forEach(el => {
          el.removeAttribute('data-tc-id');
          el.classList.remove('docx-tc-inserted', 'docx-tc-deleted');
          if (!el.className) el.removeAttribute('class');
        });
        const walker = document.createTreeWalker(article, NodeFilter.SHOW_COMMENT);
        const markers = [];
        while (walker.nextNode()) markers.push(walker.currentNode);
        markers.forEach(node => node.remove());
        projected.appendChild(article);
        const aside = document.createElement('aside');
        // Comments share one snapshot for both modes to enforce card parity.
        aside.dataset.mode = comments ? 'margin' : mode;
        const list = document.createElement('ol');
        div.querySelectorAll('.docx-tc-annotation').forEach(card => {
          const item = document.createElement('li');
          item.dataset.kind = [...card.classList].find(c => c.startsWith('docx-tc-annotation-')).replace('docx-tc-annotation-', '');
          for (const [tag, selector] of [['strong', 'strong'], ['time', 'time'], ['p', '.docx-tc-preview']]) {
            const el = document.createElement(tag);
            el.textContent = card.querySelector(selector).textContent;
            item.appendChild(el);
          }
          list.appendChild(item);
        });
        aside.appendChild(list);
        projected.appendChild(aside);
        expect(formatHTML(projected.innerHTML)).toBe(formatHTML(expected.trim()));
        expect(div.querySelector(`[data-annotation-panel="${mode}"]`)).not.toBeNull();
      } finally { div.remove(); }
    });
  }
});

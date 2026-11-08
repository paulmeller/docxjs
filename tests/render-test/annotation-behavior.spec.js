describe('Annotation behavior', function () {
  let div;
  beforeEach(() => { div = document.createElement('div'); document.body.appendChild(div); });
  afterEach(() => div.remove());

  async function load(name) {
    return fetch(`/base/tests/render-test/${name}/document.docx`).then(r => r.blob());
  }

  it('keeps comments independent from hidden revisions', async () => {
    await docx.renderAsync(await load('comments'), div, null, {
      renderComments: true, renderChanges: false, commentsMode: 'floating', trackChangesMode: 'margin'
    });
    expect(div.querySelectorAll('.docx-tc-annotation-comment').length).toBe(2);
    expect(div.querySelector('[data-annotation-panel="floating"]')).not.toBeNull();
    expect(div.querySelector('[data-annotation-panel="margin"]')).toBeNull();
  });

  it('rejects all changes with renderChanges off', async () => {
    const doc = await docx.parseAsync(await load('revision'));
    div.append(...await docx.renderDocument(doc, { renderChanges: false, trackChangeDecisions: docx.rejectAllChanges(doc) }));
    expect(div.querySelector('article').textContent).toBe('Original text  deleted');
    expect(div.querySelector('ins,del,.docx-tc-annotation')).toBeNull();
  });

  it('groups replacements but decides each member on its own', async () => {
    const doc = await docx.parseAsync(await load('revision'));
    const paragraph = doc.documentPart.body.children[0];
    // Make the two changes adjacent without the source document's intervening space run.
    paragraph.children = paragraph.children.filter((node, i) => i !== 2);
    const decisions = [];
    div.append(...await docx.renderDocument(doc, {
      renderChanges: true, trackChangesMode: 'margin',
      onTrackChangeDecision: (key, value, record) => decisions.push([key, value, record.type])
    }));
    expect(div.querySelectorAll('.docx-tc-annotation-replaced').length).toBe(1);
    const card = div.querySelector('.docx-tc-annotation-replaced');
    expect(JSON.parse(card.dataset.tcKeys)).toEqual(['body:0:ins#0', 'body:1:del#0']);
    expect(card.querySelectorAll('.docx-tc-member').length).toBe(2);
    // Only one member is decided; the other is untouched.
    card.querySelector('[data-tc-key="body:1:del#0"] button').click();
    expect(decisions).toEqual([['body:1:del#0', 'accept', 'del']]);
    card.querySelector('[data-tc-key="body:0:ins#0"] button:last-child').click();
    expect(decisions[1]).toEqual(['body:0:ins#0', 'reject', 'ins']);
  });

  it('suppresses buttons independently from revision markup', async () => {
    await docx.renderAsync(await load('revision'), div, null, {
      renderChanges: true, trackChangesMode: 'floating', showTrackChangeButtons: false,
      onTrackChangeDecision: () => { throw new Error('No button should be present'); }
    });
    expect(div.querySelector('ins')).not.toBeNull();
    expect(div.querySelector('.docx-tc-annotation button')).toBeNull();
  });

  it('positions cards without overlap under CSS zoom', async () => {
    div.style.zoom = '0.75';
    await docx.renderAsync(await load('revision'), div, null, { renderChanges: true, trackChangesMode: 'floating' });
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const cards = [...div.querySelectorAll('.docx-tc-annotation')];
    expect(cards.length).toBe(2);
    expect(cards[1].getBoundingClientRect().top).toBeGreaterThanOrEqual(cards[0].getBoundingClientRect().bottom);
    expect(cards[1].getBoundingClientRect().bottom).toBeLessThanOrEqual(div.querySelector('section').getBoundingClientRect().bottom + 1);
  });
});

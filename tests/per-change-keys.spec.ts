// @vitest-environment jsdom
import { describe, it, expect, beforeAll } from 'vitest';
import JSZip from 'jszip';
import { parseAsync, renderDocument, getTrackChanges, acceptAllChanges, selectTrackChange } from '../src/docx-preview';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const A = 'w:author="Eve" w:date="2026-01-01T00:00:00Z"';

async function docx(body: string, header = ''): Promise<Uint8Array> {
	const z = new JSZip();
	z.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>');
	z.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
	const hdrRel = header ? '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' : '';
	z.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hdrRel}</Relationships>`);
	if (header) z.file('word/header1.xml', `<?xml version="1.0"?><w:hdr ${W}>${header}</w:hdr>`);
	const sect = header ? '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/></w:sectPr>' : '';
	z.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}${sect}</w:body></w:document>`);
	return z.generateAsync({ type: 'uint8array' });
}

const ins = (id: number, t: string) => `<w:ins w:id="${id}" ${A}><w:r><w:t>${t}</w:t></w:r></w:ins>`;
const del = (id: number, t: string) => `<w:del w:id="${id}" ${A}><w:r><w:delText>${t}</w:delText></w:r></w:del>`;
const opts = { renderChanges: true, trackChangesMode: 'margin' as const, breakPages: false };

describe('per-change keys', () => {
	beforeAll(() => { (globalThis as any).requestAnimationFrame ??= (f: any) => setTimeout(f, 0); });

	it('records part, wId, type and occurrence for repeats and other parts, independent of decisions', async () => {
		const doc = await parseAsync(await docx(`<w:p>${ins(9, 'one')}${ins(9, 'two')}</w:p><w:p>${del(9, 'gone')}</w:p>`, `<w:p>${ins(9, 'head')}</w:p>`));
		const recs = getTrackChanges(doc);
		expect(recs.map(r => [r.part, r.wId, r.type, r.occurrence, r.text])).toEqual([
			['body', '9', 'ins', 0, 'one'], ['body', '9', 'ins', 1, 'two'], ['body', '9', 'del', 0, 'gone'], ['hdr1', '9', 'ins', 0, 'head']]);
		expect(recs[0]).toMatchObject({ author: 'Eve', date: '2026-01-01T00:00:00Z' });
		// Idempotent, and hiding a change does not renumber the others.
		const again = getTrackChanges(doc);
		expect(again).toEqual(recs);
	});

	it('stamps every rendered ins/del and keeps the existing attributes', async () => {
		const doc = await parseAsync(await docx(`<w:p>${ins(9, 'one')}${ins(9, 'two')}</w:p>`));
		const root = document.createElement('div');
		root.append(...await renderDocument(doc, opts));
		const els = Array.from(root.querySelectorAll('ins')) as HTMLElement[];
		expect(els.map(e => [e.dataset.tcId, e.dataset.tcPart, e.dataset.tcType, e.dataset.tcOccurrence, e.dataset.tcKey]))
			.toEqual([['9', 'body', 'ins', '0', 'body:9:ins#0'], ['9', 'body', 'ins', '1', 'body:9:ins#1']]);
		expect(els[0].getAttribute('datetime')).toBe('2026-01-01T00:00:00Z');
		// Inline mode stamps too.
		const inline = document.createElement('div');
		inline.append(...await renderDocument(doc, { ...opts, trackChangesMode: 'inline' }));
		expect((inline.querySelector('ins') as HTMLElement).dataset.tcKey).toBe('body:9:ins#0');
	});

	it('decides repeats of one w:id separately by key', async () => {
		const doc = await parseAsync(await docx(`<w:p>${ins(9, 'one')}${ins(9, 'two')}</w:p>`));
		const root = document.createElement('div');
		root.append(...await renderDocument(doc, { ...opts, trackChangeDecisions: { 'body:9:ins#1': 'reject' } }));
		expect(root.querySelector('article')!.textContent).toBe('one');
		expect(root.querySelectorAll('ins').length).toBe(1);
	});

	it('uses the caller key, and a null key is shown but not decidable', async () => {
		const doc = await parseAsync(await docx(`<w:p>${ins(1, 'a')}</w:p><w:p>${ins(2, 'b')}</w:p>`));
		const calls: any[] = [];
		const root = document.createElement('div');
		root.append(...await renderDocument(doc, { ...opts, trackChangesMode: 'floating',
			trackChangeKey: r => (r.wId === '1' ? 'server-key-1' : null),
			onTrackChangeDecision: (k, d, r) => calls.push([k, d, r.wId]),
			trackChangeDecisions: { 'server-key-1': 'accept', '2': 'reject', 'body:2:ins#0': 'reject' } }));
		// 'a' accepted (no markup), 'b' cannot be decided by any key, so stays pending.
		expect(root.querySelectorAll('ins').length).toBe(1);
		const b = root.querySelector('ins') as HTMLElement;
		expect(b.textContent).toBe('b');
		expect(b.dataset.tcUnmatched).toBe('');
		expect(b.dataset.tcKey).toBeUndefined();
		const card = root.querySelector('.docx-tc-annotation') as HTMLElement;
		expect(card.querySelector('.docx-tc-undecidable')!.textContent).toBe("Can't match this change");
		expect(card.querySelector('button')).toBeNull();
		expect(calls).toEqual([]);
	});

	it('groups del+ins visually but exposes and decides both members', async () => {
		const doc = await parseAsync(await docx(`<w:p>${del(1, 'old')}${ins(2, 'new')}</w:p>`));
		const calls: any[] = [];
		const selected: any[] = [];
		const root = document.createElement('div');
		document.body.append(root);
		root.append(...await renderDocument(doc, { ...opts, onTrackChangeDecision: (k, d) => calls.push([k, d]),
			onTrackChangeSelect: (k, r) => selected.push([k, r.type]) }));
		const card = root.querySelector('.docx-tc-annotation-replaced') as HTMLElement;
		expect(JSON.parse(card.dataset.tcKeys!)).toEqual(['body:1:del#0', 'body:2:ins#0']);
		const rows = Array.from(card.querySelectorAll('.docx-tc-member')) as HTMLElement[];
		expect(rows.map(r => r.dataset.tcKey)).toEqual(['body:1:del#0', 'body:2:ins#0']);
		(rows[1].querySelector('button') as HTMLElement).click();
		expect(calls).toEqual([['body:2:ins#0', 'accept']]);
		rows[0].click();
		expect(selected).toEqual([['body:1:del#0', 'del']]);
		selectTrackChange(root, 'body:2:ins#0');
		expect(root.querySelector('ins')!.classList.contains('docx-tc-selected')).toBe(true);
		expect(root.querySelector('del')!.classList.contains('docx-tc-selected')).toBe(false);
		expect(card.classList.contains('docx-tc-selected')).toBe(true);
		root.remove();
	});

	it('accept-all uses keys that cover repeats', async () => {
		const doc = await parseAsync(await docx(`<w:p>${ins(9, 'one')}${ins(9, 'two')}</w:p>`));
		expect(Object.keys(acceptAllChanges(doc))).toEqual(['body:9:ins#0', 'body:9:ins#1']);
	});
});

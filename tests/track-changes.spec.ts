import { describe, it, expect } from 'vitest';
import { changeVisibility } from '../src/track-changes';
import { acceptAllChanges, rejectAllChanges, getTrackChangeIds } from '../src/docx-preview';
import { DomType, WmlChange } from '../src/document/dom';
import { HtmlRenderer } from '../src/html-renderer';
import { defaultOptions } from '../src/docx-preview';

describe('revision decisions', () => {
	for (const type of [DomType.Inserted, DomType.Deleted, DomType.MoveFrom, DomType.MoveTo]) {
		it(`resolves ${type} with explicit decisions even when revisions are hidden`, () => {
			const change = { type } as WmlChange;
			const removed = type === DomType.Deleted || type === DomType.MoveFrom;
			for (const renderChanges of [true, false]) {
				expect(changeVisibility(change, 'accept', renderChanges)).toBe(removed ? 'hidden' : 'content');
				expect(changeVisibility(change, 'reject', renderChanges)).toBe(removed ? 'content' : 'hidden');
			}
			expect(changeVisibility(change, 'pending', true)).toBe('revision');
			expect(changeVisibility(change, undefined, false)).toBe(removed ? 'hidden' : 'content');
		});
	}

	it('collects nested, formatting, zero, and header keys (a w:id repeated in another part stays distinct)', () => {
		const insert = { type: DomType.Inserted, id: '0', children: [] };
		const document = { documentPart: { body: { type: DomType.Document, children: [insert,
			{ type: DomType.Run, formatChange: { type: DomType.FormatChange, id: 'f' } }] } },
			parts: [{ rootElement: { type: DomType.Header, children: [insert, { type: DomType.MoveFrom, id: 'm' }] } }] };
		// Keys are part:wId:type#occurrence, never the bare w:id.
		const keys = ['body:0:ins#0', 'body:f:rPrChange#0', 'part0:0:ins#0', 'part0:m:moveFrom#0'];
		expect(getTrackChangeIds(document)).toEqual(keys);
		expect(acceptAllChanges(document)).toEqual(Object.fromEntries(keys.map(k => [k, 'accept'])));
		expect(rejectAllChanges(document, r => `x-${r.wId}`)).toEqual({ 'x-0': 'reject', 'x-f': 'reject', 'x-m': 'reject' });
	});

	it('retains deleted text on reject-all with renderChanges off', () => {
		const renderer = new HtmlRenderer();
		renderer.options = { ...defaultOptions, trackChangeDecisions: { 'body:d:del#0': 'reject' } };
		renderer.h = value => value as any;
		const change = { type: DomType.Deleted, id: 'd', author: '', date: '', part: 'body', occurrence: 0,
			children: [{ type: DomType.DeletedText, text: 'restored' } as any] } as WmlChange;
		const result = renderer.renderDeleted(change);
		expect(result).toEqual(['restored']);
		expect(renderer.includeDeletedText).toBe(0);
	});
});

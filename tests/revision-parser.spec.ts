import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { DocumentParser } from '../src/document-parser';
import { DomType } from '../src/document/dom';

// Minimal XML element adapter: these tests exercise the real parser without a browser.
function element(localName: string, attrs: Record<string, string> = {}, children: Element[] = [], textContent = ''): Element {
	const attributes = Object.entries(attrs).map(([localName, value]) => ({ localName, value }));
	return { localName, namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', nodeType: 1, textContent,
		attributes: { length: attributes.length, item: (i: number) => attributes[i] },
		childNodes: { length: children.length, item: (i: number) => children[i] } } as any;
}

describe('extended upstream revision parser', () => {
	beforeEach(() => vi.stubGlobal('Node', { ELEMENT_NODE: 1 }));
	afterEach(() => vi.unstubAllGlobals());
	const parser = new DocumentParser({ ignoreWidth: false, debug: false });

	it('parses moves using generic change metadata and preserves range marker IDs', () => {
		const result = parser.parseParagraph(element('p', {}, [
			element('moveFromRangeStart', { id: 'r', name: 'moved' }),
			element('moveFrom', { id: '2', author: 'Author', date: '2026-01-01' }, [element('r', {}, [element('delText', {}, [], 'old')])]),
			element('moveFromRangeEnd', { id: 'r' }),
			element('moveTo', { id: '3', author: 'Author' }, [element('r', {}, [element('t', {}, [], 'new')])])
		]));
		expect(result.children[0]).toMatchObject({ type: DomType.MoveFromRangeStart, props: { id: 'r', name: 'moved' } });
		expect(result.children[1]).toMatchObject({ type: DomType.MoveFrom, changeType: 'moveFrom', id: '2', author: 'Author', date: '2026-01-01' });
		expect(result.children[1].children[0].children[0]).toMatchObject({ type: DomType.DeletedText, text: 'old' });
		expect(result.children[3]).toMatchObject({ type: DomType.MoveTo, changeType: 'moveTo', id: '3' });
	});

	it('stores run and paragraph formatting revisions on the existing model', () => {
		const change = element('rPrChange', { id: '4', author: 'Reviewer' }, [element('rPr', {}, [element('b', { val: '0' })])]);
		const run = parser.parseRun(element('r', {}, [element('rPr', {}, [element('b'), change]), element('t', {}, [], 'text')]));
		expect(run.cssStyle['font-weight']).toBe('bold');
		expect(run.formatChange).toMatchObject({ type: DomType.FormatChange, id: '4', previousStyle: { 'font-weight': 'normal' } });
		const paragraph = parser.parseParagraph(element('p', {}, [element('pPr', {}, [
			element('pPrChange', { id: '5' }, [element('pPr', {}, [element('jc', { val: 'right' })])])
		])]));
		expect(paragraph.formatChange).toMatchObject({ id: '5', previousStyle: { 'text-align': 'right' } });
	});

	it('preserves upstream nested smart tags', () => {
		const result = parser.parseParagraph(element('p', {}, [element('smartTag', {}, [element('smartTag', {}, [element('r', {}, [element('t', {}, [], 'nested')])])])]));
		expect(result.children[0].children[0].type).toBe(DomType.SmartTag);
		expect(result.children[0].children[0].children[0].children[0]).toMatchObject({ text: 'nested' });
	});
});

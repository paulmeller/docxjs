import { describe, it, expect } from 'vitest';
import { HtmlRenderer } from '../src/html-renderer';
import { defaultOptions } from '../src/docx-preview';

describe('section page boundaries', () => {
	function section(width: string, height: string, orientation: string = 'portrait', pageBreak = false) {
		return { sectProps: { pageSize: { width, height, orientation } }, elements: [], pageBreak } as any;
	}
	it('starts a page before an incoming orientation change', () => {
		const renderer = new HtmlRenderer();
		renderer.options = { ...defaultOptions, ignoreLastRenderedPageBreak: false };
		const a = section('600pt', '800pt');
		const b = section('800pt', '600pt', 'landscape');
		const c = section('800pt', '600pt', 'landscape');
		expect(renderer.groupByPageBreaks([a, b, c])).toEqual([[a], [b, c]]);
	});
	it('starts a page on size changes even with the same orientation', () => {
		const renderer = new HtmlRenderer();
		renderer.options = { ...defaultOptions, ignoreLastRenderedPageBreak: false };
		const a = section('600pt', '800pt');
		const b = section('600pt', '900pt');
		expect(renderer.groupByPageBreaks([a, b])).toEqual([[a], [b]]);
	});
	it('preserves upstream explicit page breaks and default per-section grouping', () => {
		const renderer = new HtmlRenderer();
		const a = section('600pt', '800pt', 'portrait', true);
		const b = section('600pt', '800pt');
		renderer.options = { ...defaultOptions, ignoreLastRenderedPageBreak: false };
		expect(renderer.groupByPageBreaks([a, b])).toEqual([[a], [b]]);
		renderer.options = { ...defaultOptions };
		expect(renderer.groupByPageBreaks([{ ...a, pageBreak: false }, b])).toHaveLength(2);
	});
});

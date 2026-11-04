import { describe, it, expect, vi, afterEach } from 'vitest';
import { HtmlRenderer } from '../src/html-renderer';
import { VmlElement } from '../src/vml/vml';

function fakeSvg() {
	return { attrs: {} as Record<string, string>, children: [],
		setAttribute(key: string, value: string) { this.attrs[key] = value; },
		appendChild(child: any) { this.children.push(child); } };
}

describe('VML text boxes', () => {
	afterEach(() => vi.unstubAllGlobals());
	it('clears text-box rect fills, preserves ordinary shapes, and clips the SVG', () => {
		vi.stubGlobal('requestAnimationFrame', () => 0);
		const renderer = new HtmlRenderer();
		renderer.h = () => fakeSvg() as any;
		const rect = new VmlElement();
		rect.tagName = 'rect';
		rect.attrs = { fill: '#ffffff' };
		const text = new VmlElement();
		text.tagName = 'foreignObject';
		rect.children.push(text);
		const result = renderer.renderVmlElement(rect) as any;
		expect(result.attrs.overflow).toBe('hidden');
		expect(result.children[0].attrs.fill).toBe('none');
		rect.children = [];
		expect((renderer.renderVmlChildElement(rect) as any).attrs.fill).toBe('#ffffff');
	});
});

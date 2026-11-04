import { describe, it, expect } from 'vitest';
import { emfToSvgDataUrl } from '../src/metafile';

describe('renderer EMF conversion', () => {
	it('converts filled path records to an SVG data URL', () => {
		const records = [
			[59, 8], [27, 16, 0, 0],
			[89, 36, 0, 0, 100, 100, 2, 100, (100 << 16) | 100],
			[61, 8], [60, 8], [62, 8], [14, 8]
		];
		const buffer = new ArrayBuffer(88 + records.reduce((n, r) => n + r[1], 0));
		const view = new DataView(buffer);
		[1, 88, 0, 0, 100, 100].forEach((n, i) => view.setUint32(i * 4, n, true));
		let offset = 88;
		for (const record of records) {
			record.forEach((n, i) => view.setUint32(offset + i * 4, n, true));
			offset += record[1];
		}
		const url = emfToSvgDataUrl(buffer);
		expect(url).toMatch(/^data:image\/svg\+xml/);
		const svg = decodeURIComponent(url.split(',')[1]);
		expect(svg).toContain('viewBox="0 0 100 100"');
		expect(svg).toContain('M0 0L100 0L100 100Z');
	});

	it('returns null for truncated, malformed, unsupported, or non-EMF data', () => {
		for (const bytes of [[], [1], [0, 0, 0, 0], new Array(88).fill(0)]) {
			expect(emfToSvgDataUrl(new Uint8Array(bytes).buffer)).toBeNull();
		}
	});
});

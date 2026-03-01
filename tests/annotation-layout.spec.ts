import { describe, it, expect } from 'vitest';
import { computePositions, fitWithinBounds, ANNOTATION_GAP } from '../src/annotation-layout';

describe('computePositions', () => {
	it('places cards at target positions when no overlap', () => {
		const heights = [40, 40, 40];
		const targets = [0, 100, 200];
		const result = computePositions(heights, targets, ANNOTATION_GAP);
		expect(result.tops).toEqual([0, 100, 200]);
		expect(result.lastBottom).toBe(240);
	});

	it('pushes cards down to avoid overlap', () => {
		const heights = [50, 50, 50];
		// Second target overlaps first card (0+50+8 = 58 > 30)
		const targets = [0, 30, 60];
		const result = computePositions(heights, targets, ANNOTATION_GAP);
		expect(result.tops[0]).toBe(0);
		expect(result.tops[1]).toBe(58); // pushed down: 0+50+8
		expect(result.tops[2]).toBe(116); // pushed down: 58+50+8
		expect(result.lastBottom).toBe(166);
	});

	it('calculates lastBottom correctly for single card', () => {
		const result = computePositions([80], [10], ANNOTATION_GAP);
		expect(result.tops).toEqual([10]);
		expect(result.lastBottom).toBe(90);
	});

	it('uses zero targets for tight stacking', () => {
		const heights = [40, 40, 40];
		const targets = [0, 0, 0];
		const result = computePositions(heights, targets, ANNOTATION_GAP);
		expect(result.tops).toEqual([0, 48, 96]);
		expect(result.lastBottom).toBe(136);
	});
});

describe('fitWithinBounds', () => {
	it('returns targets unchanged when cards fit within bounds', () => {
		const heights = [40, 40];
		const targets = [0, 60];
		const result = fitWithinBounds(heights, targets, ANNOTATION_GAP, 200);
		expect(result).toEqual([0, 60]);
	});

	it('shifts cards up when they overflow the bottom', () => {
		const heights = [40, 40];
		// Card 2 at 160, bottom = 200, but available = 180 → overflow by 20
		const targets = [100, 160];
		const result = fitWithinBounds(heights, targets, ANNOTATION_GAP, 180);
		expect(result).toEqual([80, 140]);
	});

	it('clamps to 0 when shift exceeds first target', () => {
		const heights = [40, 40];
		// Cards overflow by 100, but first target is only 10
		const targets = [10, 60];
		// Trial: tops=[10, 60], lastBottom=100. avail=50 → shift=50
		// Shifted: [max(0, 10-50), max(0, 60-50)] = [0, 10]
		const result = fitWithinBounds(heights, targets, ANNOTATION_GAP, 50);
		expect(result).toEqual([0, 10]);
	});

	it('handles single card at bottom of page', () => {
		const heights = [80];
		const targets = [500];
		// lastBottom = 580, avail = 550 → shift = 30
		const result = fitWithinBounds(heights, targets, ANNOTATION_GAP, 550);
		expect(result).toEqual([470]);
	});

	it('shifts ≤3 content-aligned cards near bottom to fit', () => {
		const heights = [60, 60, 60];
		const targets = [300, 400, 500];
		// Trial: tops=[300, 400, 500], lastBottom=560
		// avail=520 → shift=40
		const result = fitWithinBounds(heights, targets, ANNOTATION_GAP, 520);
		expect(result).toEqual([260, 360, 460]);
	});

	it('still overflows when total card height exceeds available (scrollable)', () => {
		const heights = [200, 200, 200];
		const targets = [0, 0, 0];
		// Tight: tops=[0, 208, 416], lastBottom=616
		// avail=400 → shift=216, shifted=[0, 0, 0] (all clamped)
		// Cards still won't fit — scrollable fallback needed
		const result = fitWithinBounds(heights, targets, ANNOTATION_GAP, 400);
		expect(result).toEqual([0, 0, 0]);
	});
});

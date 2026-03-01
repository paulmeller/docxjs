// Annotation layout constants
export const ANNOTATION_GAP = 8;
export const COMPRESSED_GAP = 2;
export const EXPAND_BUTTON_HEIGHT = 32;

/**
 * Pure math — compute card top positions without touching the DOM.
 * Each card is placed at its target position or pushed down to avoid
 * overlapping the previous card (whichever is lower).
 */
export function computePositions(
	cardHeights: number[], targets: number[], gap: number
): { tops: number[], lastBottom: number } {
	let nextTop = 0;
	const tops: number[] = [];
	for (let i = 0; i < cardHeights.length; i++) {
		const top = Math.max(targets[i], nextTop);
		tops.push(top);
		nextTop = top + cardHeights[i] + gap;
	}
	const last = cardHeights.length - 1;
	return { tops, lastBottom: tops[last] + cardHeights[last] };
}

/**
 * Pull overflowing cards up from the bottom via a backward pass.
 * Top cards that already fit keep their content-aligned positions.
 */
export function fitWithinBounds(
	cardHeights: number[], targets: number[], gap: number, availableHeight: number
): number[] {
	const trial = computePositions(cardHeights, targets, gap);
	if (trial.lastBottom <= availableHeight) return targets;

	const adjusted = [...trial.tops];
	let maxBottom = availableHeight;
	for (let i = adjusted.length - 1; i >= 0; i--) {
		if (adjusted[i] + cardHeights[i] > maxBottom) {
			adjusted[i] = Math.max(0, maxBottom - cardHeights[i]);
		}
		maxBottom = adjusted[i] - gap;
	}
	return adjusted;
}

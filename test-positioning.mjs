/**
 * Tests for hybrid annotation positioning — density-gated with trial layout.
 *
 * Algorithm (3 rules):
 *   Rule 1: ≤3 cards → always content-aligned
 *   Rule 2: density > 0.6 → tight stacking
 *   Rule 3: else → try content-aligned; if overflows, fall back to tight
 *
 * Where density = tightStackedHeight / availableHeight
 */

const ANNOTATION_GAP = 8;

// Pure-math simulation of computePositions (no DOM)
function computePositions(cardHeights, targets, gap) {
  let nextTop = 0;
  const tops = [];
  for (let i = 0; i < cardHeights.length; i++) {
    const top = Math.max(targets[i], nextTop);
    tops.push(top);
    nextTop = top + cardHeights[i] + gap;
  }
  const last = cardHeights.length - 1;
  return { tops, lastBottom: tops[last] + cardHeights[last] };
}

// Simulate the density-gated hybrid positioning algorithm
function hybridPosition(cardHeights, targets, availableHeight, gap = ANNOTATION_GAP) {
  const zeroTargets = targets.map(() => 0);
  let useTargets;
  let mode;

  if (cardHeights.length <= 3) {
    // Rule 1: few cards — always content-aligned
    useTargets = targets;
    mode = 'content-aligned';
  } else {
    const tight = computePositions(cardHeights, zeroTargets, gap);
    const density = tight.lastBottom / availableHeight;

    if (density > 0.6) {
      // Rule 2: dense — tight stack
      useTargets = zeroTargets;
      mode = 'tight-stacked';
    } else {
      // Rule 3: try content-aligned
      const aligned = computePositions(cardHeights, targets, gap);
      if (aligned.lastBottom > availableHeight) {
        useTargets = zeroTargets;
        mode = 'tight-stacked';
      } else {
        useTargets = targets;
        mode = 'content-aligned';
      }
    }
  }

  const result = computePositions(cardHeights, useTargets, gap);
  const scrollable = result.lastBottom > availableHeight;
  if (scrollable) mode = 'scrollable';

  return { ...result, mode };
}

// --- Test runner ---
let passed = 0;
let failed = 0;

function assert(condition, message) {
  if (condition) {
    passed++;
    console.log(`  ✓ ${message}`);
  } else {
    failed++;
    console.error(`  ✗ ${message}`);
  }
}

function assertEq(actual, expected, message) {
  if (actual === expected) {
    passed++;
    console.log(`  ✓ ${message}`);
  } else {
    failed++;
    console.error(`  ✗ ${message} — expected ${expected}, got ${actual}`);
  }
}

// =============================================
// Rule 1: ≤3 cards → always content-aligned
// =============================================

console.log('\n=== Rule 1: Single card — always content-aligned ===');
{
  const result = hybridPosition([60], [150], 800);
  assertEq(result.mode, 'content-aligned', 'single card is content-aligned');
  assertEq(result.tops[0], 150, 'card at target 150');
}

console.log('\n=== Rule 1: 3 scattered cards — always content-aligned ===');
{
  const cardHeights = [60, 60, 60];
  const targets = [50, 200, 400];
  const result = hybridPosition(cardHeights, targets, 800);
  assertEq(result.mode, 'content-aligned', '3 cards always content-aligned');
  assertEq(result.tops[0], 50, 'card 0 at target 50');
  assertEq(result.tops[1], 200, 'card 1 at target 200');
  assertEq(result.tops[2], 400, 'card 2 at target 400');
  assertEq(result.lastBottom, 460, 'lastBottom = 400 + 60');
}

console.log('\n=== Rule 1: 3 cards even if they overflow — still content-aligned (scrollable) ===');
{
  const cardHeights = [60, 60, 60];
  const targets = [50, 200, 400];
  const availableHeight = 200; // way too small
  const result = hybridPosition(cardHeights, targets, availableHeight);
  assertEq(result.mode, 'scrollable', '3 cards overflow → scrollable but still aligned');
  assertEq(result.tops[0], 50, 'card 0 at target 50');
  assertEq(result.tops[2], 400, 'card 2 at target 400');
}

console.log('\n=== Rule 1: 3 cards with collision — pushed down, still content-aligned ===');
{
  const cardHeights = [60, 60, 60];
  const targets = [50, 80, 110]; // too close
  const result = hybridPosition(cardHeights, targets, 800);
  assertEq(result.mode, 'content-aligned', 'still content-aligned');
  assertEq(result.tops[0], 50, 'card 0 at 50');
  assertEq(result.tops[1], 118, 'card 1 pushed to 118 (50+60+8)');
  assertEq(result.tops[2], 186, 'card 2 pushed to 186 (118+60+8)');
}

// =============================================
// Rule 2: density > 0.6 → tight stacking
// =============================================

console.log('\n=== Rule 2: Page 1 redlined — 11 cards, density 66% → tight ===');
{
  // Real-world: 11 cards × ~93px avg + 10 gaps = ~1101px in 1123px panel
  const cardHeights = Array(11).fill(93);
  const targets = Array.from({ length: 11 }, (_, i) => i * 90);
  const availableHeight = 1123;

  const tight = computePositions(cardHeights, targets.map(() => 0), ANNOTATION_GAP);
  const density = tight.lastBottom / availableHeight;
  assert(density > 0.6, `density ${(density * 100).toFixed(0)}% > 60%`);

  const result = hybridPosition(cardHeights, targets, availableHeight);
  assertEq(result.mode, 'tight-stacked', '11 cards → Rule 2 tight stack');
}

console.log('\n=== Rule 2: 8 cards, density 62% → tight ===');
{
  const cardHeights = Array(8).fill(80);
  const targets = Array.from({ length: 8 }, (_, i) => i * 120);
  const availableHeight = 1056;

  const tight = computePositions(cardHeights, targets.map(() => 0), ANNOTATION_GAP);
  const density = tight.lastBottom / availableHeight;
  assert(density > 0.6, `density ${(density * 100).toFixed(0)}% > 60%`);

  const result = hybridPosition(cardHeights, targets, availableHeight);
  assertEq(result.mode, 'tight-stacked', '8 dense cards → tight stack');
}

// =============================================
// Rule 3: try content-aligned; if overflows, fall back
// =============================================

console.log('\n=== Rule 3: Page 2 redlined — 7 cards, density 42% → try aligned → fits ===');
{
  // Real-world: 7 cards, targets well-spaced, density ~42%
  const cardHeights = Array(7).fill(60);
  const targets = [50, 200, 350, 500, 650, 800, 950];
  const availableHeight = 1123;

  const tight = computePositions(cardHeights, targets.map(() => 0), ANNOTATION_GAP);
  const density = tight.lastBottom / availableHeight;
  assert(density <= 0.6, `density ${(density * 100).toFixed(0)}% ≤ 60%`);

  const result = hybridPosition(cardHeights, targets, availableHeight);
  assertEq(result.mode, 'content-aligned', '7 cards, low density → content-aligned');
  assertEq(result.tops[0], 50, 'card 0 aligned to target');
  assertEq(result.tops[6], 950, 'last card aligned to target');
}

console.log('\n=== Rule 3: 6 clustered at bottom — try aligned → overflows → tight ===');
{
  // 6 cards all targeted near the bottom of the page
  const cardHeights = Array(6).fill(60);
  const targets = [700, 750, 800, 850, 900, 950];
  const availableHeight = 1056;

  const tight = computePositions(cardHeights, targets.map(() => 0), ANNOTATION_GAP);
  const density = tight.lastBottom / availableHeight;
  assert(density <= 0.6, `density ${(density * 100).toFixed(0)}% ≤ 60%`);

  // But content-aligned would overflow (cards pushed down from 700+)
  const aligned = computePositions(cardHeights, targets, ANNOTATION_GAP);
  assert(aligned.lastBottom > availableHeight,
    `content-aligned overflows (${aligned.lastBottom} > ${availableHeight})`);

  const result = hybridPosition(cardHeights, targets, availableHeight);
  assertEq(result.mode, 'tight-stacked', 'aligned overflows → falls back to tight');
}

console.log('\n=== Rule 3: 4 well-spaced cards — try aligned → fits ===');
{
  const cardHeights = [60, 60, 60, 60];
  const targets = [50, 200, 400, 600];
  const availableHeight = 800;

  const result = hybridPosition(cardHeights, targets, availableHeight);
  assertEq(result.mode, 'content-aligned', '4 well-spaced → content-aligned');
  assertEq(result.tops[0], 50, 'card 0 at target');
  assertEq(result.tops[3], 600, 'card 3 at target');
  assertEq(result.lastBottom, 660, 'lastBottom = 600 + 60');
}

console.log('\n=== Rule 3: 4 cards — try aligned → overflows → tight ===');
{
  const cardHeights = [60, 60, 60, 60];
  const targets = [100, 300, 500, 700]; // lastBottom = 760
  const availableHeight = 500;

  const result = hybridPosition(cardHeights, targets, availableHeight);
  assertEq(result.mode, 'tight-stacked', 'aligned overflows → tight');
  assertEq(result.tops[0], 0, 'card 0 at 0');
  assertEq(result.tops[1], 68, 'card 1 at 68');
  assertEq(result.lastBottom, 264, 'tight lastBottom = 264');
  assert(result.lastBottom <= availableHeight, 'tight fits');
}

// =============================================
// Edge cases
// =============================================

console.log('\n=== Edge: Both modes overflow → scrollable ===');
{
  const cardHeights = Array(10).fill(60);
  const targets = cardHeights.map((_, i) => i * 100);
  const availableHeight = 300;

  const result = hybridPosition(cardHeights, targets, availableHeight);
  assertEq(result.mode, 'scrollable', 'both overflow → scrollable');
  assertEq(result.lastBottom, 672, 'tight lastBottom = 672');
}

console.log('\n=== Edge: Content-aligned BARELY fits (low density) ===');
{
  // Need density ≤ 0.6 for Rule 3 to even try alignment
  // 4 cards × 60 + 3 gaps = 264; availableHeight = 500 → density 52.8%
  // Aligned lastBottom = 300 + 60 = 360 ≤ 500 → fits
  const cardHeights = [60, 60, 60, 60];
  const targets = [0, 100, 200, 300];
  const availableHeight = 500;

  const result = hybridPosition(cardHeights, targets, availableHeight);
  assertEq(result.mode, 'content-aligned', 'low density + aligned fits → content-aligned');
  assertEq(result.lastBottom, 360, 'lastBottom = 300 + 60');
}

console.log('\n=== Edge: Density exactly 0.6 — should NOT trigger Rule 2 ===');
{
  // density = tight.lastBottom / availableHeight = 0.6
  // tight: 4 cards × 60 + 3 gaps × 8 = 264; available = 264/0.6 = 440
  const cardHeights = [60, 60, 60, 60];
  const targets = [0, 100, 200, 300];
  const availableHeight = 440;

  const tight = computePositions(cardHeights, targets.map(() => 0), ANNOTATION_GAP);
  const density = tight.lastBottom / availableHeight;
  assertEq(density, 0.6, `density exactly 0.6`);

  const result = hybridPosition(cardHeights, targets, availableHeight);
  // density is not > 0.6, so Rule 3 applies (try aligned)
  assertEq(result.mode, 'content-aligned', 'density=0.6 → Rule 3, not Rule 2');
}

console.log('\n=== Edge: Zero targets (all content at top) ===');
{
  const cardHeights = [60, 60, 60];
  const targets = [0, 0, 0];
  const result = hybridPosition(cardHeights, targets, 800);
  assertEq(result.mode, 'content-aligned', 'zero targets → Rule 1 (≤3 cards)');
  assertEq(result.tops[0], 0, 'card 0 at 0');
  assertEq(result.tops[1], 68, 'card 1 at 68');
  assertEq(result.tops[2], 136, 'card 2 at 136');
}

console.log('\n=== Edge: Varying card heights ===');
{
  const cardHeights = [50, 70, 40, 60];
  const targets = [0, 0, 0, 0];
  const result = computePositions(cardHeights, targets, ANNOTATION_GAP);
  // 0+50+8=58, 58+70+8=136, 136+40+8=184, 184+60=244
  assertEq(result.lastBottom, 244, 'tight stacked = sum of heights + gaps');
  const expectedMin = cardHeights.reduce((a, b) => a + b, 0) +
    ANNOTATION_GAP * (cardHeights.length - 1);
  assertEq(result.lastBottom, expectedMin, 'minimal possible height');
}

// --- Summary ---
console.log(`\n${'='.repeat(50)}`);
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  console.log('\nFAILED tests indicate bugs to fix.');
  process.exit(1);
} else {
  console.log('\nAll tests passed.');
}

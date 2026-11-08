import { Options, TrackChangesMode } from './docx-preview';
import { WmlChange } from './document/dom';
import { TrackChangeRecord } from './track-changes';
import { h } from './html';
import { ANNOTATION_GAP, computePositions, fitWithinBounds } from './annotation-layout';

export interface Annotation {
	change: WmlChange;
	kind: string;
	anchor: HTMLElement;
	text: string;
	mode: TrackChangesMode;
	/** Caller key of the change; null means it could not be matched and cannot be decided. */
	key?: string | null;
	record?: TrackChangeRecord;
	/** A grouped ("replaced") card: both member changes, in document order, each decided on its own. */
	members?: Annotation[];
}

export interface DecisionTarget { key: string | null; record: TrackChangeRecord }

/** A single layout path for both panels, using layout pixels at every zoom level. */
export class AnnotationRenderer {
	private cards = new Map<HTMLElement, Annotation>();
	private expanded = new WeakSet<HTMLElement>();
	private observer?: ResizeObserver;
	private frame = 0;

	constructor(private options: Options, private build: typeof h) {}

	createCard(annotation: Annotation): HTMLElement {
		const c = this.options.className;
		const change = annotation.change;
		const initials = (change.author || '?').trim().split(/\s+/).slice(0, 2).map(x => x[0]).join('').toUpperCase();
		const labels = { inserted: 'Inserted', deleted: 'Deleted', moveFrom: 'Moved from', moveTo: 'Moved to', formatChange: 'Formatting', comment: 'Comment', replaced: 'Replaced' };
		const card = this.build({ tagName: 'div', className: `${c}-tc-annotation ${c}-tc-annotation-${annotation.kind}`, tabIndex: 0,
			children: [
				{ tagName: 'span', className: `${c}-tc-avatar`, children: [initials] },
				{ tagName: 'div', className: `${c}-tc-body`, children: [
					{ tagName: 'strong', children: [change.author || 'Unknown author'] },
					{ tagName: 'time', dateTime: change.date, children: [change.date || ''] },
					{ tagName: 'div', className: `${c}-tc-preview`, children: [`${labels[annotation.kind] || annotation.kind}: ${annotation.text}`] }
				] }
			] }) as HTMLElement;
		const body = card.lastElementChild;
		if (annotation.members) {
			const list = this.build({ tagName: 'div', className: `${c}-tc-members` }) as HTMLElement;
			for (const m of annotation.members) {
				const row = this.build({ tagName: 'div', className: `${c}-tc-member ${c}-tc-member-${m.kind}`, children: [
					{ tagName: 'span', className: `${c}-tc-member-text`, children: [`${labels[m.kind] || m.kind}: ${m.text}`] }] }) as HTMLElement;
				this.decorate(row, m);
				row.onclick = event => { event.stopPropagation(); this.activate(card, m, true); };
				list.appendChild(row);
			}
			body.appendChild(list);
			card.dataset.tcKeys = JSON.stringify(annotation.members.map(m => m.key));
		} else if (annotation.kind !== 'comment') {
			this.decorate(card, annotation, body);
		}
		card.dataset.tcId = change.id;
		card.setAttribute('aria-label', `${labels[annotation.kind]} by ${change.author}: ${annotation.text}`);
		const select = () => this.activate(card, annotation.members?.[0] ?? annotation);
		card.onclick = select;
		card.onkeydown = event => {
			if (event.target === card && (event.key === 'Enter' || event.key === ' ')) {
				event.preventDefault();
				select();
			}
		};
		this.cards.set(card, annotation);
		return card;
	}

	/** Keys, per-change buttons or the "can't match" note for one change (a card or a member row). */
	private decorate(el: HTMLElement, a: Annotation, host: Element = el) {
		const c = this.options.className;
		if (a.key != null) el.dataset.tcKey = a.key;
		else {
			el.dataset.tcUnmatched = '';
			host.appendChild(this.build({ tagName: 'div', className: `${c}-tc-undecidable`, children: [this.options.trackChangeUndecidableText] }));
		}
		const buttons = changeButtons(this.options, this.build, [{ key: a.key, record: a.record }]);
		if (buttons) host.appendChild(buttons);
	}

	private activate(card: HTMLElement, target: Annotation, forceOn = false) {
		const c = this.options.className;
		const active = forceOn ? (card.classList.add(`${c}-tc-active`), true) : card.classList.toggle(`${c}-tc-active`);
		card.setAttribute('aria-expanded', String(active));
		if (active) target.anchor.scrollIntoView?.({ behavior: 'smooth', block: 'center' });
		if (target.record) this.options.onTrackChangeSelect?.(active ? target.key ?? null : null, target.record);
		this.schedule();
	}

	mount(pages: HTMLElement[], annotations: Annotation[]) {
		const grouped: Annotation[] = [];
		const used = new Set<Annotation>();
		for (const a of annotations) {
			if (used.has(a)) continue;
			const partner = annotations.find(b => b !== a && !used.has(b) && a.mode === b.mode && a.change.author === b.change.author
				&& ((a.kind === 'inserted' && b.kind === 'deleted') || (a.kind === 'deleted' && b.kind === 'inserted'))
				&& (this.adjacent(a.anchor, b.anchor) || this.adjacent(b.anchor, a.anchor)));
			if (partner) {
				used.add(partner);
				const ins = a.kind === 'inserted' ? a : partner;
				const del = a.kind === 'deleted' ? a : partner;
				// Anchor and card identity follow the first member in the document; both stay decidable.
				const first = this.adjacent(a.anchor, partner.anchor) ? a : partner;
				grouped.push({ ...first, kind: 'replaced', members: first === a ? [a, partner] : [partner, a], text: `${del.text} → ${ins.text}` });
			} else grouped.push(a);
			used.add(a);
		}
		for (const annotation of grouped) {
			const page = pages.find(p => p.contains(annotation.anchor));
			if (page) this.panel(page, annotation.mode).appendChild(this.createCard(annotation));
		}
		let connected = false;
		const refresh = () => {
			if (!pages.some(p => p.isConnected)) {
				if (connected) { this.observer?.disconnect(); this.cards.clear(); }
				return;
			}
			connected = true;
			if (this.options.breakPages && !this.options.ignoreHeight) this.splitPages(pages);
			// Locate anchors again after splitting. Never infer page ownership from change IDs.
			for (const [card, annotation] of this.cards) {
				const page = pages.find(p => p.contains(annotation.anchor));
				if (page) {
					const panel = this.panel(page, annotation.mode);
					if (card.parentElement !== panel) panel.appendChild(card);
				}
			}
			for (const page of pages) {
				for (const panel of Array.from(page.querySelectorAll<HTMLElement>(':scope > [data-annotation-panel]'))) this.layout(panel, page);
			}
		};
		this.refresh = refresh;
		if (typeof ResizeObserver !== 'undefined') {
			this.observer = new ResizeObserver(() => this.schedule());
			pages.forEach(p => this.observer.observe(p));
			this.cards.forEach((_, card) => this.observer.observe(card));
		}
		this.schedule();
		globalThis.document?.fonts?.ready.then(() => this.schedule());
	}

	private refresh = () => {};

	private schedule() {
		if (!this.frame) this.frame = requestAnimationFrame(() => { this.frame = 0; this.refresh(); });
	}

	private adjacent(a: HTMLElement, b: HTMLElement) {
		let next = a.nextSibling;
		while (next?.nodeType === 3 && !next.textContent.trim()) next = next.nextSibling;
		return next === b;
	}

	private panel(page: HTMLElement, mode: TrackChangesMode): HTMLElement {
		let panel = page.querySelector<HTMLElement>(`:scope > [data-annotation-panel="${mode}"]`);
		if (panel) return panel;
		panel = this.build({ tagName: 'aside', className: `${this.options.className}-track-changes-${mode}` }) as HTMLElement;
		panel.dataset.annotationPanel = mode;
		panel.setAttribute('aria-label', 'Annotations');
		const cs = getComputedStyle(page);
		panel.style.top = page.style.paddingTop || cs.paddingTop;
		panel.style.width = this.options.trackChangesMarginWidth;
		// Margin reserves room inside an expanded sheet; floating lives outside it.
		if (mode === 'margin') {
			const right = page.style.paddingRight || cs.paddingRight;
			if (page.style.width || cs.width) page.style.width = `calc(${page.style.width || cs.width} + ${this.options.trackChangesMarginWidth} + 16px)`;
			page.style.paddingRight = `calc(${right} + ${this.options.trackChangesMarginWidth} + 16px)`;
			panel.style.right = right;
		} else {
			panel.style.left = 'calc(100% + 16px)';
		}
		page.style.overflow = 'visible';
		page.appendChild(panel);
		return panel;
	}

	private layout(panel: HTMLElement, page: HTMLElement) {
		const cards = Array.from(panel.children).filter(el => this.cards.has(el as HTMLElement)) as HTMLElement[];
		if (!cards.length) return;
		panel.querySelector('[data-expand-annotations]')?.remove();
		panel.style.maxHeight = '';
		panel.style.overflowY = '';
		cards.forEach(card => { card.hidden = false; card.style.position = 'absolute'; card.style.top = '0px'; });
		const zoom = page.offsetWidth ? page.getBoundingClientRect().width / page.offsetWidth : 1;
		const top = panel.getBoundingClientRect().top;
		cards.sort((a, b) => this.cards.get(a).anchor.getBoundingClientRect().top - this.cards.get(b).anchor.getBoundingClientRect().top);
		const targets = cards.map(card => Math.max(0, (this.cards.get(card).anchor.getBoundingClientRect().top - top) / (zoom || 1)));
		const available = page.clientHeight - panel.offsetTop - (parseFloat(getComputedStyle(page).paddingBottom) || 0);
		if (available <= 0) return;
		let heights = cards.map(card => card.offsetHeight);
		const place = () => {
			// Sparse cards follow their anchors; dense cards first try tightly stacked layout.
			const trial = computePositions(heights, targets, ANNOTATION_GAP);
			const desired = cards.length > 3 && trial.lastBottom > available ? targets.map(() => 0) : targets;
			const fitted = fitWithinBounds(heights, desired, ANNOTATION_GAP, available);
			const positions = computePositions(heights, fitted, ANNOTATION_GAP);
			cards.forEach((card, i) => card.style.top = `${positions.tops[i]}px`);
			return positions;
		};
		let positions = place();
		// Retry after text reflow at the panel's final width.
		heights = cards.map(card => card.offsetHeight);
		positions = place();
		panel.style.height = `${Math.min(available, positions.lastBottom)}px`;
		if (positions.lastBottom > available) {
			if (this.expanded.has(panel)) {
				panel.style.maxHeight = `${available}px`;
				panel.style.overflowY = 'auto';
				const button = this.build({ tagName: 'button', type: 'button', children: ['Collapse annotations'], onclick: () => { this.expanded.delete(panel); this.schedule(); } }) as HTMLElement;
				button.dataset.expandAnnotations = '';
				button.style.cssText = 'position:sticky;top:0;height:32px;width:100%;z-index:3';
				panel.appendChild(button);
			} else {
				let visible = 0;
				cards.forEach((card, i) => { card.hidden = positions.tops[i] + heights[i] > available - 32; if (!card.hidden) visible++; });
				const button = this.build({ tagName: 'button', type: 'button', children: [`Show ${cards.length - visible} more annotations`], onclick: () => { this.expanded.add(panel); this.schedule(); } }) as HTMLElement;
				button.dataset.expandAnnotations = '';
				button.style.cssText = `position:absolute;top:${Math.max(0, available - 32)}px;height:32px;width:100%`;
				panel.appendChild(button);
			}
		}
	}

	private splitPages(pages: HTMLElement[]) {
		for (let i = 0; i < pages.length; i++) {
			const page = pages[i];
			const articles = page.querySelectorAll<HTMLElement>(':scope > article');
			// Preserve upstream continuous multi-section layout; only split a single article.
			if (articles.length !== 1) continue;
			const article = articles[0];
			const children = Array.from(article.children) as HTMLElement[];
			if (children.length < 2) continue;
			const cs = getComputedStyle(page);
			const height = parseFloat(cs.minHeight);
			const zoom = page.offsetWidth ? page.getBoundingClientRect().width / page.offsetWidth : 1;
			if (!height || page.getBoundingClientRect().height / (zoom || 1) <= height + 1) continue;
			let other = 0;
			for (const el of Array.from(page.children) as HTMLElement[]) {
				if (el !== article && !el.dataset.annotationPanel) other += el.offsetHeight;
			}
			const available = height - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom) - other;
			if (available <= 0) continue;
			const start = article.getBoundingClientRect().top;
			const index = children.findIndex((child, n) => n > 0 && (child.getBoundingClientRect().bottom - start) / (zoom || 1) > available);
			if (index < 1) continue;
			const next = page.cloneNode(false) as HTMLElement;
			const nextArticle = article.cloneNode(false) as HTMLElement;
			// Copy page furniture, then move content nodes so handlers and annotation anchors survive.
			for (const el of Array.from(page.children) as HTMLElement[]) {
				if (el.tagName === 'HEADER') next.appendChild(el.cloneNode(true));
			}
			next.appendChild(nextArticle);
			children.slice(index).forEach(child => nextArticle.appendChild(child));
			for (const el of Array.from(page.children) as HTMLElement[]) {
				if (el.tagName === 'FOOTER') next.appendChild(el.cloneNode(true));
				if (el.dataset.annotationPanel) next.appendChild(el.cloneNode(false));
			}
			page.after(next);
			pages.splice(i + 1, 0, next);
			this.observer?.observe(next);
		}
	}
}

export function annotationStyle(c: string): string {
	return `
.${c}-track-changes-margin, .${c}-track-changes-floating { position:absolute; box-sizing:border-box; font:12px/1.4 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; z-index:2; }
.${c}-tc-annotation { box-sizing:border-box; width:100%; padding:10px; background:#fff; color:#374151; border:1px solid #e5e7eb; border-left:3px solid #22a06b; border-radius:6px; display:flex; gap:8px; cursor:pointer; font:12px/1.4 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif !important; }
.${c}-tc-annotation[hidden] { display:none; }
.${c}-tc-annotation * { font-family:inherit !important; font-size:12px !important; }
.${c}-tc-annotation-deleted { border-left-color:#dc3545; }
.${c}-tc-annotation-moveFrom, .${c}-tc-annotation-moveTo { border-left-color:#805ad5; }
.${c}-tc-annotation-comment { border-left-color:#e0a020; }
.${c}-tc-annotation-formatChange { border-left-color:#3182ce; }
.${c}-tc-avatar { flex:0 0 24px; height:24px; border-radius:50%; background:#edf2f7; text-align:center; line-height:24px; font-size:10px !important; }
.${c}-tc-body { min-width:0; flex:1; }
.${c}-tc-body strong { font-size:12px !important; font-weight:600; }
.${c}-tc-body time { display:block; color:#6b7280; font-size:11px !important; }
.${c}-tc-preview { margin-top:4px; font-size:12px !important; max-height:5.6em; overflow:hidden; overflow-wrap:anywhere; }
.${c}-tc-active .${c}-tc-preview { max-height:none; }
.${c}-tc-members { margin-top:6px; display:flex; flex-direction:column; gap:4px; }
.${c}-tc-member { padding:4px 6px; border-radius:4px; background:#f9fafb; }
.${c}-tc-member-deleted { border-left:2px solid #dc3545; }
.${c}-tc-member-inserted { border-left:2px solid #22a06b; }
.${c}-tc-undecidable { margin-top:4px; color:#92400e; font-style:italic; }
.${c}-tc-selected { outline:2px solid #3182ce; outline-offset:1px; }
.${c}-tc-active, .${c}-tc-annotation:focus-visible { outline:2px solid #90cdf4; }
.${c}-tc-inserted, .${c}-tc-moveTo { background:#e6f7ed; }
.${c}-tc-deleted, .${c}-tc-moveFrom { background:#fde8e8; text-decoration:line-through; }
@media print { .${c}-track-changes-floating { display:none; } }
`;
}


export function changeButtons(options: Options, build: typeof h, targets: DecisionTarget[]): HTMLElement {
	targets = targets.filter(t => t.key != null);
	if (!targets.length || !options.showTrackChangeButtons || !options.onTrackChangeDecision) return null;
	return build({ tagName: 'span', className: `${options.className}-tc-actions`, children: ['accept', 'reject'].map((decision: 'accept' | 'reject') => ({
		tagName: 'button', type: 'button', title: `${decision === 'accept' ? 'Accept' : 'Reject'} change`,
		children: [decision === 'accept' ? 'Accept' : 'Reject'],
		onclick: (event: MouseEvent) => {
			event.stopPropagation();
			targets.forEach(t => options.onTrackChangeDecision(t.key, decision, t.record));
		}
	})) }) as HTMLElement;
}

/** Marks the change(s) carrying `key` (anchors, cards, replaced-card members) as selected; null clears. */
export function selectTrackChange(root: ParentNode, key: string | null, className = 'docx') {
	const cls = `${className}-tc-selected`;
	for (const el of Array.from(root.querySelectorAll<HTMLElement>('[data-tc-key],[data-tc-keys]'))) {
		let keys: (string | null)[] = [];
		if (el.dataset.tcKeys) { try { keys = JSON.parse(el.dataset.tcKeys); } catch { keys = []; } }
		else keys = [el.dataset.tcKey];
		el.classList.toggle(cls, key != null && keys.includes(key));
	}
}

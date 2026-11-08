import { TrackChangeDecision } from './docx-preview';
import { DomType, OpenXmlElement, WmlChange } from './document/dom';

/** Revision kinds as the server names them (rPrChange and pPrChange are separate). */
export type TrackChangeRecordType = 'ins' | 'del' | 'moveFrom' | 'moveTo' | 'rPrChange' | 'pPrChange';

/**
 * One rendered-or-renderable change. `(part, wId, type, occurrence)` is its identity:
 * `occurrence` counts changes sharing the same part, w:id and type in document order
 * (w:id repeats across parts and within one). It never depends on decisions or on
 * what is currently hidden.
 */
export interface TrackChangeRecord {
	/** Part id: `body`, `hdr<n>`, `ftr<n>`, `fn`, `en` (same names as the server). */
	part: string;
	/** The w:id attribute verbatim. */
	wId: string;
	type: TrackChangeRecordType;
	author: string;
	date: string;
	/** Inserted/deleted text; empty for formatting changes. */
	text: string;
	occurrence: number;
}

export type TrackChangeKeyFn = (record: TrackChangeRecord) => string | null | undefined;

/** The key used when the caller supplies none. Never the bare w:id. */
export function defaultChangeKey(r: TrackChangeRecord): string {
	return `${r.part}:${r.wId}:${r.type}#${r.occurrence}`;
}

/** Explicit decisions take priority over whether pending revisions are displayed. */
export function changeVisibility(change: WmlChange, decision: TrackChangeDecision, renderChanges: boolean): 'content' | 'hidden' | 'revision' {
	const removed = change.type === DomType.Deleted || change.type === DomType.MoveFrom;
	if (decision === 'accept') return removed ? 'hidden' : 'content';
	if (decision === 'reject') return removed ? 'content' : 'hidden';
	return renderChanges ? 'revision' : removed ? 'hidden' : 'content';
}

const CHANGE_TYPES = [DomType.Inserted, DomType.Deleted, DomType.MoveFrom, DomType.MoveTo, DomType.FormatChange];

function recordType(change: WmlChange): TrackChangeRecordType {
	switch (change.type) {
		case DomType.Inserted: return 'ins';
		case DomType.Deleted: return 'del';
		case DomType.MoveFrom: return 'moveFrom';
		case DomType.MoveTo: return 'moveTo';
		default: return change.formatKind === 'pPrChange' ? 'pPrChange' : 'rPrChange';
	}
}

export function changeText(n: OpenXmlElement): string {
	if (!n) return '';
	let s = '';
	if ((n.type === DomType.Text || n.type === DomType.DeletedText) && typeof (n as any).text === 'string') s += (n as any).text;
	else if (n.type === DomType.Tab) s += '\t';
	else if (n.type === DomType.Break && (n as any).break === 'textWrapping') s += '\n';
	for (const c of n.children ?? []) s += changeText(c);
	return s;
}

export function partName(part: any, index: number): string {
	const path: string = part?.path ?? '';
	const m = /([A-Za-z]+?)(\d*)\.xml$/.exec(path);
	const base = m?.[1].toLowerCase();
	if (base === 'header') return `hdr${m[2] || 1}`;
	if (base === 'footer') return `ftr${m[2] || 1}`;
	if (base === 'footnotes') return 'fn';
	if (base === 'endnotes') return 'en';
	return base ? `${base}${m[2]}` : `part${index}`;
}

/**
 * Walks every part in document order, stamps `part` and `occurrence` on each change
 * node and returns one record per change. Idempotent; call it before or after rendering.
 */
export function indexTrackChanges(document: any): TrackChangeRecord[] {
	const records: TrackChangeRecord[] = [];
	const counts = new Map<string, number>();
	const visit = (el: OpenXmlElement, part: string) => {
		if (!el) return;
		if (CHANGE_TYPES.includes(el.type)) {
			const c = el as WmlChange;
			if (c.id != null) {
				const type = recordType(c);
				const k = `${part}\u0000${c.id}\u0000${type}`;
				const occurrence = counts.get(k) ?? 0;
				counts.set(k, occurrence + 1);
				c.part = part;
				c.occurrence = occurrence;
				records.push({ part, wId: c.id, type, author: c.author ?? null, date: c.date ?? null,
					text: c.type === DomType.FormatChange ? '' : changeText(c), occurrence });
			}
		}
		if (el.formatChange) visit(el.formatChange, part);
		el.children?.forEach(child => visit(child, part));
	};
	visit(document?.documentPart?.body, 'body');
	(document?.parts ?? []).forEach((part: any, i: number) => {
		const name = partName(part, i);
		if (part === document.documentPart) return;
		visit(part.rootElement, name);
		for (const note of part.notes ?? []) visit(note, name);
	});
	return records;
}

/** The record of an indexed change node. */
export function recordOf(change: WmlChange): TrackChangeRecord {
	return { part: change.part, wId: change.id, type: recordType(change), author: change.author ?? null, date: change.date ?? null,
		text: change.type === DomType.FormatChange ? '' : changeText(change), occurrence: change.occurrence ?? 0 };
}

/** Key for every record that the key function accepts, in document order, without duplicates. */
export function collectChangeKeys(document: any, keyOf: TrackChangeKeyFn = defaultChangeKey): string[] {
	const keys = new Set<string>();
	for (const r of indexTrackChanges(document)) {
		const k = keyOf(r);
		if (k != null) keys.add(k);
	}
	return Array.from(keys);
}

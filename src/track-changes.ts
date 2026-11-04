import { TrackChangeDecision } from './docx-preview';
import { DomType, OpenXmlElement, WmlChange } from './document/dom';

/** Explicit decisions take priority over whether pending revisions are displayed. */
export function changeVisibility(change: WmlChange, decision: TrackChangeDecision, renderChanges: boolean): 'content' | 'hidden' | 'revision' {
	const removed = change.type === DomType.Deleted || change.type === DomType.MoveFrom;
	if (decision === 'accept') return removed ? 'hidden' : 'content';
	if (decision === 'reject') return removed ? 'content' : 'hidden';
	return renderChanges ? 'revision' : removed ? 'hidden' : 'content';
}

export function collectChangeIds(element: OpenXmlElement, ids: Set<string>) {
	if (!element) return;
	if ([DomType.Inserted, DomType.Deleted, DomType.MoveFrom, DomType.MoveTo, DomType.FormatChange].includes(element.type)) {
		const id = (element as WmlChange).id;
		if (id != null) ids.add(id);
	}
	if (element.formatChange) collectChangeIds(element.formatChange, ids);
	element.children?.forEach(child => collectChangeIds(child, ids));
}

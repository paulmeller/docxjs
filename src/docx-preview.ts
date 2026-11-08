import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { h } from './html';
import { collectChangeKeys, indexTrackChanges, TrackChangeKeyFn, TrackChangeRecord } from './track-changes';

export { TrackChangeRecord, TrackChangeRecordType, TrackChangeKeyFn, defaultChangeKey } from './track-changes';
export { selectTrackChange } from './annotations';

export type TrackChangesMode = 'inline' | 'margin' | 'floating';
export type TrackChangeDecision = 'accept' | 'reject' | 'pending';

export interface Options {
    inWrapper: boolean;
    hideWrapperOnPrint: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
	renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
	useBase64URL: boolean;
	renderChanges: boolean;
    renderComments: boolean;
    renderAltChunks: boolean;
    trackChangesMode: TrackChangesMode;
    commentsMode: TrackChangesMode;
    trackChangesMarginWidth: string;
    showTrackChangeButtons: boolean;
    /** Decisions keyed by the key `trackChangeKey` returns (default `part:wId:type#occurrence`), never by bare w:id. */
    trackChangeDecisions: Record<string, TrackChangeDecision>;
    /** Maps a change to the caller's key. Return null/undefined and the change is shown but cannot be decided. */
    trackChangeKey?: TrackChangeKeyFn;
    onTrackChangeDecision?: (key: string, decision: 'accept' | 'reject', record: TrackChangeRecord) => void;
    /** A margin card (or one member of a "replaced" card) was clicked. */
    onTrackChangeSelect?: (key: string | null, record: TrackChangeRecord) => void;
    /** Card text for changes whose key function returned null. */
    trackChangeUndecidableText: string;
    h: typeof h;
}

export const defaultOptions: Options = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    hideWrapperOnPrint: false,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
	renderEndnotes: true,
	useBase64URL: false,
	renderChanges: false,
    renderComments: false,
    renderAltChunks: true,
    trackChangesMode: 'inline',
    commentsMode: 'inline',
    trackChangesMarginWidth: '220px',
    showTrackChangeButtons: true,
    trackChangeDecisions: {},
    trackChangeUndecidableText: "Can't match this change",
    h: h
};

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any>  {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderDocument(document: any, userOptions?: Partial<Options>): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer();
    return await renderer.render(document, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	const doc = await parseAsync(data, userOptions);
	const nodes = await renderDocument(doc, userOptions);

    styleContainer ??= bodyContainer;
    styleContainer.innerHTML = "";
    bodyContainer.innerHTML = "";

    for (let n of nodes) {
        const c = n.nodeName === "STYLE" ? styleContainer : bodyContainer;
        c.appendChild(n);
    }
    
    return doc;
}

/** Every change of the document (all parts), in document order, with part and occurrence. */
export function getTrackChanges(document: any): TrackChangeRecord[] {
	return indexTrackChanges(document);
}

/** Decision keys of the document's changes (the default key unless `keyOf` is given). */
export function getTrackChangeIds(document: any, keyOf?: TrackChangeKeyFn): string[] {
	return collectChangeKeys(document, keyOf);
}

export function acceptAllChanges(document: any, keyOf?: TrackChangeKeyFn): Record<string, TrackChangeDecision> {
	return Object.fromEntries(getTrackChangeIds(document, keyOf).map(id => [id, 'accept']));
}

export function rejectAllChanges(document: any, keyOf?: TrackChangeKeyFn): Record<string, TrackChangeDecision> {
	return Object.fromEntries(getTrackChangeIds(document, keyOf).map(id => [id, 'reject']));
}

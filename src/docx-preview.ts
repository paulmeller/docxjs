import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { DomType, OpenXmlElement, WmlTrackChange } from './document/dom';

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
    trackChangesMarginWidth: string;
    trackChangeDecisions: Record<string, TrackChangeDecision>;
    onTrackChangeDecision?: (id: string, decision: 'accept' | 'reject') => void;
    commentsMode: TrackChangesMode;
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
    trackChangesMarginWidth: '220px',
    trackChangeDecisions: {},
    commentsMode: 'inline'
}

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any>  {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderDocument(document: any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer(window.document);
	return await renderer.render(document, bodyContainer, styleContainer, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	const doc = await parseAsync(data, userOptions);
	await renderDocument(doc, bodyContainer, styleContainer, userOptions);
    return doc;
}

function collectTrackChangeIds(elem: OpenXmlElement, ids: Set<string>): void {
    const trackChangeTypes = [DomType.Inserted, DomType.Deleted, DomType.MoveFrom, DomType.MoveTo, DomType.FormatChange];
    if (trackChangeTypes.includes(elem.type)) {
        const tc = elem as WmlTrackChange;
        if (tc.id) ids.add(tc.id);
    }
    if (elem.children) {
        for (const child of elem.children) {
            collectTrackChangeIds(child, ids);
        }
    }
}

export function getTrackChangeIds(document: any): string[] {
    const ids = new Set<string>();
    const doc = document as WordDocument;
    if (doc?.documentPart?.body) {
        collectTrackChangeIds(doc.documentPart.body, ids);
    }
    return Array.from(ids);
}

export function acceptAllChanges(document: any): Record<string, TrackChangeDecision> {
    const decisions: Record<string, TrackChangeDecision> = {};
    for (const id of getTrackChangeIds(document)) {
        decisions[id] = 'accept';
    }
    return decisions;
}

export function rejectAllChanges(document: any): Record<string, TrackChangeDecision> {
    const decisions: Record<string, TrackChangeDecision> = {};
    for (const id of getTrackChangeIds(document)) {
        decisions[id] = 'reject';
    }
    return decisions;
}
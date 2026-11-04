import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { h } from './html';
import { collectChangeIds } from './track-changes';

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
    trackChangeDecisions: Record<string, TrackChangeDecision>;
    onTrackChangeDecision?: (id: string, decision: 'accept' | 'reject') => void;
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

export function getTrackChangeIds(document: any): string[] {
	const ids = new Set<string>();
	collectChangeIds(document?.documentPart?.body, ids);
	for (const part of document?.parts ?? []) {
		collectChangeIds(part.rootElement, ids);
		for (const note of part.notes ?? []) collectChangeIds(note, ids);
	}
	return Array.from(ids);
}

export function acceptAllChanges(document: any): Record<string, TrackChangeDecision> {
	return Object.fromEntries(getTrackChangeIds(document).map(id => [id, 'accept']));
}

export function rejectAllChanges(document: any): Record<string, TrackChangeDecision> {
	return Object.fromEntries(getTrackChangeIds(document).map(id => [id, 'reject']));
}

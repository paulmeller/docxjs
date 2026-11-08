/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */

export type HElement = {
    ns?: string;
    tagName: string;
    className?: string;
    style?: Record<string, string> | string;
    children?: (HElement | Node | string)[];
} & Record<string, any>;

export type TrackChangesMode = 'inline' | 'margin' | 'floating';
export type TrackChangeDecision = 'accept' | 'reject' | 'pending';

export type TrackChangeRecordType = 'ins' | 'del' | 'moveFrom' | 'moveTo' | 'rPrChange' | 'pPrChange';
export interface TrackChangeRecord {
    /** body, hdr<n>, ftr<n>, fn, en */
    part: string;
    /** the w:id attribute verbatim */
    wId: string;
    type: TrackChangeRecordType;
    author: string | null;
    date: string | null;
    text: string;
    /** index among changes with the same part, w:id and type, in document order */
    occurrence: number;
}
export type TrackChangeKeyFn = (record: TrackChangeRecord) => string | null | undefined;

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
    trackChangeKey?: TrackChangeKeyFn;
    onTrackChangeDecision?: (key: string, decision: 'accept' | 'reject', record: TrackChangeRecord) => void;
    onTrackChangeSelect?: (key: string | null, record: TrackChangeRecord) => void;
    trackChangeUndecidableText: string;
    h: (elemOrText: HElement | Node | string) => Node; //experimental, subject to change
}

//stub
export type WordDocument = any;
export declare const defaultOptions: Options;
export declare function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<WordDocument>;
export declare function renderDocument(document: WordDocument, userOptions?: Partial<Options>): Promise<Node[]>;
export declare function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any>;

export declare function defaultChangeKey(record: TrackChangeRecord): string;
export declare function getTrackChanges(document: WordDocument): TrackChangeRecord[];
export declare function getTrackChangeIds(document: WordDocument, keyOf?: TrackChangeKeyFn): string[];
export declare function acceptAllChanges(document: WordDocument, keyOf?: TrackChangeKeyFn): Record<string, TrackChangeDecision>;
export declare function rejectAllChanges(document: WordDocument, keyOf?: TrackChangeKeyFn): Record<string, TrackChangeDecision>;
export declare function selectTrackChange(root: ParentNode, key: string | null, className?: string): void;

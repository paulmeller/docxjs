import { WordDocument } from './word-document';
import {
	DomType, WmlTable, IDomNumbering,
	WmlHyperlink, IDomImage, OpenXmlElement, WmlTableColumn, WmlTableCell, WmlText, WmlSymbol, WmlBreak, WmlNoteReference,
	WmlSmartTag,
	WmlAltChunk,
	WmlTableRow,
	WmlTrackChange,
	WmlMoveRangeMarker,
	TrackChangeType
} from './document/dom';
import { CommonProperties } from './document/common';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { asArray, encloseFontFamily, escapeClassName, isString, keyBy, mergeDeep } from './utils';
import { computePixelToPoint, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties } from './document/section';
import { WmlRun, RunProperties } from './document/run';
import { WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { BaseHeaderFooterPart } from './header-footer/parts';
import { Part } from './common/part';
import { VmlElement } from './vml/vml';
import { WmlComment, WmlCommentRangeStart, WmlCommentReference } from './comments/elements';

const ns = {
	svg: "http://www.w3.org/2000/svg",
	mathML: "http://www.w3.org/1998/Math/MathML"
}

// Annotation overflow constants
const ANNOTATION_GAP = 8;
const COMPRESSED_GAP = 2;
const EXPAND_BUTTON_HEIGHT = 32;

interface CellPos {
	col: number;
	row: number;
}

interface Section {
	sectProps: SectionProperties;
	elements: OpenXmlElement[];
	pageBreak: boolean;
}

interface TrackChangeAnnotation {
	id: string;
	author: string;
	date: string;
	changeType: TrackChangeType;
	previewText: string;
	formatDescription?: string;
	element?: HTMLElement;
	contentElement?: HTMLElement;
}

interface TrackChangeEntry {
	annotation: TrackChangeAnnotation;
	contentElements: Node[];
}

declare const Highlight: any;

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

export class HtmlRenderer {

	className: string = "docx";
	rootSelector: string;
	document: WordDocument;
	options: Options;
	styleMap: Record<string, IDomStyle> = {};
	currentPart: Part = null;

	tableVerticalMerges: CellVerticalMergeType[] = [];
	currentVerticalMerge: CellVerticalMergeType = null;
	tableCellPositions: CellPos[] = [];
	currentCellPosition: CellPos = null;

	footnoteMap: Record<string, WmlFootnote> = {};
	endnoteMap: Record<string, WmlFootnote> = {};
	currentFootnoteIds: string[];
	currentEndnoteIds: string[] = [];
	usedHederFooterParts: any[] = [];

	defaultTabSize: string;
	currentTabs: any[] = [];

	commentHighlight: any;
	commentMap: Record<string, Range> = {};

	trackChangeMap: Record<string, TrackChangeEntry> = {};
	trackChangeHighlights: {
		inserted: any;
		deleted: any;
		moveFrom: any;
		moveTo: any;
		formatChange: any;
	} = null;
	currentMarginContainer: HTMLElement = null;
	currentPageElement: HTMLElement = null;

	// Floating mode: collect all annotations across pages
	floatingTrackChangeMap: Record<string, TrackChangeEntry> = {};
	floatingPanelElement: HTMLElement = null;

	tasks: Promise<any>[] = [];
	postRenderTasks: any[] = [];

	constructor(public htmlDocument: Document) {
	}

	async render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
		this.document = document;
		this.options = options;
		this.className = options.className;
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		this.styleMap = null;
		this.tasks = [];
		this.trackChangeMap = {};
		this.postRenderTasks = [];

		if (this.options.renderComments && globalThis.Highlight) {
			this.commentHighlight = new Highlight();
		}

		const needsTrackChangeHighlights = (this.options.renderChanges && (this.options.trackChangesMode === 'margin' || this.options.trackChangesMode === 'floating'))
			|| (this.options.renderComments && (this.options.commentsMode === 'margin' || this.options.commentsMode === 'floating'));
		if (needsTrackChangeHighlights && globalThis.Highlight) {
			this.trackChangeHighlights = {
				inserted: new Highlight(),
				deleted: new Highlight(),
				moveFrom: new Highlight(),
				moveTo: new Highlight(),
				formatChange: new Highlight()
			};
		}

		// Initialize floating mode collection
		this.floatingTrackChangeMap = {};
		// Clean up any existing floating panel from body
		if (this.floatingPanelElement && this.floatingPanelElement.parentNode) {
			this.floatingPanelElement.parentNode.removeChild(this.floatingPanelElement);
		}
		// Also clean up any orphaned floating panels
		this.htmlDocument.querySelectorAll(`.${this.className}-track-changes-floating`).forEach(el => el.remove());
		this.floatingPanelElement = null;

		styleContainer = styleContainer || bodyContainer;

		removeAllElements(styleContainer);
		removeAllElements(bodyContainer);

		styleContainer.appendChild(this.createComment("docxjs library predefined styles"));
		styleContainer.appendChild(this.renderDefaultStyle());

		if (document.themePart) {
			styleContainer.appendChild(this.createComment("docxjs document theme values"));
			this.renderTheme(document.themePart, styleContainer);
		}

		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);

			styleContainer.appendChild(this.createComment("docxjs document styles"));
			styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
		}

		if (document.numberingPart) {
			this.prodessNumberings(document.numberingPart.domNumberings);

			styleContainer.appendChild(this.createComment("docxjs document numbering styles"));
			styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
			//styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
		}

		if (document.footnotesPart) {
			this.footnoteMap = keyBy(document.footnotesPart.notes, x => x.id);
		}

		if (document.endnotesPart) {
			this.endnoteMap = keyBy(document.endnotesPart.notes, x => x.id);
		}

		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}

		if (!options.ignoreFonts && document.fontTablePart)
			this.renderFontTable(document.fontTablePart, styleContainer);

		var sectionElements = this.renderSections(document.documentPart.body);

		if (this.options.inWrapper) {
			bodyContainer.appendChild(this.renderWrapper(sectionElements));
		} else {
			appendChildren(bodyContainer, sectionElements);
		}

		if (this.commentHighlight && options.renderComments) {
			(CSS as any).highlights.set(`${this.className}-comments`, this.commentHighlight);
		}

		if (this.trackChangeHighlights && (
			(options.renderChanges && (options.trackChangesMode === 'margin' || options.trackChangesMode === 'floating'))
			|| (options.renderComments && (options.commentsMode === 'margin' || options.commentsMode === 'floating'))
		)) {
			(CSS as any).highlights.set(`${this.className}-tc-inserted`, this.trackChangeHighlights.inserted);
			(CSS as any).highlights.set(`${this.className}-tc-deleted`, this.trackChangeHighlights.deleted);
			(CSS as any).highlights.set(`${this.className}-tc-move-from`, this.trackChangeHighlights.moveFrom);
			(CSS as any).highlights.set(`${this.className}-tc-move-to`, this.trackChangeHighlights.moveTo);
			(CSS as any).highlights.set(`${this.className}-tc-format`, this.trackChangeHighlights.formatChange);
		}

		this.postRenderTasks.forEach(t => t());

		// Split pages whose rendered content overflows minHeight,
		// then recalculate annotation positions with overflow handling
		this.splitOverflowingPages(bodyContainer);

		await Promise.allSettled(this.tasks);

		this.refreshTabStops();
	}

	renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
		const variables = {};
		const fontScheme = themePart.theme?.fontScheme;

		if (fontScheme) {
			if (fontScheme.majorFont) {
				variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
			}

			if (fontScheme.minorFont) {
				variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
			}
		}

		const colorScheme = themePart.theme?.colorScheme;

		if (colorScheme) {
			for (let [k, v] of Object.entries(colorScheme.colors)) {
				variables[`--docx-${k}-color`] = `#${v}`;
			}
		}

		const cssText = this.styleToString(`.${this.className}`, variables);
		styleContainer.appendChild(this.createStyleElement(cssText));
	}

	renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
		for (let f of fontsPart.fonts) {
			for (let ref of f.embedFontRefs) {
				this.tasks.push(this.document.loadFont(ref.id, ref.key).then(fontData => {
					const cssValues = {
						'font-family': encloseFontFamily(f.name),
						'src': `url(${fontData})`
					};

					if (ref.type == "bold" || ref.type == "boldItalic") {
						cssValues['font-weight'] = 'bold';
					}

					if (ref.type == "italic" || ref.type == "boldItalic") {
						cssValues['font-style'] = 'italic';
					}

					const cssText = this.styleToString("@font-face", cssValues);
					styleContainer.appendChild(this.createComment(`docxjs ${f.name} font`));
					styleContainer.appendChild(this.createStyleElement(cssText));
				}));
			}
		}
	}

	processStyleName(className: string): string {
		return className ? `${this.className}_${escapeClassName(className)}` : this.className;
	}

	processStyles(styles: IDomStyle[]) {
		const stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id);

		for (const style of styles.filter(x => x.basedOn)) {
			var baseStyle = stylesMap[style.basedOn];

			if (baseStyle) {
				style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
				style.runProps = mergeDeep(style.runProps, baseStyle.runProps);

				for (const baseValues of baseStyle.styles) {
					const styleValues = style.styles.find(x => x.target == baseValues.target);

					if (styleValues) {
						this.copyStyleProperties(baseValues.values, styleValues.values);
					} else {
						style.styles.push({ ...baseValues, values: { ...baseValues.values } });
					}
				}
			}
			else if (this.options.debug)
				console.warn(`Can't find base style ${style.basedOn}`);
		}

		for (let style of styles) {
			style.cssName = this.processStyleName(style.id);
		}

		return stylesMap;
	}

	prodessNumberings(numberings: IDomNumbering[]) {
		for (let num of numberings.filter(n => n.pStyleName)) {
			const style = this.findStyle(num.pStyleName);

			if (style?.paragraphProps?.numbering) {
				style.paragraphProps.numbering.level = num.level;
			}
		}
	}

	processElement(element: OpenXmlElement) {
		if (element.children) {
			for (var e of element.children) {
				e.parent = element;

				if (e.type == DomType.Table) {
					this.processTable(e);
				}
				else {
					this.processElement(e);
				}
			}
		}
	}

	processTable(table: WmlTable) {
		for (var r of table.children) {
			for (var c of r.children) {
				c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
					"border-left", "border-right", "border-top", "border-bottom",
					"padding-left", "padding-right", "padding-top", "padding-bottom"
				]);

				this.processElement(c);
			}
		}
	}

	copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
		if (!input)
			return output;

		if (output == null) output = {};
		if (attrs == null) attrs = Object.getOwnPropertyNames(input);

		for (var key of attrs) {
			if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
				output[key] = input[key];
		}

		return output;
	}

	createPageElement(className: string, props: SectionProperties): HTMLElement {
		var elem = this.createElement("section", { className });

		if (props) {
			const isMarginMode = (this.options.renderChanges && this.options.trackChangesMode === 'margin')
				|| (this.options.renderComments && this.options.commentsMode === 'margin');

			if (props.pageMargins) {
				elem.style.paddingLeft = props.pageMargins.left;
				elem.style.paddingTop = props.pageMargins.top;
				elem.style.paddingBottom = props.pageMargins.bottom;
				elem.style.paddingRight = props.pageMargins.right;
			}

			if (props.pageSize) {
				if (!this.options.ignoreWidth) {
					let width = props.pageSize.width;
					// Expand page width for margin mode (panel is inside the section)
					if (isMarginMode) {
						const marginWidthStr = this.options.trackChangesMarginWidth || '220px';
						elem.style.width = `calc(${width} + ${marginWidthStr} + 16px)`;
					} else {
						elem.style.width = width;
					}
				}
				if (!this.options.ignoreHeight)
					elem.style.minHeight = props.pageSize.height;
			}
		}

		return elem;
	}

	createSectionContent(props: SectionProperties): HTMLElement {
		var elem = this.createElement("article")

		const isMarginMode = (this.options.renderChanges && this.options.trackChangesMode === 'margin')
			|| (this.options.renderComments && this.options.commentsMode === 'margin');

		if (isMarginMode) {
			elem.style.overflow = "hidden";
			const marginWidth = this.options.trackChangesMarginWidth || '220px';
			elem.style.maxWidth = `calc(100% - ${marginWidth} - 16px)`;
		}

		if (props.columns && props.columns.numberOfColumns) {
			elem.style.columnCount = `${props.columns.numberOfColumns}`;
			elem.style.columnGap = props.columns.space;

			if (props.columns.separator) {
				elem.style.columnRule = "1px solid black";
			}
		}

		return elem;
	}	

	renderSections(document: DocumentElement): HTMLElement[] {
		const result = [];
		const tcMargin = this.options.renderChanges && this.options.trackChangesMode === 'margin';
		const tcFloating = this.options.renderChanges && this.options.trackChangesMode === 'floating';
		const cmMargin = this.options.renderComments && this.options.commentsMode === 'margin';
		const cmFloating = this.options.renderComments && this.options.commentsMode === 'floating';
		const isMarginMode = tcMargin || cmMargin;
		const isFloatingMode = tcFloating || cmFloating;

		this.processElement(document);
		const sections = this.splitBySection(document.children, document.props);
		const pages = this.groupByPageBreaks(sections);
		let prevProps = null;

		for (let i = 0, l = pages.length; i < l; i++) {
			this.currentFootnoteIds = [];
			this.trackChangeMap = {}; // Reset for each page

			const section = pages[i][0];
			let props = section.sectProps;
			const pageElement = this.createPageElement(this.className, props);
			this.currentPageElement = pageElement;
			this.renderStyleValues(document.cssStyle, pageElement);

			// Create margin container if in margin mode (not for floating mode)
			let marginContainer: HTMLElement = null;

			if (isMarginMode) {
				marginContainer = this.createElement("aside", { className: `${this.className}-track-changes-margin` });
				marginContainer.style.width = this.options.trackChangesMarginWidth;
				this.currentMarginContainer = marginContainer;
			}

			this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs, props,
				result.length, prevProps != props, pageElement);

			for (const sect of pages[i]) {
				var contentElement = this.createSectionContent(sect.sectProps);
				this.renderElements(sect.elements, contentElement);
				pageElement.appendChild(contentElement);
				props = sect.sectProps;
			}

			if (this.options.renderFootnotes) {
				this.renderNotes(this.currentFootnoteIds, this.footnoteMap, pageElement);
			}

			if (this.options.renderEndnotes && i == l - 1) {
				this.renderNotes(this.currentEndnoteIds, this.endnoteMap, pageElement);
			}

			this.options.renderFooters && this.renderHeaderFooter(props.footerRefs, props,
				result.length, prevProps != props, pageElement);

			if (isMarginMode) {
				// Append margin directly to page element (positioned absolutely)
				pageElement.appendChild(marginContainer);
				pageElement.classList.add(`${this.className}-has-track-changes`);

				// Filter entries: only include TC entries if TC is margin, comment entries if comments is margin
				const pageTrackChangeMap: Record<string, TrackChangeEntry> = {};
				for (const [key, entry] of Object.entries(this.trackChangeMap)) {
					const isComment = entry.annotation.changeType === 'comment';
					if (isComment ? cmMargin : tcMargin) pageTrackChangeMap[key] = entry;
				}

				// Schedule annotation rendering after DOM is ready
				this.later(() => this.renderMarginAnnotationsFromMap(marginContainer, pageElement, pageTrackChangeMap));
			}

			if (isFloatingMode) {
				// Filter entries: only include TC entries if TC is floating, comment entries if comments is floating
				for (const [key, entry] of Object.entries(this.trackChangeMap)) {
					const isComment = entry.annotation.changeType === 'comment';
					if (isComment ? cmFloating : tcFloating) this.floatingTrackChangeMap[key] = entry;
				}
			}

			result.push(pageElement);
			prevProps = props;
		}

		return result;
	}

	renderHeaderFooter(refs: FooterHeaderReference[], props: SectionProperties, page: number, firstOfSection: boolean, into: HTMLElement) {
		if (!refs) return;

		var ref = (props.titlePage && firstOfSection ? refs.find(x => x.type == "first") : null)
			?? (page % 2 == 1 ? refs.find(x => x.type == "even") : null)
			?? refs.find(x => x.type == "default");

		var part = ref && this.document.findPartByRelId(ref.id, this.document.documentPart) as BaseHeaderFooterPart;

		if (part) {
			this.currentPart = part;
			if (!this.usedHederFooterParts.includes(part.path)) {
				this.processElement(part.rootElement);
				this.usedHederFooterParts.push(part.path);
			}
			const [el] = this.renderElements([part.rootElement], into) as HTMLElement[];

			if (props?.pageMargins) {
				if (part.rootElement.type === DomType.Header) {
					el.style.marginTop = `calc(${props.pageMargins.header} - ${props.pageMargins.top})`;
					el.style.minHeight = `calc(${props.pageMargins.top} - ${props.pageMargins.header})`;
				}
				else if (part.rootElement.type === DomType.Footer) {
					el.style.marginBottom = `calc(${props.pageMargins.footer} - ${props.pageMargins.bottom})`;
					el.style.minHeight = `calc(${props.pageMargins.bottom} - ${props.pageMargins.footer})`;
				}
			}

			this.currentPart = null;
		}
	}

	isPageBreakElement(elem: OpenXmlElement): boolean {
		if (elem.type != DomType.Break)
			return false;

		if ((elem as WmlBreak).break == "lastRenderedPageBreak")
			return !this.options.ignoreLastRenderedPageBreak;

		return (elem as WmlBreak).break == "page";
	}

	isPageBreakSection(prev: SectionProperties, next: SectionProperties): boolean {
		if (!prev) return false;
		if (!next) return false;

		return prev.pageSize?.orientation != next.pageSize?.orientation
			|| prev.pageSize?.width != next.pageSize?.width
			|| prev.pageSize?.height != next.pageSize?.height;
	}

	splitBySection(elements: OpenXmlElement[], defaultProps: SectionProperties): Section[] {
		var current: Section = { sectProps: null, elements: [], pageBreak: false };
		var result = [current];

		for (let elem of elements) {
			if (elem.type == DomType.Paragraph) {
				const s = this.findStyle((elem as WmlParagraph).styleName);

				if (s?.paragraphProps?.pageBreakBefore) {
					current.sectProps = sectProps;
					current.pageBreak = true;
					current = { sectProps: null, elements: [], pageBreak: false };
					result.push(current);
				}
			}

			current.elements.push(elem);

			if (elem.type == DomType.Paragraph) {
				const p = elem as WmlParagraph;

				var sectProps = p.sectionProps;
				var pBreakIndex = -1;
				var rBreakIndex = -1;

				if (this.options.breakPages && p.children) {
					pBreakIndex = p.children.findIndex(r => {
						rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
						return rBreakIndex != -1;
					});
				}

				if (sectProps || pBreakIndex != -1) {
					current.sectProps = sectProps;
					current.pageBreak = pBreakIndex != -1;
					current = { sectProps: null, elements: [], pageBreak: false };
					result.push(current);
				}

				if (pBreakIndex != -1) {
					let breakRun = p.children[pBreakIndex];
					let splitRun = rBreakIndex < breakRun.children.length - 1;

					if (pBreakIndex < p.children.length - 1 || splitRun) {
						var children = elem.children;
						var newParagraph = { ...elem, children: children.slice(pBreakIndex) };
						elem.children = children.slice(0, pBreakIndex);
						current.elements.push(newParagraph);

						if (splitRun) {
							let runChildren = breakRun.children;
							let newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
							elem.children.push(newRun);
							breakRun.children = runChildren.slice(rBreakIndex);
						}
					}
				}
			}
		}

		let currentSectProps = null;

		for (let i = result.length - 1; i >= 0; i--) {
			if (result[i].sectProps == null) {
				result[i].sectProps = currentSectProps ?? defaultProps;
			} else {
				currentSectProps = result[i].sectProps
			}
		}

		return result;
	}

	groupByPageBreaks(sections: Section[]): Section[][] {
		let current = [];
		let prev: SectionProperties;
		const result: Section[][] = [current];

		for (let s of sections) {
			// Start new page BEFORE this section if page size/orientation changed
			if (this.isPageBreakSection(prev, s.sectProps))
				result.push(current = []);

			current.push(s);

			// Start new page AFTER this section if it contains a page break marker
			if (this.options.breakPages && s.pageBreak)
				result.push(current = []);

			prev = s.sectProps;
		}

		return result.filter(x => x.length > 0);
	}

	renderWrapper(children: HTMLElement[]) {
		const isFloatingMode = (this.options.renderChanges && this.options.trackChangesMode === 'floating')
			|| (this.options.renderComments && this.options.commentsMode === 'floating');

		const wrapper = this.createElement("div", { className: `${this.className}-wrapper` }, children);

		if (isFloatingMode) {
			wrapper.classList.add(`${this.className}-has-floating-panel`);

			// Create floating panel (positioned next to first page)
			const floatingPanel = this.createElement("aside", {
				className: `${this.className}-track-changes-floating`
			});
			floatingPanel.style.width = this.options.trackChangesMarginWidth;
			this.floatingPanelElement = floatingPanel;

			// Append to first page section so left:100% positions it next to the page
			const firstPage = children[0];
			if (firstPage) {
				firstPage.style.position = 'relative';
				firstPage.style.overflow = 'visible';
				firstPage.appendChild(floatingPanel);
			}

			// Schedule annotation rendering after DOM is ready
			this.later(() => this.renderFloatingAnnotations(floatingPanel, wrapper));
		}

		return wrapper;
	}

	renderFloatingAnnotations(floatingPanel: HTMLElement, wrapper: HTMLElement): void {
		const annotations = Object.values(this.floatingTrackChangeMap)
			.map(entry => entry.annotation)
			.filter(a => a.contentElement);

		if (annotations.length === 0) return;

		// Get the reference point - top of the first page (floating panel's parent)
		const firstPage = floatingPanel.parentElement;
		if (!firstPage) return;
		const pageRect = firstPage.getBoundingClientRect();

		// Sort by vertical position of content relative to page
		const sortedAnnotations = annotations.sort((a, b) => {
			const rectA = a.contentElement.getBoundingClientRect();
			const rectB = b.contentElement.getBoundingClientRect();
			return rectA.top - rectB.top;
		});

		const GAP = 8;
		let lastBottom = -GAP;

		for (const annotation of sortedAnnotations) {
			const annotEl = this.createAnnotationElement(annotation);
			floatingPanel.appendChild(annotEl);
			annotation.element = annotEl;

			// Calculate desired position relative to the page top
			const contentRect = annotation.contentElement.getBoundingClientRect();
			const targetTop = contentRect.top - pageRect.top;
			const desiredTop = Math.max(targetTop, lastBottom + GAP);

			// Push card down with margin-top (normal flow prevents overlap)
			const currentPos = annotEl.offsetTop;
			if (desiredTop > currentPos) {
				annotEl.style.marginTop = `${desiredTop - currentPos}px`;
			}
			lastBottom = annotEl.offsetTop + annotEl.offsetHeight;

			// Click handler - select card and highlight content
			const handleActivate = () => {
				const wasActive = annotEl.classList.contains(`${this.className}-tc-annotation-active`);

				// Deselect all
				floatingPanel.querySelectorAll(`.${this.className}-tc-annotation-active`).forEach(el => {
					el.classList.remove(`${this.className}-tc-annotation-active`);
				});

				if (!wasActive) {
					annotEl.classList.add(`${this.className}-tc-annotation-active`);

					annotation.contentElement?.scrollIntoView({ behavior: 'smooth', block: 'center' });
					annotation.contentElement?.classList.add(`${this.className}-tc-highlight`);
					setTimeout(() => {
						annotation.contentElement?.classList.remove(`${this.className}-tc-highlight`);
					}, 2000);
				}
			};
			annotEl.addEventListener('click', handleActivate);
			// Keyboard accessibility: Enter/Space to activate
			annotEl.addEventListener('keydown', (e: KeyboardEvent) => {
				if (e.key === 'Enter' || e.key === ' ') {
					e.preventDefault();
					handleActivate();
				}
			});
		}
	}

	renderDefaultStyle() {
		var c = this.className;

		// Design System CSS Variables
		var designSystem = `
/* Design System Variables */
:root, .${c}-wrapper {
    --docx-bg-primary: #ffffff;
    --docx-bg-secondary: #f8f9fa;
    --docx-bg-tertiary: #f1f3f4;
    --docx-text-primary: #202124;
    --docx-text-secondary: #5f6368;
    --docx-text-muted: #80868b;
    --docx-border-color: #dadce0;
    --docx-border-light: #e8eaed;

    --docx-color-inserted: #34a853;
    --docx-color-inserted-bg: rgba(52, 168, 83, 0.12);
    --docx-color-inserted-bg-hover: rgba(52, 168, 83, 0.18);
    --docx-color-deleted: #ea4335;
    --docx-color-deleted-bg: rgba(234, 67, 53, 0.12);
    --docx-color-deleted-bg-hover: rgba(234, 67, 53, 0.18);
    --docx-color-moved: #4285f4;
    --docx-color-moved-bg: rgba(66, 133, 244, 0.12);
    --docx-color-moved-bg-hover: rgba(66, 133, 244, 0.18);
    --docx-color-format: #fbbc04;
    --docx-color-format-bg: rgba(251, 188, 4, 0.12);
    --docx-color-format-bg-hover: rgba(251, 188, 4, 0.18);
    --docx-color-comment: #f59e0b;
    --docx-color-comment-bg: rgba(245, 158, 11, 0.12);
    --docx-color-accent: #1a73e8;

    --docx-shadow-sm: 0 1px 2px rgba(60, 64, 67, 0.1);
    --docx-shadow-md: 0 2px 8px rgba(60, 64, 67, 0.15);
    --docx-shadow-lg: 0 4px 12px rgba(60, 64, 67, 0.2);
    --docx-shadow-hover: 0 4px 12px rgba(60, 64, 67, 0.18);

    --docx-radius-sm: 4px;
    --docx-radius-md: 8px;
    --docx-radius-lg: 12px;
    --docx-transition-fast: 0.15s ease;
    --docx-transition-normal: 0.2s ease;
    --docx-transition-slow: 0.3s ease;
}

/* Dark mode variables */
.dark .${c}-wrapper, .${c}-wrapper.dark,
:root.dark .${c}-wrapper {
    --docx-bg-primary: #1e1e1e;
    --docx-bg-secondary: #252526;
    --docx-bg-tertiary: #2d2d30;
    --docx-text-primary: #e8eaed;
    --docx-text-secondary: #9aa0a6;
    --docx-text-muted: #6b7280;
    --docx-border-color: #3c4043;
    --docx-border-light: #4b5563;
    --docx-shadow-sm: 0 1px 3px rgba(0, 0, 0, 0.3);
    --docx-shadow-md: 0 2px 8px rgba(0, 0, 0, 0.4);
    --docx-shadow-lg: 0 4px 12px rgba(0, 0, 0, 0.5);
    --docx-shadow-hover: 0 6px 16px rgba(0, 0, 0, 0.5);
}

/* Reduced motion */
@media (prefers-reduced-motion: reduce) {
    .${c}-wrapper, .${c}-wrapper * {
        transition-duration: 0.01ms !important;
        animation-duration: 0.01ms !important;
    }
}
`;

		var wrapperStyle = `
.${c}-wrapper { background: var(--docx-bg-tertiary, gray); padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; }
.${c}-wrapper>section.${c} { background: var(--docx-bg-primary, white); box-shadow: var(--docx-shadow-lg, 0 0 10px rgba(0, 0, 0, 0.5)); margin-bottom: 30px; }`;
		if (this.options.hideWrapperOnPrint) {
			wrapperStyle = `@media not print { ${wrapperStyle} }`;
		}
		var styleText = `${designSystem}${wrapperStyle}
.${c} { color: var(--docx-text-primary, black); hyphens: auto; text-underline-position: from-font; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>footer { z-index: 1; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c} svg { fill: transparent; }
`;

		if (this.options.renderComments) {
			styleText += `
/* Comment Reference Badge */
.${c}-comment-ref {
    cursor: default;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    vertical-align: middle;
    background: var(--docx-color-comment-bg, #fff9c4);
    border-radius: var(--docx-radius-sm);
    padding: 2px 6px;
    margin-left: 2px;
    transition: background-color var(--docx-transition-fast), transform var(--docx-transition-fast);
}
.${c}-comment-ref:hover {
    background: var(--docx-color-format-bg-hover);
    transform: scale(1.05);
}
.${c}-comment-ref svg { width: 14px; height: 14px; stroke: var(--docx-color-comment); fill: none; }

/* Comment Popover */
.${c}-comment-popover,
.${c}-comment-popover * {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important;
    font-style: normal !important;
    font-size: 11px !important;
    line-height: 1.4 !important;
}
.${c}-comment-popover {
    opacity: 0;
    visibility: hidden;
    transform: translateY(-4px);
    z-index: 1000;
    padding: 12px;
    background: var(--docx-bg-primary);
    position: absolute;
    box-shadow: var(--docx-shadow-lg);
    border: 1px solid var(--docx-border-color);
    border-radius: var(--docx-radius-md);
    width: 280px;
    transition: opacity var(--docx-transition-normal), visibility var(--docx-transition-normal), transform var(--docx-transition-normal);
    pointer-events: none;
}
.${c}-comment-popover::before {
    content: '';
    position: absolute;
    top: -6px;
    left: 16px;
    width: 10px;
    height: 10px;
    background: var(--docx-bg-primary);
    border-left: 1px solid var(--docx-border-color);
    border-top: 1px solid var(--docx-border-color);
    transform: rotate(45deg);
}
.${c}-comment-ref:hover ~ .${c}-comment-popover {
    opacity: 1;
    visibility: visible;
    transform: translateY(0);
    pointer-events: auto;
}
.${c}-comment-popover:hover {
    opacity: 1;
    visibility: visible;
    transform: translateY(0);
    pointer-events: auto;
}
.${c}-comment-author {
    font-size: 11px;
    font-weight: 600;
    color: var(--docx-text-primary);
    margin-bottom: 2px;
}
.${c}-comment-date {
    font-size: 10px;
    color: var(--docx-text-secondary);
    margin-bottom: 4px;
}
`
		}

		const needsAnnotationCSS = this.options.renderChanges
			|| (this.options.renderComments && this.options.commentsMode !== 'inline');
		if (needsAnnotationCSS) {
			// Track changes layout styles
			styleText += `
/* Track Changes Margin - page is expanded to fit annotations */
section.${c}.${c}-has-track-changes {
    position: relative;
}
.${c}-track-changes-margin {
    position: absolute;
    top: 0;
    right: 0;
    height: 100%;
    background: transparent;
    overflow-y: visible;
    box-sizing: border-box;
    padding-top: inherit;
    padding-bottom: inherit;
    border-left: 1px solid var(--docx-border-color);
}
.${c}-track-changes-margin .${c}-tc-annotation {
    position: relative;
    margin-left: 8px;
    margin-right: 8px;
    border: 1px solid var(--docx-border-light);
    border-radius: var(--docx-radius-sm);
    box-shadow: var(--docx-shadow-sm);
}

/* Track Changes Content Markers (margin mode) */
.${c}-tc-content { position: relative; }
.${c}-tc-inserted { background: var(--docx-color-inserted-bg); }
.${c}-tc-deleted { background: var(--docx-color-deleted-bg); text-decoration: line-through; }
.${c}-tc-moveFrom { background: var(--docx-color-moved-bg); text-decoration: line-through; }
.${c}-tc-moveTo { background: var(--docx-color-moved-bg); }
.${c}-tc-formatChange { background: var(--docx-color-format-bg); }

/* Highlight animation on click */
.${c}-tc-highlight {
    outline: 2px solid var(--docx-color-accent);
    outline-offset: 2px;
    animation: ${c}-pulse 0.6s ease-out;
}
@keyframes ${c}-pulse {
    0% { outline-color: var(--docx-color-accent); box-shadow: 0 0 0 0 rgba(26, 115, 232, 0.4); }
    50% { box-shadow: 0 0 0 6px rgba(26, 115, 232, 0); }
    100% { outline-color: var(--docx-color-accent); box-shadow: 0 0 0 0 rgba(26, 115, 232, 0); }
}

/* Inline mode styles */
.${c}-tc-inline-wrapper { position: relative; display: inline; }
.${c}-tc-inline {
    cursor: default;
    transition: background-color var(--docx-transition-fast);
}
.${c}-tc-inline-inserted {
    background: var(--docx-color-inserted-bg);
    border-bottom: 2px solid var(--docx-color-inserted);
}
.${c}-tc-inline-inserted:hover { background: var(--docx-color-inserted-bg-hover); }
.${c}-tc-inline-deleted {
    background: var(--docx-color-deleted-bg);
    text-decoration: line-through;
    border-bottom: 2px solid var(--docx-color-deleted);
}
.${c}-tc-inline-deleted:hover { background: var(--docx-color-deleted-bg-hover); }
.${c}-tc-inline-moveFrom {
    background: var(--docx-color-moved-bg);
    text-decoration: line-through;
    border-bottom: 2px solid var(--docx-color-moved);
}
.${c}-tc-inline-moveFrom:hover { background: var(--docx-color-moved-bg-hover); }
.${c}-tc-inline-moveTo {
    background: var(--docx-color-moved-bg);
    border-bottom: 2px solid var(--docx-color-moved);
}
.${c}-tc-inline-moveTo:hover { background: var(--docx-color-moved-bg-hover); }

/* Inline Popover - Animated */
.${c}-tc-popover,
.${c}-tc-popover * {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important;
    font-style: normal !important;
    font-size: 11px !important;
    line-height: 1.4 !important;
}
.${c}-tc-popover {
    opacity: 0;
    visibility: hidden;
    transform: translateY(-4px);
    z-index: 1000;
    padding: 12px;
    background: var(--docx-bg-primary);
    position: absolute;
    box-shadow: var(--docx-shadow-lg);
    border: 1px solid var(--docx-border-color);
    border-radius: var(--docx-radius-md);
    width: 240px;
    top: 100%;
    left: 0;
    margin-top: 8px;
    transition: opacity var(--docx-transition-normal), visibility var(--docx-transition-normal), transform var(--docx-transition-normal);
    pointer-events: none;
}
/* Arrow indicator */
.${c}-tc-popover::before {
    content: '';
    position: absolute;
    top: -6px;
    left: 16px;
    width: 10px;
    height: 10px;
    background: var(--docx-bg-primary);
    border-left: 1px solid var(--docx-border-color);
    border-top: 1px solid var(--docx-border-color);
    transform: rotate(45deg);
}
.${c}-tc-inline:hover ~ .${c}-tc-popover,
.${c}-tc-popover:hover {
    opacity: 1;
    visibility: visible;
    transform: translateY(0);
    pointer-events: auto;
}

.${c}-tc-popover-header {
    margin-bottom: 2px;
}
.${c}-tc-popover-author {
    font-family: inherit;
    font-size: 11px;
    font-weight: 600;
    color: var(--docx-text-primary);
    margin-bottom: 2px;
}
.${c}-tc-popover-date {
    font-family: inherit;
    font-size: 10px;
    font-weight: 400;
    color: var(--docx-text-secondary);
    display: block;
    margin-bottom: 4px;
}
.${c}-tc-popover-content {
    font-family: inherit;
    font-size: 11px;
    font-weight: 400;
    color: var(--docx-text-secondary);
}
.${c}-tc-popover-type {
    font-family: inherit;
    font-size: 10px;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.3px;
}
.${c}-tc-popover-type-inserted { color: var(--docx-color-inserted) !important; }
.${c}-tc-popover-type-deleted { color: var(--docx-color-deleted) !important; }
.${c}-tc-popover-type-moveFrom, .${c}-tc-popover-type-moveTo { color: var(--docx-color-moved) !important; }
.${c}-tc-popover-type-formatChange { color: var(--docx-color-format) !important; }

/* Track Changes Annotation Cards */
.${c}-tc-annotation,
.${c}-tc-annotation * {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important;
    font-style: normal !important;
    font-size: 11px !important;
    line-height: 1.4 !important;
}
.${c}-tc-annotation {
    background: var(--docx-bg-primary);
    border-left: 3px solid var(--docx-border-light);
    padding: 8px 10px;
    cursor: pointer;
    transition: background-color var(--docx-transition-fast), box-shadow var(--docx-transition-fast);
}
.${c}-tc-annotation:hover {
    background: var(--docx-bg-secondary);
}
.${c}-tc-annotation:focus-visible {
    outline: 2px solid var(--docx-color-accent);
    outline-offset: -2px;
}
/* Colored left border per change type */
.${c}-tc-annotation-inserted { border-left-color: var(--docx-color-inserted); }
.${c}-tc-annotation-deleted { border-left-color: var(--docx-color-deleted); }
.${c}-tc-annotation-moveFrom,
.${c}-tc-annotation-moveTo { border-left-color: var(--docx-color-moved); }
.${c}-tc-annotation-formatChange { border-left-color: var(--docx-color-format); }
.${c}-tc-annotation-comment { border-left-color: var(--docx-color-comment); }
/* Active/selected card */
.${c}-tc-annotation-active {
    background: var(--docx-bg-secondary);
    box-shadow: 0 1px 4px rgba(60, 64, 67, 0.2);
    z-index: 1;
    position: relative;
}

.${c}-comment-marker { border-bottom: 2px solid var(--docx-color-comment); display: inline-block; cursor: pointer; vertical-align: middle; line-height: 1; }
.${c}-comment-marker svg { width: 14px; height: 14px; stroke: var(--docx-color-comment); fill: none; }

/* Annotation card layout */
.${c}-tc-annotation-row {
    display: block;
}

/* Body: header + content */
.${c}-tc-annotation-body {
}

/* Header: author + date inline */
.${c}-tc-annotation-header {
    display: flex;
    align-items: baseline;
    gap: 6px;
    margin-bottom: 2px;
}
.${c}-tc-annotation-author {
    font-family: inherit;
    font-size: 11px !important;
    font-weight: 600;
    color: var(--docx-text-primary);
    white-space: nowrap;
}
.${c}-tc-annotation-date {
    font-family: inherit;
    font-size: 10px !important;
    font-weight: 400;
    color: var(--docx-text-muted);
    white-space: nowrap;
}

/* Content area — clamped to 3 lines by default */
.${c}-tc-annotation-content {
    font-family: inherit;
    font-size: 11px;
    font-weight: 400;
    color: var(--docx-text-secondary);
    line-height: 1.4;
    display: -webkit-box;
    -webkit-line-clamp: 3;
    -webkit-box-orient: vertical;
    overflow: hidden;
}
/* Active card shows full content */
.${c}-tc-annotation-active .${c}-tc-annotation-content {
    -webkit-line-clamp: unset;
    overflow: visible;
}
.${c}-tc-annotation-type {
    font-family: inherit;
    font-size: 10px !important;
    font-weight: 500;
    text-transform: none;
    letter-spacing: 0;
}
.${c}-tc-annotation-inserted .${c}-tc-annotation-type { color: var(--docx-color-inserted); }
.${c}-tc-annotation-deleted .${c}-tc-annotation-type { color: var(--docx-color-deleted); }
.${c}-tc-annotation-moveFrom .${c}-tc-annotation-type,
.${c}-tc-annotation-moveTo .${c}-tc-annotation-type { color: var(--docx-color-moved); }
.${c}-tc-annotation-formatChange .${c}-tc-annotation-type { color: var(--docx-color-format); }
.${c}-tc-annotation-comment .${c}-tc-annotation-type { color: var(--docx-color-comment); }

/* CSS Highlight API styles */
::highlight(${c}-tc-inserted) { background: var(--docx-color-inserted-bg); }
::highlight(${c}-tc-deleted) { background: var(--docx-color-deleted-bg); text-decoration: line-through; }
::highlight(${c}-tc-move-from) { background: var(--docx-color-moved-bg); text-decoration: line-through; }
::highlight(${c}-tc-move-to) { background: var(--docx-color-moved-bg); }
::highlight(${c}-tc-format) { background: var(--docx-color-format-bg); }

/* Floating Panel Mode - positioned like margin mode but persistent across pages */
.${c}-wrapper.${c}-has-floating-panel {
    position: relative;
}
.${c}-wrapper.${c}-has-floating-panel > section.${c}:first-of-type {
    position: relative;
}
.${c}-track-changes-floating {
    position: absolute;
    top: 0;
    left: calc(100% + 16px);
    height: 100%;
    background: transparent;
    padding-left: 12px;
    box-sizing: border-box;
    overflow: hidden;
}
.${c}-track-changes-floating .${c}-tc-annotation {
    position: relative;
    border: 1px solid var(--docx-border-light);
    border-left: 3px solid var(--docx-border-light);
    border-radius: var(--docx-radius-sm);
    box-shadow: var(--docx-shadow-sm);
    margin-bottom: 4px;
}
/* Re-apply colored left border in floating mode */
.${c}-track-changes-floating .${c}-tc-annotation-inserted { border-left-color: var(--docx-color-inserted); }
.${c}-track-changes-floating .${c}-tc-annotation-deleted { border-left-color: var(--docx-color-deleted); }
.${c}-track-changes-floating .${c}-tc-annotation-moveFrom,
.${c}-track-changes-floating .${c}-tc-annotation-moveTo { border-left-color: var(--docx-color-moved); }
.${c}-track-changes-floating .${c}-tc-annotation-formatChange { border-left-color: var(--docx-color-format); }
.${c}-track-changes-floating .${c}-tc-annotation-comment { border-left-color: var(--docx-color-comment); }

/* Accessibility: Focus styles for keyboard navigation */
.${c}-tc-inline:focus-visible {
    outline: 2px solid var(--docx-color-accent);
    outline-offset: 1px;
    border-radius: 2px;
}
.${c}-tc-annotation[tabindex]:focus-visible {
    outline: 2px solid var(--docx-color-accent);
    outline-offset: 2px;
}

/* Annotation panel: scrollable when content overflows page height */
.${c}-track-changes-floating[data-scrollable] {
    overflow-y: auto;
    scrollbar-width: thin;
    scrollbar-color: var(--docx-border-color) transparent;
    -webkit-mask-image: linear-gradient(to bottom,
        transparent 0px, black 20px,
        black calc(100% - 20px), transparent 100%);
    mask-image: linear-gradient(to bottom,
        transparent 0px, black 20px,
        black calc(100% - 20px), transparent 100%);
}
.${c}-track-changes-floating[data-scrollable]::-webkit-scrollbar { width: 4px; }
.${c}-track-changes-floating[data-scrollable]::-webkit-scrollbar-thumb {
    background: var(--docx-border-color);
    border-radius: 2px;
}
`;
		}

		return this.createStyleElement(styleText);
	}

	// renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
	//     let css = "";
	//     const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
	//     const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
	//     const topCounters = [];

	//     for(let num of numberingPart.numberings) {
	//         const absNum = numberingMap[num.abstractId];

	//         for(let lvl of absNum.levels) {
	//             const className = this.numberingClass(num.id, lvl.level);
	//             let listStyleType = "none";

	//             if(lvl.text && lvl.format == 'decimal') {
	//                 const counter = this.numberingCounter(num.id, lvl.level);

	//                 if (lvl.level > 0) {
	//                     css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
	//                         "counter-reset": counter
	//                     });
	//                 } else {
	//                     topCounters.push(counter);
	//                 }

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": this.levelTextToContent(lvl.text, num.id),
	//                     "counter-increment": counter
	//                 });
	//             } else if(lvl.bulletPictureId) {
	//                 let pict = bulletMap[lvl.bulletPictureId];
	//                 let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": "' '",
	//                     "display": "inline-block",
	//                     "background": `var(${variable})`
	//                 }, pict.style);

	//                 this.document.loadNumberingImage(pict.referenceId).then(data => {
	//                     var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
	//                     container.appendChild(createStyleElement(text));
	//                 });
	//             } else {
	//                 listStyleType = this.numFormatToCssValue(lvl.format);
	//             }

	//             css += this.styleToString(`p.${className}`, {
	//                 "display": "list-item",
	//                 "list-style-position": "inside",
	//                 "list-style-type": listStyleType,
	//                 //TODO
	//                 //...num.style
	//             });
	//         }
	//     }

	//     if (topCounters.length > 0) {
	//         css += this.styleToString(`.${this.className}-wrapper`, {
	//             "counter-reset": topCounters.join(" ")
	//         });
	//     }

	//     return createStyleElement(css);
	// }

	renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
		var styleText = "";
		var resetCounters = [];

		for (var num of numberings) {
			var selector = `p.${this.numberingClass(num.id, num.level)}`;
			var listStyleType = "none";

			if (num.bullet) {
				let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				}, num.bullet.style);

				this.tasks.push(this.document.loadNumberingImage(num.bullet.src).then(data => {
					var text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
					styleContainer.appendChild(this.createStyleElement(text));
				}));
			}
			else if (num.levelText) {
				let counter = this.numberingCounter(num.id, num.level);
				const counterReset = counter + " " + (num.start - 1);
				if (num.level > 0) {
					styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
						"counter-set": counterReset
					});
				}
				// reset all level counters with start value
				resetCounters.push(counterReset);

				styleText += this.styleToString(`${selector}:before`, {
					"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
					"counter-increment": counter,
					...num.rStyle,
				});
			}
			else {
				listStyleType = this.numFormatToCssValue(num.format);
			}

			styleText += this.styleToString(selector, {
				"display": "list-item",
				"list-style-position": "inside",
				"list-style-type": listStyleType,
				...num.pStyle
			});
		}

		if (resetCounters.length > 0) {
			styleText += this.styleToString(this.rootSelector, {
				"counter-reset": resetCounters.join(" ")
			});
		}

		return this.createStyleElement(styleText);
	}

	renderStyles(styles: IDomStyle[]): HTMLElement {
		var styleText = "";
		const stylesMap = this.styleMap;
		const defautStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);

		for (const style of styles) {
			var subStyles = style.styles;

			if (style.linked) {
				var linkedStyle = style.linked && stylesMap[style.linked];

				if (linkedStyle)
					subStyles = subStyles.concat(linkedStyle.styles);
				else if (this.options.debug)
					console.warn(`Can't find linked style ${style.linked}`);
			}

			for (const subStyle of subStyles) {
				//TODO temporary disable modificators until test it well
				var selector = `${style.target ?? ''}.${style.cssName}`; //${subStyle.mod ?? ''} 

				if (style.target != subStyle.target)
					selector += ` ${subStyle.target}`;

				if (defautStyles[style.target] == style)
					selector = `.${this.className} ${style.target}, ` + selector;

				styleText += this.styleToString(selector, subStyle.values);
			}
		}

		return this.createStyleElement(styleText);
	}

	renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>, into: HTMLElement) {
		var notes = noteIds.map(id => notesMap[id]).filter(x => x);

		if (notes.length > 0) {
			var result = this.createElement("ol", null, this.renderElements(notes));
			into.appendChild(result);
		}
	}

	renderElement(elem: OpenXmlElement): Node | Node[] {
		switch (elem.type) {
			case DomType.Paragraph:
				return this.renderParagraph(elem as WmlParagraph);

			case DomType.BookmarkStart:
				return this.renderBookmarkStart(elem as WmlBookmarkStart);

			case DomType.BookmarkEnd:
				return null; //ignore bookmark end

			case DomType.Run:
				return this.renderRun(elem as WmlRun);

			case DomType.Table:
				return this.renderTable(elem);

			case DomType.Row:
				return this.renderTableRow(elem);

			case DomType.Cell:
				return this.renderTableCell(elem);

			case DomType.Hyperlink:
				return this.renderHyperlink(elem);
			
			case DomType.SmartTag:
				return this.renderSmartTag(elem);

			case DomType.Drawing:
				return this.renderDrawing(elem);

			case DomType.Image:
				return this.renderImage(elem as IDomImage);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.DeletedText:
				return this.renderDeletedText(elem as WmlText);
	
			case DomType.Tab:
				return this.renderTab(elem);

			case DomType.Symbol:
				return this.renderSymbol(elem as WmlSymbol);

			case DomType.Break:
				return this.renderBreak(elem as WmlBreak);

			case DomType.Footer:
				return this.renderContainer(elem, "footer");

			case DomType.Header:
				return this.renderContainer(elem, "header");

			case DomType.Footnote:
			case DomType.Endnote:
				return this.renderContainer(elem, "li");

			case DomType.FootnoteReference:
				return this.renderFootnoteReference(elem as WmlNoteReference);

			case DomType.EndnoteReference:
				return this.renderEndnoteReference(elem as WmlNoteReference);

			case DomType.NoBreakHyphen:
				return this.createElement("wbr");

			case DomType.VmlPicture:
				return this.renderVmlPicture(elem);

			case DomType.VmlElement:
				return this.renderVmlElement(elem as VmlElement);
	
			case DomType.MmlMath:
				return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });
	
			case DomType.MmlMathParagraph:
				return this.renderContainer(elem, "span");

			case DomType.MmlFraction:
				return this.renderContainerNS(elem, ns.mathML, "mfrac");

			case DomType.MmlBase:
				return this.renderContainerNS(elem, ns.mathML, 
					elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");

			case DomType.MmlNumerator:
			case DomType.MmlDenominator:
			case DomType.MmlFunction:
			case DomType.MmlLimit:
			case DomType.MmlBox:
				return this.renderContainerNS(elem, ns.mathML, "mrow");

			case DomType.MmlGroupChar:
				return this.renderMmlGroupChar(elem);

			case DomType.MmlLimitLower:
				return this.renderContainerNS(elem, ns.mathML, "munder");

			case DomType.MmlMatrix:
				return this.renderContainerNS(elem, ns.mathML, "mtable");

			case DomType.MmlMatrixRow:
				return this.renderContainerNS(elem, ns.mathML, "mtr");
	
			case DomType.MmlRadical:
				return this.renderMmlRadical(elem);

			case DomType.MmlSuperscript:
				return this.renderContainerNS(elem, ns.mathML, "msup");

			case DomType.MmlSubscript:
				return this.renderContainerNS(elem, ns.mathML, "msub");

			case DomType.MmlDegree:
			case DomType.MmlSuperArgument:
			case DomType.MmlSubArgument:
				return this.renderContainerNS(elem, ns.mathML, "mn");

			case DomType.MmlFunctionName:
				return this.renderContainerNS(elem, ns.mathML, "ms");
	
			case DomType.MmlDelimiter:
				return this.renderMmlDelimiter(elem);

			case DomType.MmlRun:
				return this.renderMmlRun(elem);

			case DomType.MmlNary:
				return this.renderMmlNary(elem);

			case DomType.MmlPreSubSuper:
				return this.renderMmlPreSubSuper(elem);

			case DomType.MmlBar:
				return this.renderMmlBar(elem);
	
			case DomType.MmlEquationArray:
				return this.renderMllList(elem);

			case DomType.Inserted:
				return this.renderInserted(elem as WmlTrackChange);

			case DomType.Deleted:
				return this.renderDeleted(elem as WmlTrackChange);

			case DomType.MoveFrom:
				return this.renderMoveFrom(elem as WmlTrackChange);

			case DomType.MoveTo:
				return this.renderMoveTo(elem as WmlTrackChange);

			case DomType.MoveFromRangeStart:
			case DomType.MoveFromRangeEnd:
			case DomType.MoveToRangeStart:
			case DomType.MoveToRangeEnd:
				return null; // Range markers are not rendered directly

			case DomType.CommentRangeStart:
				return this.renderCommentRangeStart(elem);

			case DomType.CommentRangeEnd:
				return this.renderCommentRangeEnd(elem);

			case DomType.CommentReference:
				return this.renderCommentReference(elem);

			case DomType.AltChunk:
				return this.renderAltChunk(elem);
		}

		return null;
	}
	renderElements(elems: OpenXmlElement[], into?: Node): Node[] {
		if (elems == null)
			return null;

		var result = elems.flatMap(e => this.renderElement(e)).filter(e => e != null);

		if (into)
			appendChildren(into, result);

		return result;
	}

	renderContainer<T extends keyof HTMLElementTagNameMap>(elem: OpenXmlElement, tagName: T, props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>): HTMLElementTagNameMap[T] {
		return this.createElement<T>(tagName, props, this.renderElements(elem.children));
	}

	renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
		return this.createElementNS(ns, tagName, props, this.renderElements(elem.children));
	}

	renderParagraph(elem: WmlParagraph) {
		var result = this.renderContainer(elem, "p");

		const style = this.findStyle(elem.styleName);
		elem.tabs ??= style?.paragraphProps?.tabs;  //TODO

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);
		this.renderCommonProperties(result.style, elem);

		const numbering = elem.numbering ?? style?.paragraphProps?.numbering;

		if (numbering) {
			result.classList.add(this.numberingClass(numbering.id, numbering.level));
		}

		// Handle paragraph formatting changes
		if (elem.formatChange && this.options.renderChanges) {
			this.registerFormatChange(elem.formatChange, result);
		}

		return result;
	}

	registerFormatChange(formatChange: WmlTrackChange, element: HTMLElement): void {
		if (this.options.trackChangesMode !== 'margin' && this.options.trackChangesMode !== 'floating') return;

		element.classList.add(`${this.className}-tc-content`, `${this.className}-tc-formatChange`);
		element.dataset.tcId = formatChange.id;

		const annotation: TrackChangeAnnotation = {
			id: formatChange.id,
			author: formatChange.author || 'Unknown',
			date: formatChange.date || '',
			changeType: 'formatChange',
			previewText: '',
			formatDescription: formatChange.formatDescription || 'Formatting changed',
			contentElement: element
		};

		this.trackChangeMap[formatChange.id] = {
			annotation,
			contentElements: [element]
		};

		if (this.trackChangeHighlights) {
			this.later(() => {
				const rng = new Range();
				rng.selectNodeContents(element);
				this.trackChangeHighlights.formatChange.add(rng);
			});
		}
	}

	renderRunProperties(style: any, props: RunProperties) {
		this.renderCommonProperties(style, props);
	}

	renderCommonProperties(style: any, props: CommonProperties) {
		if (props == null)
			return;

		if (props.color) {
			style["color"] = props.color;
		}

		if (props.fontSize) {
			style["font-size"] = props.fontSize;
		}
	}

	renderHyperlink(elem: WmlHyperlink) {
		var result = this.renderContainer(elem, "a");

		this.renderStyleValues(elem.cssStyle, result);

		let href = '';

		if (elem.id) {
			const rel = this.document.documentPart.rels.find(it => it.id == elem.id && it.targetMode === "External");
			href = rel?.target ?? href;
		}

		if (elem.anchor) {
			href += `#${elem.anchor}`;
		}

		result.href = href;

		return result;
	}
	
	renderSmartTag(elem: WmlSmartTag) {
		return this.renderContainer(elem, "span");
	}
	
	renderCommentRangeStart(commentStart: WmlCommentRangeStart) {
		if (!this.options.renderComments)
			return null;

		const rng = new Range();
		this.commentHighlight?.add(rng);

		const result = this.createComment(`start of comment #${commentStart.id}`);
		this.later(() => rng.setStart(result, 0));
		this.commentMap[commentStart.id] = rng;

		return result
	}

	renderCommentRangeEnd(commentEnd: WmlCommentRangeStart) {
		if (!this.options.renderComments)
			return null;

		const rng = this.commentMap[commentEnd.id];
		const result = this.createComment(`end of comment #${commentEnd.id}`);
		this.later(() => rng?.setEnd(result, 0));

		return result;
	}

	renderCommentReference(commentRef: WmlCommentReference) {
		if (!this.options.renderComments)
			return null;

		var comment = this.document.commentsPart?.commentMap[commentRef.id];

		if (!comment)
			return null;

		// Margin or floating mode - render comment in the sidebar/panel
		if (this.options.commentsMode === 'margin' || this.options.commentsMode === 'floating') {
			return this.renderCommentWithMargin(comment);
		}

		// Inline mode - render as popover
		const frg = new DocumentFragment();
		const commentRefEl = this.createElement("span", { className: `${this.className}-comment-ref` });
		// Lucide message-square icon as inline SVG
		commentRefEl.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg>`;
		const commentsContainerEl = this.createElement("div", { className: `${this.className}-comment-popover` });

		this.renderCommentContent(comment, commentsContainerEl);

		frg.appendChild(this.createComment(`comment #${comment.id} by ${comment.author} on ${comment.date}`));
		frg.appendChild(commentRefEl);
		frg.appendChild(commentsContainerEl);

		return frg;
	}

	renderCommentWithMargin(comment: WmlComment): Node {
		// Create a marker span for the comment location
		const marker = this.createElement("span", {
			className: `${this.className}-comment-marker`
		});
		marker.dataset.commentId = comment.id;
		marker.dataset.tcId = `comment-${comment.id}`;
		// Lucide message-square icon so the anchor point is visible
		marker.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg>`;

		// Extract comment text for preview
		let commentText = '';
		const extractText = (el: any) => {
			if (el.text) commentText += el.text;
			if (el.children) el.children.forEach(extractText);
		};
		comment.children?.forEach(extractText);

		// Register for margin annotation
		const annotation: TrackChangeAnnotation = {
			id: `comment-${comment.id}`,
			author: comment.author || 'Unknown',
			date: comment.date || '',
			changeType: 'comment',
			previewText: commentText.substring(0, 100) || '(no text)',
			contentElement: marker
		};

		this.trackChangeMap[`comment-${comment.id}`] = {
			annotation,
			contentElements: [marker]
		};

		return marker;
	}

	renderAltChunk(elem: WmlAltChunk) {
		if (!this.options.renderAltChunks)
			return null;

		var result = this.createElement("iframe");
		
		this.tasks.push(this.document.loadAltChunk(elem.id, this.currentPart).then(x => {
			result.srcdoc = x;
		}));

		return result;
	}

	renderCommentContent(comment: WmlComment, container: Node) {
		container.appendChild(this.createElement('div', { className: `${this.className}-comment-author` }, [comment.author]));
		container.appendChild(this.createElement('div', { className: `${this.className}-comment-date` }, [new Date(comment.date).toLocaleString()]));

		this.renderElements(comment.children, container);
	}

	renderDrawing(elem: OpenXmlElement) {
		var result = this.renderContainer(elem, "div");

		result.style.display = "inline-block";
		result.style.position = "relative";
		result.style.textIndent = "0px";

		this.renderStyleValues(elem.cssStyle, result);

		return result;
	}

	renderImage(elem: IDomImage) {
		let result = this.createElement("img");
		let transform = elem.cssStyle?.transform;

		this.renderStyleValues(elem.cssStyle, result);

		if (elem.srcRect && elem.srcRect.some(x => x != 0)) {
			var [left, top, right, bottom] = elem.srcRect;
			transform = `scale(${1 / (1 - left - right)}, ${1 / (1 - top - bottom)})`;
			result.style['clip-path'] = `rect(${(100 * top).toFixed(2)}% ${(100 * (1 - right)).toFixed(2)}% ${(100 * (1 - bottom)).toFixed(2)}% ${(100 * left).toFixed(2)}%)`;
		}

		if (elem.rotation)
			transform = `rotate(${elem.rotation}deg) ${transform ?? ''}`;

		result.style.transform = transform?.trim();

		// Hide images that fail to load (e.g. WMF/EMF formats unsupported by browsers)
		result.onerror = () => {
			result.style.display = 'none';
			// Also hide the drawing wrapper div that has explicit dimensions
			if (result.parentElement) {
				result.parentElement.style.display = 'none';
			}
		};

		if (this.document) {
			this.tasks.push(this.document.loadDocumentImage(elem.src, this.currentPart).then(x => {
				result.src = x;
			}));
		}

		return result;
	}

	renderText(elem: WmlText) {
		return this.htmlDocument.createTextNode(elem.text);
	}

	renderDeletedText(elem: WmlText) {
		return this.options.renderChanges ? this.renderText(elem) : null;
	}

	renderBreak(elem: WmlBreak) {
		if (elem.break == "textWrapping") {
			return this.createElement("br");
		}

		return null;
	}

	renderInserted(elem: WmlTrackChange): Node | Node[] {
		if (!this.options.renderChanges) {
			return this.renderElements(elem.children);
		}

		if (this.options.trackChangesMode === 'margin' || this.options.trackChangesMode === 'floating') {
			return this.renderTrackChangeWithMargin(elem, 'inserted');
		}

		return this.renderTrackChangeInline(elem, 'inserted');
	}

	renderDeleted(elem: WmlTrackChange): Node | Node[] {
		if (!this.options.renderChanges) {
			return null;
		}

		if (this.options.trackChangesMode === 'margin' || this.options.trackChangesMode === 'floating') {
			return this.renderTrackChangeWithMargin(elem, 'deleted');
		}

		return this.renderTrackChangeInline(elem, 'deleted');
	}

	renderMoveFrom(elem: WmlTrackChange): Node | Node[] {
		if (!this.options.renderChanges) {
			return null;
		}

		if (this.options.trackChangesMode === 'margin' || this.options.trackChangesMode === 'floating') {
			return this.renderTrackChangeWithMargin(elem, 'moveFrom');
		}

		return this.renderTrackChangeInline(elem, 'moveFrom');
	}

	renderMoveTo(elem: WmlTrackChange): Node | Node[] {
		if (!this.options.renderChanges) {
			return this.renderElements(elem.children);
		}

		if (this.options.trackChangesMode === 'margin' || this.options.trackChangesMode === 'floating') {
			return this.renderTrackChangeWithMargin(elem, 'moveTo');
		}

		return this.renderTrackChangeInline(elem, 'moveTo');
	}

	renderTrackChangeInline(elem: WmlTrackChange, changeType: TrackChangeType): Node {
		const typeLabels: Record<string, string> = {
			'inserted': 'Inserted',
			'deleted': 'Deleted',
			'moveFrom': 'Moved from',
			'moveTo': 'Moved to'
		};

		// Create wrapper with position relative for popover
		const wrapper = this.createElement("span", {
			className: `${this.className}-tc-inline-wrapper`
		});
		wrapper.dataset.tcId = elem.id;

		// Render the content with styling
		const contentWrapper = this.createElement("span", {
			className: `${this.className}-tc-inline ${this.className}-tc-inline-${changeType}`
		});
		this.renderElements(elem.children, contentWrapper);
		wrapper.appendChild(contentWrapper);

		// Create popover (shows on hover over content)
		const popover = this.createElement("div", {
			className: `${this.className}-tc-popover`
		});

		// Header: author + date inline
		const header = this.createElement('div', {
			className: `${this.className}-tc-popover-header`
		});

		header.appendChild(this.createElement('span', {
			className: `${this.className}-tc-popover-author`
		}, [elem.author || 'Unknown']));

		if (elem.date) {
			try {
				const d = new Date(elem.date);
				header.appendChild(this.createElement('span', {
					className: `${this.className}-tc-popover-date`
				}, [d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })]));
			} catch {
				header.appendChild(this.createElement('span', {
					className: `${this.className}-tc-popover-date`
				}, [elem.date]));
			}
		}

		popover.appendChild(header);

		// Content: type + preview
		const content = this.createElement('div', {
			className: `${this.className}-tc-popover-content`
		});

		const typeSpan = this.createElement('span', {
			className: `${this.className}-tc-popover-type ${this.className}-tc-popover-type-${changeType}`
		}, [typeLabels[changeType] + ': ']);
		content.appendChild(typeSpan);

		const previewText = this.extractPreviewText(elem);
		content.appendChild(this.htmlDocument.createTextNode(previewText));

		popover.appendChild(content);
		wrapper.appendChild(popover);

		return wrapper;
	}

	renderTrackChangeWithMargin(elem: WmlTrackChange, changeType: TrackChangeType): Node[] {
		const children = this.renderElements(elem.children);
		if (!children || children.length === 0) return children;

		// Create wrapper span for the content
		const wrapper = this.createElement("span", {
			className: `${this.className}-tc-content ${this.className}-tc-${changeType}`
		});
		wrapper.dataset.tcId = elem.id;

		appendChildren(wrapper, children);

		// Extract preview text for the annotation
		const previewText = this.extractPreviewText(elem);

		// Register for margin annotation
		const annotation: TrackChangeAnnotation = {
			id: elem.id,
			author: elem.author || 'Unknown',
			date: elem.date || '',
			changeType: changeType,
			previewText: previewText,
			formatDescription: elem.formatDescription,
			contentElement: wrapper
		};

		this.trackChangeMap[elem.id] = {
			annotation,
			contentElements: [wrapper]
		};

		// Add to CSS Highlight API if available
		if (this.trackChangeHighlights) {
			this.later(() => {
				const rng = new Range();
				rng.selectNodeContents(wrapper);

				switch (changeType) {
					case 'inserted':
						this.trackChangeHighlights.inserted.add(rng);
						break;
					case 'deleted':
						this.trackChangeHighlights.deleted.add(rng);
						break;
					case 'moveFrom':
						this.trackChangeHighlights.moveFrom.add(rng);
						break;
					case 'moveTo':
						this.trackChangeHighlights.moveTo.add(rng);
						break;
					case 'formatChange':
						this.trackChangeHighlights.formatChange.add(rng);
						break;
				}
			});
		}

		return [wrapper];
	}

	extractPreviewText(elem: OpenXmlElement, maxLength: number = 50): string {
		let text = '';

		const extractFromElement = (el: OpenXmlElement) => {
			if (text.length >= maxLength) return;

			if (el.type === DomType.Text || el.type === DomType.DeletedText) {
				text += (el as WmlText).text;
			} else if (el.children) {
				for (const child of el.children) {
					extractFromElement(child);
					if (text.length >= maxLength) break;
				}
			}
		};

		extractFromElement(elem);

		if (text.length > maxLength) {
			text = text.substring(0, maxLength) + '...';
		}

		return text || '(empty)';
	}

	renderMarginAnnotationsFromMap(marginContainer: HTMLElement, contentWrapper: HTMLElement, trackChangeMap: Record<string, TrackChangeEntry>): void {
		const annotations = Object.values(trackChangeMap)
			.map(entry => entry.annotation)
			.filter(a => a.contentElement);

		if (annotations.length === 0) return;

		// Sort by vertical position of content
		const sortedAnnotations = annotations.sort((a, b) => {
			const rectA = a.contentElement.getBoundingClientRect();
			const rectB = b.contentElement.getBoundingClientRect();
			return rectA.top - rectB.top;
		});

		const containerRect = marginContainer.getBoundingClientRect();
		const GAP = 8;
		let lastBottom = -GAP;

		for (const annotation of sortedAnnotations) {
			const contentRect = annotation.contentElement.getBoundingClientRect();
			const targetTop = contentRect.top - containerRect.top;

			// Create annotation element
			const annotEl = this.createAnnotationElement(annotation);
			marginContainer.appendChild(annotEl);

			// Calculate desired position (avoid overlaps)
			const desiredTop = Math.max(targetTop, lastBottom + GAP);

			// Push card down with margin-top (normal flow prevents overlap)
			const currentPos = annotEl.offsetTop;
			if (desiredTop > currentPos) {
				annotEl.style.marginTop = `${desiredTop - currentPos}px`;
			}
			lastBottom = annotEl.offsetTop + annotEl.offsetHeight;
			annotation.element = annotEl;

			// Add click handler to highlight content
			const handleActivate = () => {
				annotation.contentElement?.scrollIntoView({ behavior: 'smooth', block: 'center' });
				annotation.contentElement?.classList.add(`${this.className}-tc-highlight`);
				setTimeout(() => {
					annotation.contentElement?.classList.remove(`${this.className}-tc-highlight`);
				}, 2000);
			};
			annotEl.addEventListener('click', handleActivate);
			// Keyboard accessibility: Enter/Space to activate
			annotEl.addEventListener('keydown', (e: KeyboardEvent) => {
				if (e.key === 'Enter' || e.key === ' ') {
					e.preventDefault();
					handleActivate();
				}
			});
		}
	}

	createAnnotationElement(annotation: TrackChangeAnnotation): HTMLElement {
		const typeLabels: Record<TrackChangeType, string> = {
			'inserted': 'Inserted',
			'deleted': 'Deleted',
			'moveFrom': 'Moved from',
			'moveTo': 'Moved to',
			'formatChange': 'Formatted',
			'comment': 'Comment'
		};

		const el = this.createElement("div", {
			className: `${this.className}-tc-annotation ${this.className}-tc-annotation-${annotation.changeType}`
		});

		// Link annotation card to its content element via data attribute
		el.dataset.tcId = annotation.id;

		// Accessibility: make annotation focusable and interactive
		el.setAttribute('role', 'button');
		el.setAttribute('tabindex', '0');
		el.setAttribute('aria-label', `${typeLabels[annotation.changeType]} by ${annotation.author}: ${annotation.previewText}`);

		// Row container: avatar + body
		const row = this.createElement("div", {
			className: `${this.className}-tc-annotation-row`
		});

		// Body: header + content
		const body = this.createElement("div", {
			className: `${this.className}-tc-annotation-body`
		});

		// Header: author + date
		const header = this.createElement("div", {
			className: `${this.className}-tc-annotation-header`
		});

		const authorEl = this.createElement("span", {
			className: `${this.className}-tc-annotation-author`
		});
		authorEl.textContent = annotation.author;
		header.appendChild(authorEl);

		if (annotation.date) {
			const dateEl = this.createElement("span", {
				className: `${this.className}-tc-annotation-date`
			});
			try {
				dateEl.textContent = this.formatRelativeDate(new Date(annotation.date));
			} catch {
				dateEl.textContent = annotation.date;
			}
			header.appendChild(dateEl);
		}

		body.appendChild(header);

		// Content: change type + preview
		const content = this.createElement("div", {
			className: `${this.className}-tc-annotation-content`
		});

		if (annotation.changeType === 'comment') {
			content.textContent = annotation.previewText;
		} else if (annotation.changeType === 'formatChange' && annotation.formatDescription) {
			const typeSpan = this.createElement("span", {
				className: `${this.className}-tc-annotation-type`
			});
			typeSpan.textContent = typeLabels[annotation.changeType] + ': ';
			content.appendChild(typeSpan);
			content.appendChild(this.htmlDocument.createTextNode(annotation.formatDescription));
		} else {
			const typeSpan = this.createElement("span", {
				className: `${this.className}-tc-annotation-type`
			});
			typeSpan.textContent = typeLabels[annotation.changeType] + ': ';
			content.appendChild(typeSpan);
			content.appendChild(this.htmlDocument.createTextNode(annotation.previewText));
		}

		body.appendChild(content);
		row.appendChild(body);
		el.appendChild(row);

		return el;
	}

	formatRelativeDate(date: Date): string {
		const now = new Date();
		const diffMs = now.getTime() - date.getTime();
		const diffSec = Math.floor(diffMs / 1000);
		const diffMin = Math.floor(diffSec / 60);
		const diffHr = Math.floor(diffMin / 60);
		const diffDay = Math.floor(diffHr / 24);

		if (diffSec < 60) return 'just now';
		if (diffMin < 60) return `${diffMin}m ago`;
		if (diffHr < 24) return `${diffHr}h ago`;
		if (diffDay < 7) return `${diffDay}d ago`;
		return date.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
	}

	renderSymbol(elem: WmlSymbol) {
		var span = this.createElement("span");
		span.style.fontFamily = elem.font;
		span.innerHTML = `&#x${elem.char};`
		return span;
	}

	renderFootnoteReference(elem: WmlNoteReference) {
		var result = this.createElement("sup");
		this.currentFootnoteIds.push(elem.id);
		result.textContent = `${this.currentFootnoteIds.length}`;
		return result;
	}

	renderEndnoteReference(elem: WmlNoteReference) {
		var result = this.createElement("sup");
		this.currentEndnoteIds.push(elem.id);
		result.textContent = `${this.currentEndnoteIds.length}`;
		return result;
	}

	renderTab(elem: OpenXmlElement) {
		var tabSpan = this.createElement("span");

		tabSpan.innerHTML = "&emsp;";//"&nbsp;";

		if (this.options.experimental) {
			tabSpan.className = this.tabStopClass();
			var stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
			this.currentTabs.push({ stops, span: tabSpan });
		}

		return tabSpan;
	}

	renderBookmarkStart(elem: WmlBookmarkStart): HTMLElement {
		return this.createElement("span", { id: elem.name });
	}

	renderRun(elem: WmlRun) {
		if (elem.fieldRun)
			return null;

		const result = this.createElement("span");

		if (elem.id)
			result.id = elem.id;

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.verticalAlign) {
			const wrapper = this.createElement(elem.verticalAlign as any);
			this.renderElements(elem.children, wrapper);
			result.appendChild(wrapper);
		}
		else {
			this.renderElements(elem.children, result);
		}

		// Handle run formatting changes
		if (elem.formatChange && this.options.renderChanges) {
			this.registerFormatChange(elem.formatChange, result);
		}

		return result;
	}

	renderTable(elem: WmlTable) {
		let result = this.createElement("table");

		this.tableCellPositions.push(this.currentCellPosition);
		this.tableVerticalMerges.push(this.currentVerticalMerge);
		this.currentVerticalMerge = {};
		this.currentCellPosition = { col: 0, row: 0 };

		if (elem.columns)
			result.appendChild(this.renderTableColumns(elem.columns));

		this.renderClass(elem, result);
		this.renderElements(elem.children, result);
		this.renderStyleValues(elem.cssStyle, result);

		this.currentVerticalMerge = this.tableVerticalMerges.pop();
		this.currentCellPosition = this.tableCellPositions.pop();

		return result;
	}

	renderTableColumns(columns: WmlTableColumn[]) {
		let result = this.createElement("colgroup");

		for (let col of columns) {
			let colElem = this.createElement("col");

			if (col.width)
				colElem.style.width = col.width;

			result.appendChild(colElem);
		}

		return result;
	}

	renderTableRow(elem: WmlTableRow) {
		let result = this.createElement("tr");

		this.currentCellPosition.col = 0;

		if (elem.gridBefore)
			result.appendChild(this.renderTableCellPlaceholder(elem.gridBefore));

		this.renderClass(elem, result);
		this.renderElements(elem.children, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.gridAfter)
			result.appendChild(this.renderTableCellPlaceholder(elem.gridAfter));

		this.currentCellPosition.row++;

		return result;
	}

	renderTableCellPlaceholder(colSpan: number) {
		const result = this.createElement("td", { colSpan })
		result.style['border'] = 'none';
		return result;
	}

	renderTableCell(elem: WmlTableCell) {
		let result = this.renderContainer(elem, "td");

		const key = this.currentCellPosition.col;

		if (elem.verticalMerge) {
			if (elem.verticalMerge == "restart") {
				this.currentVerticalMerge[key] = result;
				result.rowSpan = 1;
			} else if (this.currentVerticalMerge[key]) {
				this.currentVerticalMerge[key].rowSpan += 1;
				result.style.display = "none";
			}
		} else {
			this.currentVerticalMerge[key] = null;
		}

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		if (elem.span)
			result.colSpan = elem.span;

		this.currentCellPosition.col += result.colSpan;

		return result;
	}

	renderVmlPicture(elem: OpenXmlElement) {
		return this.renderContainer(elem, "div");
	}

	renderVmlElement(elem: VmlElement): SVGElement {
		var container = this.createSvgElement("svg");

		container.setAttribute("style", elem.cssStyleText);
		container.setAttribute("overflow", "hidden");

		const result = this.renderVmlChildElement(elem);

		// Clear fill on rect text boxes — white fill creates visible artifacts
		// when the SVG extends past the page boundary (section overflow:visible)
		if (elem.tagName === "rect" && elem.children.some(
			c => c.type == DomType.VmlElement && (c as VmlElement).tagName === "foreignObject"
		)) {
			result.setAttribute("fill", "none");
		}

		if (elem.imageHref?.id) {
			this.tasks.push(this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
				.then(x => result.setAttribute("href", x)));
		}

		container.appendChild(result);

		requestAnimationFrame(() => {
			const bb = (container.firstElementChild as any).getBBox();

			container.setAttribute("width", `${Math.ceil(bb.x +  bb.width)}`);
			container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
		});

		return container;
	}

	renderVmlChildElement(elem: VmlElement): any {
		const result = this.createSvgElement(elem.tagName as any);
		Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));

		for (let child of elem.children) {
			if (child.type == DomType.VmlElement) {
				result.appendChild(this.renderVmlChildElement(child as VmlElement));
			} else {
				result.appendChild(...asArray(this.renderElement(child as any)));
			}
		}

		return result;
	}

	renderMmlRadical(elem: OpenXmlElement): HTMLElement {
		const base = elem.children.find(el => el.type == DomType.MmlBase);

		if (elem.props?.hideDegree) {
			return this.createElementNS(ns.mathML, "msqrt", null, this.renderElements([base]));
		}

		const degree = elem.children.find(el => el.type == DomType.MmlDegree);
		return this.createElementNS(ns.mathML, "mroot", null, this.renderElements([base, degree]));
	}

	renderMmlDelimiter(elem: OpenXmlElement): HTMLElement {		
		const children = [];

		children.push(this.createElementNS(ns.mathML, "mo", null, [elem.props.beginChar ?? '(']));
		children.push(...this.renderElements(elem.children));
		children.push(this.createElementNS(ns.mathML, "mo", null, [elem.props.endChar ?? ')']));

		return this.createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlNary(elem: OpenXmlElement): HTMLElement {		
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;

		const charElem = this.createElementNS(ns.mathML, "mo", null, [elem.props?.char ?? '\u222B']);

		if (supElem || subElem) {
			children.push(this.createElementNS(ns.mathML, "munderover", null, [charElem, subElem, supElem]));
		} else if(supElem) {
			children.push(this.createElementNS(ns.mathML, "mover", null, [charElem, supElem]));
		} else if(subElem) {
			children.push(this.createElementNS(ns.mathML, "munder", null, [charElem, subElem]));
		} else {
			children.push(charElem);
		}

		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlPreSubSuper(elem: OpenXmlElement) {
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createElementNS(ns.mathML, "mo", null, asArray(this.renderElement(sub))) : null;
		const stubElem = this.createElementNS(ns.mathML, "mo", null);

		children.push(this.createElementNS(ns.mathML, "msubsup", null, [stubElem, subElem, supElem]));
		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createElementNS(ns.mathML, "mrow", null, children);
	}

	renderMmlGroupChar(elem: OpenXmlElement) {
		const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
		const result = this.renderContainerNS(elem, ns.mathML, tagName);

		if (elem.props.char) {
			result.appendChild(this.createElementNS(ns.mathML, "mo", null, [elem.props.char]));
		}

		return result;
	}

	renderMmlBar(elem: OpenXmlElement) {
		const result = this.renderContainerNS(elem, ns.mathML, "mrow");

		switch(elem.props.position) {
			case "top": result.style.textDecoration = "overline"; break
			case "bottom": result.style.textDecoration = "underline"; break
		}

		return result;
	}

	renderMmlRun(elem: OpenXmlElement) {
		const result = this.createElementNS(ns.mathML, "ms", null, this.renderElements(elem.children));

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		return result;
	}

	renderMllList(elem: OpenXmlElement) {
		const result = this.createElementNS(ns.mathML, "mtable");

		this.renderClass(elem, result);
		this.renderStyleValues(elem.cssStyle, result);

		for (let child of this.renderElements(elem.children)) {
			result.appendChild(this.createElementNS(ns.mathML, "mtr", null, [
				this.createElementNS(ns.mathML, "mtd", null, [child])
			]));
		}

		return result;
	}


	renderStyleValues(style: Record<string, string>, ouput: HTMLElement) {
		for (let k in style) {
			if (k.startsWith("$")) {
				ouput.setAttribute(k.slice(1), style[k]);
			} else {
				ouput.style[k] = style[k];
			}
		}
	}

	renderClass(input: OpenXmlElement, ouput: HTMLElement) {
		if (input.className)
			ouput.className = input.className;

		if (input.styleName)
			ouput.classList.add(this.processStyleName(input.styleName));
	}

	findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	numberingClass(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	tabStopClass() {
		return `${this.className}-tab-stop`;
	}

	styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
		let result = `${selectors} {\r\n`;

		for (const key in values) {
			if (key.startsWith('$'))
				continue;
			
			result += `  ${key}: ${values[key]};\r\n`;
		}

		if (cssText)
			result += cssText;

		return result + "}\r\n";
	}

	numberingCounter(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		const suffMap = {
			"tab": "\\9",
			"space": "\\a0",
		};

		var result = text.replace(/%\d*/g, s => {
			let lvl = parseInt(s.substring(1), 10) - 1;
			return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
		});

		return `"${result}${suffMap[suff] ?? ""}"`;
	}

	numFormatToCssValue(format: string) {
		var mapping = {
			none: "none",
			bullet: "disc",
			decimal: "decimal",
			lowerLetter: "lower-alpha",
			upperLetter: "upper-alpha",
			lowerRoman: "lower-roman",
			upperRoman: "upper-roman",
			decimalZero: "decimal-leading-zero", // 01,02,03,...
			// ordinal: "", // 1st, 2nd, 3rd,...
			// ordinalText: "", //First, Second, Third, ...
			// cardinalText: "", //One,Two Three,...
			// numberInDash: "", //-1-,-2-,-3-, ...
			// hex: "upper-hexadecimal",
			aiueo: "katakana",
			aiueoFullWidth: "katakana",
			chineseCounting: "simp-chinese-informal",
			chineseCountingThousand: "simp-chinese-informal",
			chineseLegalSimplified: "simp-chinese-formal", // 中文大写
			chosung: "hangul-consonant",
			ideographDigital: "cjk-ideographic",
			ideographTraditional: "cjk-heavenly-stem", // 十天干
			ideographLegalTraditional: "trad-chinese-formal",
			ideographZodiac: "cjk-earthly-branch", // 十二地支
			iroha: "katakana-iroha",
			irohaFullWidth: "katakana-iroha",
			japaneseCounting: "japanese-informal",
			japaneseDigitalTenThousand: "cjk-decimal",
			japaneseLegal: "japanese-formal",
			thaiNumbers: "thai",
			koreanCounting: "korean-hangul-formal",
			koreanDigital: "korean-hangul-formal",
			koreanDigital2: "korean-hanja-informal",
			hebrew1: "hebrew",
			hebrew2: "hebrew",
			hindiNumbers: "devanagari",
			ganada: "hangul",
			taiwaneseCounting: "cjk-ideographic",
			taiwaneseCountingThousand: "cjk-ideographic",
			taiwaneseDigital:  "cjk-decimal",
		};

		return mapping[format] ?? format;
	}

	refreshTabStops() {
		if (!this.options.experimental)
			return;

		setTimeout(() => {
			const pixelToPoint = computePixelToPoint();

			for (let tab of this.currentTabs) {
				updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
			}
		}, 500);
	}

	createElementNS(ns: string, tagName: string, props?: Partial<Record<any, any>>, children?: ChildType[]): any {
		var result = ns ? this.htmlDocument.createElementNS(ns, tagName) : this.htmlDocument.createElement(tagName);
		Object.assign(result, props);
		children && appendChildren(result, children);
		return result;
	}

	createElement<T extends keyof HTMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>, children?: ChildType[]): HTMLElementTagNameMap[T] {
		return this.createElementNS(undefined, tagName, props, children);
	}

	createSvgElement<T extends keyof SVGElementTagNameMap>(tagName: T, props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>, children?: ChildType[]): SVGElementTagNameMap[T] {
		return this.createElementNS(ns.svg, tagName, props, children);
	}

	createStyleElement(cssText: string) {
		return this.createElement("style", { innerHTML: cssText });
	}
	
	createComment(text: string) {
		return this.htmlDocument.createComment(text);
	}

	later(func: Function) {
		this.postRenderTasks.push(func);
	}

	// ── Post-render page splitting ──────────────────────────────────────

	/**
	 * Walk rendered sections and split any that exceed their intended
	 * page height (minHeight) into multiple pages. Redistributes
	 * annotation cards to the correct pages and recalculates positions.
	 */
	splitOverflowingPages(bodyContainer: HTMLElement): void {
		const wrapper = bodyContainer.querySelector(`.${this.className}-wrapper`);
		if (!wrapper) return;

		const sections = Array.from(
			wrapper.querySelectorAll(`:scope > section.${this.className}`)
		) as HTMLElement[];
		let i = 0;

		while (i < sections.length) {
			const newSection = this.splitSectionIfNeeded(sections[i], wrapper);
			if (newSection) {
				sections.splice(i + 1, 0, newSection);
			}
			i++;
		}

		// After all splitting, redistribute cards to correct sections, then recalc positions
		this.redistributeFloatingCards(wrapper as HTMLElement);
		this.recalcAllAnnotations(wrapper as HTMLElement);
	}

	private splitSectionIfNeeded(section: HTMLElement, wrapper: Element): HTMLElement | null {
		const article = section.querySelector(':scope > article') as HTMLElement;
		if (!article || article.children.length <= 1) return null;

		const cs = getComputedStyle(section);
		const intendedHeight = parseFloat(cs.minHeight);
		if (!intendedHeight || isNaN(intendedHeight)) return null;

		const actualHeight = section.getBoundingClientRect().height;
		if (actualHeight <= intendedHeight + 1) return null;

		const paddingTop = parseFloat(cs.paddingTop) || 0;
		const paddingBottom = parseFloat(cs.paddingBottom) || 0;

		// Account for non-article, in-flow siblings (headers, footers)
		let nonArticleHeight = 0;
		for (const child of Array.from(section.children)) {
			if (child === article) continue;
			const pos = getComputedStyle(child as HTMLElement).position;
			if (pos === 'absolute' || pos === 'fixed') continue;
			nonArticleHeight += (child as HTMLElement).getBoundingClientRect().height;
		}

		const availableHeight = intendedHeight - paddingTop - paddingBottom - nonArticleHeight;
		if (availableHeight <= 0) return null;

		const children = Array.from(article.children) as HTMLElement[];
		const articleTop = article.getBoundingClientRect().top;
		let breakIndex = -1;

		for (let i = 0; i < children.length; i++) {
			const childRect = children[i].getBoundingClientRect();
			const childBottom = childRect.bottom - articleTop;
			if (childBottom > availableHeight && i > 0) {
				breakIndex = i;
				break;
			}
		}

		if (breakIndex <= 0) return null;

		const newSection = section.cloneNode(false) as HTMLElement;
		const newArticle = article.cloneNode(false) as HTMLElement;

		const overflow = children.slice(breakIndex);
		for (const child of overflow) {
			newArticle.appendChild(child);
		}
		newSection.appendChild(newArticle);

		// Handle margin aside
		const marginAside = section.querySelector(
			`:scope > .${this.className}-track-changes-margin`
		) as HTMLElement;
		if (marginAside) {
			const newAside = marginAside.cloneNode(false) as HTMLElement;
			newAside.innerHTML = '';
			newSection.appendChild(newAside);
			this.distributeAnnotationCards(marginAside, newAside, newArticle);
		}

		// Handle floating panel
		const floatingPanel = section.querySelector(
			`:scope > .${this.className}-track-changes-floating`
		) as HTMLElement;
		if (floatingPanel) {
			const newPanel = floatingPanel.cloneNode(false) as HTMLElement;
			newPanel.innerHTML = '';
			newSection.appendChild(newPanel);
			newSection.style.position = 'relative';
			this.distributeAnnotationCards(floatingPanel, newPanel, newArticle);
		}

		wrapper.insertBefore(newSection, section.nextSibling);
		return newSection;
	}

	/** Move annotation cards whose content element is in newArticle */
	private distributeAnnotationCards(
		originalPanel: HTMLElement,
		newPanel: HTMLElement,
		newArticle: HTMLElement
	): void {
		const cards = Array.from(originalPanel.querySelectorAll('[data-tc-id]')) as HTMLElement[];
		for (const card of cards) {
			const tcId = card.dataset.tcId;
			if (!tcId) continue;
			if (newArticle.querySelector(`[data-tc-id="${CSS.escape(tcId)}"]`)) {
				card.style.marginTop = '';
				newPanel.appendChild(card);
			}
		}
	}

	/**
	 * After page splitting, redistribute annotation cards so each card
	 * lives in the section that actually contains its content element.
	 * Cards from later DOCX sections (section breaks) may have been placed
	 * on the first section's panel; this moves them to the correct one.
	 */
	private redistributeFloatingCards(wrapper: HTMLElement): void {
		const c = this.className;
		const sections = Array.from(
			wrapper.querySelectorAll(`:scope > section.${c}`)
		) as HTMLElement[];

		// Collect all floating cards from all panels
		const allCards: HTMLElement[] = [];
		for (const section of sections) {
			const panel = section.querySelector(`:scope > .${c}-track-changes-floating`);
			if (!panel) continue;
			allCards.push(...Array.from(panel.querySelectorAll(':scope > [data-tc-id]')) as HTMLElement[]);
		}

		for (const card of allCards) {
			const tcId = card.dataset.tcId;
			if (!tcId) continue;

			// Find which section contains this card's content element
			let targetSection: HTMLElement | null = null;
			for (const section of sections) {
				if (section.querySelector(`article [data-tc-id="${CSS.escape(tcId)}"]`)) {
					targetSection = section;
					break;
				}
			}
			if (!targetSection) continue;

			// Get or create floating panel on the target section
			let targetPanel = targetSection.querySelector(
				`:scope > .${c}-track-changes-floating`
			) as HTMLElement;
			if (!targetPanel) {
				const existingPanel = wrapper.querySelector(`.${c}-track-changes-floating`) as HTMLElement;
				if (!existingPanel) continue;
				targetPanel = existingPanel.cloneNode(false) as HTMLElement;
				targetPanel.innerHTML = '';
				targetSection.appendChild(targetPanel);
				targetSection.style.position = 'relative';
				targetSection.style.overflow = 'visible';
			}

			if (card.parentElement !== targetPanel) {
				card.style.marginTop = '';
				targetPanel.appendChild(card);
			}
		}
	}

	// ── Annotation position recalculation with overflow handling ────────

	private recalcAllAnnotations(wrapper: HTMLElement): void {
		const panels = wrapper.querySelectorAll(
			`.${this.className}-track-changes-margin, .${this.className}-track-changes-floating`
		);
		for (const panel of Array.from(panels)) {
			this.recalcPanelPositions(panel as HTMLElement);
		}
	}

	/** Walk offsetParent chain from el up to ancestor for zoom-independent Y offset */
	private offsetTopRelativeTo(el: HTMLElement, ancestor: HTMLElement): number {
		let y = 0;
		let cur: HTMLElement | null = el;
		while (cur && cur !== ancestor) {
			y += cur.offsetTop;
			cur = cur.offsetParent as HTMLElement | null;
		}
		return y;
	}

	/**
	 * Pure math — compute card top positions without touching the DOM.
	 * Used by the density gate to trial-layout before committing.
	 */
	private computePositions(
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
	 * Position cards with marginTop so each aligns with its content element
	 * (or stacks below the previous card). Uses offsetTop/offsetHeight
	 * (zoom-independent). Returns the bottom of the last card.
	 */
	private reposWithTargets(
		cards: HTMLElement[],
		panel: HTMLElement,
		targets: number[],
		gap: number
	): number {
		// Remove all cards first, then re-append in sorted order.
		// This avoids inconsistent intermediate DOM states when cards
		// are reordered (appending to end while unprocessed cards remain).
		for (const card of cards) card.remove();

		let nextTop = 0;

		for (let i = 0; i < cards.length; i++) {
			const card = cards[i];
			card.style.marginTop = '';
			card.style.marginBottom = '0px'; // Override CSS margin-bottom to prevent collapse
			panel.appendChild(card);

			const desiredTop = Math.max(targets[i], nextTop);
			const naturalTop = card.offsetTop;
			const delta = desiredTop - naturalTop;
			if (Math.abs(delta) > 1) {
				card.style.marginTop = `${delta}px`;
			}
			const actualTop = Math.abs(delta) > 1 ? desiredTop : naturalTop;
			nextTop = actualTop + card.offsetHeight + gap;
		}

		const last = cards[cards.length - 1];
		return last.offsetTop + last.offsetHeight;
	}

	/**
	 * Find the index of the first card whose bottom exceeds available space
	 * (reserving room for the expand button). Always returns at least 1.
	 */
	private findCutoffIndex(
		cards: HTMLElement[],
		availableHeight: number,
		buttonHeight: number
	): number {
		const limit = availableHeight - buttonHeight;
		for (let i = 0; i < cards.length; i++) {
			const bottom = cards[i].offsetTop + cards[i].offsetHeight;
			if (bottom > limit && i > 0) return i;
		}
		return cards.length;
	}

	private recalcPanelPositions(panel: HTMLElement): void {
		const c = this.className;

		// Clean up previous state
		panel.removeAttribute('data-overflow');
		panel.removeAttribute('data-scrollable');
		panel.style.overflowY = '';
		panel.style.maxHeight = '';

		const cards = Array.from(
			panel.querySelectorAll(`:scope > [data-tc-id]`)
		) as HTMLElement[];
		if (cards.length === 0) return;

		for (const card of cards) {
			card.style.display = '';
			card.style.marginTop = '';
		}

		const section = panel.closest(`section.${c}`) as HTMLElement;
		if (!section) return;

		// Panel is position:absolute — its clientHeight grows with content.
		// Use the section's height minus the panel's vertical padding as the true constraint.
		const panelStyle = getComputedStyle(panel);
		const panelPaddingY = (parseFloat(panelStyle.paddingTop) || 0)
			+ (parseFloat(panelStyle.paddingBottom) || 0);
		const availableHeight = section.clientHeight - panelPaddingY;
		if (!availableHeight) return;

		// Sort cards by content element vertical position
		cards.sort((a, b) => {
			const aEl = section.querySelector(
				`article [data-tc-id="${CSS.escape(a.dataset.tcId!)}"]`
			) as HTMLElement;
			const bEl = section.querySelector(
				`article [data-tc-id="${CSS.escape(b.dataset.tcId!)}"]`
			) as HTMLElement;
			if (!aEl || !bEl) return 0;
			return this.offsetTopRelativeTo(aEl, section)
				- this.offsetTopRelativeTo(bEl, section);
		});

		// Compute target positions for each card (content-aligned)
		const panelOffset = this.offsetTopRelativeTo(panel, section);
		const targets: number[] = cards.map(card => {
			const tcId = card.dataset.tcId;
			if (!tcId) return 0;
			const contentEl = section.querySelector(
				`article [data-tc-id="${CSS.escape(tcId)}"]`
			) as HTMLElement;
			if (!contentEl) return 0;
			return this.offsetTopRelativeTo(contentEl, section) - panelOffset;
		});

		const cardHeights = cards.map(c => c.offsetHeight);
		const zeroTargets = targets.map(() => 0);

		let useTargets: number[];

		if (cards.length <= 3) {
			// Rule 1: few cards — always content-aligned
			useTargets = targets;
		} else {
			const tight = this.computePositions(cardHeights, zeroTargets, ANNOTATION_GAP);
			const density = tight.lastBottom / availableHeight;

			if (density > 0.6) {
				// Rule 2: dense — tight stack
				useTargets = zeroTargets;
			} else {
				// Rule 3: try content-aligned
				const aligned = this.computePositions(cardHeights, targets, ANNOTATION_GAP);
				useTargets = aligned.lastBottom > availableHeight ? zeroTargets : targets;
			}
		}

		const lastBottom = this.reposWithTargets(cards, panel, useTargets, ANNOTATION_GAP);

		if (lastBottom > availableHeight) {
			panel.setAttribute('data-scrollable', '');
			panel.style.overflowY = 'auto';
			panel.style.maxHeight = `${availableHeight}px`;
		}
	}
}

type ChildType = Node | string;

function removeAllElements(elem: HTMLElement) {
	elem.innerHTML = '';
}

function appendChildren(elem: Node, children: (Node | string)[]) {
	children.forEach(c => elem.appendChild(isString(c) ? document.createTextNode(c) : c));
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
	var parent = elem.parent;

	while (parent != null && parent.type != type)
		parent = parent.parent;

	return <T>parent;
}

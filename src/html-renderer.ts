import { WordDocument } from './word-document';
import {
	DomType, WmlTable, IDomNumbering,
	WmlHyperlink, IDomImage, OpenXmlElement, WmlTableColumn, WmlTableCell, WmlText, WmlSymbol, WmlBreak, WmlNoteReference,
	WmlSmartTag,
	WmlAltChunk,
	WmlTableRow,
	IDomChart,
	IDomShape
} from './document/dom';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { asArray, encloseFontFamily, escapeClassName, isString, keyBy, mergeDeep, parseCssRules } from './utils';
import { computePixelToPoint, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties } from './document/section';
import { WmlRun } from './document/run';
import { WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { BaseHeaderFooterPart } from './header-footer/parts';
import { Part } from './common/part';
import { VmlElement } from './vml/vml';
import { WmlComment, WmlCommentRangeStart, WmlCommentReference } from './comments/elements';
import { cx, h, ns } from './html';
import { ChartPart } from './chart/chart-part';
import { ChartData, ChartSeriesData } from './chart/chart';

interface CellPos {
	col: number;
	row: number;
}

interface Section {
	sectProps: SectionProperties;
	elements: OpenXmlElement[];
	pageBreak: boolean;
}

declare const Highlight: any;

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

const defaultChartPalette = [
	"#4472C4",
	"#ED7D31",
	"#A5A5A5",
	"#FFC000",
	"#5B9BD5",
	"#70AD47",
	"#264478",
	"#9E480E",
	"#636363",
	"#997300",
];

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
	vmlTextPathIndex: number = 0;

	tasks: Promise<any>[] = [];
	postRenderTasks: any[] = [];
	h = h;

	async render(document: WordDocument, options: Options): Promise<Node[]> {
		this.document = document;
		this.options = options;
		this.className = options.className;
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		this.h = options.h ?? h;
		this.styleMap = null;
		this.tasks = [];

		if (this.options.renderComments && globalThis.Highlight) {
			this.commentHighlight = new Highlight();
		}

		const result: Node[] = [...this.renderDefaultStyle()];

		if (document.themePart) {
			result.push(...this.renderTheme(document.themePart));
		}

		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);
			result.push(...this.renderStyles(document.stylesPart.styles));
		}

		if (document.numberingPart) {
			this.prodessNumberings(document.numberingPart.domNumberings);

			result.push(...await this.renderNumbering(document.numberingPart.domNumberings));
			//result.push(...await this.renderNumbering2(document.numberingPart.domNumberings));
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
			result.push(...await this.renderFontTable(document.fontTablePart));

		var sectionElements = this.renderSections(document.documentPart.body);

		if (this.options.inWrapper) {
			result.push(this.renderWrapper(sectionElements));
		} else {
			result.push(...sectionElements);
		}

		if (this.commentHighlight && options.renderComments) {
			(CSS as any).highlights.set(`${this.className}-comments`, this.commentHighlight);
		}

		this.postRenderTasks.forEach(t => t());

		await Promise.allSettled(this.tasks);

		this.refreshTabStops();

		return result;
	}

	renderTheme(themePart: ThemePart) {
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
		return [
			this.h({ tagName: "#comment", children: ["docxjs document theme values"] }),
			this.h({ tagName: "style", children: [cssText] })
		];
	}

	async renderFontTable(fontsPart: FontTablePart) {
		const result = [];

		for (let f of fontsPart.fonts) {
			for (let ref of f.embedFontRefs) {
				try {
					const fontData = await this.document.loadFont(ref.id, ref.key);
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

					result.push(this.h({ tagName: "#comment", children: [`docxjs ${f.name} font`] }));
					result.push(this.h({ tagName: "style", children: [this.styleToString(`@font-face`, cssValues)] }));
				} catch (e) {
					if (this.options.debug) console.warn(`Can't load font with id ${ref.id} and key ${ref.key}`);
				}
			}
		}

		return result;
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

	createPageElement(className: string, props: SectionProperties, docStyle: Record<string, any>) {
		const style: Record<string, string> = { ...docStyle };

		if (props) {
			if (props.pageMargins) {
				style.paddingLeft = props.pageMargins.left;
				style.paddingRight = props.pageMargins.right;
				style.paddingTop = props.pageMargins.top;
				style.paddingBottom = props.pageMargins.bottom;
			}

			if (props.pageSize) {
				if (!this.options.ignoreWidth)
					style.width = props.pageSize.width;
				if (!this.options.ignoreHeight)
					style.minHeight = props.pageSize.height;
			}
		}

		return this.h({ tagName: "section", className, style }) as HTMLElement;
	}

	createSectionContent(props: SectionProperties) {
		const style: Record<string, string> = {};

		if (props.columns && props.columns.numberOfColumns) {
			style.columnCount = `${props.columns.numberOfColumns}`;
			style.columnGap = props.columns.space;

			if (props.columns.separator) {
				style.columnRule = "1px solid black";
			}
		}

		return this.h({ tagName: "article", style });
	}

	renderSections(document: DocumentElement): HTMLElement[] {
		const result = [];

		this.processElement(document);
		const sections = this.splitBySection(document.children, document.props);
		const pages = this.groupByPageBreaks(sections);
		let prevProps = null;

		for (let i = 0, l = pages.length; i < l; i++) {
			this.currentFootnoteIds = [];

			const section = pages[i][0];
			let props = section.sectProps;
			const pageElement = this.createPageElement(this.className, props, document.cssStyle);

			this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs, props,
				result.length, prevProps != props, pageElement);

			for (const sect of pages[i]) {
				var contentElement = this.createSectionContent(sect.sectProps);
				this.renderElements(sect.elements, contentElement);
				pageElement.appendChild(contentElement);
				props = sect.sectProps;
			}

			if (this.options.renderFootnotes) {
				const notes = this.renderNotes(this.currentFootnoteIds, this.footnoteMap);
				notes && pageElement.appendChild(notes);
			}

			if (this.options.renderEndnotes && i == l - 1) {
				const notes = this.renderNotes(this.currentEndnoteIds, this.endnoteMap);
				notes && pageElement.appendChild(notes);
			}

			this.options.renderFooters && this.renderHeaderFooter(props.footerRefs, props,
				result.length, prevProps != props, pageElement);

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
			current.push(s);

			if (this.options.ignoreLastRenderedPageBreak || s.pageBreak || this.isPageBreakSection(prev, s.sectProps))
				result.push(current = []);

			prev = s.sectProps;
		}

		return result.filter(x => x.length > 0);
	}

	renderWrapper(children: HTMLElement[]) {
		return this.h({ tagName: "div", className: `${this.className}-wrapper`, children });
	}

	renderDefaultStyle() {
		var c = this.className;
		var wrapperStyle = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }`;
		if (this.options.hideWrapperOnPrint) {
			wrapperStyle = `@media not print { ${wrapperStyle} }`;
		}
		var styleText = `${wrapperStyle}
.${c} { color: black; hyphens: auto; text-underline-position: from-font; }
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
.${c}-comment-ref { cursor: default; }
.${c}-comment-popover { display: none; z-index: 1000; padding: 0.5rem; background: white; position: absolute; box-shadow: 0 0 0.25rem rgba(0, 0, 0, 0.25); width: 30ch; }
.${c}-comment-ref:hover~.${c}-comment-popover { display: block; }
.${c}-comment-author,.${c}-comment-date { font-size: 0.875rem; color: #888; }
`
		};

		return [
			this.h({ tagName: "#comment", children: ["docxjs library predefined styles"] }),
			this.h({ tagName: "style", children: [styleText] })
		];
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

	async renderNumbering(numberings: IDomNumbering[]) {
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

				try {
					const imgData = await this.document.loadNumberingImage(num.bullet.src);
					styleText += `${this.rootSelector} { ${valiable}: url(${imgData}) }`;
				} catch (e) {
					if (this.options.debug) console.warn(`Can't load numbering image with src ${num.bullet.src}`);
				}
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

		return [
			this.h({ tagName: "#comment", children: ["docxjs document numbering styles"] }),
			this.h({ tagName: "style", children: [styleText] })
		];
	}

	renderStyles(styles: IDomStyle[]) {
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

		return [
			this.h({ tagName: "#comment", children: ["docxjs document styles"] }),
			this.h({ tagName: "style", children: [styleText] })
		];
	}

	renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>) {
		var notes = noteIds.map(id => notesMap[id]).filter(x => x);

		if (notes.length > 0) {
			return this.h({ tagName: "ol", children: this.renderElements(notes) });
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

			case DomType.Chart:
				return this.renderChart(elem as IDomChart);

			case DomType.Shape:
				return this.renderShape(elem as IDomShape);

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
				return this.h({ tagName: "wbr" });

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
				return this.renderInserted(elem);

			case DomType.Deleted:
				return this.renderDeleted(elem);

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
			result.forEach(c => into.appendChild(isString(c) ? document.createTextNode(c) : c));

		return result;
	}

	renderContainer<T extends keyof HTMLElementTagNameMap>(elem: OpenXmlElement, tagName: T): HTMLElementTagNameMap[T] {
		return this.h({ tagName, children: this.renderElements(elem.children) }) as any;
	}

	renderContainerNS(elem: OpenXmlElement, ns: ns, tagName: string, props?: Record<string, any>) {
		return this.h({ ns, tagName, children: this.renderElements(elem.children), ...props });
	}

	renderParagraph(elem: WmlParagraph) {
		var result = this.toHTML(elem, ns.html, "p");

		const style = this.findStyle(elem.styleName);
		elem.tabs ??= style?.paragraphProps?.tabs;  //TODO

		const numbering = elem.numbering ?? style?.paragraphProps?.numbering;

		if (numbering) {
			result.classList.add(this.numberingClass(numbering.id, numbering.level));
		}

		return result;
	}

	renderHyperlink(elem: WmlHyperlink) {
		const res = this.toH(elem, ns.html, "a");
		res.href = '';

		if (elem.id) {
			const rel = this.document.documentPart.rels.find(it => it.id == elem.id && it.targetMode === "External");
			res.href = rel?.target ?? res.href;
		}

		if (elem.anchor) {
			res.href += `#${elem.anchor}`;
		}

		return this.h(res);
	}

	renderSmartTag(elem: WmlSmartTag) {
		return this.renderContainer(elem, "span");
	}

	renderCommentRangeStart(commentStart: WmlCommentRangeStart) {
		if (!this.options.renderComments)
			return null;

		const rng = new Range();
		this.commentHighlight?.add(rng);

		const result = this.h({ tagName: "#comment", children: [`start of comment #${commentStart.id}`] });
		this.later(() => rng.setStart(result, 0));
		this.commentMap[commentStart.id] = rng;

		return result
	}

	renderCommentRangeEnd(commentEnd: WmlCommentRangeStart) {
		if (!this.options.renderComments)
			return null;

		const rng = this.commentMap[commentEnd.id];
		const result = this.h({ tagName: "#comment", children: [`end of comment #${commentEnd.id}`] });
		this.later(() => rng?.setEnd(result, 0));

		return result;
	}

	renderCommentReference(commentRef: WmlCommentReference) {
		if (!this.options.renderComments)
			return null;

		var comment = this.document.commentsPart?.commentMap[commentRef.id];

		if (!comment)
			return null;

		const commentRefEl = this.h({ tagName: "span", className: `${this.className}-comment-ref`, children: ['💬'] });
		const commentsContainerEl = this.h({
			tagName: "div", className: `${this.className}-comment-popover`, children: [
				this.h({ tagName: 'div', className: `${this.className}-comment-author`, children: [comment.author] }),
				this.h({ tagName: 'div', className: `${this.className}-comment-date`, children: [new Date(comment.date).toLocaleString()] }),
				...this.renderElements(comment.children)
			]
		});

		return this.h({
			tagName: "#fragment", children: [
				this.h({ tagName: "#comment", children: [`comment #${comment.id} by ${comment.author} on ${comment.date}`] }),
				commentRefEl,
				commentsContainerEl
			]
		});
	}

	renderAltChunk(elem: WmlAltChunk) {
		if (!this.options.renderAltChunks)
			return null;

		var result = this.h({ tagName: "iframe" }) as HTMLIFrameElement;

		this.tasks.push(this.document.loadAltChunk(elem.id, this.currentPart).then(x => {
			result.srcdoc = x;
		}));

		return result;
	}

	renderDrawing(elem: OpenXmlElement) {
		var result = this.toHTML(elem, ns.html, "div");

		result.style.display = "inline-block";
		result.style.position = "relative";
		result.style.textIndent = "0px";

		return result;
	}

	renderImage(elem: IDomImage) {
		let result = this.toHTML(elem, ns.html, "img", []);
		let transform = elem.cssStyle?.transform;

		if (elem.srcRect && elem.srcRect.some(x => x != 0)) {
			var [left, top, right, bottom] = elem.srcRect;
			transform = `scale(${1 / (1 - left - right)}, ${1 / (1 - top - bottom)})`;
			result.style['clip-path'] = `rect(${(100 * top).toFixed(2)}% ${(100 * (1 - right)).toFixed(2)}% ${(100 * (1 - bottom)).toFixed(2)}% ${(100 * left).toFixed(2)}%)`;
		}

		if (elem.rotation)
			transform = `rotate(${elem.rotation}deg) ${transform ?? ''}`;

		result.style.transform = transform?.trim();

		if (this.document) {
			this.tasks.push(this.document.loadDocumentImage(elem.src, this.currentPart).then(x => {
				result.src = x;
			}));
		}

		return result;
	}

	// 渲染图表引用，优先输出常见图表 SVG，复杂或未支持类型回退为数据表避免空白
	renderChart(elem: IDomChart) {
		var basePart = this.currentPart ?? this.document.documentPart; // 当前图表所属部件
		var chartPart = this.document.findPartByRelId(elem.id, basePart) as ChartPart; // 图表部件对象
		var chartData = chartPart?.chart; // 图表解析结果
		var result = this.h({
			tagName: "div",
			className: cx(elem.className, `${this.className}-chart`),
			style: {
				boxSizing: "border-box",
				display: "flex",
				flexDirection: "column",
				gap: "8px",
				color: "inherit",
				position: "absolute",
				left: "0px",
				top: "0px",
				width: "100%",
				height: "100%",
				padding: "6px 8px",
				overflow: "hidden",
				...elem.cssStyle
			}
		}) as HTMLDivElement;

		if (!chartData) {
			result.appendChild(this.renderChartNote("图表数据未找到，当前仅保留图表容器。"));
			return result;
		}

		if (chartData.title) {
			result.appendChild(this.h({
				tagName: "div",
				className: `${this.className}-chart-title`,
				style: {
					fontSize: "11pt",
					fontWeight: "600",
					lineHeight: "1.35",
					textAlign: "center"
				},
				children: [chartData.title]
			}));
		}

		var visualNode = this.h({
			tagName: "div",
			className: `${this.className}-chart-visual`,
			style: {
				flex: "1 1 auto",
				minHeight: chartData.title ? "96px" : "120px"
			}
		}) as HTMLDivElement;
		var fallbackReason = this.getChartFallbackReason(chartData); // 当前图表回退原因

		if (fallbackReason) {
			visualNode.appendChild(this.renderChartTable(chartData));
			result.appendChild(visualNode);
			result.appendChild(this.renderChartNote(fallbackReason));
			return result;
		}

		visualNode.appendChild(this.renderChartSvg(chartData));
		result.appendChild(visualNode);

		var legendNode = this.renderChartLegend(chartData); // 图例节点

		if (legendNode)
			result.appendChild(legendNode);

		return result;
	}

	getChartFallbackReason(chart: ChartData): string {
		if (!chart)
			return "图表数据缺失。";

		if ((chart.plotTypes?.length ?? 0) > 1)
			return `组合图表暂未完整绘制，已回退显示数据表：${chart.plotTypes.join(" + ")}。`;

		if (!chart.series?.length)
			return "图表缺少可渲染的系列数据。";

		switch (chart.type) {
			case "barChart":
			case "lineChart":
			case "areaChart":
			case "pieChart":
			case "doughnutChart":
			case "scatterChart":
				return null;

			default:
				return `暂未完整支持 ${chart.type}，已回退显示图表数据表。`;
		}
	}

	renderChartSvg(chart: ChartData): SVGElement {
		switch (chart.type) {
			case "barChart":
				return this.renderBarChartSvg(chart);

			case "lineChart":
				return this.renderLineLikeChartSvg(chart, false);

			case "areaChart":
				return this.renderLineLikeChartSvg(chart, true);

			case "pieChart":
				return this.renderPieChartSvg(chart, false);

			case "doughnutChart":
				return this.renderPieChartSvg(chart, true);

			case "scatterChart":
				return this.renderScatterChartSvg(chart);

			default:
				return this.createChartSvgRoot();
		}
	}

	renderChartLegend(chart: ChartData): HTMLElement {
		var legendItems = this.getChartLegendItems(chart); // 图例项集合

		if (!chart.legendPosition || legendItems.length === 0)
			return null;

		return this.h({
			tagName: "div",
			className: `${this.className}-chart-legend`,
			style: {
				display: "flex",
				flexWrap: "wrap",
				justifyContent: "center",
				gap: "6px 14px",
				fontSize: "9pt",
				lineHeight: "1.3"
			},
			children: legendItems.map(item => this.h({
				tagName: "span",
				className: `${this.className}-chart-legend-item`,
				style: {
					display: "inline-flex",
					alignItems: "center",
					gap: "6px"
				},
				children: [
					this.h({
						tagName: "span",
						className: `${this.className}-chart-swatch`,
						style: {
							backgroundColor: item.color,
							width: "10px",
							height: "10px",
							borderRadius: "2px",
							flex: "0 0 auto"
						}
					}),
					item.label
				]
			}))
		}) as HTMLDivElement;
	}

	getChartLegendItems(chart: ChartData): { label: string; color: string }[] {
		if (!chart)
			return [];

		if ((chart.type === "pieChart" || chart.type === "doughnutChart") && chart.series.length > 0) {
			var series = chart.series[0]; // 饼图只读取首个系列
			var categories = this.getChartCategories(chart, series.values?.length ?? 0); // 饼图分类标签

			return categories.map((label, index) => ({
				label,
				color: this.getChartPointColor(series, index),
			}));
		}

		return chart.series.map((series, index) => ({
			label: series.name || `系列 ${index + 1}`,
			color: this.getChartSeriesColor(series, index),
		}));
	}

	renderChartNote(message: string): HTMLElement {
		return this.h({
			tagName: "div",
			className: `${this.className}-chart-note`,
			style: {
				fontSize: "9pt",
				color: "#666666",
				lineHeight: "1.35"
			},
			children: [message]
		}) as HTMLDivElement;
	}

	renderChartTable(chart: ChartData): HTMLTableElement {
		var tableNode = this.h({
			tagName: "table",
			className: `${this.className}-chart-table`,
			style: {
				width: "100%",
				borderCollapse: "collapse",
				fontSize: "9pt",
				lineHeight: "1.35"
			}
		}) as HTMLTableElement; // 图表回退数据表
		var headerNode = tableNode.createTHead().insertRow(); // 表头行
		var bodyNode = tableNode.createTBody(); // 表体节点

		if (chart.type === "scatterChart") {
			headerNode.appendChild(this.h({ tagName: "th", children: ["点位"] }));

			chart.series.forEach((series, index) => {
				headerNode.appendChild(this.h({ tagName: "th", children: [`${series.name || `系列 ${index + 1}`} X`] }));
				headerNode.appendChild(this.h({ tagName: "th", children: [`${series.name || `系列 ${index + 1}`} Y`] }));
			});

			var pointCount = chart.series.reduce((maxValue, series) => Math.max(maxValue, series.values?.length ?? 0, series.xValues?.length ?? 0), 0); // 散点图最大点数

			for (let pointIndex = 0; pointIndex < pointCount; pointIndex++) {
				var rowNode = bodyNode.insertRow(); // 当前散点数据行
				rowNode.appendChild(this.h({ tagName: "th", children: [`${pointIndex + 1}`] }));

				chart.series.forEach(series => {
					rowNode.appendChild(this.h({ tagName: "td", children: [this.formatChartValue(series.xValues?.[pointIndex])] }));
					rowNode.appendChild(this.h({ tagName: "td", children: [this.formatChartValue(series.values?.[pointIndex])] }));
				});
			}

			return tableNode;
		}

		var categories = this.getChartCategories(chart); // 通用分类标签
		var rowCount = Math.max(categories.length, ...chart.series.map(series => series.values?.length ?? 0)); // 数据行数量

		headerNode.appendChild(this.h({ tagName: "th", children: [chart.categoryAxisTitle || "分类"] }));
		chart.series.forEach((series, index) => {
			headerNode.appendChild(this.h({
				tagName: "th",
				children: [series.name || `系列 ${index + 1}`]
			}));
		});

		for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
			var rowNode = bodyNode.insertRow(); // 当前数据行
			var categoryLabel = categories[rowIndex] ?? `数据 ${rowIndex + 1}`; // 当前分类标签
			rowNode.appendChild(this.h({ tagName: "th", children: [categoryLabel] }));

			chart.series.forEach(series => {
				rowNode.appendChild(this.h({
					tagName: "td",
					children: [this.formatChartValue(series.values?.[rowIndex])]
				}));
			});
		}

		Array.from(tableNode.querySelectorAll("th, td")).forEach(cellNode => {
			var isHeaderCell = cellNode.tagName.toLowerCase() === "th"; // 是否表头单元格
			var tableCellNode = cellNode as HTMLTableCellElement; // 当前表格单元格节点
			(cellNode as HTMLElement).style.border = "1px solid #BFBFBF";
			(cellNode as HTMLElement).style.padding = "2pt 4pt";
			(cellNode as HTMLElement).style.textAlign = isHeaderCell && tableCellNode.cellIndex === 0 ? "left" : "center";
			(cellNode as HTMLElement).style.verticalAlign = "middle";
		});

		Array.from(tableNode.querySelectorAll("th:first-child, td:first-child")).forEach(cellNode => {
			(cellNode as HTMLElement).style.textAlign = "left";
		});

		return tableNode;
	}

	renderBarChartSvg(chart: ChartData): SVGElement {
		var svgNode = this.createChartSvgRoot(); // 柱状图 SVG 根节点
		var categories = this.getChartCategories(chart); // 分类标签集合
		var seriesList = chart.series.map(series => ({ ...series, values: this.normalizeChartSeriesValues(series.values, categories.length) })); // 归一化系列数据
		var values = seriesList.flatMap(series => series.values).filter(value => Number.isFinite(value)); // 所有数值集合
		var hasData = categories.length > 0 && values.length > 0; // 是否存在有效数据

		if (!hasData)
			return svgNode;

		var isHorizontal = chart.barDirection === "bar"; // 是否横向条形图
		var isStacked = chart.grouping === "stacked" || chart.grouping === "percentStacked"; // 是否堆积柱图
		var usesPercentScale = chart.grouping === "percentStacked"; // 是否百分比堆积
		var negativeSeriesExists = seriesList.some(series => series.values.some(value => value < 0)); // 是否包含负值

		if (isStacked && negativeSeriesExists)
			return svgNode;

		if (usesPercentScale) {
			for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
				var categoryTotal = seriesList.reduce((sumValue, series) => sumValue + Math.max(series.values[categoryIndex] ?? 0, 0), 0); // 当前分类总值

				seriesList.forEach(series => {
					series.values[categoryIndex] = categoryTotal > 0
						? (Math.max(series.values[categoryIndex] ?? 0, 0) / categoryTotal) * 100
						: 0;
				});
			}
		}

		var normalizedValues = seriesList.flatMap(series => series.values).filter(value => Number.isFinite(value)); // 归一化后的数值
		var minValue = Math.min(0, ...normalizedValues); // 数值最小值
		var maxValue = Math.max(...normalizedValues, usesPercentScale ? 100 : 0); // 数值最大值
		var safeMaxValue = maxValue === minValue ? maxValue + 1 : maxValue; // 安全最大值
		var viewWidth = 800; // SVG 视口宽度
		var viewHeight = 520; // SVG 视口高度
		var margin = { top: 24, right: 26, bottom: chart.categoryAxisTitle ? 110 : 88, left: chart.valueAxisTitle ? 92 : 74 }; // 图表边距
		var plotWidth = viewWidth - margin.left - margin.right; // 绘图区宽度
		var plotHeight = viewHeight - margin.top - margin.bottom; // 绘图区高度
		var categoryBand = (isHorizontal ? plotHeight : plotWidth) / Math.max(categories.length, 1); // 单个分类带宽
		var gridCount = 5; // 网格线数量
		var valueToPosition = (value: number) => isHorizontal
			? margin.left + ((value - minValue) / (safeMaxValue - minValue)) * plotWidth
			: margin.top + plotHeight - ((value - minValue) / (safeMaxValue - minValue)) * plotHeight; // 数值坐标换算
		var zeroPosition = valueToPosition(0); // 零轴位置

		for (let gridIndex = 0; gridIndex <= gridCount; gridIndex++) {
			var gridValue = minValue + ((safeMaxValue - minValue) / gridCount) * gridIndex; // 当前网格值
			var labelNode = this.createSvgElement("text"); // 当前网格标签
			var lineNode = this.createSvgElement("line"); // 当前网格线

			if (isHorizontal) {
				var xPosition = valueToPosition(gridValue); // 当前网格横坐标
				lineNode.setAttribute("x1", `${xPosition}`);
				lineNode.setAttribute("y1", `${margin.top}`);
				lineNode.setAttribute("x2", `${xPosition}`);
				lineNode.setAttribute("y2", `${margin.top + plotHeight}`);
				labelNode.setAttribute("x", `${xPosition}`);
				labelNode.setAttribute("y", `${margin.top + plotHeight + 20}`);
				labelNode.setAttribute("text-anchor", "middle");
			} else {
				var yPosition = valueToPosition(gridValue); // 当前网格纵坐标
				lineNode.setAttribute("x1", `${margin.left}`);
				lineNode.setAttribute("y1", `${yPosition}`);
				lineNode.setAttribute("x2", `${margin.left + plotWidth}`);
				lineNode.setAttribute("y2", `${yPosition}`);
				labelNode.setAttribute("x", `${margin.left - 10}`);
				labelNode.setAttribute("y", `${yPosition + 4}`);
				labelNode.setAttribute("text-anchor", "end");
			}

			lineNode.setAttribute("stroke", "#D9D9D9");
			lineNode.setAttribute("stroke-width", "1");
			labelNode.setAttribute("font-size", "14");
			labelNode.setAttribute("fill", "#666666");
			labelNode.textContent = this.formatChartValue(gridValue);
			svgNode.appendChild(lineNode);
			svgNode.appendChild(labelNode);
		}

		var axisNode = this.createSvgElement("line"); // 零轴线
		axisNode.setAttribute("stroke", "#7F7F7F");
		axisNode.setAttribute("stroke-width", "1.25");

		if (isHorizontal) {
			axisNode.setAttribute("x1", `${zeroPosition}`);
			axisNode.setAttribute("y1", `${margin.top}`);
			axisNode.setAttribute("x2", `${zeroPosition}`);
			axisNode.setAttribute("y2", `${margin.top + plotHeight}`);
		} else {
			axisNode.setAttribute("x1", `${margin.left}`);
			axisNode.setAttribute("y1", `${zeroPosition}`);
			axisNode.setAttribute("x2", `${margin.left + plotWidth}`);
			axisNode.setAttribute("y2", `${zeroPosition}`);
		}

		svgNode.appendChild(axisNode);

		for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
			var categoryOffset = categoryBand * categoryIndex; // 当前分类偏移
			var categoryLabelNode = this.createSvgElement("text"); // 当前分类文字节点

			if (isHorizontal) {
				categoryLabelNode.setAttribute("x", `${margin.left - 10}`);
				categoryLabelNode.setAttribute("y", `${margin.top + categoryOffset + categoryBand / 2 + 5}`);
				categoryLabelNode.setAttribute("text-anchor", "end");
			} else {
				categoryLabelNode.setAttribute("x", `${margin.left + categoryOffset + categoryBand / 2}`);
				categoryLabelNode.setAttribute("y", `${margin.top + plotHeight + 44}`);
				categoryLabelNode.setAttribute("text-anchor", "middle");
			}

			categoryLabelNode.setAttribute("font-size", "14");
			categoryLabelNode.setAttribute("fill", "#444444");
			categoryLabelNode.textContent = categories[categoryIndex];
			svgNode.appendChild(categoryLabelNode);
		}

		if (isHorizontal) {
			var barBandSize = isStacked ? categoryBand * 0.62 : (categoryBand * 0.72) / Math.max(seriesList.length, 1); // 横向柱宽

			for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
				var stackedCursor = zeroPosition; // 当前分类堆积起点

				for (let seriesIndex = 0; seriesIndex < seriesList.length; seriesIndex++) {
					var series = seriesList[seriesIndex]; // 当前系列
					var value = series.values[categoryIndex] ?? 0; // 当前柱值
					var color = this.getChartSeriesColor(series, seriesIndex); // 当前柱颜色
					var rectNode = this.createSvgElement("rect"); // 当前矩形柱节点
					var yPosition = margin.top + categoryIndex * categoryBand + (isStacked ? categoryBand * 0.19 : categoryBand * 0.14 + seriesIndex * barBandSize); // 当前柱纵坐标
					var startPosition = isStacked ? stackedCursor : Math.min(zeroPosition, valueToPosition(value)); // 起始横坐标
					var endPosition = isStacked ? valueToPosition(value + (stackedCursor === zeroPosition ? 0 : 0)) : Math.max(zeroPosition, valueToPosition(value)); // 结束横坐标占位
					var width = isStacked ? Math.max(valueToPosition(value) - zeroPosition, 0) : Math.abs(valueToPosition(value) - zeroPosition); // 当前柱宽度

					rectNode.setAttribute("x", `${isStacked ? stackedCursor : startPosition}`);
					rectNode.setAttribute("y", `${yPosition}`);
					rectNode.setAttribute("width", `${Math.max(width, 1)}`);
					rectNode.setAttribute("height", `${Math.max(barBandSize, 8)}`);
					rectNode.setAttribute("fill", color);
					rectNode.setAttribute("rx", "2");
					rectNode.setAttribute("ry", "2");
					svgNode.appendChild(rectNode);

					if (isStacked)
						stackedCursor += width;
				}
			}
		} else {
			var columnBandSize = isStacked ? categoryBand * 0.62 : (categoryBand * 0.72) / Math.max(seriesList.length, 1); // 纵向柱宽

			for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
				var stackedHeight = 0; // 当前分类堆积高度

				for (let seriesIndex = 0; seriesIndex < seriesList.length; seriesIndex++) {
					var series = seriesList[seriesIndex]; // 当前系列
					var value = series.values[categoryIndex] ?? 0; // 当前柱值
					var color = this.getChartSeriesColor(series, seriesIndex); // 当前柱颜色
					var rectNode = this.createSvgElement("rect"); // 当前矩形柱节点
					var xPosition = margin.left + categoryIndex * categoryBand + (isStacked ? categoryBand * 0.19 : categoryBand * 0.14 + seriesIndex * columnBandSize); // 当前柱横坐标
					var currentHeight = Math.abs(valueToPosition(value) - zeroPosition); // 当前柱高度
					var yPosition = isStacked
						? zeroPosition - stackedHeight - currentHeight
						: Math.min(zeroPosition, valueToPosition(value)); // 当前柱纵坐标

					rectNode.setAttribute("x", `${xPosition}`);
					rectNode.setAttribute("y", `${yPosition}`);
					rectNode.setAttribute("width", `${Math.max(columnBandSize, 8)}`);
					rectNode.setAttribute("height", `${Math.max(currentHeight, 1)}`);
					rectNode.setAttribute("fill", color);
					rectNode.setAttribute("rx", "2");
					rectNode.setAttribute("ry", "2");
					svgNode.appendChild(rectNode);

					if (isStacked)
						stackedHeight += currentHeight;
				}
			}
		}

		this.appendChartAxisTitles(svgNode, chart, viewWidth, viewHeight, margin, isHorizontal);
		return svgNode;
	}

	renderLineLikeChartSvg(chart: ChartData, isAreaChart: boolean): SVGElement {
		var svgNode = this.createChartSvgRoot(); // 折线或面积图 SVG 根节点
		var categories = this.getChartCategories(chart); // 分类标签集合
		var seriesList = chart.series.map(series => ({ ...series, values: this.normalizeChartSeriesValues(series.values, categories.length) })); // 归一化系列数据
		var values = seriesList.flatMap(series => series.values).filter(value => Number.isFinite(value)); // 所有数值

		if (categories.length === 0 || values.length === 0)
			return svgNode;

		var minValue = Math.min(0, ...values); // 最小值
		var maxValue = Math.max(...values); // 最大值
		var safeMaxValue = maxValue === minValue ? maxValue + 1 : maxValue; // 安全最大值
		var viewWidth = 800; // SVG 视口宽度
		var viewHeight = 520; // SVG 视口高度
		var margin = { top: 24, right: 26, bottom: chart.categoryAxisTitle ? 110 : 88, left: chart.valueAxisTitle ? 92 : 74 }; // 图表边距
		var plotWidth = viewWidth - margin.left - margin.right; // 绘图区宽度
		var plotHeight = viewHeight - margin.top - margin.bottom; // 绘图区高度
		var stepX = categories.length > 1 ? plotWidth / (categories.length - 1) : 0; // 分类横向步长
		var gridCount = 5; // 网格线数量
		var valueToY = (value: number) => margin.top + plotHeight - ((value - minValue) / (safeMaxValue - minValue)) * plotHeight; // 数值纵坐标换算
		var baselineY = valueToY(0); // 零轴位置

		for (let gridIndex = 0; gridIndex <= gridCount; gridIndex++) {
			var gridValue = minValue + ((safeMaxValue - minValue) / gridCount) * gridIndex; // 当前网格值
			var yPosition = valueToY(gridValue); // 当前网格纵坐标
			var lineNode = this.createSvgElement("line"); // 当前网格线
			var labelNode = this.createSvgElement("text"); // 当前网格标签

			lineNode.setAttribute("x1", `${margin.left}`);
			lineNode.setAttribute("y1", `${yPosition}`);
			lineNode.setAttribute("x2", `${margin.left + plotWidth}`);
			lineNode.setAttribute("y2", `${yPosition}`);
			lineNode.setAttribute("stroke", "#D9D9D9");
			lineNode.setAttribute("stroke-width", "1");
			labelNode.setAttribute("x", `${margin.left - 10}`);
			labelNode.setAttribute("y", `${yPosition + 4}`);
			labelNode.setAttribute("text-anchor", "end");
			labelNode.setAttribute("font-size", "14");
			labelNode.setAttribute("fill", "#666666");
			labelNode.textContent = this.formatChartValue(gridValue);
			svgNode.appendChild(lineNode);
			svgNode.appendChild(labelNode);
		}

		var axisNode = this.createSvgElement("line"); // 横轴线
		axisNode.setAttribute("x1", `${margin.left}`);
		axisNode.setAttribute("y1", `${baselineY}`);
		axisNode.setAttribute("x2", `${margin.left + plotWidth}`);
		axisNode.setAttribute("y2", `${baselineY}`);
		axisNode.setAttribute("stroke", "#7F7F7F");
		axisNode.setAttribute("stroke-width", "1.25");
		svgNode.appendChild(axisNode);

		for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
			var xPosition = margin.left + (categories.length > 1 ? categoryIndex * stepX : plotWidth / 2); // 当前分类横坐标
			var labelNode = this.createSvgElement("text"); // 当前分类标签

			labelNode.setAttribute("x", `${xPosition}`);
			labelNode.setAttribute("y", `${margin.top + plotHeight + 44}`);
			labelNode.setAttribute("text-anchor", "middle");
			labelNode.setAttribute("font-size", "14");
			labelNode.setAttribute("fill", "#444444");
			labelNode.textContent = categories[categoryIndex];
			svgNode.appendChild(labelNode);
		}

		seriesList.forEach((series, seriesIndex) => {
			var strokeColor = series.strokeColor || series.color || this.getChartSeriesColor(series, seriesIndex); // 当前系列描边色
			var fillColor = series.color || this.getChartSeriesColor(series, seriesIndex); // 当前系列填充色
			var points = series.values.map((value, pointIndex) => ({
				x: margin.left + (categories.length > 1 ? pointIndex * stepX : plotWidth / 2),
				y: valueToY(value),
			})); // 当前系列点集合

			if (points.length === 0)
				return;

			if (isAreaChart) {
				var areaNode = this.createSvgElement("path"); // 面积区域节点
				var areaPath = `M ${points[0].x} ${baselineY} L ${points.map(point => `${point.x} ${point.y}`).join(" L ")} L ${points[points.length - 1].x} ${baselineY} Z`; // 面积路径
				areaNode.setAttribute("d", areaPath);
				areaNode.setAttribute("fill", fillColor);
				areaNode.setAttribute("fill-opacity", "0.24");
				areaNode.setAttribute("stroke", "none");
				svgNode.appendChild(areaNode);
			}

			var lineNode = this.createSvgElement("path"); // 系列折线节点
			lineNode.setAttribute("d", `M ${points.map(point => `${point.x} ${point.y}`).join(" L ")}`);
			lineNode.setAttribute("fill", "none");
			lineNode.setAttribute("stroke", strokeColor);
			lineNode.setAttribute("stroke-width", "3");
			lineNode.setAttribute("stroke-linejoin", "round");
			lineNode.setAttribute("stroke-linecap", "round");
			svgNode.appendChild(lineNode);

			points.forEach(point => {
				var markerNode = this.createSvgElement("circle"); // 数据点标记节点
				markerNode.setAttribute("cx", `${point.x}`);
				markerNode.setAttribute("cy", `${point.y}`);
				markerNode.setAttribute("r", "4.25");
				markerNode.setAttribute("fill", fillColor);
				markerNode.setAttribute("stroke", "#FFFFFF");
				markerNode.setAttribute("stroke-width", "1.5");
				svgNode.appendChild(markerNode);
			});
		});

		this.appendChartAxisTitles(svgNode, chart, viewWidth, viewHeight, margin, false);
		return svgNode;
	}

	renderScatterChartSvg(chart: ChartData): SVGElement {
		var svgNode = this.createChartSvgRoot(); // 散点图 SVG 根节点
		var pointSets = chart.series
			.map(series => {
				var sourceXValues = series.xValues?.length ? series.xValues : series.values?.map((_, index) => index + 1) ?? []; // 当前系列横轴原始值
				var pairs = (series.values ?? []).map((value, index) => ({ // 当前系列点对集合
					x: sourceXValues[index],
					y: value,
				})).filter(pair => Number.isFinite(pair.x) && Number.isFinite(pair.y));

				return {
					...series,
					xValues: pairs.map(pair => pair.x),
					values: pairs.map(pair => pair.y),
				};
			})
			.filter(series => series.xValues.length > 0 && series.values.length > 0); // 有效系列集合

		if (pointSets.length === 0)
			return svgNode;

		var xValues = pointSets.flatMap(series => series.xValues).filter(value => Number.isFinite(value)); // 全部横轴值
		var yValues = pointSets.flatMap(series => series.values).filter(value => Number.isFinite(value)); // 全部纵轴值
		var minX = Math.min(0, ...xValues); // 横轴最小值
		var maxX = Math.max(...xValues); // 横轴最大值
		var minY = Math.min(0, ...yValues); // 纵轴最小值
		var maxY = Math.max(...yValues); // 纵轴最大值
		var safeMaxX = maxX === minX ? maxX + 1 : maxX; // 安全横轴最大值
		var safeMaxY = maxY === minY ? maxY + 1 : maxY; // 安全纵轴最大值
		var viewWidth = 800; // SVG 视口宽度
		var viewHeight = 520; // SVG 视口高度
		var margin = { top: 24, right: 26, bottom: chart.categoryAxisTitle ? 100 : 78, left: chart.valueAxisTitle ? 92 : 74 }; // 图表边距
		var plotWidth = viewWidth - margin.left - margin.right; // 绘图区宽度
		var plotHeight = viewHeight - margin.top - margin.bottom; // 绘图区高度
		var valueToX = (value: number) => margin.left + ((value - minX) / (safeMaxX - minX)) * plotWidth; // 横轴坐标换算
		var valueToY = (value: number) => margin.top + plotHeight - ((value - minY) / (safeMaxY - minY)) * plotHeight; // 纵轴坐标换算
		var gridCount = 5; // 网格线数量

		for (let gridIndex = 0; gridIndex <= gridCount; gridIndex++) {
			var gridXValue = minX + ((safeMaxX - minX) / gridCount) * gridIndex; // 当前横轴网格值
			var gridXPosition = valueToX(gridXValue); // 当前横轴网格坐标
			var gridXLine = this.createSvgElement("line"); // 当前横轴网格线
			var gridXLabel = this.createSvgElement("text"); // 当前横轴网格标签

			gridXLine.setAttribute("x1", `${gridXPosition}`);
			gridXLine.setAttribute("y1", `${margin.top}`);
			gridXLine.setAttribute("x2", `${gridXPosition}`);
			gridXLine.setAttribute("y2", `${margin.top + plotHeight}`);
			gridXLine.setAttribute("stroke", "#E0E0E0");
			gridXLine.setAttribute("stroke-width", "1");
			gridXLabel.setAttribute("x", `${gridXPosition}`);
			gridXLabel.setAttribute("y", `${margin.top + plotHeight + 20}`);
			gridXLabel.setAttribute("text-anchor", "middle");
			gridXLabel.setAttribute("font-size", "14");
			gridXLabel.setAttribute("fill", "#666666");
			gridXLabel.textContent = this.formatChartValue(gridXValue);
			svgNode.appendChild(gridXLine);
			svgNode.appendChild(gridXLabel);
		}

		for (let gridIndex = 0; gridIndex <= gridCount; gridIndex++) {
			var gridYValue = minY + ((safeMaxY - minY) / gridCount) * gridIndex; // 当前纵轴网格值
			var gridYPosition = valueToY(gridYValue); // 当前纵轴网格坐标
			var gridYLine = this.createSvgElement("line"); // 当前纵轴网格线
			var gridYLabel = this.createSvgElement("text"); // 当前纵轴网格标签

			gridYLine.setAttribute("x1", `${margin.left}`);
			gridYLine.setAttribute("y1", `${gridYPosition}`);
			gridYLine.setAttribute("x2", `${margin.left + plotWidth}`);
			gridYLine.setAttribute("y2", `${gridYPosition}`);
			gridYLine.setAttribute("stroke", "#E0E0E0");
			gridYLine.setAttribute("stroke-width", "1");
			gridYLabel.setAttribute("x", `${margin.left - 10}`);
			gridYLabel.setAttribute("y", `${gridYPosition + 4}`);
			gridYLabel.setAttribute("text-anchor", "end");
			gridYLabel.setAttribute("font-size", "14");
			gridYLabel.setAttribute("fill", "#666666");
			gridYLabel.textContent = this.formatChartValue(gridYValue);
			svgNode.appendChild(gridYLine);
			svgNode.appendChild(gridYLabel);
		}

		pointSets.forEach((series, seriesIndex) => {
			var strokeColor = series.strokeColor || series.color || this.getChartSeriesColor(series, seriesIndex); // 当前系列描边色
			var fillColor = series.color || this.getChartSeriesColor(series, seriesIndex); // 当前系列填充色
			var points = series.values.map((value, pointIndex) => ({
				x: valueToX(series.xValues[pointIndex] ?? pointIndex + 1),
				y: valueToY(value),
			})); // 当前系列点集

			if (points.length === 0)
				return;

			var lineNode = this.createSvgElement("path"); // 点集连接线
			lineNode.setAttribute("d", `M ${points.map(point => `${point.x} ${point.y}`).join(" L ")}`);
			lineNode.setAttribute("fill", "none");
			lineNode.setAttribute("stroke", strokeColor);
			lineNode.setAttribute("stroke-width", "2");
			lineNode.setAttribute("stroke-opacity", "0.8");
			svgNode.appendChild(lineNode);

			points.forEach(point => {
				var pointNode = this.createSvgElement("circle"); // 散点节点
				pointNode.setAttribute("cx", `${point.x}`);
				pointNode.setAttribute("cy", `${point.y}`);
				pointNode.setAttribute("r", "4.25");
				pointNode.setAttribute("fill", fillColor);
				pointNode.setAttribute("stroke", "#FFFFFF");
				pointNode.setAttribute("stroke-width", "1.25");
				svgNode.appendChild(pointNode);
			});
		});

		this.appendChartAxisTitles(svgNode, chart, viewWidth, viewHeight, margin, false);
		return svgNode;
	}

	renderPieChartSvg(chart: ChartData, isDoughnutChart: boolean): SVGElement {
		var svgNode = this.createChartSvgRoot(); // 饼图或圆环图 SVG 根节点
		var series = chart.series[0]; // 饼图首个系列
		var categories = this.getChartCategories(chart, series?.values?.length ?? 0); // 扇区分类标签
		var slices = (series?.values ?? []).map((value, index) => ({ // 饼图扇区数据集合
			value,
			label: categories[index] ?? `数据 ${index + 1}`,
			color: this.getChartPointColor(series, index),
		})).filter(slice => Number.isFinite(slice.value) && slice.value > 0);

		if (!series || slices.length === 0 || chart.series.length !== 1)
			return svgNode;

		var totalValue = slices.reduce((sumValue, currentValue) => sumValue + currentValue.value, 0); // 总值
		var viewWidth = 520; // SVG 视口宽度
		var viewHeight = 420; // SVG 视口高度
		var centerX = 210; // 圆心横坐标
		var centerY = 210; // 圆心纵坐标
		var outerRadius = 150; // 外半径
		var innerRadius = isDoughnutChart ? outerRadius * ((chart.holeSize ?? 50) / 100) : 0; // 内半径
		var currentAngle = -Math.PI / 2; // 当前起始角度

		svgNode.setAttribute("viewBox", `0 0 ${viewWidth} ${viewHeight}`);

		for (let pointIndex = 0; pointIndex < slices.length; pointIndex++) {
			var value = slices[pointIndex].value; // 当前扇区数值
			var angleSize = totalValue > 0 ? (value / totalValue) * Math.PI * 2 : 0; // 当前扇区弧度
			var startAngle = currentAngle; // 当前扇区起始角
			var endAngle = currentAngle + angleSize; // 当前扇区结束角
			var pathNode = this.createSvgElement("path"); // 当前扇区节点
			var percentValue = totalValue > 0 ? (value / totalValue) * 100 : 0; // 当前百分比
			var middleAngle = startAngle + angleSize / 2; // 当前扇区中间角
			var labelRadius = innerRadius > 0 ? innerRadius + (outerRadius - innerRadius) / 2 : outerRadius * 0.62; // 标签半径

			pathNode.setAttribute("d", this.describePieArc(centerX, centerY, outerRadius, innerRadius, startAngle, endAngle));
			pathNode.setAttribute("fill", slices[pointIndex].color);
			pathNode.setAttribute("stroke", "#FFFFFF");
			pathNode.setAttribute("stroke-width", "1.5");
			svgNode.appendChild(pathNode);

			if (percentValue >= 7) {
				var labelNode = this.createSvgElement("text"); // 当前扇区标签
				labelNode.setAttribute("x", `${centerX + Math.cos(middleAngle) * labelRadius}`);
				labelNode.setAttribute("y", `${centerY + Math.sin(middleAngle) * labelRadius + 5}`);
				labelNode.setAttribute("text-anchor", "middle");
				labelNode.setAttribute("font-size", "14");
				labelNode.setAttribute("fill", "#FFFFFF");
				labelNode.setAttribute("font-weight", "600");
				labelNode.textContent = `${Math.round(percentValue)}%`;
				svgNode.appendChild(labelNode);
			}

			currentAngle = endAngle;
		}

		return svgNode;
	}

	appendChartAxisTitles(
		svgNode: SVGElement,
		chart: ChartData,
		viewWidth: number,
		viewHeight: number,
		margin: { top: number; right: number; bottom: number; left: number },
		isHorizontalBarChart: boolean
	) {
		var bottomAxisTitle = isHorizontalBarChart ? chart.valueAxisTitle : chart.categoryAxisTitle; // 底部轴标题
		var leftAxisTitle = isHorizontalBarChart ? chart.categoryAxisTitle : chart.valueAxisTitle; // 左侧轴标题

		if (bottomAxisTitle) {
			var categoryNode = this.createSvgElement("text"); // 分类轴标题节点
			categoryNode.setAttribute("x", `${viewWidth / 2}`);
			categoryNode.setAttribute("y", `${viewHeight - 24}`);
			categoryNode.setAttribute("text-anchor", "middle");
			categoryNode.setAttribute("font-size", "15");
			categoryNode.setAttribute("fill", "#555555");
			categoryNode.textContent = bottomAxisTitle;
			svgNode.appendChild(categoryNode);
		}

		if (leftAxisTitle) {
			var valueNode = this.createSvgElement("text"); // 数值轴标题节点
			valueNode.setAttribute("x", `${28}`);
			valueNode.setAttribute("y", `${margin.top + (viewHeight - margin.top - margin.bottom) / 2}`);
			valueNode.setAttribute("text-anchor", "middle");
			valueNode.setAttribute("font-size", "15");
			valueNode.setAttribute("fill", "#555555");
			valueNode.setAttribute("transform", `rotate(-90 28 ${margin.top + (viewHeight - margin.top - margin.bottom) / 2})`);
			valueNode.textContent = leftAxisTitle;
			svgNode.appendChild(valueNode);
		}
	}

	createChartSvgRoot(): SVGElement {
		var svgNode = this.createSvgElement("svg"); // 图表 SVG 根节点
		svgNode.setAttribute("viewBox", "0 0 800 520");
		svgNode.setAttribute("preserveAspectRatio", "xMidYMid meet");
		Object.assign(svgNode.style, {
			width: "100%",
			height: "100%",
			display: "block",
			overflow: "visible"
		});
		return svgNode;
	}

	describePieArc(
		centerX: number,
		centerY: number,
		outerRadius: number,
		innerRadius: number,
		startAngle: number,
		endAngle: number
	): string {
		var largeArcFlag = endAngle - startAngle > Math.PI ? 1 : 0; // 是否大圆弧
		var startOuterPoint = this.getCirclePoint(centerX, centerY, outerRadius, startAngle); // 外圆起点
		var endOuterPoint = this.getCirclePoint(centerX, centerY, outerRadius, endAngle); // 外圆终点

		if (innerRadius <= 0) {
			return [
				`M ${centerX} ${centerY}`,
				`L ${startOuterPoint.x} ${startOuterPoint.y}`,
				`A ${outerRadius} ${outerRadius} 0 ${largeArcFlag} 1 ${endOuterPoint.x} ${endOuterPoint.y}`,
				"Z"
			].join(" ");
		}

		var endInnerPoint = this.getCirclePoint(centerX, centerY, innerRadius, endAngle); // 内圆终点
		var startInnerPoint = this.getCirclePoint(centerX, centerY, innerRadius, startAngle); // 内圆起点

		return [
			`M ${startOuterPoint.x} ${startOuterPoint.y}`,
			`A ${outerRadius} ${outerRadius} 0 ${largeArcFlag} 1 ${endOuterPoint.x} ${endOuterPoint.y}`,
			`L ${endInnerPoint.x} ${endInnerPoint.y}`,
			`A ${innerRadius} ${innerRadius} 0 ${largeArcFlag} 0 ${startInnerPoint.x} ${startInnerPoint.y}`,
			"Z"
		].join(" ");
	}

	getCirclePoint(centerX: number, centerY: number, radius: number, angle: number) {
		return {
			x: centerX + Math.cos(angle) * radius,
			y: centerY + Math.sin(angle) * radius
		};
	}

	getChartCategories(chart: ChartData, fallbackLength: number = 0): string[] {
		var categoryList = chart.categories?.length ? [...chart.categories] : []; // 图表分类标签集合
		var requiredLength = Math.max(
			categoryList.length,
			fallbackLength,
			...chart.series.map(series => series.categories?.length ?? 0),
			...chart.series.map(series => series.values?.length ?? 0)
		); // 分类标签目标长度

		if (categoryList.length === 0) {
			var seriesCategories = chart.series.find(series => series.categories?.length > 0)?.categories ?? []; // 系列中的分类标签
			categoryList = [...seriesCategories];
		}

		while (categoryList.length < requiredLength) {
			categoryList.push(`数据 ${categoryList.length + 1}`);
		}

		return categoryList;
	}

	normalizeChartSeriesValues(values: number[] = [], targetLength: number): number[] {
		var result = [...values]; // 当前归一化数据集合

		while (result.length < targetLength) {
			result.push(0);
		}

		return result.map(value => Number.isFinite(value) ? value : 0);
	}

	getChartSeriesColor(series: ChartSeriesData, seriesIndex: number): string {
		return series.color || series.strokeColor || defaultChartPalette[seriesIndex % defaultChartPalette.length];
	}

	getChartPointColor(series: ChartSeriesData, pointIndex: number): string {
		return series.pointColors?.[pointIndex]
			|| series.color
			|| defaultChartPalette[pointIndex % defaultChartPalette.length];
	}

	formatChartValue(value: number | string): string {
		if (value == null || value === "")
			return "";

		if (typeof value === "string")
			return value;

		if (!Number.isFinite(value))
			return "";

		if (Math.abs(value) >= 1000 || Number.isInteger(value))
			return `${Math.round(value * 100) / 100}`;

		return `${value.toFixed(2).replace(/\.?0+$/, "")}`;
	}

	// 渲染 DrawingML 形状，使用 SVG 承载几何图形并叠加文本层
	renderShape(elem: IDomShape) {
		var result = this.h({
			tagName: "div",
			className: cx(elem.className, `${this.className}-shape`),
			style: {
				position: "absolute",
				left: "0px",
				top: "0px",
				width: "100%",
				height: "100%",
				...elem.cssStyle
			}
		}) as HTMLDivElement;
		var svgElement = this.createSvgElement("svg"); // 形状 SVG 容器
		var shapeElement = this.createDrawingShapeElement(elem); // 形状几何图形节点
		var textPadding = elem.textPadding ?? {}; // 形状文本内边距

		svgElement.setAttribute("viewBox", "0 0 100 100");
		svgElement.setAttribute("preserveAspectRatio", "none");
		Object.assign(svgElement.style, {
			width: "100%",
			height: "100%",
			display: "block",
			overflow: "visible"
		});

		svgElement.appendChild(shapeElement);
		result.appendChild(svgElement);

		if (elem.children?.length > 0) {
			var textLayer = this.h({
				tagName: "div",
				style: {
					position: "absolute",
					left: "0px",
					top: "0px",
					right: "0px",
					bottom: "0px",
					display: "flex",
					flexDirection: "column",
					justifyContent: elem.textVerticalAlign ?? "flex-start",
					boxSizing: "border-box",
					paddingLeft: textPadding.left ?? "0pt",
					paddingTop: textPadding.top ?? "0pt",
					paddingRight: textPadding.right ?? "0pt",
					paddingBottom: textPadding.bottom ?? "0pt",
					overflow: "hidden"
				},
				children: this.renderElements(elem.children)
			}) as HTMLDivElement;

			result.appendChild(textLayer);
		}

		return result;
	}

	// 根据预设几何类型创建 SVG 图形，优先覆盖常见 Office 绘图形状
	createDrawingShapeElement(elem: IDomShape): SVGElement {
		var geometry = this.getDrawingShapeGeometry(elem.geometry); // 当前形状的 SVG 几何定义
		var result = this.createSvgElement(geometry.tagName as any);

		Object.entries(geometry.attrs).forEach(([key, value]) => result.setAttribute(key, value));
		result.setAttribute("fill", geometry.tagName === "line" ? "none" : (elem.fill ?? "none"));
		result.setAttribute("stroke", elem.stroke ?? "none");
		result.setAttribute("stroke-width", elem.strokeWidth ?? "1pt");
		result.setAttribute("stroke-linejoin", "round");

		return result;
	}

	// 将 Word 预设几何图形映射为最小可用的 SVG 形状
	getDrawingShapeGeometry(geometry: string): { tagName: string; attrs: Record<string, string> } {
		switch (geometry) {
			case "ellipse":
				return { tagName: "ellipse", attrs: { cx: "50", cy: "50", rx: "50", ry: "50" } };

			case "line":
			case "straightConnector1":
				return { tagName: "line", attrs: { x1: "0", y1: "50", x2: "100", y2: "50" } };

			case "triangle":
				return { tagName: "polygon", attrs: { points: "50,0 100,100 0,100" } };

			case "rtTriangle":
				return { tagName: "polygon", attrs: { points: "0,0 100,100 0,100" } };

			case "diamond":
				return { tagName: "polygon", attrs: { points: "50,0 100,50 50,100 0,50" } };

			case "parallelogram":
				return { tagName: "polygon", attrs: { points: "25,0 100,0 75,100 0,100" } };

			case "hexagon":
				return { tagName: "polygon", attrs: { points: "25,0 75,0 100,50 75,100 25,100 0,50" } };

			case "octagon":
				return { tagName: "polygon", attrs: { points: "30,0 70,0 100,30 100,70 70,100 30,100 0,70 0,30" } };

			case "roundRect":
				return { tagName: "rect", attrs: { x: "0", y: "0", width: "100", height: "100", rx: "12", ry: "12" } };

			default:
				return { tagName: "rect", attrs: { x: "0", y: "0", width: "100", height: "100" } };
		}
	}

	renderText(elem: WmlText) {
		return this.h(elem.text);
	}

	renderDeletedText(elem: WmlText) {
		return this.options.renderChanges ? this.renderText(elem) : null;
	}

	renderBreak(elem: WmlBreak) {
		return elem.break == "textWrapping" ? this.h({ tagName: "br" }) : null;
	}

	renderInserted(elem: OpenXmlElement): Node | Node[] {
		if (this.options.renderChanges)
			return this.renderContainer(elem, "ins");

		return this.renderElements(elem.children);
	}

	renderDeleted(elem: OpenXmlElement): Node {
		if (this.options.renderChanges)
			return this.renderContainer(elem, "del");

		return null;
	}

	renderSymbol(elem: WmlSymbol) {
		return this.h({ tagName: "span", children: [String.fromCharCode(elem.char)], style: { fontFamily: elem.font } });
	}

	renderFootnoteReference(elem: WmlNoteReference) {
		this.currentFootnoteIds.push(elem.id);
		return this.h({ tagName: "sup", children: [`${this.currentFootnoteIds.length}`] });
	}

	renderEndnoteReference(elem: WmlNoteReference) {
		this.currentEndnoteIds.push(elem.id);
		return this.h({ tagName: "sup", children: [`${this.currentEndnoteIds.length}`] });
	}

	renderTab(elem: OpenXmlElement) {
		var tabSpan = this.h({ tagName: "span", children: ["\u2003"] }) as HTMLElement;//"&nbsp;";

		if (this.options.experimental) {
			tabSpan.className = this.tabStopClass();
			var stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
			this.currentTabs.push({ stops, span: tabSpan });
		}

		return tabSpan;
	}

	renderBookmarkStart(elem: WmlBookmarkStart) {
		return this.h({ tagName: "span", id: elem.name });
	}

	renderRun(elem: WmlRun) {
		if (elem.fieldRun)
			return null;

		let children = this.renderElements(elem.children);

		if (elem.verticalAlign) {
			children = [this.h({ tagName: elem.verticalAlign, children: this.renderElements(elem.children) })];
		}

		const result = this.toHTML(elem, ns.html, "span", children);

		if (elem.id)
			result.id = elem.id;

		return result;
	}

	renderTable(elem: WmlTable) {
		this.tableCellPositions.push(this.currentCellPosition);
		this.tableVerticalMerges.push(this.currentVerticalMerge);
		this.currentVerticalMerge = {};
		this.currentCellPosition = { col: 0, row: 0 };

		const children = [];

		if (elem.columns)
			children.push(this.renderTableColumns(elem.columns));

		children.push(...this.renderElements(elem.children));

		this.currentVerticalMerge = this.tableVerticalMerges.pop();
		this.currentCellPosition = this.tableCellPositions.pop();
		return this.toHTML(elem, ns.html, "table", children);
	}

	renderTableColumns(columns: WmlTableColumn[]) {
		const children = columns.map(x => this.h({ tagName: "col", style: { width: x.width } }));
		return this.h({ tagName: "colgroup", children });
	}

	renderTableRow(elem: WmlTableRow) {
		this.currentCellPosition.col = 0;

		const children = [];

		if (elem.gridBefore)
			children.push(this.renderTableCellPlaceholder(elem.gridBefore));

		children.push(...this.renderElements(elem.children));

		if (elem.gridAfter)
			children.push(this.renderTableCellPlaceholder(elem.gridAfter));

		this.currentCellPosition.row++;

		return this.toHTML(elem, ns.html, "tr", children) as HTMLTableRowElement;
	}

	renderTableCellPlaceholder(colSpan: number) {
		return this.h({ tagName: "td", colSpan, style: { border: "none" } });
	}

	renderTableCell(elem: WmlTableCell) {
		let result = this.toHTML(elem, ns.html, "td");

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

		if (elem.span)
			result.colSpan = elem.span;

		this.currentCellPosition.col += result.colSpan;

		return result;
	}

	renderVmlPicture(elem: OpenXmlElement) {
		return this.h({
			tagName: "span",
			style: {
				position: "relative",
				display: "inline-block",
				width: "0px",
				height: "0px",
				fontSize: "0px",
				lineHeight: "0px",
				verticalAlign: "top",
				overflow: "visible",
				zIndex: "1"
			},
			children: this.renderElements(elem.children)
		});
	}

	renderVmlElement(elem: VmlElement): SVGElement {
		var style = this.parseVmlStyleText(elem.cssStyleText); // VML 外层 SVG 样式
		var container = this.h({ ns: ns.svg, tagName: "svg", style }) as SVGElement;

		const result = this.renderVmlChildElement(elem);

		container.setAttribute("preserveAspectRatio", "none");

		if (elem.coordOrigin?.length === 2 && elem.coordSize?.length === 2) {
			container.setAttribute("viewBox", `${elem.coordOrigin[0]} ${elem.coordOrigin[1]} ${elem.coordSize[0]} ${elem.coordSize[1]}`);
		}

		if (elem.imageHref?.id) {
			this.tasks.push(this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
				.then(x => result.setAttribute("href", x)));
		}

		container.appendChild(result);

		requestAnimationFrame(() => {
			if (elem.coordSize?.length === 2)
				return;

			try {
				const bb = (container.firstElementChild as any)?.getBBox?.();

				if (!bb)
					return;

				container.setAttribute("width", `${Math.ceil(bb.x + bb.width)}`);
				container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
			} catch {
				// TODO: 保留 VML 边界计算失败场景，后续如有样例再细化
			}
		});

		return container;
	}

	renderVmlChildElement(elem: VmlElement): any {
		if (elem.textPath) {
			return this.renderVmlTextPathElement(elem);
		}

		if (elem.shapeType) {
			return this.renderVmlShapeElement(elem);
		}

		const result = this.createSvgElement((elem.tagName || "g") as any);
		this.applyVmlCommonAttributes(result, elem);

		for (let child of elem.children) {
			if (child.type == DomType.VmlElement) {
				result.appendChild(this.renderVmlChildElement(child as VmlElement));
			} else {
				for (const renderedChild of asArray(this.renderElement(child as any))) {
					result.appendChild(renderedChild);
				}
			}
		}

		return result;
	}

	parseVmlStyleText(styleText: string): Record<string, string> {
		var result = parseCssRules(styleText ?? ""); // VML 样式对象

		result.overflow ??= "visible";
		result.display ??= "block";

		return result;
	}

	getVmlBounds(elem: VmlElement): { left: number; top: number; width: number; height: number } {
		var style = this.parseVmlStyleText(elem.cssStyleText); // 当前 VML 节点样式
		var left = this.parseVmlNumber(style.left); // 原始左侧坐标
		var top = this.parseVmlNumber(style.top); // 原始顶部坐标
		var width = this.parseVmlNumber(style.width); // 宽度
		var height = this.parseVmlNumber(style.height); // 高度
		var isTopLevelPositionedShape = elem.parent?.type === DomType.VmlPicture && this.getVmlShapeKind(elem.shapeType) !== "group"; // 顶层 VML 形状由外层 SVG 容器负责定位

		if (isTopLevelPositionedShape) {
			left = 0; // 顶层形状内部坐标归零，避免与外层容器定位重复叠加
			top = 0; // 顶层形状内部坐标归零，避免与外层容器定位重复叠加
		}

		return { left, top, width, height };
	}

	parseVmlNumber(value: string): number {
		if (!value)
			return 0;

		var matchedValue = value.match(/-?\d+(\.\d+)?/); // 数值片段
		return matchedValue ? parseFloat(matchedValue[0]) : 0;
	}

	normalizeVmlShapeType(shapeType: string): string {
		return shapeType?.replace(/^#_x0000_t/i, ""); // 统一清洗内建 VML 图形类型
	}

	getVmlShapeKind(shapeType: string): string {
		switch (this.normalizeVmlShapeType(shapeType)) {
			case "3":
				return "ellipse";

			case "136":
				return "textPlain";

			case "144":
				return "textArchUpCurve";

			default:
				return "group";
		}
	}

	applyVmlCommonAttributes(result: SVGElement, elem: VmlElement) {
		Object.entries(elem.attrs).forEach(([key, value]) => {
			if (value != null)
				result.setAttribute(key, value);
		});
	}

	renderVmlShapeElement(elem: VmlElement): SVGElement {
		var shapeKind = this.getVmlShapeKind(elem.shapeType); // VML 形状类别
		var bounds = this.getVmlBounds(elem); // 形状边界数据

		switch (shapeKind) {
			case "ellipse":
				return this.renderVmlEllipseShape(elem, bounds);

			default:
				return this.renderVmlGroupElement(elem);
		}
	}

	renderVmlGroupElement(elem: VmlElement): SVGElement {
		var result = this.createSvgElement("g"); // SVG 分组节点

		this.applyVmlCommonAttributes(result, elem);

		for (let child of elem.children) {
			if (child.type == DomType.VmlElement) {
				result.appendChild(this.renderVmlChildElement(child as VmlElement));
			} else {
				for (const renderedChild of asArray(this.renderElement(child as any))) {
					result.appendChild(renderedChild);
				}
			}
		}

		return result;
	}

	renderVmlEllipseShape(elem: VmlElement, bounds: { left: number; top: number; width: number; height: number }): SVGElement {
		var result = this.createSvgElement("ellipse"); // 椭圆轮廓节点
		var unitsPerPoint = this.getVmlUnitsPerPoint(elem); // 当前椭圆所在坐标系每磅对应单位数
		var strokeWidth = this.parseVmlStrokeWidth(elem.attrs["stroke-width"], unitsPerPoint); // 椭圆描边宽度
		var halfStrokeWidth = strokeWidth ? parseFloat(strokeWidth) / 2 : 0; // 椭圆半描边宽度
		var radiusX = Math.max(bounds.width / 2 - halfStrokeWidth, 0); // 扣除描边后的横向半径
		var radiusY = Math.max(bounds.height / 2 - halfStrokeWidth, 0); // 扣除描边后的纵向半径

		result.setAttribute("cx", `${bounds.left + bounds.width / 2}`);
		result.setAttribute("cy", `${bounds.top + bounds.height / 2}`);
		result.setAttribute("rx", `${radiusX}`);
		result.setAttribute("ry", `${radiusY}`);

		this.applyVmlCommonAttributes(result, elem);
		result.setAttribute("vector-effect", "none");
		result.setAttribute("shape-rendering", "geometricPrecision");

		if (strokeWidth)
			result.setAttribute("stroke-width", strokeWidth);

		return result;
	}

	renderVmlTextPathElement(elem: VmlElement): SVGElement {
		var shapeKind = this.getVmlShapeKind(elem.shapeType); // 文本路径图形类别
		var bounds = this.getVmlBounds(elem); // 文本路径边界
		var group = this.createSvgElement("g"); // 文本路径分组节点

		if (shapeKind === "textArchUpCurve") {
			var defs = this.createSvgElement("defs"); // 路径定义区域
			var path = this.createSvgElement("path"); // 文本弧线路径
			var text = this.createSvgElement("text"); // 文本节点
			var textPath = this.createSvgElement("textPath" as any); // SVG 文本路径节点
			var pathId = `${this.className}-vml-textpath-${++this.vmlTextPathIndex}`; // 唯一路径标识

			path.setAttribute("id", pathId);
			path.setAttribute("d", `M ${bounds.left} ${bounds.top + bounds.height / 2} A ${bounds.width / 2} ${bounds.height / 2} 0 0 1 ${bounds.left + bounds.width} ${bounds.top + bounds.height / 2}`);
			path.setAttribute("fill", "none");
			path.setAttribute("stroke", "none");

			defs.appendChild(path);
			group.appendChild(defs);

			this.applyVmlTextAttributes(text, elem);
			textPath.setAttribute("href", `#${pathId}`);
			textPath.setAttribute("startOffset", "50%");
			textPath.setAttribute("text-anchor", "middle");
			textPath.textContent = elem.textPath.text;
			text.appendChild(textPath);
			group.appendChild(text);

			return group;
		}

		var text = this.createSvgElement("text"); // 普通文字节点

		text.setAttribute("x", `${bounds.left + bounds.width / 2}`);
		text.setAttribute("y", `${bounds.top + bounds.height / 2}`);
		text.setAttribute("text-anchor", "middle");
		text.setAttribute("dominant-baseline", "middle");
		this.applyVmlTextAttributes(text, elem);
		text.textContent = elem.textPath.text;
		group.appendChild(text);

		return group;
	}

	applyVmlTextAttributes(text: SVGTextElement, elem: VmlElement) {
		var textStyle = elem.textPath?.style ?? {}; // 文本路径样式
		var textColor = elem.attrs.stroke ?? elem.attrs.fill ?? "#000000"; // 文本颜色
		var unitsPerPoint = this.getVmlUnitsPerPoint(elem); // 当前 VML 坐标系每磅对应的用户单位数
		var fontSize = this.parseVmlFontSize(textStyle["font-size"], unitsPerPoint); // 换算后的 SVG 字号
		var strokeWidth = this.parseVmlStrokeWidth(elem.attrs["stroke-width"], unitsPerPoint); // 文字描边粗细

		text.setAttribute("fill", textColor);
		text.setAttribute("stroke", textColor);
		text.setAttribute("xml:space", "preserve");
		text.setAttribute("paint-order", "stroke fill");
		text.setAttribute("stroke-linejoin", "round");
		text.setAttribute("stroke-width", strokeWidth ?? "1");
		text.setAttribute("dominant-baseline", "middle");

		if (textStyle["font-family"])
			text.setAttribute("font-family", textStyle["font-family"]);

		if (fontSize)
			text.setAttribute("font-size", fontSize);
	}

	getVmlUnitsPerPoint(elem: VmlElement): number {
		var parent = elem.parent as VmlElement; // 当前图形的父级 VML 节点

		if (!parent || parent.type !== DomType.VmlElement)
			return 1;

		return this.resolveVmlCoordinateScale(parent);
	}

	resolveVmlCoordinateScale(elem: VmlElement): number {
		var style = this.parseVmlStyleText(elem.cssStyleText); // 当前 VML 节点样式
		var parent = elem.parent as VmlElement; // 上级 VML 节点
		var width = this.parseVmlNumber(style.width); // 当前节点宽度
		var coordWidth = elem.coordSize?.length === 2 ? this.parseVmlNumber(elem.coordSize[0]) : 0; // 当前坐标宽度
		var parentScale = parent?.type === DomType.VmlElement ? this.resolveVmlCoordinateScale(parent) : 1; // 父级坐标缩放比例
		var hasAbsoluteUnit = /pt|px|cm|mm|in|pc/i.test(style.width ?? ""); // 是否为绝对长度单位

		if (!coordWidth || !width)
			return parentScale;

		if (hasAbsoluteUnit)
			return coordWidth / width;

		return parentScale * (coordWidth / width);
	}

	parseVmlFontSize(fontSize: string, unitsPerPoint: number): string {
		var size = this.parseVmlNumber(fontSize); // 原始字号数值

		if (!size)
			return null;

		if (/pt$/i.test(fontSize ?? ""))
			return `${size * unitsPerPoint}`;

		return `${size}`;
	}

	parseVmlStrokeWidth(strokeWidth: string, unitsPerPoint: number): string {
		var size = this.parseVmlNumber(strokeWidth); // 原始描边宽度

		if (!size)
			return null;

		if (/pt$/i.test(strokeWidth ?? ""))
			return `${size * unitsPerPoint}`;

		return `${size}`;
	}

	renderMmlRadical(elem: OpenXmlElement) {
		const base = elem.children.find(el => el.type == DomType.MmlBase);

		if (elem.props?.hideDegree) {
			return this.createMathMLElement("msqrt", null, this.renderElements([base]));
		}

		const degree = elem.children.find(el => el.type == DomType.MmlDegree);
		return this.createMathMLElement("mroot", null, this.renderElements([base, degree]));
	}

	renderMmlDelimiter(elem: OpenXmlElement) {
		const children = [];

		children.push(this.createMathMLElement("mo", null, [elem.props.beginChar ?? '(']));
		children.push(...this.renderElements(elem.children));
		children.push(this.createMathMLElement("mo", null, [elem.props.endChar ?? ')']));

		return this.createMathMLElement("mrow", null, children);
	}

	renderMmlNary(elem: OpenXmlElement) {
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;

		const charElem = this.createMathMLElement("mo", null, [elem.props?.char ?? '\u222B']);

		if (supElem || subElem) {
			children.push(this.createMathMLElement("munderover", null, [charElem, subElem, supElem]));
		} else if (supElem) {
			children.push(this.createMathMLElement("mover", null, [charElem, supElem]));
		} else if (subElem) {
			children.push(this.createMathMLElement("munder", null, [charElem, subElem]));
		} else {
			children.push(charElem);
		}

		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createMathMLElement("mrow", null, children);
	}

	renderMmlPreSubSuper(elem: OpenXmlElement) {
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;
		const stubElem = this.createMathMLElement("mo", null);

		children.push(this.createMathMLElement("msubsup", null, [stubElem, subElem, supElem]));
		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createMathMLElement("mrow", null, children);
	}

	renderMmlGroupChar(elem: OpenXmlElement) {
		const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
		const result = this.renderContainerNS(elem, ns.mathML, tagName);

		if (elem.props.char) {
			result.appendChild(this.createMathMLElement("mo", null, [elem.props.char]));
		}

		return result;
	}

	renderMmlBar(elem: OpenXmlElement) {
		const style = {} as any;

		switch (elem.props.position) {
			case "top": style.textDecoration = "overline"; break
			case "bottom": style.textDecoration = "underline"; break
		}

		return this.renderContainerNS(elem, ns.mathML, "mrow", { style }) as MathMLElement;
	}

	renderMmlRun(elem: OpenXmlElement) {
		return this.toHTML(elem, ns.mathML, "ms");
	}

	renderMllList(elem: OpenXmlElement) {
		const children = this.renderElements(elem.children).map(x => this.createMathMLElement("mtr", null, [
			this.createMathMLElement("mtd", null, [x])
		]));

		return this.toHTML(elem, ns.mathML, "mtable", children);
	}

	toH(elem: OpenXmlElement, ns: ns, tagName: string, children: Node[] = null) {
		const { "$lang": lang, ...style } = elem.cssStyle ?? {};
		const className = cx(elem.className, elem.styleName && this.processStyleName(elem.styleName));
		return { ns, tagName, className, lang, style, children: children ?? this.renderElements(elem.children) } as any;
	}

	toHTML(elem: OpenXmlElement, ns: ns, tagName: string, children: Node[] = null) {
		return this.h(this.toH(elem, ns, tagName, children)) as any;
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
			taiwaneseDigital: "cjk-decimal",
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

	createElementNS(ns: any, tagName: string, props?: Partial<Record<any, any>>, children?: any[]) {
		return this.h({ ns, tagName, children, ...props }) as any;
	}

	createElement<T extends keyof HTMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>, children?: any[]): HTMLElementTagNameMap[T] {
		return this.createElementNS(ns.html, tagName, props, children);
	}

	createMathMLElement<T extends keyof MathMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof MathMLElementTagNameMap[T], any>>, children?: any[]): MathMLElementTagNameMap[T] {
		return this.createElementNS(ns.mathML, tagName, props, children);
	}

	createSvgElement<T extends keyof SVGElementTagNameMap>(tagName: T, props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>, children?: any[]): SVGElementTagNameMap[T] {
		return this.createElementNS(ns.svg, tagName, props, children);
	}

	later(func: Function) {
		this.postRenderTasks.push(func);
	}
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
	var parent = elem.parent;

	while (parent != null && parent.type != type)
		parent = parent.parent;

	return <T>parent;
}


import { getLanguage } from 'obsidian';
import type { PluginSettings } from './settings';

export type Scalar = string | number | boolean | null;

export const NUMERIC_RE = /^-?(0|[1-9]\d*)(\.\d+)?$/;

export interface CellBorders {
	top?: boolean;
	right?: boolean;
	bottom?: boolean;
	left?: boolean;
	/** Hex color such as "#ff0000". Missing = follows the Obsidian text color. */
	color?: string;
}

/**
 * What happens to text that is longer than the cell:
 * - overflow: spills into the empty neighbouring cell (default)
 * - wrap:     continues on a new line inside the same cell
 * - expand:   the column is widened to fit the text
 */
export type WrapMode = 'overflow' | 'wrap' | 'expand';

/** Width in px of a column that has no explicit width. Columns never auto-size on their own. */
export const DEFAULT_COL_WIDTH = 100;
export type VerticalAlign = 'top' | 'middle' | 'bottom';

export interface CellFormat {
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	align?: 'left' | 'center' | 'right';
	valign?: VerticalAlign;
	wrap?: WrapMode;
	/** Text rotation in degrees: 0, 45, 90 (up) or -45, -90 (down). */
	rotate?: number;
	/** Font size in px. Missing = the default font size from the settings. */
	fontSize?: number;
	bgColor?: string;
	textColor?: string;
	borders?: CellBorders;
	/** The cell shows a check box; its value is TRUE (checked) or FALSE. */
	checkbox?: boolean;
}

/** A note on a cell, marked by a colored corner and shown next to the cell. */
export interface CellComment {
	text: string;
	/** Color of the corner mark, "#rrggbb". */
	color: string;
	/** When it was written or last changed (ms since 1970). */
	time: number;
}

export const DEFAULT_COMMENT_COLOR = '#f2a900';

/** An image that floats over the grid, like a chart, instead of living in a cell. */
export interface FloatingImage {
	id: string;
	/** Vault path of the image. */
	path: string;
	/** Position and size in px at 100% zoom, measured from the top-left corner of cell A1. */
	x: number;
	y: number;
	w: number;
	h: number;
	/** Stacking order among charts and floating images: higher is drawn on top (missing = 0). */
	z?: number;
}

/** Remembers a cell's size from before an image resized it, so deleting the image can restore it. */
export interface ImageSizeRecord {
	prevW: number | null;
	prevH: number | null;
	/** The size the image set (restored only if the user has not resized the cell since). */
	w: number;
	h: number;
}

export interface MergeInfo {
	row: number;
	col: number;
	rowspan: number;
	colspan: number;
}

export type LegendPosition = 'bottom' | 'top' | 'right' | 'none';

export interface AxisStyle {
	title?: string;
	/** Fixed ends of a numeric axis; missing = from the data. */
	min?: number;
	max?: number;
	/** Missing = shown. */
	showLabels?: boolean;
}

export interface SeriesStyle {
	color?: string;
	/** Scatter charts; missing = the chart's own trendline setting. */
	trendline?: boolean;
	trendColor?: string;
}

/** Chart types that can be inserted from the toolbar. */
export type ChartType = 'column' | 'bar' | 'line' | 'pie' | 'scatter';

export interface ChartSpec {
	id: string;
	type: ChartType;
	title: string;
	/** The values the chart shows (0-based, inclusive): a series per column. */
	range: SelRange;
	/**
	 * The labels (slice names of a pie chart, X values of a scatter chart). When missing and
	 * `range` has two or more columns, its first column holds them.
	 */
	xRange?: SelRange;
	/** There are no labels (the labels field was cleared): every column is a series. */
	noLabels?: boolean;
	/** The first row of the range holds the series names. */
	headerRow: boolean;
	/** Scatter charts: a linear trendline for every series. */
	trendline: boolean;
	/** Color of the trendline ("#rrggbb"); missing = the color of its series. */
	trendColor?: string;
	/** Trendline text in the legend. Missing = shown. */
	showEquation?: boolean;
	showR2?: boolean;
	/* Look (Customize tab of the chart editor). Missing values use the defaults. */
	bg?: string;
	fontSize?: number;
	titleColor?: string;
	pointSize?: number;
	lineWidth?: number;
	/** Pie charts: size of the hole in the middle, in percent of the radius (0 = a full pie). */
	donutHole?: number;
	legend?: LegendPosition;
	/** The axes as they are drawn: horizontal (bottom) and vertical (left). */
	xAxis?: AxisStyle;
	yAxis?: AxisStyle;
	/** Lines across the plot at the value ticks (default on). */
	gridlines?: boolean;
	/** Scatter charts: lines at the X ticks too (default off). */
	xGridlines?: boolean;
	gridColor?: string;
	/** Per series (per slice for pie charts): color and trendline. */
	series?: SeriesStyle[];
	/** Position and size in px at 100% zoom, measured from the top-left corner of cell A1. */
	x: number;
	y: number;
	w: number;
	h: number;
	/** Stacking order among charts and floating images: higher is drawn on top (missing = 0). */
	z?: number;
}

export interface SheetModel {
	version: number;
	/** The name on the sheet's tab (unique within the file). */
	name: string;
	data: Scalar[][];
	cellFormats: Record<string, CellFormat>;
	/** "row,col" -> vault path of an image shown inside the cell. */
	images: Record<string, string>;
	imageSizes: Record<string, ImageSizeRecord>;
	/** "row,col" -> comment. Replaced, never changed in place (undo compares them). */
	comments: Record<string, CellComment>;
	/** "row,col" -> link target: a web address or the name of a note in the vault. */
	links: Record<string, string>;
	/** column index -> width the column had before "widen column" changed it. */
	widenRecords?: Record<string, { prevW: number | null; w: number }>;
	merges: MergeInfo[];
	colWidths: (number | null)[];
	rowHeights: (number | null)[];
	rtl: boolean;
	charts: ChartSpec[];
	floatingImages: FloatingImage[];
}

/** A file holds one or more sheets (tabs). */
export interface WorkbookModel {
	sheets: SheetModel[];
	/** Index of the sheet that was shown when the file was saved. */
	active: number;
	/**
	 * Answers of CLAUDE() and GEMINI() formulas by request, so opening the file again does
	 * not ask (and pay) again. Not part of the undo history.
	 */
	aiResults?: Record<string, string>;
}

/** File format version that stores several sheets. Older files hold one sheet at the top level. */
export const WORKBOOK_VERSION = 3;

/** A normalized (0-based, inclusive, top-left to bottom-right) selection range. */
export interface SelRange {
	r1: number;
	c1: number;
	r2: number;
	c2: number;
}

export const FONT_SIZES: number[] = [
	8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72,
];

export const ROTATIONS: { value: number; label: string }[] = [
	{ value: 0, label: 'No rotation' },
	{ value: 45, label: 'Up 45°' },
	{ value: 90, label: 'Up 90°' },
	{ value: -45, label: 'Down 45°' },
	{ value: -90, label: 'Down 90°' },
];

export const key = (r: number, c: number): string => `${r},${c}`;

/** 0 -> "A", 25 -> "Z", 26 -> "AA" */
export function indexToColumnLabel(index: number): string {
	let n = index + 1;
	let out = '';
	while (n > 0) {
		const rem = (n - 1) % 26;
		out = String.fromCharCode(65 + rem) + out;
		n = Math.floor((n - 1) / 26);
	}
	return out;
}

/**
 * Direction for new sheets: right-to-left when Obsidian (or the system) uses a
 * right-to-left language. Each sheet can be switched with the toolbar button, and
 * imports let you pick the direction.
 */
export function defaultSheetRtl(): boolean {
	const lang = getLanguage() || 'en';
	return /^(he|iw|ar|fa|ur|yi)(-|_|$)/i.test(lang);
}

/** Heights are derived from the font so text is never clipped vertically. */
export const ROW_LINE_HEIGHT = 1.4;
export const ROW_VERTICAL_PADDING = 8;
export const MIN_ZOOM = 0.5;
export const MAX_ZOOM = 3;

export function rowBaseHeight(fontSizePx: number): number {
	return Math.ceil(fontSizePx * ROW_LINE_HEIGHT) + ROW_VERTICAL_PADDING;
}

export function createDefaultModel(
	s: PluginSettings,
	rtl: boolean,
	name = 'Sheet1',
): SheetModel {
	return {
		version: 2,
		name,
		data: Array.from({ length: s.defaultRowCount }, () =>
			Array<Scalar>(s.defaultColCount).fill(null),
		),
		cellFormats: {},
		images: {},
		imageSizes: {},
		comments: {},
		links: {},
		merges: [],
		colWidths: [],
		rowHeights: [],
		rtl,
		charts: [],
		floatingImages: [],
	};
}

export function serializeWorkbook(wb: WorkbookModel): string {
	return JSON.stringify({
		version: WORKBOOK_VERSION,
		active: wb.active,
		sheets: wb.sheets,
		aiResults: wb.aiResults ?? {},
	});
}

/** Characters that Excel does not allow in sheet names, plus the quote used in references. */
const SHEET_NAME_INVALID = /[\\/?*[\]:'!]/g;

export function cleanSheetName(name: string): string {
	return name.replace(SHEET_NAME_INVALID, ' ').replace(/\s+/g, ' ').trim().slice(0, 31);
}

/** "Sheet2", "Sheet3", ... or "Name (2)": the first name that no sheet uses yet. */
export function uniqueSheetName(taken: string[], base: string): string {
	const used = new Set(taken.map((n) => n.toLowerCase()));
	if (!used.has(base.toLowerCase())) return base;
	const numbered = /^(.*?)(\d+)$/.exec(base);
	for (let n = 2; ; n++) {
		const candidate = numbered
			? `${numbered[1] ?? ''}${n}`
			: `${base} (${n})`;
		if (!used.has(candidate.toLowerCase())) return candidate;
	}
}

/** How a sheet is named inside a formula: quoted when it is not a plain word. */
export function sheetRefName(name: string): string {
	return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) &&
		!/^[A-Za-z]{1,3}\d+$/.test(name)
		? name
		: `'${name}'`;
}

/** "B3" -> { row: 2, col: 1 }; null when the text is not a cell reference. */
export function parseCellRef(ref: string): { row: number; col: number } | null {
	const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
	if (!m) return null;
	let col = 0;
	for (const ch of (m[1] ?? '').toUpperCase())
		col = col * 26 + (ch.charCodeAt(0) - 64);
	const row = parseInt(m[2] ?? '0', 10) - 1;
	return row < 0 ? null : { row, col: col - 1 };
}

/** "A1:C5" or "B2" (an optional "Sheet!" in front is left out) -> a normalized range. */
export function parseRangeRef(ref: string): SelRange | null {
	const bare = ref.includes('!') ? ref.slice(ref.lastIndexOf('!') + 1) : ref;
	const [a = '', b] = bare.split(':');
	const from = parseCellRef(a);
	const to = b === undefined ? from : parseCellRef(b);
	if (!from || !to) return null;
	return {
		r1: Math.min(from.row, to.row),
		c1: Math.min(from.col, to.col),
		r2: Math.max(from.row, to.row),
		c2: Math.max(from.col, to.col),
	};
}

export function rangeLabel(rg: SelRange): string {
	const first = `${indexToColumnLabel(rg.c1)}${rg.r1 + 1}`;
	const last = `${indexToColumnLabel(rg.c2)}${rg.r2 + 1}`;
	return first === last ? first : `${first}:${last}`;
}

export function isEmptyFormat(f: CellFormat): boolean {
	if (
		f.bold ||
		f.italic ||
		f.underline ||
		f.strike ||
		f.align ||
		f.valign ||
		f.wrap ||
		f.rotate ||
		f.fontSize ||
		f.bgColor ||
		f.textColor ||
		f.checkbox
	) {
		return false;
	}
	const b = f.borders;
	return !(b && (b.top || b.right || b.bottom || b.left));
}

export function cloneFormat(f: CellFormat): CellFormat {
	return JSON.parse(JSON.stringify(f)) as CellFormat;
}

export type FilterCondition =
	| 'none'
	| 'contains'
	| 'not-contains'
	| 'equals'
	| 'starts'
	| 'ends'
	| 'greater'
	| 'less'
	| 'empty'
	| 'not-empty';

export const FILTER_CONDITIONS: { value: FilterCondition; label: string }[] = [
	{ value: 'none', label: 'No condition' },
	{ value: 'contains', label: 'Contains' },
	{ value: 'not-contains', label: 'Does not contain' },
	{ value: 'equals', label: 'Equals' },
	{ value: 'starts', label: 'Starts with' },
	{ value: 'ends', label: 'Ends with' },
	{ value: 'greater', label: 'Greater than' },
	{ value: 'less', label: 'Less than' },
	{ value: 'empty', label: 'Is empty' },
	{ value: 'not-empty', label: 'Is not empty' },
];

/** A filter on one column. Filters only change the view: they are not saved in the file. */
export interface ColumnFilter {
	/** Cell texts that stay visible. null = every value. */
	values: Set<string> | null;
	condition: FilterCondition;
	operand: string;
}

/** Text order used for sorting and for value lists: "item 2" comes before "item 10". */
export function compareText(a: string, b: string): number {
	return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
}

/** True when a cell showing `text` passes the filter. Text is compared without case. */
export function filterAccepts(filter: ColumnFilter, text: string): boolean {
	if (filter.values && !filter.values.has(text)) return false;
	const value = text.trim().toLowerCase();
	const operand = filter.operand.trim().toLowerCase();
	switch (filter.condition) {
		case 'none':
			return true;
		case 'empty':
			return value === '';
		case 'not-empty':
			return value !== '';
		case 'contains':
			return value.includes(operand);
		case 'not-contains':
			return !value.includes(operand);
		case 'equals':
			return value === operand;
		case 'starts':
			return value.startsWith(operand);
		case 'ends':
			return value.endsWith(operand);
		case 'greater':
		case 'less': {
			if (value === '') return false;
			const bothNumbers =
				NUMERIC_RE.test(value) && NUMERIC_RE.test(operand);
			const order = bothNumbers
				? Number(value) - Number(operand)
				: compareText(value, operand);
			return filter.condition === 'greater' ? order > 0 : order < 0;
		}
	}
}

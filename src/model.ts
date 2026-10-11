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

export interface SheetModel {
	version: number;
	data: Scalar[][];
	cellFormats: Record<string, CellFormat>;
	/** "row,col" -> vault path of an image shown inside the cell. */
	images: Record<string, string>;
	imageSizes: Record<string, ImageSizeRecord>;
	/** column index -> width the column had before "widen column" changed it. */
	widenRecords?: Record<string, { prevW: number | null; w: number }>;
	merges: MergeInfo[];
	colWidths: (number | null)[];
	rowHeights: (number | null)[];
	rtl: boolean;
}

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
	let lang = 'en';
	try {
		lang =
			window.localStorage.getItem('language') ||
			navigator.language ||
			'en';
	} catch {
		lang = typeof navigator !== 'undefined' ? navigator.language : 'en';
	}
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
): SheetModel {
	return {
		version: 2,
		data: Array.from({ length: s.defaultRowCount }, () =>
			Array<Scalar>(s.defaultColCount).fill(null),
		),
		cellFormats: {},
		images: {},
		imageSizes: {},
		merges: [],
		colWidths: [],
		rowHeights: [],
		rtl,
	};
}

export function serializeModel(model: SheetModel): string {
	return JSON.stringify(model);
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
		f.textColor
	) {
		return false;
	}
	const b = f.borders;
	return !(b && (b.top || b.right || b.bottom || b.left));
}

export function cloneFormat(f: CellFormat): CellFormat {
	return JSON.parse(JSON.stringify(f)) as CellFormat;
}

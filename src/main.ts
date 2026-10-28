import {
	debounce,
	KeymapEventHandler,
	Menu,
	Notice,
	Platform,
	Plugin,
	Scope,
	TFile,
	TFolder,
	TextFileView,
	WorkspaceLeaf,
	normalizePath,
	requestUrl,
	setIcon,
} from 'obsidian';
import {
	AI_ERROR_TEXT,
	AiAssistantPanel,
	AiFormulaRunner,
	AiTool,
	AiToolResult,
	JsonSchema,
	aiApiKey,
	generateGeminiImage,
} from './ai';
import Handsontable from 'handsontable';
import 'handsontable/styles/handsontable.min.css';
import 'handsontable/styles/ht-theme-main.min.css';
import './plugin.css';
import ExcelJS from 'exceljs';
import {
	DEFAULT_SETTINGS,
	MAX_RECENT_EMOJIS,
	MAX_SAVED_COLORS,
	PluginSettings,
	SettingsTab,
	sanitizeSettings,
} from './settings';
import {
	CellBorders,
	CellComment,
	CellFormat,
	ChartSpec,
	ChartType,
	DEFAULT_COMMENT_COLOR,
	FloatingImage,
	ColumnFilter,
	DEFAULT_COL_WIDTH,
	FONT_SIZES,
	ImageSizeRecord,
	MergeInfo,
	NUMERIC_RE,
	ROTATIONS,
	Scalar,
	SelRange,
	SheetModel,
	WorkbookModel,
	cleanSheetName,
	cloneFormat,
	MAX_ZOOM,
	MIN_ZOOM,
	ROW_LINE_HEIGHT,
	ROW_VERTICAL_PADDING,
	compareText,
	createDefaultModel,
	defaultSheetRtl,
	filterAccepts,
	indexToColumnLabel,
	isEmptyFormat,
	key,
	parseRangeRef,
	rangeLabel,
	rowBaseHeight,
	serializeWorkbook,
	sheetRefName,
	uniqueSheetName,
} from './model';
import {
	CHART_DEFAULT_H,
	CHART_DEFAULT_W,
	CHART_HEAD_H,
	CHART_MIN_H,
	CHART_MIN_W,
	CHART_TYPES,
	ChartEditor,
	guessHeaderRow,
	normalizeChartExtras,
	readChartData,
	renderChart,
} from './charts';
import {
	FORMULA_FUNCTIONS,
	FormulaFunctionInfo,
	buildFunctionCall,
	evaluateFormula,
	formulaReferences,
	moveFormula,
	normalizeSeparators,
	referenceSlot,
	setAiFormulaHandler,
	shiftFormulaRows,
} from './formulas';
import { ImportedSheet, parseCsv, parseXlsx } from './xlsx';
import {
	ColumnFilterModal,
	CommentModal,
	ConfirmModal,
	EmojiPicker,
	FormulaAutocomplete,
	FunctionPickerModal,
	ImportChoice,
	ImportDirectionModal,
	ImportPrompt,
	LinkModal,
	TextPromptModal,
	VaultFilePickerModal,
} from './ui';

/* ------------------------------------------------------------------ */
/* Constants                                                           */
/* ------------------------------------------------------------------ */

const VIEW_TYPE = 'rtl-spreadsheet-view';
const SHEET_EXT = 'sheet';
/** Used automatically when another plugin already owns the .sheet extension. */
const FALLBACK_SHEET_EXT = 'rtlsheet';
const IMAGE_EXTENSIONS = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg'];
/** Most image paths the find_images tool returns. */
const AI_MAX_FOUND_IMAGES = 30;
/** Results of one web image search, and the width of the copies it links to (in px). */
const AI_MAX_WEB_IMAGES = 12;
const AI_WEB_IMAGE_WIDTH = 800;
const MAX_IMAGE_CELL_PX = 240;
/** Limits (unzoomed px) for a cell that is resized by dragging a corner of its image. */
const MIN_IMAGE_RESIZE_PX = 24;
const MAX_IMAGE_RESIZE_PX = 2000;
/** The corners of an image cell that have a resize handle (physical sides). */
const IMAGE_CORNERS = ['nw', 'ne', 'sw', 'se'] as const;
/** Dragging a selection handle this close (px) to an edge of the grid scrolls the sheet. */
const SEL_SCROLL_EDGE = 36;
/** Px scrolled per frame while a selection handle is held near an edge. */
const SEL_SCROLL_STEP = 8;
/** Extra px added to measured text when a column is widened to fit it (cell padding, border, rounding). */
const EXPAND_PADDING = 24;
const MAX_EXPAND_WIDTH = 3000;
/** Horizontal padding inside a cell, used when measuring text. */
const CELL_PADDING = 8;
/** Width of the column of row numbers at 100% (Handsontable's default). */
const ROW_HEADER_WIDTH = 50;

type BorderMode =
	| 'all'
	| 'outer'
	| 'top'
	| 'bottom'
	| 'left'
	| 'right'
	| 'none';

/** True on Obsidian mobile and on any device whose primary pointer is touch. */
const IS_TOUCH: boolean =
	Platform.isMobile ||
	(typeof window !== 'undefined' &&
		window.matchMedia('(pointer: coarse)').matches);

/** The text of a cell's source value, as it is shown for editing. */
function sourceText(value: unknown): string {
	if (typeof value === 'string') return value;
	if (typeof value === 'number' || typeof value === 'boolean')
		return String(value);
	return '';
}

/** A copied block of cells (Ctrl+C or Ctrl+X): the text that went to the clipboard and the formats. */
interface CopiedCells {
	text: string[][];
	formats: (CellFormat | undefined)[][];
	/** Where the cells came from, outlined with moving dashes until a paste, Esc or typing. */
	mark: { view: SheetView; sheet: number; range: SelRange } | null;
}

/**
 * The cells copied last, in any sheet. The clipboard only carries text, so a paste whose
 * text is the same puts these formats (check boxes too) on the pasted cells.
 */
let copiedCells: CopiedCells | null = null;

const clipboardText = (data: unknown[][]): string[][] =>
	data.map((row) => row.map((v) => sourceText(v).trim()));

/** Handsontable theme that follows the active Obsidian theme (light/dark). */
function handsontableThemeName(): string {
	return document.body.classList.contains('theme-dark')
		? 'ht-theme-main-dark'
		: 'ht-theme-main';
}

function toArgb(hex: string): string {
	let h = hex.replace('#', '');
	if (h.length === 3)
		h = h
			.split('')
			.map((ch) => ch + ch)
			.join('');
	return `FF${h.toUpperCase()}`;
}

function sanitizeSheetName(name: string): string {
	const cleaned = name.replace(/[\\/?*[\]:]/g, ' ').trim();
	return (cleaned || 'Sheet1').slice(0, 31);
}

type BorderSide = 'top' | 'right' | 'bottom' | 'left';

/** What the border renderer needs to know about the grid around a cell; built once per render. */
interface BorderGrid {
	/** The merge that covers each merged cell, by "row,col". */
	merges: Map<string, MergeInfo>;
	isHidden(row: number): boolean;
	/** The next row down that is not hidden by a filter, or null. */
	nextRow(row: number): number | null;
	/** The first row that is not hidden: it draws the line under the column headers. */
	firstRow: number;
	/** Turns a stored color (or '' for the grid line color) into one the corner image can use. */
	resolve(color: string): string;
}

/** The color of one side of a cell's borders, or null when that side has no border. */
function borderColor(b: CellBorders | undefined, side: BorderSide): string | null {
	return b?.[side] ? (b.color ?? 'var(--text-normal)') : null;
}

function firstColor(colors: (string | null)[]): string | null {
	return colors.find((c): c is string => c !== null) ?? null;
}

function span(from: number, to: number): number[] {
	const out: number[] = [];
	for (let i = from; i <= to; i++) out.push(i);
	return out;
}

/**
 * Draws a cell's borders on the grid lines themselves, like Excel. Every cell owns its bottom
 * and inline-end grid line (the cell's own borders), so a border between two cells is one line,
 * whichever cell it was set on. A filled cell keeps its grid lines (they only give way to a
 * border, or to text that spills over them). The first row also owns the line under the column headers and the first column the line next to
 * the row headers (they come from the cells' top and inline-start borders). A merged cell owns
 * the lines along its whole bottom and inline-end side, and rows hidden by a filter are skipped.
 */
function applyBorderStyles(
	td: HTMLTableCellElement,
	grid: BorderGrid,
	formats: Record<string, CellFormat>,
	row: number,
	col: number,
	rtl: boolean,
): void {
	const st = td.style;
	const start: BorderSide = rtl ? 'right' : 'left';
	const end: BorderSide = rtl ? 'left' : 'right';
	const side = (r: number, c: number, s: BorderSide): string | null =>
		borderColor(formats[key(r, c)]?.borders, s);
	const mergeOf = (r: number, c: number): MergeInfo | undefined =>
		grid.merges.get(key(r, c));
	const sameMerge = (r1: number, c1: number, r2: number, c2: number): boolean => {
		const m = mergeOf(r1, c1);
		return m !== undefined && m === mergeOf(r2, c2);
	};
	// The grid line under cell (r, c) and the one at its inline end, shared with the neighbors.
	// There is no line between two cells of the same merge.
	const lineBelow = (r: number, c: number): string | null => {
		const next = grid.nextRow(r);
		if (next === null) return side(r, c, 'bottom');
		if (sameMerge(r, c, next, c)) return null;
		return side(r, c, 'bottom') ?? side(next, c, 'top');
	};
	const lineAtEnd = (r: number, c: number): string | null => {
		if (sameMerge(r, c, r, c + 1)) return null;
		return side(r, c, end) ?? side(r, c + 1, start);
	};

	const m = mergeOf(row, col);
	const isAnchor = m !== undefined && m.row === row && m.col === col;
	const lastRow = isAnchor ? m.row + m.rowspan - 1 : row;
	const lastCol = isAnchor ? m.col + m.colspan - 1 : col;
	const below = grid.nextRow(lastRow);
	const cols = span(col, lastCol);
	const rows = span(row, lastRow).filter((r) => !grid.isHidden(r));
	const isTop = row === grid.firstRow;

	const bottom = firstColor(cols.map((c) => lineBelow(lastRow, c)));
	const atEnd = firstColor(rows.map((r) => lineAtEnd(r, lastCol)));
	const top = isTop ? firstColor(cols.map((c) => side(row, c, 'top'))) : null;
	const atStart = col === 0 ? firstColor(rows.map((r) => side(r, 0, start))) : null;

	const bottomColor = bottom ?? '';
	// Spilled text has made the inline-end line transparent; a border still shows over it.
	const endColor =
		atEnd ?? (st.borderInlineEndColor === 'transparent' ? 'transparent' : '');
	st.borderBottomColor = bottomColor;
	st.borderInlineEndColor = endColor;
	st.borderTopColor = top ?? '';
	st.borderInlineStartColor = atStart ?? '';
	st.borderImage = '';

	// The browser splits a corner pixel between the cell's two sides: when they differ in color
	// the corner is a blend, and a border that only the neighbors draw never reaches it. When a
	// border meets a corner the sides cannot paint whole, draw the frame as a 3×3 image instead.
	const corners = {
		bottomEnd: firstColor([
			bottom,
			atEnd,
			lineBelow(lastRow, lastCol + 1),
			below === null ? null : lineAtEnd(below, lastCol),
		]),
		topEnd: isTop ? firstColor([top, atEnd, side(row, lastCol + 1, 'top')]) : null,
		bottomStart:
			col === 0
				? firstColor([atStart, bottom, below === null ? null : side(below, 0, start)])
				: null,
		topStart: isTop && col === 0 ? firstColor([top, atStart]) : null,
	};
	const whole = (want: string | null, a: string | null, b: string | null): boolean =>
		want === null || (a === want && b === want);
	if (
		whole(corners.bottomEnd, bottom, atEnd) &&
		whole(corners.topEnd, top, atEnd) &&
		whole(corners.bottomStart, atStart, bottom) &&
		whole(corners.topStart, top, atStart)
	)
		return;

	// The colors the sides are drawn in, as set above (no computed style, which is slow here).
	const startColor = grid.resolve(atStart ?? '');
	const edge = {
		top: grid.resolve(top ?? ''),
		bottom: grid.resolve(bottomColor),
		left: rtl ? grid.resolve(endColor) : startColor,
		right: rtl ? startColor : grid.resolve(endColor),
	};
	const corner = (want: string | null, fallback: string): string =>
		want === null ? fallback : grid.resolve(want);
	const px = {
		topLeft: corner(rtl ? corners.topEnd : corners.topStart, edge.top),
		topRight: corner(rtl ? corners.topStart : corners.topEnd, edge.top),
		bottomLeft: corner(rtl ? corners.bottomEnd : corners.bottomStart, edge.bottom),
		bottomRight: corner(rtl ? corners.bottomStart : corners.bottomEnd, edge.bottom),
	};
	const dot = (x: number, y: number, color: string): string =>
		`<rect x="${x}" y="${y}" width="1" height="1" fill="${color}"/>`;
	const svg =
		'<svg xmlns="http://www.w3.org/2000/svg" width="3" height="3" shape-rendering="crispEdges">' +
		dot(0, 0, px.topLeft) +
		dot(1, 0, edge.top) +
		dot(2, 0, px.topRight) +
		dot(0, 1, edge.left) +
		dot(2, 1, edge.right) +
		dot(0, 2, px.bottomLeft) +
		dot(1, 2, edge.bottom) +
		dot(2, 2, px.bottomRight) +
		'</svg>';
	st.borderImage = `url("data:image/svg+xml,${encodeURIComponent(svg)}") 1 stretch`;
}

/** Our rotation values (+up / -down) to Excel's textRotation (0-90 up, 91-180 down). */
function rotationToExcel(rotate: number | undefined): number | undefined {
	if (!rotate) return undefined;
	if (rotate > 0) return rotate;
	return 90 + Math.abs(rotate);
}

/** True for fields the user types text into, where Ctrl+Z belongs to the field itself. */
function isTextEntry(el: HTMLElement): boolean {
	if (el.isContentEditable || el.tagName === 'TEXTAREA') return true;
	if (el.tagName !== 'INPUT') return false;
	return !['color', 'checkbox', 'radio', 'button', 'range', 'file'].includes(
		(el as HTMLInputElement).type,
	);
}

/** Sort order like Excel: numbers, then text, then TRUE/FALSE; empty cells always come last. */
function compareForSort(a: Scalar, b: Scalar, descending: boolean): number {
	const blankA = a === null || a === '';
	const blankB = b === null || b === '';
	if (blankA || blankB) return Number(blankA) - Number(blankB);
	const rank = (v: Scalar): number =>
		typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
	let order = rank(a) - rank(b);
	if (order === 0) {
		order =
			typeof a === 'string' && typeof b === 'string'
				? compareText(a, b)
				: Number(a) - Number(b);
	}
	return descending ? -order : order;
}

/** The image type ExcelJS can embed for a vault file, or null (webp and svg cannot go in .xlsx). */
function excelImageType(file: unknown): 'png' | 'jpeg' | 'gif' | null {
	if (!(file instanceof TFile)) return null;
	const ext = file.extension.toLowerCase();
	if (ext === 'png') return 'png';
	if (ext === 'jpg' || ext === 'jpeg') return 'jpeg';
	if (ext === 'gif') return 'gif';
	return null;
}

function decodeText(data: ArrayBuffer): string {
	try {
		return new TextDecoder('utf-8', { fatal: true }).decode(data);
	} catch {
		// Excel on Windows often saves Hebrew CSV files as Windows-1255.
		try {
			return new TextDecoder('windows-1255').decode(data);
		} catch {
			return new TextDecoder().decode(data);
		}
	}
}

/* ------------------------------------------------------------------ */
/* Plugin                                                              */
/* ------------------------------------------------------------------ */

export default class RtlSpreadsheetPlugin extends Plugin {
	settings: PluginSettings = { ...DEFAULT_SETTINGS };
	/** The file extension actually registered for spreadsheet files. */
	sheetExt: string = SHEET_EXT;

	async onload(): Promise<void> {
		await this.loadSettings();

		this.registerView(VIEW_TYPE, (leaf) => new SheetView(leaf, this));
		this.registerSheetExtension();

		this.addRibbonIcon('table', 'New spreadsheet', () => {
			void this.createNewSheet();
		});

		this.addCommand({
			id: 'create-new-sheet',
			name: 'Create new spreadsheet',
			callback: () => {
				void this.createNewSheet();
			},
		});

		// "New spreadsheet" next to "New note" and "New canvas" in the file explorer menu.
		this.registerEvent(
			this.app.workspace.on('file-menu', (menu, file) => {
				const folder = file instanceof TFolder ? file : file.parent;
				if (!folder) return;
				menu.addItem((item) =>
					item
						.setTitle('New spreadsheet')
						.setIcon('table')
						.setSection('action-primary')
						.onClick(() => {
							void this.createNewSheet(folder);
						}),
				);
			}),
		);

		this.addCommand({
			id: 'find-in-active-sheet',
			name: 'Find in spreadsheet',
			checkCallback: (checking) => {
				const view = this.app.workspace.getActiveViewOfType(SheetView);
				if (!view) return false;
				if (!checking) view.openFind();
				return true;
			},
		});

		this.addCommand({
			id: 'save-active-sheet',
			name: 'Save active spreadsheet',
			checkCallback: (checking) => {
				const view = this.app.workspace.getActiveViewOfType(SheetView);
				if (!view) return false;
				if (!checking) void view.saveNow();
				return true;
			},
		});

		this.addCommand({
			id: 'export-active-sheet-xlsx',
			name: 'Export active spreadsheet to Excel (.xlsx)',
			checkCallback: (checking) => {
				const view = this.app.workspace.getActiveViewOfType(SheetView);
				if (!view) return false;
				if (!checking) void view.exportToXlsx();
				return true;
			},
		});

		this.addSettingTab(new SettingsTab(this.app, this));
	}

	/**
	 * registerExtensions throws if the extension is already owned by another
	 * plugin (or by Obsidian). In that case fall back to a unique extension
	 * instead of failing to load the whole plugin.
	 */
	private registerSheetExtension(): void {
		try {
			this.registerExtensions([SHEET_EXT], VIEW_TYPE);
			this.sheetExt = SHEET_EXT;
		} catch (e) {
			console.warn(
				`[Free-Excel] ".${SHEET_EXT}" is already registered, using ".${FALLBACK_SHEET_EXT}"`,
				e,
			);
			try {
				this.registerExtensions([FALLBACK_SHEET_EXT], VIEW_TYPE);
				this.sheetExt = FALLBACK_SHEET_EXT;
				new Notice(
					`Free-Excel: the ".${SHEET_EXT}" extension is used by another plugin. ` +
						`New spreadsheets will use ".${FALLBACK_SHEET_EXT}" instead.`,
					10000,
				);
			} catch (e2) {
				this.reportError(
					'Could not register the spreadsheet file extension',
					e2,
				);
			}
		}
	}

	onunload(): void {
		for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE)) {
			if (leaf.view instanceof SheetView) leaf.view.destroyGrid();
		}
	}

	async loadSettings(): Promise<void> {
		const stored =
			(await this.loadData()) as Partial<PluginSettings> | null;
		this.settings = sanitizeSettings(stored);
	}

	async saveSettings(): Promise<void> {
		this.settings = sanitizeSettings(this.settings);
		await this.saveData(this.settings);
		for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE)) {
			if (leaf.view instanceof SheetView) leaf.view.applySettings();
		}
	}

	reportError(message: string, error: unknown): void {
		console.error(`[Free-Excel] ${message}`, error);
		const detail = error instanceof Error ? error.message : String(error);
		new Notice(`${message}: ${detail}`);
	}

	/** Creates a sheet file in `folder`, or where Obsidian puts new notes. */
	async createNewSheet(target?: TFolder): Promise<void> {
		try {
			const active = this.app.workspace.getActiveFile();
			const folder =
				target ??
				this.app.fileManager.getNewFileParent(active?.path ?? '');
			let path = normalizePath(
				`${folder.path}/Untitled sheet.${this.sheetExt}`,
			);
			let n = 1;
			while (this.app.vault.getAbstractFileByPath(path)) {
				path = normalizePath(
					`${folder.path}/Untitled sheet ${n++}.${this.sheetExt}`,
				);
			}
			const file = await this.app.vault.create(
				path,
				serializeWorkbook({
					sheets: [
						createDefaultModel(this.settings, defaultSheetRtl()),
					],
					active: 0,
				}),
			);
			await this.app.workspace.getLeaf(true).openFile(file);
		} catch (e) {
			this.reportError('Could not create spreadsheet', e);
		}
	}
}

/* ------------------------------------------------------------------ */
/* View                                                                */
/* ------------------------------------------------------------------ */

const baseTextRenderer = Handsontable.renderers.getRenderer('text');

const CHART_ICONS: Record<ChartSpec['type'], string> = {
	column: 'chart-column',
	bar: 'chart-bar',
	line: 'chart-line',
	pie: 'chart-pie',
	scatter: 'chart-scatter',
};

function newChartId(prefix = 'chart'): string {
	return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/** Smallest floating image (unzoomed px). */
const FLOAT_MIN_PX = 16;
/** Widest a floating image is when it is inserted (unzoomed px). */
const FLOAT_DEFAULT_MAX_PX = 320;

/** True for a web address (or another URL scheme) rather than a note name. */
function isExternalLink(target: string): boolean {
	return /^[a-z][a-z0-9+.-]*:/i.test(target);
}

/** A cell value as a check box reads it: TRUE, "TRUE", 1 and "1" are checked. */
function isChecked(value: unknown): boolean {
	if (typeof value === 'boolean') return value;
	if (typeof value === 'number') return value !== 0;
	if (typeof value === 'string')
		return /^(true|1)$/i.test(value.trim());
	return false;
}

/* ---------------- AI assistant tools ---------------- */

/** How much of the sheet goes with every question to the assistant. */
const AI_CONTEXT_ROWS = 100;
const AI_CONTEXT_COLS = 30;
const AI_CONTEXT_CELL_CHARS = 120;
/** Most cells one read_range call returns. */
const AI_MAX_READ_CELLS = 3000;

const RANGE_PARAM: JsonSchema = {
	type: 'string',
	description: 'Cells in A1 notation, such as "B2:D10" or "C5".',
};
const CELL_PARAM: JsonSchema = {
	type: 'string',
	description: 'One cell in A1 notation, such as "C5".',
};

/** What the assistant can do with the shown sheet (see SheetView.runAiTool). */
const AI_TOOLS: AiTool[] = [
	{
		name: 'read_range',
		description:
			'Reads cells. Returns the rows of the range; a formula cell is shown as "=FORMULA → result".',
		parameters: {
			type: 'object',
			properties: {
				range: RANGE_PARAM,
				sheet: {
					type: 'string',
					description: 'Name of another sheet of the file; the shown sheet when missing.',
				},
			},
			required: ['range'],
		},
	},
	{
		name: 'write_cells',
		description:
			'Writes values or formulas into the shown sheet, starting at a cell. Each inner list is one row, left to right. "" empties a cell; formulas start with "=". The grid grows when needed.',
		parameters: {
			type: 'object',
			properties: {
				start: { type: 'string', description: 'Top-left cell, such as "A1".' },
				rows: {
					type: 'array',
					description: 'Rows of cell contents.',
					items: { type: 'array', items: { type: 'string' } },
				},
			},
			required: ['start', 'rows'],
		},
	},
	{
		name: 'clear_range',
		description: 'Empties the cells of a range of the shown sheet (formats stay).',
		parameters: {
			type: 'object',
			properties: { range: RANGE_PARAM },
			required: ['range'],
		},
	},
	{
		name: 'format_range',
		description:
			'Formats cells of the shown sheet. Only the given properties change; "" removes a color or an alignment.',
		parameters: {
			type: 'object',
			properties: {
				range: RANGE_PARAM,
				bold: { type: 'boolean' },
				italic: { type: 'boolean' },
				underline: { type: 'boolean' },
				strikethrough: { type: 'boolean' },
				font_size: { type: 'number', description: 'In px; 0 = the default size.' },
				text_color: { type: 'string', description: '"#rrggbb"' },
				fill_color: { type: 'string', description: 'Background, "#rrggbb"' },
				align: {
					type: 'string',
					description: 'left, center or right; "" for the default.',
				},
				vertical_align: {
					type: 'string',
					description: 'top, middle or bottom; "" for the default.',
				},
				wrap: {
					type: 'string',
					enum: ['overflow', 'wrap', 'expand'],
					description: 'Long text: spill into empty neighbours, wrap inside the cell, or widen the column.',
				},
				borders: { type: 'string', enum: ['all', 'outer', 'none'] },
				checkbox: {
					type: 'boolean',
					description: 'Show check boxes; their values are TRUE (checked) and FALSE.',
				},
			},
			required: ['range'],
		},
	},
	{
		name: 'set_comment',
		description: 'Adds, changes or (with empty text) removes the comment of a cell.',
		parameters: {
			type: 'object',
			properties: {
				cell: CELL_PARAM,
				text: { type: 'string' },
				color: { type: 'string', description: 'Corner mark color, "#rrggbb" (optional).' },
			},
			required: ['cell', 'text'],
		},
	},
	{
		name: 'set_link',
		description:
			'Makes a cell a link to a web address or to a note in the vault; an empty url removes the link.',
		parameters: {
			type: 'object',
			properties: {
				cell: CELL_PARAM,
				url: { type: 'string', description: 'https://… or the name of a note.' },
				text: { type: 'string', description: 'What the cell shows (optional).' },
			},
			required: ['cell', 'url'],
		},
	},
	{
		name: 'insert_chart',
		description:
			'Inserts a chart of a range of the shown sheet (with its header row and label column).',
		parameters: {
			type: 'object',
			properties: {
				type: { type: 'string', enum: ['column', 'bar', 'line', 'pie', 'scatter'] },
				range: RANGE_PARAM,
				title: { type: 'string' },
			},
			required: ['type', 'range'],
		},
	},
	{
		name: 'find_images',
		description:
			'Looks for image files in the vault whose path contains the given words (in any order). Returns vault paths for insert_image; an empty query lists the newest images.',
		parameters: {
			type: 'object',
			properties: {
				query: { type: 'string', description: 'Words of the file name or folder, such as "money".' },
			},
			required: ['query'],
		},
	},
	{
		name: 'search_web_images',
		description:
			'Searches Wikimedia Commons (free pictures) and returns image addresses for insert_image, with their titles and sizes. Search words in English work best.',
		parameters: {
			type: 'object',
			properties: {
				query: { type: 'string', description: 'What the picture shows, such as "banknotes".' },
			},
			required: ['query'],
		},
	},
	{
		name: 'insert_image',
		description:
			'Puts an image from the vault (a path from find_images) or from a web address (https://…, downloaded into the vault) in a cell, or over the sheet as a floating image.',
		parameters: {
			type: 'object',
			properties: {
				source: { type: 'string', description: 'Vault path of an image, or an https:// address of an image file.' },
				cell: {
					type: 'string',
					description: 'The cell, such as "C5"; the selected cell when missing. A floating image is placed at its top-left corner.',
				},
				floating: { type: 'boolean', description: 'true: float over the sheet; false or missing: inside the cell.' },
			},
			required: ['source'],
		},
	},
	{
		name: 'generate_image',
		description:
			'Makes a new picture from a description with the Gemini image model (needs the Gemini API key), saves it in the vault and puts it in a cell or over the sheet. Use only when the user wants a new picture made.',
		parameters: {
			type: 'object',
			properties: {
				prompt: { type: 'string', description: 'What the picture shows, in detail (in English works best).' },
				cell: {
					type: 'string',
					description: 'The cell, such as "C5"; the selected cell when missing.',
				},
				floating: { type: 'boolean', description: 'true: float over the sheet; false or missing: inside the cell.' },
			},
			required: ['prompt'],
		},
	},
	{
		name: 'select_range',
		description: 'Selects and scrolls to cells, to show them to the user.',
		parameters: {
			type: 'object',
			properties: { range: RANGE_PARAM },
			required: ['range'],
		},
	},
];

/** The color pickers in the toolbar; fill and text also have saved colors. */
type ColorKind = 'fill' | 'text' | 'border';
type PaletteKind = 'fill' | 'text';

interface ToolbarRefs {
	toggles: Record<string, HTMLButtonElement>;
	fontSelect?: HTMLSelectElement;
	rotateSelect?: HTMLSelectElement;
	brush?: HTMLButtonElement;
	address?: HTMLElement;
	formulaInput?: HTMLInputElement;
	zoomLabel?: HTMLButtonElement;
	/** The saved-color swatches next to the fill and text color pickers. */
	palettes: Partial<
		Record<
			PaletteKind,
			{
				el: HTMLElement;
				input: HTMLInputElement;
				setColor: (f: CellFormat, color: string) => void;
			}
		>
	>;
	filterBtn?: HTMLButtonElement;
	findInput?: HTMLInputElement;
	findCount?: HTMLElement;
}

interface MenuEntry {
	title: string;
	icon: string;
	onClick: () => void;
	/** Shows a check mark next to an on/off entry. */
	checked?: boolean;
}

/** A step on Handsontable's undo stack: one of its own, or one that this plugin added. */
interface UndoStep {
	actionType: string;
	undo(hot: Handsontable.Core, done: () => void): void;
	redo(hot: Handsontable.Core, done: () => void): void;
}

/** The parts of Handsontable's UndoRedo plugin that are used here. */
interface UndoRedoPlugin {
	undo(): void;
	redo(): void;
	clear(): void;
	done(wrappedAction: () => UndoStep, source?: string): void;
	ignoreNewActions: boolean;
	doneActions: unknown[];
	undoneActions: unknown[];
}

/**
 * The saved parts of the model that Handsontable knows nothing about, and so cannot undo:
 * everything except the cell values, the merges and the direction.
 */
interface SideState {
	formats: Record<string, CellFormat>;
	images: Record<string, string>;
	imageSizes: Record<string, ImageSizeRecord>;
	comments: Record<string, CellComment>;
	links: Record<string, string>;
	/** As JSON: these records are changed in place. */
	widenRecords: string;
	/** As JSON: a chart is moved by changing its position. */
	charts: string;
	floatingImages: string;
	colWidths: (number | null)[];
	rowHeights: (number | null)[];
}

/** Formats and image sizes are compared by identity: they are replaced, never changed in place. */
function sameSideState(a: SideState, b: SideState): boolean {
	const sameRecord = (
		x: Record<string, unknown>,
		y: Record<string, unknown>,
	): boolean => {
		const keys = Object.keys(x);
		return (
			keys.length === Object.keys(y).length &&
			keys.every((k) => x[k] === y[k])
		);
	};
	const sameSizes = (
		x: (number | null)[],
		y: (number | null)[],
	): boolean => {
		for (let i = Math.max(x.length, y.length) - 1; i >= 0; i--)
			if ((x[i] || null) !== (y[i] || null)) return false;
		return true;
	};
	return (
		sameRecord(a.formats, b.formats) &&
		sameRecord(a.images, b.images) &&
		sameRecord(a.imageSizes, b.imageSizes) &&
		sameRecord(a.comments, b.comments) &&
		sameRecord(a.links, b.links) &&
		a.widenRecords === b.widenRecords &&
		a.charts === b.charts &&
		a.floatingImages === b.floatingImages &&
		sameSizes(a.colWidths, b.colWidths) &&
		sameSizes(a.rowHeights, b.rowHeights)
	);
}

/** Everything that belongs to a set of rows, so the rows can be moved and put back. */
interface RowsSnapshot {
	rows: number[];
	data: Scalar[][];
	formats: Record<string, CellFormat>;
	images: Record<string, string>;
	imageSizes: Record<string, ImageSizeRecord>;
	comments: Record<string, CellComment>;
	links: Record<string, string>;
	heights: (number | null)[];
}

/** How far a cell's text spills past its own cell. */
interface SpillInfo {
	visualRight: boolean;
	/** Px (zoomed) the text reaches past its own cell. */
	extra: number;
}

/** The spills of one row, and the columns whose inline-end grid line spilled text crosses. */
interface RowSpill {
	spills: Map<number, SpillInfo>;
	crossed: Set<number>;
}

export class SheetView extends TextFileView {
	private plugin: RtlSpreadsheetPlugin;
	/** The sheet that is shown: always `sheets[activeSheet]`. */
	private model: SheetModel;
	/** Every sheet (tab) of the file. */
	private sheets: SheetModel[];
	private activeSheet = 0;
	private tabBarEl: HTMLElement;
	/** Charts of the shown sheet, drawn inside the scrolled part of the grid. */
	private chartLayer: HTMLElement | null = null;
	private chartFrame = 0;
	/** Panel above the grid where a chart is created or changed. */
	private chartPanelEl: HTMLElement;
	private chartEditor: ChartEditor | null = null;
	/** The chart that the editor changes. */
	private chartEditorId: string | null = null;
	/** The ranges in the chart editor, outlined in the sheet. */
	private chartPickRanges: SelRange[] = [];
	/** Ctrl+G, registered in the view's scope while the setting is on. */
	private findHotkey: KeymapEventHandler | null = null;
	private hot: Handsontable | null = null;
	/** Grid instance used for formula lookups while Handsontable is still constructing. */
	private gridForCalc: Handsontable.Core | null = null;
	private formulaBarEl: HTMLElement;
	private toolbarEl: HTMLElement;
	private findBarEl: HTMLElement;
	private gridEl: HTMLElement;
	private loading = false;
	private dirty = false;
	private appliedSettings: PluginSettings;
	/** File content this view last loaded; a different content on disk means an outside change. */
	private diskData = '';
	private refreshTimer = 0;
	private refreshing = false;

	private readonly ui: ToolbarRefs = { toggles: {}, palettes: {} };
	private readonly autocomplete = new FormulaAutocomplete();
	/** Top-left cell of the current selection (shown in the formula bar). */
	private formulaCell: { row: number; col: number } | null = null;
	/** Cell that the formula bar was editing when it received focus. */
	private formulaEditCell: { row: number; col: number } | null = null;
	/** Reference that was put into the formula being typed by clicking (or dragging over) cells. */
	private refPick: {
		field: HTMLInputElement | HTMLTextAreaElement;
		from: number;
		to: number;
		anchor: { row: number; col: number };
		range: SelRange | null;
	} | null = null;
	private writingRefPick = false;
	/** The cells that the formula of the selected cell refers to; they get the same dashed outline. */
	private formulaRefs: SelRange[] = [];
	/** Formatting copied by the format painter; null when the painter is off. */
	private painterSource: CellFormat | null = null;
	/** The side state as it was after the last recorded action (see recordSideChanges). */
	private sideBaseline: SideState | null = null;
	private sideCheckQueued = false;
	/** True while one of our undo steps is running: nothing is recorded meanwhile. */
	private sideBusy = false;
	/** Undo steps that were already looked at, so new ones can be told apart. */
	private knownSteps = new WeakSet<object>();
	/** Set before markDirty() when the change may extend the previous step (same tag, in time). */
	private coalesce: { tag: string; windowMs: number } | null = null;
	private lastSideStep: {
		step: UndoStep;
		tag: string;
		time: number;
		state: { before: SideState; after: SideState };
	} | null = null;

	/** Column filters (column index -> filter). They hide rows in the view and are not saved. */
	private filters = new Map<number, ColumnFilter>();
	/** Sorting and filtering leave the first row in place when it holds the column titles. */
	private hasHeaderRow = true;
	/** Cells that match the find bar: all of them in row order, and their "row,col" keys. */
	private findMatches: { row: number; col: number }[] = [];
	private findHits = new Set<string>();
	private findIndex = -1;
	/** Set when cells changed while the find bar was open; the next step searches again. */
	private findStale = false;

	/** View zoom (0.5-3) of the column widths; row heights and fonts use zoomY. Not saved in the file. */
	private zoom = 1;
	private pinch: { startDistance: number; startZoom: number } | null = null;
	private pendingZoom = 1;
	private zoomFrame = 0;
	/** Touch only: the corner handles of the selection (see installSelectionHandles). */
	private selHandles: { el: HTMLElement; rowEnd: boolean; colEnd: boolean }[] =
		[];
	private selHandleFrame = 0;
	private selDragging = false;
	/** row -> columns that have a format (rebuilt lazily; used to compute row heights). */
	private rowIndex: Map<number, number[]> | null = null;
	/** Computed row heights (unzoomed), kept until the sheet changes: Handsontable asks for them all the time. */
	private rowHeightCache = new Map<number, number>();
	/** Formula results by cell, kept until the sheet changes, so a formula is worked out once and not once per cell that uses it. */
	private formulaCache = new Map<string, { raw: string; value: Scalar }>();
	/** Counts evaluations that were cut short by the depth limit; such results are not cached. */
	private formulaLimitHits = 0;
	/** Spilled text by row, kept until the sheet changes (like the formula results). */
	private spillCache = new Map<number, RowSpill>();
	/** Merges, hidden rows and colors for drawing borders, kept like the spilled text. */
	private borderGridCache: BorderGrid | null = null;
	private fontFamilyCache: string | null = null;
	private measureCtx: CanvasRenderingContext2D | null = null;

	/** Answers of AI formulas, saved with the file (see WorkbookModel.aiResults). */
	private aiResults: Record<string, string> = {};
	private readonly aiRunner: AiFormulaRunner;
	/** Next to the grid: the list of comments or the AI assistant. */
	private sidePanelEl: HTMLElement;
	private commentsPanelEl: HTMLElement;
	private aiPanelEl: HTMLElement;
	private sidePanel: 'comments' | 'ai' | null = null;
	/** Kept while the view is open, so closing the panel does not lose the conversation. */
	private aiPanel: AiAssistantPanel | null = null;
	private readonly emojiPicker = new EmojiPicker();
	/** Popup next to a cell that shows its comment and link. */
	private cellInfoEl: HTMLElement;
	/** The cell the popup shows. */
	private cellInfoCell: { row: number; col: number } | null = null;
	/** The cell under the mouse pointer. */
	private hoverCell: { row: number; col: number } | null = null;
	/** True while the pointer is over the popup (so its links can be clicked). */
	private cellInfoHovered = false;
	private cellInfoTimer = 0;
	/** The floating image that was clicked last; Delete removes it. */
	private activeFloatId: string | null = null;
	/** A press on a check box: where it started, to tell a tap from a scroll. */
	private checkboxPress: {
		row: number;
		col: number;
		x: number;
		y: number;
		pointerId: number;
	} | null = null;
	/** Set while a paste of copied cells runs: their formats go on the pasted cells. */
	private pasteFormats: CopiedCells | null = null;

	constructor(leaf: WorkspaceLeaf, plugin: RtlSpreadsheetPlugin) {
		super(leaf);
		this.plugin = plugin;
		this.model = createDefaultModel(plugin.settings, defaultSheetRtl());
		this.sheets = [this.model];
		this.appliedSettings = { ...plugin.settings };
		this.aiRunner = new AiFormulaRunner(
			this.app,
			() => this.plugin.settings,
			() => this.aiResults,
			(stored) => {
				// Cells show the answer now; a new answer is saved with the file.
				this.invalidateRowSizes();
				if (this.hot && !this.hot.isDestroyed) this.hot.render();
				if (stored) this.markDirty();
			},
		);
		// Hotkeys registered here win over Obsidian's while this view has the focus.
		this.scope = new Scope(this.app.scope);
		this.applyShortcutSettings();

		this.contentEl.empty();
		this.contentEl.addClass('rtl-spreadsheet-view');
		this.contentEl.toggleClass('is-touch', IS_TOUCH);

		this.formulaBarEl = this.contentEl.createDiv({
			cls: 'rtl-sheet-formula-bar',
		});
		this.toolbarEl = this.contentEl.createDiv({ cls: 'rtl-sheet-toolbar' });
		this.findBarEl = this.contentEl.createDiv({
			cls: 'rtl-sheet-find-bar',
		});
		this.chartPanelEl = this.contentEl.createDiv({
			cls: 'rtl-sheet-chart-panel',
		});
		this.chartPanelEl.hide();
		// Like the formula bar: typing in the panel must not reach the grid.
		this.chartPanelEl.addEventListener('focusin', (e) => {
			if (e.target instanceof HTMLElement && isTextEntry(e.target))
				this.hot?.unlisten();
		});
		// The grid and, next to it, the chart editor (below it on narrow screens).
		const body = this.contentEl.createDiv({ cls: 'rtl-sheet-body' });
		this.gridEl = body.createDiv({ cls: 'rtl-sheet-grid' });
		body.appendChild(this.chartPanelEl);
		this.sidePanelEl = body.createDiv({ cls: 'rtl-sheet-side-panel' });
		this.sidePanelEl.hide();
		this.commentsPanelEl = this.sidePanelEl.createDiv({
			cls: 'rtl-sheet-comments-panel',
		});
		this.aiPanelEl = this.sidePanelEl.createDiv({ cls: 'rtl-sheet-ai-host' });
		this.sidePanelEl.addEventListener('focusin', (e) => {
			if (e.target instanceof HTMLElement && isTextEntry(e.target))
				this.hot?.unlisten();
		});
		this.cellInfoEl = this.contentEl.createDiv({ cls: 'rtl-sheet-cell-info' });
		this.cellInfoEl.hide();
		for (const type of ['mousedown', 'pointerdown', 'touchstart'])
			this.cellInfoEl.addEventListener(type, (e) => e.stopPropagation());
		this.cellInfoEl.addEventListener('mouseenter', () => {
			this.cellInfoHovered = true;
		});
		this.cellInfoEl.addEventListener('mouseleave', () => {
			this.cellInfoHovered = false;
			this.queueCellInfo();
		});
		this.tabBarEl = this.contentEl.createDiv({ cls: 'rtl-sheet-tabs' });

		this.buildFormulaBar();
		this.buildToolbar();
		this.buildFindBar();
		// Capture phase, so the keys are seen before Handsontable and whatever has the focus.
		this.registerDomEvent(document, 'keydown', (e) => this.onViewKeyDown(e), {
			capture: true,
		});
		this.applyToolbarScale();
		this.applySelectionHandleSettings();
		this.applyCopyMarkSettings();
		this.installAutosave();
		this.installRefreshTimer();
		// Sync plugins that write through the vault are picked up right away, the timer covers the rest.
		this.registerEvent(
			this.app.vault.on('modify', (file) => {
				if (file === this.file && this.plugin.settings.autoRefresh)
					void this.refreshFromDisk();
			}),
		);
		this.register(() => window.clearInterval(this.refreshTimer));

		// Keep the grid theme in sync with Obsidian's light/dark theme.
		this.registerEvent(
			this.app.workspace.on('css-change', () => this.updateTheme()),
		);

		// Paste (Ctrl+V) or drop an image straight into the grid. Capture phase so that
		// Handsontable's own paste handling does not also run.
		this.registerDomEvent(this.gridEl, 'paste', (e) => this.onPaste(e), {
			capture: true,
		});
		this.registerDomEvent(this.gridEl, 'dragover', (e) => {
			if (this.hasImageFile(e)) e.preventDefault();
		});
		this.registerDomEvent(this.gridEl, 'drop', (e) => this.onDrop(e));

		// Corner handles of images. Registered before the other press handlers, in the capture
		// phase, so neither they nor Handsontable see a press on a handle.
		this.registerDomEvent(
			this.gridEl,
			'pointerdown',
			(e) => this.onImageHandlePress(e),
			{ capture: true },
		);
		// Corner handles of the selection (touch screens), handled the same way.
		this.registerDomEvent(
			this.gridEl,
			'pointerdown',
			(e) => this.onSelectionHandlePress(e),
			{ capture: true },
		);
		// Scrolling brings cells into the rendered part of the grid: a handle whose cell was
		// not rendered can be placed now. (Scroll events do not bubble, hence the capture.)
		this.registerDomEvent(
			this.gridEl,
			'scroll',
			() => {
				this.queueSelectionHandles();
				this.positionCellInfo();
			},
			{ capture: true, passive: true },
		);
		// Check boxes: a tap toggles them (a drag that scrolls the sheet does not).
		this.registerDomEvent(
			this.gridEl,
			'pointerdown',
			(e) => this.onCheckboxPress(e),
			{ capture: true },
		);
		this.registerDomEvent(this.gridEl, 'pointerup', (e) =>
			this.onCheckboxRelease(e),
		);
		this.registerDomEvent(
			this.gridEl,
			'click',
			(e) => this.onGridClick(e),
			{ capture: true },
		);
		this.registerDomEvent(
			this.gridEl,
			'dblclick',
			(e) => {
				// A double click on a check box toggles it twice; it does not open the editor.
				if (
					e.target instanceof HTMLElement &&
					e.target.closest('.rtl-sheet-checkbox')
				)
					e.stopPropagation();
			},
			{ capture: true },
		);
		// Comments and links show up next to the cell under the pointer (when the setting is on).
		this.registerDomEvent(this.gridEl, 'mouseover', (e) => {
			if (!this.plugin.settings.commentsOnHover) return;
			const cell = this.cellFromEvent(e);
			const same =
				cell?.row === this.hoverCell?.row &&
				cell?.col === this.hoverCell?.col;
			if (same) return;
			this.hoverCell = cell;
			this.queueCellInfo(cell && this.hasCellInfo(cell) ? 0 : 150);
		});
		this.registerDomEvent(this.gridEl, 'mouseleave', () => {
			this.hoverCell = null;
			this.queueCellInfo(250);
		});
		for (const type of ['mousedown', 'touchstart', 'touchend'] as const)
			this.registerDomEvent(
				this.gridEl,
				type,
				(e) => this.keepHandlePress(e),
				{ capture: true, passive: type !== 'mousedown' },
			);

		// While a formula is typed, a press on a cell inserts its reference. Capture phase so
		// Handsontable does not see the press (it would close the editor and move the selection).
		this.registerDomEvent(
			this.gridEl,
			'mousedown',
			(e) => this.onGridPress(e),
			{ capture: true },
		);
		// A tap ends in a mousedown that is handled above; keep the touch itself away from
		// Handsontable. Scrolling and pinching (two fingers) are not affected.
		const keepTapForPick = (e: TouchEvent): void => {
			const touches =
				e.type === 'touchstart' ? e.touches : e.changedTouches;
			if (touches.length === 1 && this.refPickTarget(e)) e.stopPropagation();
		};
		this.registerDomEvent(this.gridEl, 'touchstart', keepTapForPick, {
			capture: true,
			passive: true,
		});
		this.registerDomEvent(this.gridEl, 'touchend', keepTapForPick, {
			capture: true,
			passive: true,
		});

		// Pinch to zoom. Touches that start in the grid are not passed on to Obsidian, so
		// dragging sideways scrolls the columns instead of opening the sidebar.
		const touchDistance = (e: TouchEvent): number => {
			const a = e.touches[0];
			const b = e.touches[1];
			return a && b
				? Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY)
				: 0;
		};
		this.registerDomEvent(
			this.gridEl,
			'touchstart',
			(e) => {
				e.stopPropagation();
				if (e.touches.length === 2)
					this.pinch = {
						startDistance: touchDistance(e),
						startZoom: this.zoom,
					};
			},
			{ passive: true },
		);
		this.registerDomEvent(
			this.gridEl,
			'touchmove',
			(e) => {
				e.stopPropagation();
				const pinch = this.pinch;
				if (
					!pinch ||
					e.touches.length !== 2 ||
					pinch.startDistance <= 0
				)
					return;
				e.preventDefault();
				this.pendingZoom =
					(pinch.startZoom * touchDistance(e)) / pinch.startDistance;
				if (!this.zoomFrame) {
					this.zoomFrame = window.requestAnimationFrame(() => {
						this.zoomFrame = 0;
						this.applyZoom(this.pendingZoom);
					});
				}
			},
			{ passive: false },
		);
		const endTouch = (e: TouchEvent): void => {
			e.stopPropagation();
			if (e.touches.length < 2) this.pinch = null;
		};
		this.registerDomEvent(this.gridEl, 'touchend', endTouch, {
			passive: true,
		});
		this.registerDomEvent(this.gridEl, 'touchcancel', endTouch, {
			passive: true,
		});
		this.registerDomEvent(
			this.gridEl,
			'wheel',
			(e) => {
				if (!e.ctrlKey) return;
				e.preventDefault();
				this.applyZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
			},
			{ passive: false },
		);
	}

	private updateTheme(): void {
		this.fontFamilyCache = null;
		if (this.hot && !this.hot.isDestroyed) {
			this.hot.updateSettings({ themeName: handsontableThemeName() });
		}
	}

	/* ---------------- TextFileView contract ---------------- */

	getViewType(): string {
		return VIEW_TYPE;
	}

	getDisplayText(): string {
		return this.file?.basename ?? 'Spreadsheet';
	}

	getIcon(): string {
		return 'table';
	}

	getViewData(): string {
		this.captureGridState();
		return serializeWorkbook({
			sheets: this.sheets,
			active: this.activeSheet,
			aiResults: this.aiResults,
		});
	}

	setViewData(data: string, clear: boolean): void {
		this.loading = true;
		try {
			let parsed: unknown = null;
			if (data.trim().length > 0) {
				try {
					parsed = JSON.parse(data);
				} catch (e) {
					this.plugin.reportError(
						'Invalid sheet file, opened an empty sheet',
						e,
					);
				}
			}
			const next = this.normalizeWorkbook(parsed);
			this.aiResults = next.aiResults ?? {};
			this.diskData = data;
			if (clear) this.filters.clear(); // another file: its columns mean something else
			// Same file with the same tabs: stay on the sheet that is shown.
			const sameTabs =
				!clear &&
				next.sheets.length === this.sheets.length &&
				next.sheets.every((sh, i) => sh.name === this.sheets[i]?.name);
			const index = sameTabs ? this.activeSheet : next.active;
			const shown = next.sheets[index] ?? next.sheets[0];
			if (!shown) return;
			// Same file, new content: update the live grid so selection and scroll stay put.
			if (sameTabs && this.applyModelInPlace(shown)) {
				this.sheets = next.sheets;
				this.renderTabs();
			} else {
				// Drop the old grid first, rebuildGrid() would copy its stale cells into the new model.
				this.destroyGrid();
				if (!sameTabs) this.filters.clear();
				this.sheets = next.sheets;
				this.activeSheet = next.sheets.indexOf(shown);
				this.model = shown;
				this.rebuildGrid();
			}
			this.dirty = false;
		} finally {
			this.loading = false;
		}
	}

	clear(): void {
		this.destroyGrid();
		this.filters.clear();
		this.model = createDefaultModel(
			this.plugin.settings,
			defaultSheetRtl(),
		);
		this.sheets = [this.model];
		this.activeSheet = 0;
		this.aiResults = {};
		this.dirty = false;
	}

	async save(clear?: boolean): Promise<void> {
		await super.save(clear);
		this.dirty = false;
	}

	async saveNow(): Promise<void> {
		try {
			await this.save();
			new Notice('Sheet saved');
		} catch (e) {
			this.plugin.reportError('Could not save sheet', e);
		}
	}

	async onClose(): Promise<void> {
		try {
			if (this.dirty && this.file) await this.save();
		} catch (e) {
			this.plugin.reportError('Could not save sheet on close', e);
		}
		window.clearInterval(this.refreshTimer);
		window.cancelAnimationFrame(this.chartFrame);
		window.clearTimeout(this.cellInfoTimer);
		this.autocomplete.destroy();
		this.emojiPicker.close();
		this.aiPanel?.destroy();
		this.aiPanel = null;
		this.destroyGrid();
	}

	onResize(): void {
		if (this.hot && !this.hot.isDestroyed) this.hot.refreshDimensions();
	}

	/* ---------------- Settings, autosave ---------------- */

	private installAutosave(): void {
		// Replace the built-in 2s debounce with one that follows the plugin setting.
		this.requestSave = debounce(
			() => {
				void this.save().catch((e) =>
					this.plugin.reportError('Autosave failed', e),
				);
			},
			this.plugin.settings.autosaveDelayMs,
			true,
		);
	}

	private installRefreshTimer(): void {
		window.clearInterval(this.refreshTimer);
		this.refreshTimer = 0;
		const s = this.plugin.settings;
		if (!s.autoRefresh) return;
		this.refreshTimer = window.setInterval(
			() => void this.refreshFromDisk(),
			s.refreshIntervalSec * 1000,
		);
	}

	/** True while the user is typing in a cell or in the formula bar. */
	private isEditing(): boolean {
		const input = this.ui.formulaInput;
		if (input && input.ownerDocument.activeElement === input) return true;
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return false;
		const editor = hot.getActiveEditor() as unknown as
			| { isOpened?: () => boolean }
			| undefined;
		return !!editor?.isOpened?.();
	}

	/**
	 * Loads the file again when it was changed outside this view. Waits (until the next
	 * call) while there are unsaved local changes or the user is typing.
	 */
	private async refreshFromDisk(): Promise<void> {
		const file = this.file;
		const busy = (): boolean =>
			this.file !== file ||
			!this.hot ||
			this.hot.isDestroyed ||
			this.loading ||
			this.dirty ||
			this.isEditing();
		if (!file || this.refreshing || busy()) return;
		if (!this.containerEl.isShown()) return;
		this.refreshing = true;
		try {
			const disk = await this.app.vault.read(file);
			if (disk === this.diskData || busy()) return;
			if (disk === this.getViewData()) {
				this.diskData = disk; // our own save
				return;
			}
			this.data = disk;
			this.setViewData(disk, false);
		} catch (e) {
			console.warn('[Free-Excel] Could not refresh the sheet', e);
		} finally {
			this.refreshing = false;
		}
	}

	private markDirty(): void {
		this.invalidateRowSizes();
		this.findStale = true;
		if (this.loading) return;
		this.queueSideCheck();
		this.dirty = true;
		if (this.plugin.settings.autosave) this.requestSave();
	}

	applySettings(): void {
		const s = this.plugin.settings;
		const prev = this.appliedSettings;
		this.appliedSettings = { ...s };
		this.installAutosave();
		if (
			prev.autoRefresh !== s.autoRefresh ||
			prev.refreshIntervalSec !== s.refreshIntervalSec
		)
			this.installRefreshTimer();

		if (!prev.autosave && s.autosave && this.dirty) this.requestSave();
		if (prev.ctrlGFind !== s.ctrlGFind) this.applyShortcutSettings();
		if (!s.commentsOnHover && this.hoverCell) {
			this.hoverCell = null;
			this.refreshCellInfo();
		}
		this.applyToolbarScale();
		this.applySelectionHandleSettings();
		this.applyCopyMarkSettings();
		this.renderPalette('fill');
		this.renderPalette('text');

		if (prev.zoomAxes !== s.zoomAxes) this.applyZoomSizes();
		else if (
			prev.defaultFontSize !== s.defaultFontSize ||
			prev.enableFormulas !== s.enableFormulas
		) {
			this.invalidateRowSizes();
			if (this.hot && !this.hot.isDestroyed) this.hot.render();
		}
	}

	/* ---------------- Model helpers ---------------- */

	/** Reads a file: several sheets (version 3), or one sheet at the top level (older files). */
	private normalizeWorkbook(raw: unknown): WorkbookModel {
		const r = (raw && typeof raw === 'object' ? raw : {}) as {
			sheets?: unknown;
			active?: unknown;
		};
		const list: unknown[] =
			Array.isArray(r.sheets) && r.sheets.length > 0 ? r.sheets : [raw];
		const sheets: SheetModel[] = [];
		for (const item of list) {
			const sheet = this.normalizeModel(item);
			// Every tab needs a name of its own: formulas refer to sheets by name.
			sheet.name = uniqueSheetName(
				sheets.map((sh) => sh.name),
				sheet.name || `Sheet${sheets.length + 1}`,
			);
			sheets.push(sheet);
		}
		const active = Math.round(Number(r.active));
		const ai = (raw as { aiResults?: unknown } | null)?.aiResults;
		const aiResults: Record<string, string> = {};
		if (ai && typeof ai === 'object')
			for (const [k, v] of Object.entries(ai as Record<string, unknown>))
				if (typeof v === 'string') aiResults[k] = v;
		return {
			sheets,
			aiResults,
			active:
				Number.isFinite(active) && active >= 0 && active < sheets.length
					? active
					: 0,
		};
	}

	private normalizeCharts(raw: unknown): ChartSpec[] {
		if (!Array.isArray(raw)) return [];
		const num = (v: unknown, fallback: number): number =>
			typeof v === 'number' && Number.isFinite(v) ? v : fallback;
		const types: string[] = ['column', 'bar', 'line', 'pie', 'scatter'];
		const out: ChartSpec[] = [];
		for (const item of raw as (Partial<ChartSpec> | null)[]) {
			if (!item || typeof item !== 'object' || !item.range) continue;
			const rg = item.range;
			const xr = item.xRange;
			out.push({
				id: typeof item.id === 'string' ? item.id : newChartId(),
				type: types.includes(item.type ?? '')
					? (item.type as ChartSpec['type'])
					: 'column',
				title: typeof item.title === 'string' ? item.title : '',
				range: {
					r1: Math.max(0, num(rg.r1, 0)),
					c1: Math.max(0, num(rg.c1, 0)),
					r2: Math.max(0, num(rg.r2, 0)),
					c2: Math.max(0, num(rg.c2, 0)),
				},
				xRange:
					xr && typeof xr === 'object'
						? {
								r1: Math.max(0, num(xr.r1, 0)),
								c1: Math.max(0, num(xr.c1, 0)),
								r2: Math.max(0, num(xr.r2, 0)),
								c2: Math.max(0, num(xr.c2, 0)),
							}
						: undefined,
				headerRow: !!item.headerRow,
				trendline: !!item.trendline,
				trendColor:
					typeof item.trendColor === 'string' &&
					/^#[0-9a-f]{6}$/i.test(item.trendColor)
						? item.trendColor
						: undefined,
				x: Math.max(0, num(item.x, 0)),
				y: Math.max(0, num(item.y, 0)),
				w: Math.max(CHART_MIN_W, num(item.w, CHART_DEFAULT_W)),
				h: Math.max(CHART_MIN_H, num(item.h, CHART_DEFAULT_H)),
				...normalizeChartExtras(item),
			});
		}
		return out;
	}

	private normalizeModel(raw: unknown): SheetModel {
		const base = createDefaultModel(
			this.plugin.settings,
			defaultSheetRtl(),
		);
		if (!raw || typeof raw !== 'object') return base;
		const r = raw as Partial<SheetModel>;

		let data = base.data;
		if (
			Array.isArray(r.data) &&
			r.data.length > 0 &&
			r.data.every((row) => Array.isArray(row))
		) {
			const width = Math.max(1, ...r.data.map((row) => row.length));
			data = r.data.map((row) => {
				const copy = row.slice();
				while (copy.length < width) copy.push(null);
				return copy;
			});
		}

		return {
			version: 2,
			name: typeof r.name === 'string' ? cleanSheetName(r.name) : '',
			data,
			cellFormats:
				r.cellFormats && typeof r.cellFormats === 'object'
					? r.cellFormats
					: {},
			images: r.images && typeof r.images === 'object' ? r.images : {},
			imageSizes:
				r.imageSizes && typeof r.imageSizes === 'object'
					? r.imageSizes
					: {},
			widenRecords:
				r.widenRecords && typeof r.widenRecords === 'object'
					? r.widenRecords
					: {},
			comments: this.normalizeComments(r.comments),
			links: this.normalizeLinks(r.links),
			merges: Array.isArray(r.merges) ? r.merges : [],
			colWidths: Array.isArray(r.colWidths) ? r.colWidths : [],
			rowHeights: Array.isArray(r.rowHeights) ? r.rowHeights : [],
			rtl: typeof r.rtl === 'boolean' ? r.rtl : defaultSheetRtl(),
			charts: this.normalizeCharts(r.charts),
			floatingImages: this.normalizeFloatingImages(r.floatingImages),
		};
	}

	private normalizeComments(raw: unknown): Record<string, CellComment> {
		const out: Record<string, CellComment> = {};
		if (!raw || typeof raw !== 'object') return out;
		for (const [k, v] of Object.entries(raw as Record<string, unknown>)) {
			const c = v as Partial<CellComment> | null;
			if (!c || typeof c.text !== 'string' || !/^\d+,\d+$/.test(k)) continue;
			out[k] = {
				text: c.text,
				color:
					typeof c.color === 'string' && /^#[0-9a-f]{6}$/i.test(c.color)
						? c.color
						: DEFAULT_COMMENT_COLOR,
				time: typeof c.time === 'number' ? c.time : 0,
			};
		}
		return out;
	}

	private normalizeLinks(raw: unknown): Record<string, string> {
		const out: Record<string, string> = {};
		if (!raw || typeof raw !== 'object') return out;
		for (const [k, v] of Object.entries(raw as Record<string, unknown>))
			if (typeof v === 'string' && v && /^\d+,\d+$/.test(k)) out[k] = v;
		return out;
	}

	private normalizeFloatingImages(raw: unknown): FloatingImage[] {
		if (!Array.isArray(raw)) return [];
		const num = (v: unknown, fallback: number): number =>
			typeof v === 'number' && Number.isFinite(v) ? v : fallback;
		const out: FloatingImage[] = [];
		for (const item of raw as (Partial<FloatingImage> | null)[]) {
			if (!item || typeof item.path !== 'string' || !item.path) continue;
			out.push({
				id: typeof item.id === 'string' ? item.id : newChartId('img'),
				path: item.path,
				x: Math.max(0, num(item.x, 0)),
				y: Math.max(0, num(item.y, 0)),
				w: Math.max(FLOAT_MIN_PX, num(item.w, 200)),
				h: Math.max(FLOAT_MIN_PX, num(item.h, 150)),
				...(typeof item.z === 'number' && Number.isFinite(item.z)
					? { z: item.z }
					: {}),
			});
		}
		return out;
	}

	/**
	 * Replaces the model under the live grid without rebuilding it, so an outside change
	 * shows up without losing the selection or the scroll position. Returns false when
	 * the grid has to be rebuilt instead.
	 */
	private applyModelInPlace(next: SheetModel): boolean {
		const hot = this.hot;
		if (!hot || hot.isDestroyed || next.rtl !== this.model.rtl) return false;

		this.captureGridState(); // current merges, to compare with the new ones
		const prev = this.model;
		this.model = next;
		this.invalidateRowSizes();

		hot.updateData(next.data);
		if (JSON.stringify(prev.merges) !== JSON.stringify(next.merges))
			hot.updateSettings({ mergeCells: next.merges });

		const nextCols = next.colWidths.slice();
		const cols = Math.max(prev.colWidths.length, nextCols.length);
		for (let i = 0; i < cols; i++) {
			const w = nextCols[i] ?? null;
			if ((prev.colWidths[i] ?? null) === w) continue;
			if (w) this.setColumnWidth(i, w);
			else this.clearColumnWidth(i);
		}
		const nextRows = next.rowHeights.slice();
		const rows = Math.max(prev.rowHeights.length, nextRows.length);
		for (let i = 0; i < rows; i++) {
			const h = nextRows[i] ?? null;
			if ((prev.rowHeights[i] ?? null) === h) continue;
			if (h) this.setRowHeight(i, h);
			else this.clearRowHeight(i);
		}

		// The undo history describes the sheet as it was before the outside change.
		this.undoPlugin()?.clear();
		this.resetSideBaseline();
		hot.render();
		this.queueCharts();
		this.chartEditor?.refresh();
		this.refreshToolbarState();
		return true;
	}

	/** Copies the live grid state (data and merges) into the model. */
	private captureGridState(): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		this.model.data = hot.getSourceDataArray() as Scalar[][];
		this.model.merges = this.liveMerges(hot).map((m) => ({
			row: m.row,
			col: m.col,
			rowspan: m.rowspan,
			colspan: m.colspan,
		}));
	}

	/** The grid's current merges, or the model's when there is no live grid. */
	private liveMerges(hot: Handsontable.Core | null): MergeInfo[] {
		if (!hot || hot.isDestroyed) return this.model.merges;
		// The merged-cells collection is not part of Handsontable's public typings.
		const mc = hot.getPlugin('mergeCells') as unknown as {
			mergedCellsCollection?: { mergedCells?: MergeInfo[] };
		};
		return mc.mergedCellsCollection?.mergedCells ?? [];
	}

	/** The merge that covers a cell, if any. */
	private mergeAt(
		hot: Handsontable.Core | null,
		row: number,
		col: number,
	): MergeInfo | null {
		for (const m of this.liveMerges(hot)) {
			if (
				row >= m.row &&
				row < m.row + m.rowspan &&
				col >= m.col &&
				col < m.col + m.colspan
			)
				return m;
		}
		return null;
	}

	private captureGridStateIfLive(): void {
		if (this.hot && !this.hot.isDestroyed) this.captureGridState();
	}

	/** Shifts per-cell data (formats, images) after rows/columns are inserted or removed. */
	private remapCellMaps(
		axis: 'row' | 'col',
		index: number,
		amount: number,
		mode: 'insert' | 'remove',
	): void {
		const remap = <T>(source: Record<string, T>): Record<string, T> => {
			const next: Record<string, T> = {};
			for (const [k, value] of Object.entries(source)) {
				const parts = k.split(',').map(Number);
				const r = parts[0] ?? 0;
				const c = parts[1] ?? 0;
				let pos = axis === 'row' ? r : c;
				if (mode === 'insert') {
					if (pos >= index) pos += amount;
				} else {
					if (pos >= index && pos < index + amount) continue;
					if (pos >= index + amount) pos -= amount;
				}
				next[axis === 'row' ? key(pos, c) : key(r, pos)] = value;
			}
			return next;
		};
		if (axis === 'col' && this.model.widenRecords) {
			const next: Record<string, { prevW: number | null; w: number }> =
				{};
			for (const [k, rec] of Object.entries(this.model.widenRecords)) {
				let c = Number(k);
				if (mode === 'insert') {
					if (c >= index) c += amount;
				} else {
					if (c >= index && c < index + amount) continue;
					if (c >= index + amount) c -= amount;
				}
				next[String(c)] = rec;
			}
			this.model.widenRecords = next;
		}
		this.model.cellFormats = remap(this.model.cellFormats);
		this.model.images = remap(this.model.images);
		this.model.imageSizes = remap(this.model.imageSizes);
		this.model.comments = remap(this.model.comments);
		this.model.links = remap(this.model.links);
		// Charts keep showing the same cells (replaced, not changed: undo compares them).
		const shiftRange = (rg: SelRange): SelRange => {
			const lo = axis === 'row' ? rg.r1 : rg.c1;
			const hi = axis === 'row' ? rg.r2 : rg.c2;
			let a = lo;
			let b = hi;
			if (mode === 'insert') {
				if (index <= lo) a += amount;
				if (index <= hi) b += amount;
			} else {
				const end = index + amount; // first index after the removed ones
				const shift = (p: number): number =>
					p < index ? p : p >= end ? p - amount : index;
				a = shift(lo);
				// The last cell moves back to the last one that is left before it.
				b = hi < index ? hi : hi >= end ? hi - amount : index - 1;
				if (b < a) b = a;
			}
			if (a === lo && b === hi) return rg;
			return axis === 'row'
				? { ...rg, r1: a, r2: b }
				: { ...rg, c1: a, c2: b };
		};
		this.model.charts = this.model.charts.map((chart) => {
			const range = shiftRange(chart.range);
			const xRange = chart.xRange && shiftRange(chart.xRange);
			return range === chart.range && xRange === chart.xRange
				? chart
				: { ...chart, range, xRange };
		});
	}

	private spliceSizes(
		arr: (number | null)[],
		index: number,
		amount: number,
		mode: 'insert' | 'remove',
	): void {
		while (arr.length < index) arr.push(null);
		if (mode === 'insert')
			arr.splice(index, 0, ...Array<null>(amount).fill(null));
		else arr.splice(index, amount);
	}

	/* ---------------- Formula resolution ---------------- */

	/** The stored value of a cell on any sheet of the file (the shown one is read from the grid). */
	private rawCell(sheet: number, row: number, col: number): unknown {
		if (sheet === this.activeSheet) {
			const hot: Handsontable.Core | null = this.gridForCalc ?? this.hot;
			if (hot && !hot.isDestroyed) return hot.getSourceDataAtCell(row, col);
		}
		return this.sheets[sheet]?.data[row]?.[col] ?? null;
	}

	/** Index of the sheet a formula names ("Sheet2" or "'My sheet'"), or -1. */
	private sheetIndexByName(name: string): number {
		const wanted = name
			.replace(/^'(.*)'$/, '$1')
			.replace(/''/g, "'")
			.trim()
			.toLowerCase();
		return this.sheets.findIndex((s) => s.name.toLowerCase() === wanted);
	}

	/**
	 * `at` is the cell the value lives in (ROW() and COLUMN() without an argument report it),
	 * `sheet` the index of its sheet (the shown one by default).
	 */
	private resolveRaw(
		raw: unknown,
		depth: number,
		coerce: boolean,
		at?: { row: number; col: number },
		sheet = this.activeSheet,
	): Scalar {
		if (raw === undefined || raw === null) return null;
		if (typeof raw === 'string') {
			if (
				this.plugin.settings.enableFormulas &&
				raw.length > 1 &&
				raw.startsWith('=')
			) {
				if (depth > 40) {
					this.formulaLimitHits++;
					return '#CYCLE!';
				}
				const cacheKey = at ? `${sheet}!${key(at.row, at.col)}` : null;
				if (cacheKey) {
					const hit = this.formulaCache.get(cacheKey);
					if (hit && hit.raw === raw) return hit.value;
				}
				const limitHits = this.formulaLimitHits;
				// CLAUDE() and GEMINI() ask through this view; the cell is remembered for "ask again".
				const previous = setAiFormulaHandler((provider, prompt, context) =>
					this.aiRunner.resolve(provider, prompt, context, cacheKey),
				);
				let value: Scalar;
				try {
					value = evaluateFormula(
						raw,
						(r, c, sheetName) => {
							const target =
								sheetName === undefined
									? sheet
									: this.sheetIndexByName(sheetName);
							if (target < 0) return '#REF!';
							return this.resolveRaw(
								this.rawCell(target, r, c),
								depth + 1,
								true,
								{ row: r, col: c },
								target,
							);
						},
						at,
					);
				} finally {
					setAiFormulaHandler(previous);
				}
				// A result that ran into the depth limit depends on where the evaluation started.
				if (cacheKey && this.formulaLimitHits === limitHits)
					this.formulaCache.set(cacheKey, { raw, value });
				return value;
			}
			if (coerce && NUMERIC_RE.test(raw)) return Number(raw);
			return raw;
		}
		if (typeof raw === 'number' || typeof raw === 'boolean') return raw;
		return null;
	}

	/* ---------------- Grid lifecycle ---------------- */

	destroyGrid(): void {
		this.closeChartEditor();
		this.clearRefPick();
		this.autocomplete.detach();
		if (this.hot && !this.hot.isDestroyed) this.hot.destroy();
		this.hot = null;
		this.gridForCalc = null;
		this.chartLayer = null;
		this.gridEl.empty();
	}

	/**
	 * `keepHistory` carries the undo history over to the new grid. Only for a rebuild of the
	 * same sheet (a direction change), where the old steps still describe the cells.
	 */
	private rebuildGrid(keepHistory = false): void {
		this.captureGridStateIfLive();
		const old = keepHistory ? this.undoPlugin() : null;
		const history = old
			? { done: [...old.doneActions], undone: [...old.undoneActions] }
			: null;
		this.destroyGrid();

		const m = this.model;
		this.gridEl.setAttr('dir', m.rtl ? 'rtl' : 'ltr');
		this.gridEl.toggleClass('is-rtl', m.rtl);
		this.applyZoomStyles();

		this.fontFamilyCache = null;
		this.invalidateRowSizes();
		const toSizes = (
			arr: (number | null)[],
			zoom: number,
		): (number | undefined)[] =>
			arr.map((v) => (v ? Math.round(v * zoom) : undefined));

		try {
			const options: Handsontable.GridSettings = {
				data: m.data,
				licenseKey: 'non-commercial-and-evaluation',
				themeName: handsontableThemeName(),
				layoutDirection: m.rtl ? 'rtl' : 'ltr',
				rowHeaders: true,
				rowHeaderWidth: this.rowHeaderWidth(),
				colHeaders: true,
				width: '100%',
				height: '100%',
				manualColumnResize: toSizes(m.colWidths, this.zoom) as number[],
				manualRowResize: toSizes(m.rowHeights, this.zoomY) as number[],
				// Functions (instead of auto-sizing) so the user decides column widths, and row heights
				// always follow the font size; manual resizing still overrides both.
				colWidths: () => Math.round(DEFAULT_COL_WIDTH * this.zoom),
				rowHeights: (row: number) =>
					Math.round(this.computeRowHeight(row) * this.zoomY),
				mergeCells: m.merges,
				// Filters hide rows instead of trimming them, so a row keeps its index and the
				// "row,col" keys and formula references stay valid.
				hiddenRows: {
					rows: [],
					indicators: true,
					copyPasteEnabled: false,
				},
				afterGetColHeader: (col, th) => {
					th.classList.toggle(
						'rtl-sheet-filtered',
						this.filters.has(col),
					);
				},
				contextMenu: [
					'row_above',
					'row_below',
					'col_left',
					'col_right',
					'---------',
					'remove_row',
					'remove_col',
					'---------',
					'undo',
					'redo',
					'---------',
					'copy',
					'cut',
					'---------',
					'mergeCells',
				],
				copyPaste: true,
				undo: true,
				fillHandle: !IS_TOUCH,
				selectionMode: 'multiple',
				outsideClickDeselects: false,
				autoWrapRow: true,
				autoWrapCol: true,
				renderer: this.cellRenderer,
				// Cells that hold an image cannot be typed into (the image would hide the text).
				cells: (row, col) => ({
					readOnly: !!this.model.images[key(row, col)],
				}),
				beforeKeyDown: (event) => this.onBeforeKeyDown(event),
				// A full render follows every change to the sheet, so it starts from fresh row
				// heights and formula results; moving the selection and scrolling reuse them.
				beforeRender: (isForced) => {
					if (!isForced) return;
					this.rowHeightCache.clear();
					this.formulaCache.clear();
					this.spillCache.clear();
					this.borderGridCache = null;
					this.invalidateGridRowHeights();
				},
				// Charts follow their data: draw them again after every full render.
				afterRender: (isForced) => {
					if (isForced) this.queueCharts();
				},
				afterChange: (changes, source) => {
					this.invalidateRowSizes();
					if (source === 'loadData' || this.loading) return;
					// Typing in a cell ends the copy outline; a paste and the cut itself do not.
					if (source !== 'CopyPaste.paste' && source !== 'CopyPaste.cut')
						this.clearCopyMark();
					const expandCols: number[] = [];
					for (const change of changes ?? []) {
						const [row, prop] = change as unknown as [
							number,
							string | number,
						];
						const col = Number(prop);
						if (
							this.model.cellFormats[key(row, col)]?.wrap ===
							'expand'
						)
							expandCols.push(col);
					}
					if (expandCols.length > 0)
						this.fitExpandColumns(expandCols);
					this.readFormulaRefs(); // the selected cell's formula may be the one that changed
					if (this.plugin.settings.enableFormulas && this.hot)
						this.hot.render();
					this.markDirty();
					this.refreshToolbarState();
				},
				// A drag of a selection handle reports once, when it ends.
				afterSelectionEnd: () => {
					if (!this.selDragging) this.onSelectionChanged();
				},
				afterBeginEditing: () => this.onBeginEditing(),
				afterCreateRow: (index, amount) => {
					if (this.loading) return;
					this.remapCellMaps('row', index, amount, 'insert');
					this.spliceSizes(
						this.model.rowHeights,
						index,
						amount,
						'insert',
					);
					this.markDirty();
				},
				afterRemoveRow: (index, amount) => {
					if (this.loading) return;
					this.remapCellMaps('row', index, amount, 'remove');
					this.spliceSizes(
						this.model.rowHeights,
						index,
						amount,
						'remove',
					);
					this.markDirty();
				},
				afterCreateCol: (index, amount) => {
					if (this.loading) return;
					this.remapCellMaps('col', index, amount, 'insert');
					this.remapFilters(index, amount, 'insert');
					this.spliceSizes(
						this.model.colWidths,
						index,
						amount,
						'insert',
					);
					this.markDirty();
				},
				afterRemoveCol: (index, amount) => {
					if (this.loading) return;
					this.remapCellMaps('col', index, amount, 'remove');
					this.remapFilters(index, amount, 'remove');
					this.spliceSizes(
						this.model.colWidths,
						index,
						amount,
						'remove',
					);
					this.markDirty();
				},
				afterColumnResize: (newSize, column) => {
					while (this.model.colWidths.length <= column)
						this.model.colWidths.push(null);
					this.model.colWidths[column] = Math.round(
						newSize / this.zoom,
					);
					// Wrapped text in the column now takes a different number of lines.
					this.invalidateRowSizes();
					this.markDirty();
				},
				afterRowResize: (newSize, row) => {
					while (this.model.rowHeights.length <= row)
						this.model.rowHeights.push(null);
					this.model.rowHeights[row] = Math.round(
						newSize / this.zoomY,
					);
					this.markDirty();
				},
				// Handsontable's own steps move formats and sizes through the hooks above:
				// that is part of the undo, not a new change to record.
				afterUndo: () => {
					if (!this.sideBusy) this.resetSideBaseline();
				},
				afterRedo: () => {
					if (!this.sideBusy) this.resetSideBaseline();
				},
				afterMergeCells: () => this.markDirty(),
				afterUnmergeCells: () => this.markDirty(),
				afterAutofill: (_fillData, sourceRange, targetRange) => {
					this.copyFormatsOnFill(sourceRange, targetRange);
				},
				beforeCopy: (data, coords) => this.rememberCopiedCells(data, coords),
				beforeCut: (data, coords) => this.rememberCopiedCells(data, coords),
				beforePaste: (data) => this.beforePasteCells(data),
				afterPaste: (_data, coords) => this.pasteCopiedFormats(coords),
			};

			this.hot = new Handsontable(this.gridEl, options);
			// The first render ran before `this.hot` existed; render again so formulas see all cells.
			if (this.plugin.settings.enableFormulas) this.hot.render();
			this.installSelectionHandles(this.hot);
			this.chartLayer =
				this.gridEl
					.querySelector('.ht_master .wtHider')
					?.createDiv({ cls: 'rtl-sheet-chart-layer' }) ?? null;
			this.queueCharts();
			this.formulaCell = null;
			const undo = this.undoPlugin();
			if (history && undo) {
				undo.doneActions.push(...history.done);
				undo.undoneActions.push(...history.undone);
			} else {
				this.knownSteps = new WeakSet();
				this.lastSideStep = null;
			}
			this.resetSideBaseline();
			this.applyFilters(); // the new grid starts with every row visible
			if (this.findBarEl.isShown()) this.runFind();
			this.refreshToolbarState();
		} catch (e) {
			this.plugin.reportError('Could not create the spreadsheet grid', e);
		}
		this.renderTabs();
		// Another sheet (or the same one again): its comments and floating images.
		this.hoverCell = null;
		this.activeFloatId = null;
		this.refreshCellInfo();
		this.refreshSidePanel();
	}

	private copyFormatsOnFill(
		source: {
			from: { row: number | null; col: number | null };
			to: { row: number | null; col: number | null };
		},
		target: {
			from: { row: number | null; col: number | null };
			to: { row: number | null; col: number | null };
		},
	): void {
		const toRange = (rg: typeof source): SelRange => {
			const fr = rg.from.row ?? 0;
			const tr = rg.to.row ?? fr;
			const fc = rg.from.col ?? 0;
			const tc = rg.to.col ?? fc;
			return {
				r1: Math.min(fr, tr),
				r2: Math.max(fr, tr),
				c1: Math.min(fc, tc),
				c2: Math.max(fc, tc),
			};
		};
		const s = toRange(source);
		const t = toRange(target);
		const h = s.r2 - s.r1 + 1;
		const w = s.c2 - s.c1 + 1;
		for (let r = t.r1; r <= t.r2; r++) {
			for (let c = t.c1; c <= t.c2; c++) {
				const sr = s.r1 + ((((r - s.r1) % h) + h) % h);
				const sc = s.c1 + ((((c - s.c1) % w) + w) % w);
				const f = this.model.cellFormats[key(sr, sc)];
				if (f) this.model.cellFormats[key(r, c)] = cloneFormat(f);
				else delete this.model.cellFormats[key(r, c)];
			}
		}
	}

	/** Keeps the formats of the cells that are being copied, for the paste (see `copiedCells`). */
	private rememberCopiedCells(
		data: unknown[][],
		coords: { startRow: number; startCol: number; endRow: number; endCol: number }[],
	): void {
		this.clearCopyMark();
		copiedCells = null;
		const hot = this.hot;
		const rg = coords.length === 1 ? coords[0] : undefined;
		if (!hot || hot.isDestroyed || !rg) return;
		const r1 = Math.max(0, Math.min(rg.startRow, rg.endRow));
		const r2 = Math.max(rg.startRow, rg.endRow);
		const c1 = Math.max(0, Math.min(rg.startCol, rg.endCol));
		const c2 = Math.max(rg.startCol, rg.endCol);
		// Rows hidden by a filter are not copied.
		const hidden = hot.getPlugin('hiddenRows');
		const rows: number[] = [];
		for (let r = r1; r <= r2; r++) if (!hidden.isHidden(r)) rows.push(r);
		if (rows.length !== data.length) return;
		const formats = rows.map((r) => {
			const line: (CellFormat | undefined)[] = [];
			for (let c = c1; c <= c2; c++) {
				const f = this.model.cellFormats[key(r, c)];
				line.push(f ? cloneFormat(f) : undefined);
			}
			return line;
		});
		copiedCells = {
			text: clipboardText(data),
			formats,
			mark: {
				view: this,
				sheet: this.activeSheet,
				range: { r1, r2, c1, c2 },
			},
		};
		this.queueCharts(); // draws the mark
	}

	/** Removes the moving outline of the copied cells (the copy itself can still be pasted). */
	clearCopyMark(): void {
		const copied = copiedCells;
		if (!copied?.mark) return;
		const view = copied.mark.view;
		copied.mark = null;
		view.queueCharts();
	}

	/** The moving dashed outline of the copied cells, drawn on the chart layer. */
	private renderCopyMark(
		hot: Handsontable.Core,
		layer: HTMLElement,
		origin: { x: number; y: number },
	): void {
		const mark = copiedCells?.mark;
		if (!mark || mark.view !== this || mark.sheet !== this.activeSheet) return;
		const { r1, r2, c1, c2 } = mark.range;
		const hidden = hot.getPlugin('hiddenRows');
		const rowSpan = (from: number, to: number): number => {
			let h = 0;
			for (let r = from; r < to; r++)
				if (!hidden.isHidden(r)) h += hot.getRowHeight(r) ?? 0;
			return h;
		};
		const colSpan = (from: number, to: number): number => {
			let w = 0;
			for (let c = from; c < to; c++) w += hot.getColWidth(c);
			return w;
		};
		const x = colSpan(0, c1);
		const y = rowSpan(0, r1);
		const box = layer.createDiv({ cls: 'rtl-sheet-copy-mark' });
		const st = box.style;
		// Over the grid line before the first cell, like the selection outline.
		st.top = `${origin.y + y - 1}px`;
		st.left = this.model.rtl ? '' : `${origin.x + x - 1}px`;
		st.right = this.model.rtl ? `${origin.x + x - 1}px` : '';
		st.width = `${colSpan(c1, c2 + 1) + 1}px`;
		st.height = `${rowSpan(r1, r2 + 1) + 1}px`;
		// The layer is drawn again after every change; starting at the same point of the cycle
		// keeps the dashes moving instead of jumping back to the start each time.
		const cycle = 2.5 / this.plugin.settings.copyMarkSpeed;
		st.animationDelay = `-${((performance.now() / 1000) % cycle).toFixed(3)}s`;
	}

	/** A paste of the cells copied last: check boxes get their TRUE/FALSE values back. */
	private beforePasteCells(data: unknown[][]): void {
		this.clearCopyMark(); // any paste ends the outline
		const copied = copiedCells;
		const text = clipboardText(data);
		const same =
			!!copied &&
			copied.text.length === text.length &&
			copied.text.every(
				(row, r) =>
					row.length === text[r]?.length &&
					row.every((v, c) => v === text[r]?.[c]),
			);
		this.pasteFormats = same ? copied : null;
		if (!copied || !same) return;
		data.forEach((row, r) =>
			row.forEach((v, c) => {
				if (!copied.formats[r]?.[c]?.checkbox) return;
				const t = sourceText(v).trim();
				if (/^(true|false)$/i.test(t)) row[c] = /^true$/i.test(t);
			}),
		);
	}

	/** Puts the copied formats on the pasted cells, repeated when the paste fills a larger area. */
	private pasteCopiedFormats(
		coords: { startRow: number; startCol: number; endRow: number; endCol: number }[],
	): void {
		const copied = this.pasteFormats;
		this.pasteFormats = null;
		const h = copied?.formats.length ?? 0;
		const w = copied?.formats[0]?.length ?? 0;
		if (!copied || h === 0 || w === 0) return;
		for (const rg of coords) {
			const r1 = Math.min(rg.startRow, rg.endRow);
			const c1 = Math.min(rg.startCol, rg.endCol);
			for (let r = r1; r <= Math.max(rg.startRow, rg.endRow); r++)
				for (let c = c1; c <= Math.max(rg.startCol, rg.endCol); c++) {
					const f = copied.formats[(r - r1) % h]?.[(c - c1) % w];
					if (f) this.model.cellFormats[key(r, c)] = cloneFormat(f);
					else delete this.model.cellFormats[key(r, c)];
				}
		}
		this.markDirty(); // the same undo step as the pasted values
		this.hot?.render();
		this.refreshToolbarState();
	}

	/** Row heights are computed from the fonts; call after anything that can change them. */
	private invalidateRowSizes(): void {
		this.rowIndex = null;
		this.rowHeightCache.clear();
		this.formulaCache.clear();
		this.spillCache.clear();
		this.borderGridCache = null;
		this.invalidateGridRowHeights();
	}

	/**
	 * Handsontable caches the results of the `rowHeights` function and only drops them on
	 * `updateSettings`, so a row whose wrapped text got fewer lines (a wider column) kept its
	 * old height. Ask it to read the heights again.
	 */
	private invalidateGridRowHeights(): void {
		const hot = this.hot;
		if (hot && !hot.isDestroyed) hot.view?.invalidateRowHeightCache?.();
	}

	/**
	 * Sizes in the grid's CSS that follow the zoom of the vertical axis. The cells' vertical
	 * padding is part of the row height. Handsontable never makes a row lower than the theme's
	 * line height plus its padding, so those are zoomed too, or rows could not get smaller than
	 * at 100% (they also size the headers' text). The values are the main theme's.
	 */
	private applyZoomStyles(): void {
		const z = this.zoomY;
		this.gridEl.setCssProps({
			'--rtl-sheet-cell-pad-y': `${(ROW_VERTICAL_PADDING / 2) * z}px`,
			'--ht-line-height': `${Math.round(20 * z)}px`,
			'--ht-cell-vertical-padding': `${Math.round(4 * z * 10) / 10}px`,
			'--ht-font-size': `${Math.round(14 * z * 10) / 10}px`,
		});
	}

	private borderGrid(hot: Handsontable.Core): BorderGrid {
		if (this.borderGridCache) return this.borderGridCache;
		const merges = new Map<string, MergeInfo>();
		for (const m of this.liveMerges(hot))
			for (let r = m.row; r < m.row + m.rowspan; r++)
				for (let c = m.col; c < m.col + m.colspan; c++)
					merges.set(key(r, c), m);
		const hidden = hot.getPlugin('hiddenRows');
		const isHidden = (row: number): boolean => hidden.isHidden(row);
		const rows = hot.countRows();
		const next = new Map<number, number | null>();
		const nextRow = (row: number): number | null => {
			const known = next.get(row);
			if (known !== undefined) return known;
			let r = row + 1;
			while (r < rows && isHidden(r)) r++;
			const found = r < rows ? r : null;
			next.set(row, found);
			return found;
		};
		let firstRow = 0;
		while (firstRow < rows && isHidden(firstRow)) firstRow++;
		// var(--…) and light-dark() colors are worked out inside the grid, where the theme is.
		const root = hot.rootElement;
		const colors = new Map<string, string>();
		const resolve = (color: string): string => {
			const value = color || 'var(--ht-border-color)';
			if (value === 'transparent' || /^(#|rgba?\()/i.test(value)) return value;
			const known = colors.get(value);
			if (known !== undefined) return known;
			const probe =
				root.querySelector<HTMLElement>(':scope > .rtl-sheet-color-probe') ??
				root.createSpan({ cls: 'rtl-sheet-color-probe' });
			probe.style.color = value;
			const out = root.win.getComputedStyle(probe).color || value;
			colors.set(value, out);
			return out;
		};
		this.borderGridCache = { merges, isHidden, nextRow, firstRow, resolve };
		return this.borderGridCache;
	}

	private getRowIndex(): Map<number, number[]> {
		if (this.rowIndex) return this.rowIndex;
		const index = new Map<number, number[]>();
		for (const k of Object.keys(this.model.cellFormats)) {
			const parts = k.split(',').map(Number);
			const r = parts[0];
			const c = parts[1];
			if (r === undefined || c === undefined) continue;
			const list = index.get(r);
			if (list) list.push(c);
			else index.set(r, [c]);
		}
		this.rowIndex = index;
		return index;
	}

	/**
	 * Height of a row in unzoomed px: tall enough for the biggest font in it (like Google
	 * Sheets), for the wrapped lines of "wrap" cells and for vertical text.
	 */
	private computeRowHeight(row: number): number {
		const cached = this.rowHeightCache.get(row);
		if (cached !== undefined) return cached;
		const hot: Handsontable.Core | null = this.gridForCalc ?? this.hot;
		const height = this.measureRowHeight(row, hot);
		// Without a grid the wrapped text was not measured: do not keep that height.
		if (hot && !hot.isDestroyed) this.rowHeightCache.set(row, height);
		return height;
	}

	private measureRowHeight(
		row: number,
		hot: Handsontable.Core | null,
	): number {
		const defaultSize = this.plugin.settings.defaultFontSize;
		let height = rowBaseHeight(defaultSize);
		const cols = this.getRowIndex().get(row);
		if (!cols) return height;

		const rowMerges = this.liveMerges(hot).filter(
			(m) => row >= m.row && row < m.row + m.rowspan,
		);
		for (const col of cols) {
			const f = this.model.cellFormats[key(row, col)];
			if (!f) continue;
			const mg =
				rowMerges.find(
					(m) => col >= m.col && col < m.col + m.colspan,
				) ?? null;
			// Cells hidden inside a merge keep their format but show nothing.
			if (mg && (mg.row !== row || mg.col !== col)) continue;
			const size = f.fontSize ?? defaultSize;
			let need = rowBaseHeight(size);
			const vertical = f.rotate === 90 || f.rotate === -90;
			if ((f.wrap === 'wrap' || vertical) && hot && !hot.isDestroyed) {
				const display = this.resolveRaw(
					hot.getSourceDataAtCell(row, col),
					0,
					false,
					{ row, col },
				);
				const text = display === null ? '' : String(display);
				if (text) {
					if (vertical) {
						need = Math.max(
							need,
							Math.ceil(this.measureTextWidth(text, f)) + 16,
						);
					} else {
						let colWidth = 0;
						for (let c = col; c < col + (mg?.colspan ?? 1); c++)
							colWidth +=
								this.model.colWidths[c] || DEFAULT_COL_WIDTH;
						const lines = this.countWrappedLines(
							text,
							Math.max(20, colWidth - 2 * CELL_PADDING),
							f,
						);
						need = Math.max(
							need,
							lines * Math.ceil(size * ROW_LINE_HEIGHT) +
								ROW_VERTICAL_PADDING,
						);
					}
				}
			}
			// A merged cell also has the height of the other rows it covers.
			if (mg) {
				for (let r = mg.row + 1; r < mg.row + mg.rowspan; r++)
					need -=
						this.model.rowHeights[r] || rowBaseHeight(defaultSize);
			}
			height = Math.max(height, need);
		}
		return height;
	}

	private countWrappedLines(
		text: string,
		maxWidth: number,
		f: CellFormat,
	): number {
		const space = this.measureTextWidth(' ', f);
		let total = 0;
		for (const paragraph of text.split('\n')) {
			const words = paragraph.split(/\s+/).filter((w) => w !== '');
			if (words.length === 0) {
				total += 1;
				continue;
			}
			let lines = 1;
			let lineWidth = 0;
			for (const word of words) {
				const w = this.measureTextWidth(word, f);
				if (lineWidth === 0) {
					lineWidth = w;
				} else if (lineWidth + space + w <= maxWidth) {
					lineWidth += space + w;
				} else {
					lines += 1;
					lineWidth = w;
				}
				if (w > maxWidth) {
					// A word longer than the cell is broken across lines.
					lines += Math.floor(w / maxWidth);
					lineWidth = w % maxWidth;
				}
			}
			total += lines;
		}
		return total;
	}

	/**
	 * The font the cells are drawn with. Handsontable's theme sets its own font on the grid,
	 * so it is read from there: measuring with Obsidian's font gave widths that were too small.
	 */
	private fontFamily(): string {
		if (this.fontFamilyCache) return this.fontFamilyCache;
		const grid = this.gridEl.querySelector('.handsontable');
		const family =
			getComputedStyle(grid ?? this.gridEl).fontFamily || 'sans-serif';
		// Before the grid exists the font is not final: do not keep it.
		if (grid) this.fontFamilyCache = family;
		return family;
	}

	/** Width of text in unzoomed px. */
	private measureTextWidth(text: string, f: CellFormat | undefined): number {
		if (!this.measureCtx)
			this.measureCtx = createEl('canvas').getContext('2d');
		const ctx = this.measureCtx;
		if (!ctx) return 0;
		const size = f?.fontSize ?? this.plugin.settings.defaultFontSize;
		ctx.font = `${f?.italic ? 'italic ' : ''}${f?.bold ? 'bold ' : ''}${size}px ${this.fontFamily()}`;
		return ctx.measureText(text).width;
	}

	/* ---------------- Cell editor ---------------- */

	private onBeginEditing(): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const editor = hot.getActiveEditor() as unknown as
			| { TEXTAREA?: HTMLTextAreaElement }
			| undefined;
		const textarea = editor?.TEXTAREA;
		if (!textarea) return;

		// dir="auto" makes formulas such as =SUM(A1) render left-to-right even in an RTL sheet,
		// while Hebrew text still renders right-to-left.
		textarea.setAttribute('dir', 'auto');
		textarea.setAttribute('autocapitalize', 'off');
		textarea.setAttribute('autocorrect', 'off');
		textarea.spellcheck = false;
		this.autocomplete.attach(textarea);
	}

	/* ---------------- Picking references with the mouse ---------------- */

	/** The field a formula is being typed in (formula bar or cell editor) and the cell it belongs to. */
	private activeFormulaField(): {
		field: HTMLInputElement | HTMLTextAreaElement;
		cell: { row: number; col: number } | null;
	} | null {
		const input = this.ui.formulaInput;
		if (input && input.ownerDocument.activeElement === input)
			return { field: input, cell: this.formulaEditCell };
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return null;
		const editor = hot.getActiveEditor() as unknown as
			| {
					isOpened?: () => boolean;
					TEXTAREA?: HTMLTextAreaElement;
					row?: number;
					col?: number;
			  }
			| undefined;
		if (!editor?.isOpened?.() || !editor.TEXTAREA) return null;
		return {
			field: editor.TEXTAREA,
			cell:
				editor.row === undefined || editor.col === undefined
					? null
					: { row: editor.row, col: editor.col },
		};
	}

	private cellFromEvent(e: Event): { row: number; col: number } | null {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return null;
		const td = (e.target as Element | null)?.closest?.('td');
		if (!td || !this.gridEl.contains(td)) return null;
		const coords = hot.getCoords(td);
		const row = coords?.row ?? -1;
		const col = coords?.col ?? -1;
		return row < 0 || col < 0 ? null : { row, col };
	}

	/**
	 * When a press on the grid should put a reference into the formula being typed: the cell
	 * that was pressed, the field and the part of its text to replace. Null otherwise.
	 */
	private refPickTarget(e: Event): {
		field: HTMLInputElement | HTMLTextAreaElement;
		slot: { from: number; to: number };
		cell: { row: number; col: number };
	} | null {
		if (!this.plugin.settings.enableFormulas) return null;
		const cell = this.cellFromEvent(e);
		const active = cell ? this.activeFormulaField() : null;
		if (!cell || !active) return null;
		// The cell being edited cannot refer to itself.
		if (active.cell?.row === cell.row && active.cell.col === cell.col)
			return null;
		const { field } = active;
		const caret = field.selectionStart ?? field.value.length;
		const slot = referenceSlot(
			field.value,
			caret,
			field.selectionEnd ?? caret,
		);
		return slot ? { field, slot, cell } : null;
	}

	private onGridPress(e: MouseEvent): void {
		if (e.button !== 0) return;
		const target = this.refPickTarget(e);
		if (!target) return;
		// Keep the focus (and the caret) in the field, and keep the press away from Handsontable.
		e.preventDefault();
		e.stopPropagation();

		const { field, slot, cell } = target;
		this.clearRefPick();
		this.refPick = { field, ...slot, anchor: cell, range: null };
		const el: HTMLElement = field;
		el.addEventListener('input', this.onRefPickFieldInput);
		el.addEventListener('keydown', this.onRefPickFieldKey);
		el.addEventListener('blur', this.clearRefPick);
		this.writePickedReference(cell);

		// Dragging over more cells turns the reference into a range such as A1:B5.
		const doc = field.ownerDocument;
		const onMove = (ev: MouseEvent): void => {
			const over = this.cellFromEvent(ev);
			if (over) this.writePickedReference(over);
		};
		const onUp = (): void => {
			doc.removeEventListener('mousemove', onMove, true);
			doc.removeEventListener('mouseup', onUp, true);
		};
		doc.addEventListener('mousemove', onMove, true);
		doc.addEventListener('mouseup', onUp, true);
	}

	/** Writes the reference from the pick's first cell to the given cell into the formula. */
	private writePickedReference(to: { row: number; col: number }): void {
		const pick = this.refPick;
		if (!pick) return;
		const range: SelRange = {
			r1: Math.min(pick.anchor.row, to.row),
			r2: Math.max(pick.anchor.row, to.row),
			c1: Math.min(pick.anchor.col, to.col),
			c2: Math.max(pick.anchor.col, to.col),
		};
		const old = pick.range;
		if (
			old &&
			old.r1 === range.r1 &&
			old.r2 === range.r2 &&
			old.c1 === range.c1 &&
			old.c2 === range.c2
		)
			return;
		pick.range = range;

		const first = `${indexToColumnLabel(range.c1)}${range.r1 + 1}`;
		const last = `${indexToColumnLabel(range.c2)}${range.r2 + 1}`;
		const text = first === last ? first : `${first}:${last}`;
		const { field } = pick;
		field.value = `${field.value.slice(0, pick.from)}${text}${field.value.slice(pick.to)}`;
		pick.to = pick.from + text.length;
		// The caret ends up right after the reference, so the next click replaces it.
		field.setSelectionRange(pick.to, pick.to);
		this.writingRefPick = true;
		try {
			field.dispatchEvent(new Event('input', { bubbles: true }));
		} finally {
			this.writingRefPick = false;
		}
		this.hot?.render(); // moves the highlight of the picked cells
	}

	private readonly onRefPickFieldInput = (): void => {
		if (!this.writingRefPick) this.clearRefPick();
	};

	private readonly onRefPickFieldKey = (e: KeyboardEvent): void => {
		if (e.key === 'Enter' || e.key === 'Escape' || e.key === 'Tab')
			this.clearRefPick();
	};

	/** Forgets the picked reference and removes its highlight (the formula text stays). */
	private readonly clearRefPick = (): void => {
		const pick = this.refPick;
		if (!pick) return;
		this.refPick = null;
		const el: HTMLElement = pick.field;
		el.removeEventListener('input', this.onRefPickFieldInput);
		el.removeEventListener('keydown', this.onRefPickFieldKey);
		el.removeEventListener('blur', this.clearRefPick);
		if (this.hot && !this.hot.isDestroyed) this.hot.render();
	};

	/* ---------------- Renderer ---------------- */

	private readonly cellRenderer = (
		instance: Handsontable.Core,
		td: HTMLTableCellElement,
		row: number,
		col: number,
		prop: string | number,
		value: unknown,
		cellProperties: Handsontable.CellProperties,
	): void => {
		this.gridForCalc = instance;

		const display = this.resolveRaw(value, 0, false, { row, col });
		const isNumber =
			typeof display === 'number' ||
			(typeof display === 'string' && NUMERIC_RE.test(display));
		const text =
			display === null
				? ''
				: typeof display === 'boolean'
					? display
						? 'TRUE'
						: 'FALSE'
					: String(display);

		baseTextRenderer(instance, td, row, col, prop, text, cellProperties);
		// Image cells are read-only, but they should not be greyed out like other read-only cells.
		td.classList.remove(
			'htDimmed',
			'rtl-sheet-img-cell',
			'rtl-sheet-img-active',
			'rtl-sheet-checkbox-cell',
		);
		td.classList.toggle('rtl-sheet-ref-pick', this.isOutlined(row, col));
		const found = this.findMatches[this.findIndex];
		td.classList.toggle(
			'rtl-sheet-find-hit',
			this.findHits.has(key(row, col)),
		);
		td.classList.toggle(
			'rtl-sheet-find-current',
			!!found && found.row === row && found.col === col,
		);

		const f = this.model.cellFormats[key(row, col)];
		const imagePath = this.model.images[key(row, col)];
		const comment = this.model.comments[key(row, col)];
		const link = this.model.links[key(row, col)];
		const rtl = this.model.rtl;
		const wrap = f?.wrap ?? 'overflow';
		const st = td.style;

		st.fontSize = `${Math.round((f?.fontSize ?? this.plugin.settings.defaultFontSize) * this.zoomY * 10) / 10}px`;
		// Line height follows the font size; the theme's fixed line height made big wrapped text overlap.
		st.lineHeight = String(ROW_LINE_HEIGHT);
		st.fontWeight = f?.bold ? 'bold' : '';
		st.fontStyle = f?.italic ? 'italic' : '';
		st.textDecoration = [
			f?.underline ? 'underline' : '',
			f?.strike ? 'line-through' : '',
		]
			.filter((part) => part !== '')
			.join(' ');
		st.backgroundColor = f?.bgColor ?? '';
		st.color = f?.textColor ?? '';
		st.textAlign =
			f?.align ?? (isNumber ? 'right' : rtl ? 'right' : 'left');
		st.verticalAlign = f?.valign ?? '';
		// The text direction follows the content (numbers such as -5 read left-to-right). Never set
		// `direction` on the cell: the grid line is the cell's inline-end border, so it would move
		// to the other side and leave a gap next to every number in a right-to-left sheet.
		st.direction = '';
		st.unicodeBidi = rtl ? 'plaintext' : '';
		st.whiteSpace = wrap === 'wrap' ? 'pre-wrap' : 'nowrap';
		st.wordBreak = wrap === 'wrap' ? 'break-word' : '';
		st.overflow = '';
		st.position = '';
		st.zIndex = '';
		st.clipPath = '';
		// Text that spills over this cell hides the grid line it crosses (like Google Sheets).
		// The border keeps its width, so the columns do not move; only its color goes.
		st.borderInlineEndColor = this.rowSpill(instance, row).crossed.has(col)
			? 'transparent'
			: '';
		applyBorderStyles(
			td,
			this.borderGrid(instance),
			this.model.cellFormats,
			row,
			col,
			rtl,
		);
		// The colored corner of a cell with a comment.
		td.classList.toggle('rtl-sheet-has-comment', !!comment);
		if (comment) st.setProperty('--rtl-sheet-comment-color', comment.color);
		else st.removeProperty('--rtl-sheet-comment-color');

		if (imagePath) {
			this.renderImage(td, imagePath);
			const active = this.formulaCell;
			td.classList.toggle(
				'rtl-sheet-img-active',
				!!active && active.row === row && active.col === col,
			);
		} else if (f?.checkbox) {
			this.renderCheckbox(td, display);
			st.textAlign = f.align ?? 'center';
		} else if (f?.rotate) {
			this.renderRotated(td, text, f.rotate);
		} else if (wrap === 'overflow') {
			const spill = this.rowSpill(instance, row).spills.get(col);
			if (spill) {
				st.overflow = 'visible';
				st.position = 'relative';
				st.zIndex = '1';
				// Paint the overflow only as far as the empty cells it spills into.
				st.clipPath = spill.visualRight
					? `inset(0 -${spill.extra}px 0 0)`
					: `inset(0 0 0 -${spill.extra}px)`;
			}
		}
		if (link && !imagePath && !f?.checkbox && !f?.rotate && text) {
			td.textContent = '';
			td.createSpan({ cls: 'rtl-sheet-link', text });
		}
	};

	private renderCheckbox(td: HTMLTableCellElement, value: Scalar): void {
		td.textContent = '';
		td.addClass('rtl-sheet-checkbox-cell');
		const box = td.createEl('input', {
			type: 'checkbox',
			cls: 'rtl-sheet-checkbox',
			attr: { tabindex: '-1', 'aria-label': 'Check box' },
		});
		box.checked = isChecked(value);
		box.setCssProps({
			'--rtl-sheet-checkbox-size': `${Math.round(16 * this.zoomY)}px`,
		});
	}

	/**
	 * Every spill of a row. A cell has to know whether text from another cell crosses its grid
	 * line, so the whole row is worked out at once and kept until the sheet changes.
	 */
	private rowSpill(hot: Handsontable.Core, row: number): RowSpill {
		const cached = this.spillCache.get(row);
		if (cached) return cached;
		const result: RowSpill = { spills: new Map(), crossed: new Set() };
		this.spillCache.set(row, result);
		const cols = hot.countCols();
		for (let col = 0; col < cols; col++) {
			const raw: unknown = hot.getSourceDataAtCell(row, col);
			if (raw === null || raw === undefined || raw === '') continue;
			if (this.model.images[key(row, col)]) continue;
			const f = this.model.cellFormats[key(row, col)];
			if ((f?.wrap ?? 'overflow') !== 'overflow' || f?.rotate || f?.checkbox)
				continue;
			const mg = this.mergeAt(hot, row, col);
			// Cells hidden inside a merge show nothing; merges over several rows do not spill.
			if (mg && (mg.row !== row || mg.col !== col || mg.rowspan > 1))
				continue;
			const display = this.resolveRaw(raw, 0, false, { row, col });
			if (display === null || typeof display !== 'string') continue;
			if (NUMERIC_RE.test(display)) continue;
			const crossed: number[] = [];
			const spill = this.spillInfo(
				hot,
				row,
				col,
				display,
				f,
				mg?.colspan ?? 1,
				crossed,
			);
			if (!spill) continue;
			result.spills.set(col, spill);
			for (const c of crossed) result.crossed.add(c);
		}
		return result;
	}

	/**
	 * Text that is wider than its cell spills into the empty cells next to it (like Excel).
	 * Numbers and numeric formula results never spill: only text, and only as far as needed.
	 * `crossed` receives the columns whose inline-end grid line the text crosses.
	 */
	private spillInfo(
		hot: Handsontable.Core,
		row: number,
		col: number,
		text: string,
		f: CellFormat | undefined,
		colspan: number,
		crossed: number[],
	): SpillInfo | null {
		if (!text) return null;

		// A merged cell is as wide as all the columns it covers.
		let own = 0;
		for (let c = col; c < col + colspan; c++) own += hot.getColWidth(c);
		const needed =
			this.measureTextWidth(text, f) * this.zoomY + 2 * CELL_PADDING + 2;
		if (needed <= own) return null; // it fits: nothing crosses the cell border

		const visualRight = !/[\u0590-\u08FF]/.test(text);
		// In an RTL sheet, column index + 1 is the visually-left neighbour.
		const delta = visualRight === !this.model.rtl ? 1 : -1;

		let reach = own;
		let extra = 0;
		for (
			let c = delta > 0 ? col + colspan : col - 1;
			c >= 0 && c < hot.countCols();
			c += delta
		) {
			const raw: unknown = hot.getSourceDataAtCell(row, c);
			if (
				!(raw === null || raw === undefined || raw === '') ||
				this.model.images[key(row, c)] ||
				this.model.cellFormats[key(row, c)]?.checkbox ||
				this.mergeAt(hot, row, c)
			)
				break;
			// The line between two columns is the inline-end border of the one with the lower
			// index: this cell's own (or the last neighbour's) going up, the neighbour's going down.
			crossed.push(delta > 0 ? (c === col + colspan ? col : c - 1) : c);
			const w = hot.getColWidth(c);
			extra += w;
			reach += w;
			if (needed <= reach) break;
		}
		return extra > 0 ? { visualRight, extra } : null;
	}

	private renderRotated(
		td: HTMLTableCellElement,
		text: string,
		rotate: number,
	): void {
		td.textContent = '';
		const span = td.createSpan({ cls: 'rtl-sheet-rotated', text });
		if (rotate === 90) {
			span.addClasses(['is-vertical', 'is-flipped']);
		} else if (rotate === -90) {
			span.addClass('is-vertical');
		} else {
			span.addClass('is-angled');
			span.setCssProps({ '--rtl-sheet-rotate': `${-rotate}deg` });
		}
	}

	private renderImage(td: HTMLTableCellElement, path: string): void {
		td.textContent = '';
		const file = this.app.vault.getAbstractFileByPath(path);
		if (!(file instanceof TFile)) {
			td.setText('Missing image');
			return;
		}
		td.addClass('rtl-sheet-img-cell');
		td.createEl('img', {
			cls: 'rtl-sheet-img',
			attr: {
				src: this.app.vault.getResourcePath(file),
				alt: file.name,
				draggable: 'false',
			},
		});
		// Drag a corner to resize the cell, and the image with it (see onImageHandlePress).
		for (const corner of IMAGE_CORNERS)
			td.createDiv({
				cls: `rtl-sheet-img-handle is-${corner}`,
				attr: { 'data-corner': corner },
			});
	}

	/** Marks the selected image cell: on touch screens its resize handles show while it is selected. */
	private markActiveImageCell(): void {
		const hot = this.hot;
		this.gridEl
			.querySelectorAll('td.rtl-sheet-img-active')
			.forEach((el) => el.classList.remove('rtl-sheet-img-active'));
		const cell = this.formulaCell;
		if (!hot || hot.isDestroyed || !cell) return;
		if (!this.model.images[key(cell.row, cell.col)]) return;
		hot.getCell(cell.row, cell.col)?.classList.add('rtl-sheet-img-active');
	}

	/**
	 * A press on a corner handle of an image starts resizing its cell. The row and the column
	 * grow together, so the cell keeps its proportions and the image keeps filling it.
	 */
	private onImageHandlePress(e: PointerEvent): void {
		const handle =
			e.target instanceof HTMLElement
				? e.target.closest<HTMLElement>('.rtl-sheet-img-handle')
				: null;
		const hot = this.hot;
		const td = handle?.closest('td');
		if (!handle || !td || !hot || hot.isDestroyed) return;
		if (e.pointerType === 'mouse' && e.button !== 0) return;
		const coords = hot.getCoords(td);
		const row = Number(coords?.row ?? -1);
		const col = Number(coords?.col ?? -1);
		if (row < 0 || col < 0) return;
		// Keep the press away from Handsontable (it would start selecting cells).
		e.preventDefault();
		e.stopImmediatePropagation();

		const corner = handle.dataset.corner ?? 'se';
		const signX = corner.endsWith('e') ? 1 : -1;
		const signY = corner.startsWith('s') ? 1 : -1;
		const startX = e.clientX;
		const startY = e.clientY;
		const startW =
			this.model.colWidths[col] || hot.getColWidth(col) / this.zoom;
		const startH =
			this.model.rowHeights[row] ||
			(hot.getRowHeight(row) ?? 0) / this.zoomY ||
			this.computeRowHeight(row);
		let width = startW;
		let height = startH;
		let frame = 0;
		let changed = false;

		const apply = (): void => {
			frame = 0;
			if (hot.isDestroyed) return;
			this.setColumnWidth(col, width);
			this.setRowHeight(row, height);
			this.invalidateRowSizes();
			hot.render();
		};
		const onMove = (ev: PointerEvent): void => {
			if (ev.pointerId !== e.pointerId) return;
			ev.preventDefault();
			const dx = ((ev.clientX - startX) * signX) / this.zoom;
			const dy = ((ev.clientY - startY) * signY) / this.zoomY;
			const factor = Math.min(
				MAX_IMAGE_RESIZE_PX / Math.max(startW, startH),
				Math.max(
					MIN_IMAGE_RESIZE_PX / Math.min(startW, startH),
					(startW + dx) / startW,
					(startH + dy) / startH,
				),
			);
			width = Math.round(startW * factor);
			height = Math.round(startH * factor);
			changed = true;
			if (!frame) frame = window.requestAnimationFrame(apply);
		};
		const doc = this.gridEl.ownerDocument;
		const onEnd = (ev: PointerEvent): void => {
			if (ev.pointerId !== e.pointerId) return;
			doc.removeEventListener('pointermove', onMove);
			doc.removeEventListener('pointerup', onEnd);
			doc.removeEventListener('pointercancel', onEnd);
			if (frame) window.cancelAnimationFrame(frame);
			if (!changed || hot.isDestroyed) return;
			apply();
			// The size is still "the image's size", so removing the image restores the old one.
			const k = key(row, col);
			const rec = this.model.imageSizes[k];
			if (rec && this.model.images[k])
				this.model.imageSizes[k] = { ...rec, w: width, h: height };
			this.markDirty(); // one undo step for the whole drag
		};
		doc.addEventListener('pointermove', onMove, { passive: false });
		doc.addEventListener('pointerup', onEnd);
		doc.addEventListener('pointercancel', onEnd);
	}

	/* ---------------- Selection handles (touch) ---------------- */

	/**
	 * Touch screens cannot drag over cells to select them (a drag scrolls), so the selection
	 * gets a handle on each corner: dragging one stretches the selection from the opposite
	 * corner. The handles live inside the scrolled part of the grid, so they move with the
	 * cells and pass under the headers.
	 */
	private installSelectionHandles(hot: Handsontable): void {
		this.selHandles = [];
		if (!IS_TOUCH) return;
		const hider = this.gridEl.querySelector('.ht_master .wtHider');
		if (!hider) return;
		for (const rowEnd of [false, true])
			for (const colEnd of [false, true])
				this.selHandles.push({
					el: hider.createDiv({ cls: 'rtl-sheet-sel-handle' }),
					rowEnd,
					colEnd,
				});
		const queue = (): void => this.queueSelectionHandles();
		hot.addHook('afterSelection', queue);
		hot.addHook('afterDeselect', queue);
		hot.addHook('afterViewRender', queue);
		this.placeSelectionHandles();
	}

	/** Which way the dashes around copied cells move, and how fast, are settings. */
	private applyCopyMarkSettings(): void {
		const s = this.plugin.settings;
		this.gridEl.setCssProps({
			'--rtl-sheet-ant-duration': `${2.5 / s.copyMarkSpeed}s`,
			'--rtl-sheet-ant-direction':
				s.copyMarkMotion === 'left' ? 'reverse' : 'normal',
			'--rtl-sheet-ant-state':
				s.copyMarkMotion === 'still' ? 'paused' : 'running',
		});
	}

	/** The size of the handles and how many of them show are settings. */
	private applySelectionHandleSettings(): void {
		this.gridEl.setCssProps({
			'--rtl-sheet-sel-scale': String(
				this.plugin.settings.selectionHandleScale / 100,
			),
		});
		this.queueSelectionHandles();
	}

	private queueSelectionHandles(): void {
		if (this.selHandles.length === 0 || this.selHandleFrame) return;
		this.selHandleFrame = window.requestAnimationFrame(() => {
			this.selHandleFrame = 0;
			this.placeSelectionHandles();
		});
	}

	/** Puts each handle on its corner of the selection; a corner that is not rendered has none. */
	private placeSelectionHandles(): void {
		const hot = this.hot;
		const hider = this.selHandles[0]?.el.parentElement;
		const ranges = this.getSelectedRanges();
		const rg = ranges[ranges.length - 1];
		const hideAll = (): void => {
			for (const h of this.selHandles) h.el.hide();
		};
		if (!hot || hot.isDestroyed || !hider || !hider.isConnected || !rg)
			return hideAll();

		const hidden = hot.getPlugin('hiddenRows');
		const c1 = rg.c1;
		const c2 = Math.min(rg.c2, hot.countCols() - 1);
		let r1 = rg.r1;
		let r2 = Math.min(rg.r2, hot.countRows() - 1);
		// Rows hidden by a filter have no cells: use the nearest visible row.
		while (r1 < r2 && hidden.isHidden(r1)) r1++;
		while (r2 > r1 && hidden.isHidden(r2)) r2--;
		// A selected image cell shows its own resize handles on the same corners.
		if (
			this.model.images[key(rg.r1, rg.c1)] &&
			hot.getCell(r1, c1) === hot.getCell(r2, c2)
		)
			return hideAll();

		const base = hider.getBoundingClientRect();
		const rtl = this.model.rtl;
		const onlyTwo = this.plugin.settings.selectionHandles === 'two';
		for (const h of this.selHandles) {
			// In a right-to-left sheet the first column is on the right.
			const onRight = h.colEnd !== rtl;
			// With two handles: the top right and the bottom left corners.
			if (onlyTwo && onRight === h.rowEnd) {
				h.el.hide();
				continue;
			}
			const td = hot.getCell(h.rowEnd ? r2 : r1, h.colEnd ? c2 : c1);
			const rect = td?.getBoundingClientRect();
			if (
				!td ||
				!rect ||
				!hider.contains(td) ||
				(rect.width === 0 && rect.height === 0)
			) {
				h.el.hide();
				continue;
			}
			const x = onRight ? rect.right : rect.left;
			const y = h.rowEnd ? rect.bottom : rect.top;
			h.el.style.left = `${x - base.left}px`;
			h.el.style.top = `${y - base.top}px`;
			h.el.show();
		}
	}

	/** A press on a selection handle starts stretching the selection from the opposite corner. */
	private onSelectionHandlePress(e: PointerEvent): void {
		const hot = this.hot;
		const handle = this.selHandles.find((h) => h.el === e.target);
		const ranges = this.getSelectedRanges();
		const rg = ranges[ranges.length - 1];
		if (!handle || !rg || !hot || hot.isDestroyed) return;
		if (e.pointerType === 'mouse' && e.button !== 0) return;
		// Keep the press away from Handsontable and from the scrolling of the grid.
		e.preventDefault();
		e.stopImmediatePropagation();

		const anchorRow = Math.min(
			handle.rowEnd ? rg.r1 : rg.r2,
			hot.countRows() - 1,
		);
		const anchorCol = Math.min(
			handle.colEnd ? rg.c1 : rg.c2,
			hot.countCols() - 1,
		);
		const holder = this.gridEl.querySelector('.ht_master .wtHolder');
		const doc = this.gridEl.ownerDocument;
		let x = e.clientX;
		let y = e.clientY;
		let lastRow = -1;
		let lastCol = -1;
		let frame = 0;
		this.selDragging = true;
		// The handles let the press through, so the cell under the finger can be found.
		this.gridEl.addClass('is-sel-dragging');

		const step = (): void => {
			frame = 0;
			if (hot.isDestroyed) return;
			// Near an edge of the grid the sheet scrolls, so the selection can grow past the screen.
			if (holder) {
				const box = holder.getBoundingClientRect();
				const dx =
					x < box.left + SEL_SCROLL_EDGE
						? -SEL_SCROLL_STEP
						: x > box.right - SEL_SCROLL_EDGE
							? SEL_SCROLL_STEP
							: 0;
				const dy =
					y < box.top + SEL_SCROLL_EDGE
						? -SEL_SCROLL_STEP
						: y > box.bottom - SEL_SCROLL_EDGE
							? SEL_SCROLL_STEP
							: 0;
				if (dx !== 0 || dy !== 0) {
					holder.scrollBy(dx, dy);
					frame = window.requestAnimationFrame(step);
				}
			}
			const cell = doc
				.elementFromPoint(x, y)
				?.closest<HTMLTableCellElement>('td, th');
			if (!cell || !this.gridEl.contains(cell)) return;
			const coords = hot.getCoords(cell);
			if (!coords) return;
			// Over a header the selection reaches the first row or column.
			const row = Math.max(0, coords.row ?? 0);
			const col = Math.max(0, coords.col ?? 0);
			if (row === lastRow && col === lastCol) return;
			lastRow = row;
			lastCol = col;
			hot.selectCell(anchorRow, anchorCol, row, col, false);
		};
		const onMove = (ev: PointerEvent): void => {
			if (ev.pointerId !== e.pointerId) return;
			ev.preventDefault();
			x = ev.clientX;
			y = ev.clientY;
			if (!frame) frame = window.requestAnimationFrame(step);
		};
		const onEnd = (ev: PointerEvent): void => {
			if (ev.pointerId !== e.pointerId) return;
			doc.removeEventListener('pointermove', onMove);
			doc.removeEventListener('pointerup', onEnd);
			doc.removeEventListener('pointercancel', onEnd);
			if (frame) window.cancelAnimationFrame(frame);
			this.selDragging = false;
			this.gridEl.removeClass('is-sel-dragging');
			if (lastRow < 0 || hot.isDestroyed) return;
			this.onSelectionChanged();
			this.queueSelectionHandles();
		};
		doc.addEventListener('pointermove', onMove, { passive: false });
		doc.addEventListener('pointerup', onEnd);
		doc.addEventListener('pointercancel', onEnd);
	}

	/** Mouse and touch events that follow a press on a handle must not reach the grid. */
	private keepHandlePress(e: Event): void {
		if (
			e.target instanceof HTMLElement &&
			e.target.closest('.rtl-sheet-img-handle, .rtl-sheet-sel-handle')
		) {
			if (e.cancelable && e.type === 'mousedown') e.preventDefault();
			e.stopImmediatePropagation();
		}
	}

	/* ---------------- Formula bar ---------------- */

	private buildFormulaBar(): void {
		const bar = this.formulaBarEl;
		bar.empty();

		this.ui.address = bar.createDiv({ cls: 'rtl-sheet-address' });
		this.addButton(bar, '', 'fx', 'Insert function', () =>
			this.openFunctionPicker(),
		);

		const input = bar.createEl('input', {
			type: 'text',
			cls: 'rtl-sheet-formula-input',
			attr: {
				placeholder: 'Value or formula, e.g. =SUM(A1:A3)',
				dir: 'auto',
				spellcheck: 'false',
				autocomplete: 'off',
				autocapitalize: 'off',
				autocorrect: 'off',
			},
		});
		this.ui.formulaInput = input;
		this.autocomplete.attach(input);

		input.addEventListener('focus', () => {
			this.formulaEditCell = this.formulaCell;
			this.hot?.unlisten();
		});
		input.addEventListener('blur', () => {
			this.commitFormulaBar();
			this.hot?.listen();
		});
		input.addEventListener('keydown', (e) => {
			if (e.key === 'Enter') {
				e.preventDefault();
				this.commitFormulaBar();
				input.blur();
			} else if (e.key === 'Escape') {
				this.refreshToolbarState(true);
				input.blur();
			}
		});
	}

	private commitFormulaBar(): void {
		const hot = this.hot;
		const input = this.ui.formulaInput;
		const cell = this.formulaEditCell ?? this.formulaCell;
		if (!hot || hot.isDestroyed || !input || !cell) return;
		const current = hot.getSourceDataAtCell(cell.row, cell.col);
		const currentText = sourceText(current);
		if (input.value === currentText) return;
		hot.setDataAtCell(
			cell.row,
			cell.col,
			input.value === '' ? null : input.value,
		);
	}

	private openFunctionPicker(): void {
		new FunctionPickerModal(this.app, (fn) => {
			// Let the modal finish closing before touching focus.
			window.setTimeout(() => this.insertFunction(fn), 50);
		}).open();
	}

	private insertFunction(fn: FormulaFunctionInfo): void {
		const input = this.ui.formulaInput;
		if (!input) return;
		if (!this.formulaCell) {
			new Notice('Select a cell first.');
			return;
		}
		this.formulaEditCell = this.formulaCell;
		input.focus();

		let value = input.value;
		let start = input.selectionStart ?? value.length;
		let end = input.selectionEnd ?? start;
		if (!value.startsWith('=')) {
			value = '=';
			start = 1;
			end = 1;
		}
		const call = buildFunctionCall(fn);
		input.value = `${value.slice(0, start)}${call.text}${value.slice(end)}`;
		input.setSelectionRange(start + call.selectFrom, start + call.selectTo);
	}

	/* ---------------- Toolbar ---------------- */

	/** The toolbar size is a setting, separate for mobile and desktop. */
	private applyToolbarScale(): void {
		const s = this.plugin.settings;
		const percent = Platform.isMobile
			? s.toolbarScaleMobile
			: s.toolbarScaleDesktop;
		this.toolbarEl.setCssProps({
			'--rtl-sheet-toolbar-zoom': String(percent / 100),
		});
	}

	private buildToolbar(): void {
		const tb = this.toolbarEl;
		tb.empty();
		this.ui.toggles = {};
		this.ui.palettes = {};
		const group = (): HTMLElement =>
			tb.createDiv({ cls: 'rtl-sheet-toolbar-group' });

		let g = group();
		this.addButton(g, 'undo', 'Undo', 'Undo', () => this.undoRedo('undo'));
		this.addButton(g, 'redo', 'Redo', 'Redo', () => this.undoRedo('redo'));
		this.ui.brush = this.addButton(
			g,
			'paintbrush',
			'Brush',
			'Format painter: copy the formatting of the selected cell, then select the cells to apply it to',
			() => this.toggleFormatPainter(),
		);

		g = group();
		this.ui.fontSelect = this.addSelect(
			g,
			'Font size',
			[
				{ value: '', label: 'Size' },
				...FONT_SIZES.map((s) => ({
					value: String(s),
					label: String(s),
				})),
			],
			(ranges, value) =>
				this.applyFormat(ranges, (f) => {
					if (value) f.fontSize = Number(value);
					else delete f.fontSize;
				}),
		);
		this.addToggle(g, 'bold', 'bold', 'B', 'Bold', () =>
			this.toggleFlag('bold'),
		);
		this.addToggle(g, 'italic', 'italic', 'I', 'Italic', () =>
			this.toggleFlag('italic'),
		);
		this.addToggle(g, 'underline', 'underline', 'U', 'Underline', () =>
			this.toggleFlag('underline'),
		);
		this.addToggle(g, 'strike', 'strikethrough', 'S', 'Strikethrough', () =>
			this.toggleFlag('strike'),
		);

		g = group();
		this.addColorPicker(
			g,
			'Fill',
			'Cell fill color',
			'#ffff00',
			'fill',
			(f, color) => {
				f.bgColor = color;
			},
		);
		this.addColorPicker(
			g,
			'Text',
			'Text color',
			'#d00000',
			'text',
			(f, color) => {
				f.textColor = color;
			},
		);
		this.addButton(g, 'eraser', 'x', 'Clear fill and text color', () =>
			this.applyFormat(this.requireSelection(), (f) => {
				delete f.bgColor;
				delete f.textColor;
			}),
		);

		g = group();
		this.addToggle(
			g,
			'align-left',
			'align-left',
			'Left',
			'Align left',
			() => this.setFormatProp('align', 'left'),
		);
		this.addToggle(
			g,
			'align-center',
			'align-center',
			'Center',
			'Align center',
			() => this.setFormatProp('align', 'center'),
		);
		this.addToggle(
			g,
			'align-right',
			'align-right',
			'Right',
			'Align right',
			() => this.setFormatProp('align', 'right'),
		);

		g = group();
		this.addToggle(
			g,
			'valign-top',
			'align-vertical-justify-start',
			'Top',
			'Align top',
			() => this.setFormatProp('valign', 'top'),
		);
		this.addToggle(
			g,
			'valign-middle',
			'align-vertical-justify-center',
			'Middle',
			'Align middle',
			() => this.setFormatProp('valign', 'middle'),
		);
		this.addToggle(
			g,
			'valign-bottom',
			'align-vertical-justify-end',
			'Bottom',
			'Align bottom',
			() => this.setFormatProp('valign', 'bottom'),
		);

		g = group();
		this.addToggle(
			g,
			'wrap-expand',
			'move-horizontal',
			'Widen',
			'Excess text: widen the column to fit the text',
			() => this.setFormatProp('wrap', 'expand'),
		);
		this.addToggle(
			g,
			'wrap-overflow',
			'arrow-right-from-line',
			'Spill',
			'Excess text: spill into the neighbouring cell (only if it is empty)',
			() => this.setFormatProp('wrap', 'overflow'),
		);
		this.addToggle(
			g,
			'wrap-wrap',
			'wrap-text',
			'Wrap',
			'Excess text: continue on a new line inside the same cell (the column keeps its width)',
			() => this.setFormatProp('wrap', 'wrap'),
		);
		this.ui.rotateSelect = this.addSelect(
			g,
			'Text rotation',
			ROTATIONS.map((r) => ({ value: String(r.value), label: r.label })),
			(ranges, value) =>
				this.applyFormat(ranges, (f) => {
					const n = Number(value);
					if (n) f.rotate = n;
					else delete f.rotate;
				}),
		);

		g = group();
		this.addButton(g, 'merge', 'Merge', 'Merge cells', () =>
			this.mergeSelection(),
		);
		this.addButton(g, 'split', 'Unmerge', 'Unmerge cells', () =>
			this.unmergeSelection(),
		);

		g = group();
		const bordersBtn: HTMLButtonElement = this.addButton(
			g,
			'grid-3x3',
			'Borders',
			'Borders',
			() =>
				this.showMenu(bordersBtn, [
					{
						title: 'All borders (every selected cell)',
						// grid-3x3 is not in Obsidian's icon set, so the item showed no icon.
						icon: 'table',
						onClick: () => this.applyBorders('all'),
					},
					{
						title: 'Outer frame only',
						icon: 'square',
						onClick: () => this.applyBorders('outer'),
					},
					{
						title: 'Top edge of the selection',
						icon: 'panel-top',
						onClick: () => this.applyBorders('top'),
					},
					{
						title: 'Bottom edge of the selection',
						icon: 'panel-bottom',
						onClick: () => this.applyBorders('bottom'),
					},
					{
						title: 'Left edge of the selection',
						icon: 'panel-left',
						onClick: () => this.applyBorders('left'),
					},
					{
						title: 'Right edge of the selection',
						icon: 'panel-right',
						onClick: () => this.applyBorders('right'),
					},
					{
						title: 'No borders',
						icon: 'ban',
						onClick: () => this.applyBorders('none'),
					},
				]),
		);
		this.addColorPicker(
			g,
			'Border',
			'Border color',
			'#000000',
			'border',
			(f, color) => {
				f.borders = {
					top: true,
					right: true,
					bottom: true,
					left: true,
					...(f.borders ?? {}),
					color,
				};
			},
		);

		g = group();
		const sortBtn: HTMLButtonElement = this.addButton(
			g,
			'arrow-up-down',
			'Sort',
			'Sort rows by the column of the selected cell',
			() => {
				const column = this.activeColumnName();
				this.showMenu(sortBtn, [
					{
						title: `Sort by ${column}, A to Z`,
						icon: 'arrow-down-a-z',
						onClick: () => this.sortByColumn(false),
					},
					{
						title: `Sort by ${column}, Z to A`,
						icon: 'arrow-up-a-z',
						onClick: () => this.sortByColumn(true),
					},
					this.headerRowEntry(),
				]);
			},
		);
		const filterBtn: HTMLButtonElement = this.addButton(
			g,
			'filter',
			'Filter',
			'Filter rows by the column of the selected cell',
			() => {
				const column = this.activeColumnName();
				const col = this.formulaCell?.col ?? -1;
				const entries: MenuEntry[] = [
					{
						title: `Filter ${column}…`,
						icon: 'filter',
						onClick: () => this.openColumnFilter(),
					},
				];
				if (this.filters.has(col))
					entries.push({
						title: `Remove the filter from ${column}`,
						icon: 'filter-x',
						onClick: () => this.setColumnFilter(col, null),
					});
				if (this.filters.size > 0)
					entries.push(
						{
							title: 'Apply filters again',
							icon: 'refresh-cw',
							onClick: () => this.applyFilters(),
						},
						{
							title: 'Remove all filters',
							icon: 'filter-x',
							onClick: () => {
								this.filters.clear();
								this.applyFilters();
							},
						},
					);
				entries.push(this.headerRowEntry());
				this.showMenu(filterBtn, entries);
			},
		);
		this.ui.filterBtn = filterBtn;
		this.addButton(g, 'search', 'Find', 'Find in sheet (Ctrl+F)', () => {
			if (this.findBarEl.isShown()) this.closeFind();
			else this.openFind();
		});

		g = group();
		const imageBtn: HTMLButtonElement = this.addButton(
			g,
			'image',
			'Image',
			'Insert image into the selected cell',
			() =>
				this.showMenu(imageBtn, [
					{
						title: 'Choose from vault…',
						icon: 'folder-open',
						onClick: () => this.pickImageFromVault(),
					},
					{
						title: 'Upload from device…',
						icon: 'upload',
						onClick: () => void this.pickImageFromDevice(),
					},
					{
						title: 'Paste from clipboard',
						icon: 'clipboard-paste',
						onClick: () => void this.pasteImageFromClipboard(),
					},
					{
						title: 'Floating image from vault…',
						icon: 'image-plus',
						onClick: () => this.pickFloatingImageFromVault(),
					},
					{
						title: 'Floating image from device…',
						icon: 'image-up',
						onClick: () => void this.insertFloatingImageFrom('device'),
					},
					{
						title: 'Floating image from clipboard',
						icon: 'clipboard-paste',
						onClick: () =>
							void this.insertFloatingImageFrom('clipboard'),
					},
				]),
		);
		this.addButton(
			g,
			'image-off',
			'No image',
			'Remove image from the selected cells',
			() => this.removeImages(),
		);

		g = group();
		const chartBtn: HTMLButtonElement = this.addButton(
			g,
			'chart-column',
			'Chart',
			'Insert a chart: column, bar, line, pie or scatter',
			() =>
				this.showMenu(
					chartBtn,
					CHART_TYPES.map((t) => ({
						title: t.label,
						icon: CHART_ICONS[t.value],
						onClick: () => this.insertChart(t.value),
					})),
				),
		);

		g = group();
		const emojiBtn: HTMLButtonElement = this.addButton(
			g,
			'smile',
			'Emoji',
			'Insert an emoji',
			() => this.openEmojiPicker(emojiBtn),
		);
		this.addButton(
			g,
			'square-check',
			'Checkbox',
			'Check boxes in the selected cells (Space checks them; again to remove)',
			() => this.toggleCheckboxes(),
		);
		this.addButton(
			g,
			'link',
			'Link',
			'Insert link (Ctrl+click a link to open it)',
			() => this.editLink(),
		);

		g = group();
		this.addButton(
			g,
			'message-square-plus',
			'Add comment',
			'Add or edit the comment of the selected cell',
			() => this.editComment(),
		);
		this.addButton(
			g,
			'messages-square',
			'Show comments',
			'Show all comments in this sheet',
			() => this.toggleSidePanel('comments'),
		);

		g = group();
		const aiBtn: HTMLButtonElement = this.addButton(
			g,
			'sparkles',
			'AI',
			'AI assistant and AI formulas (Claude and Gemini)',
			() =>
				this.showMenu(aiBtn, [
					{
						title: 'AI assistant',
						icon: 'sparkles',
						onClick: () => this.toggleSidePanel('ai'),
						checked: this.sidePanel === 'ai',
					},
					{
						title: 'Insert CLAUDE() formula',
						icon: 'square-function',
						onClick: () => this.insertAiFormula('CLAUDE'),
					},
					{
						title: 'Insert GEMINI() formula',
						icon: 'square-function',
						onClick: () => this.insertAiFormula('GEMINI'),
					},
					{
						title: 'Ask again for the AI formulas in the selection',
						icon: 'refresh-cw',
						onClick: () => this.refreshAiFormulas(),
					},
					{
						title: 'AI settings…',
						icon: 'settings',
						onClick: () => this.openPluginSettings(),
					},
				]),
		);

		g = group();
		const importBtn: HTMLButtonElement = this.addButton(
			g,
			'file-input',
			'Import',
			'Import an Excel (.xlsx) or CSV file into this sheet',
			() =>
				this.showMenu(importBtn, [
					{
						title: 'From this device…',
						icon: 'upload',
						onClick: () => void this.importFromDevice(),
					},
					{
						title: 'From vault…',
						icon: 'folder-open',
						onClick: () => this.importFromVault(),
					},
				]),
		);
		this.addButton(
			g,
			'download',
			'Export',
			'Export to Excel (.xlsx)',
			() => {
				void this.exportToXlsx();
			},
		);

		if (IS_TOUCH) {
			// Long-press context menus are not reliable on every touch platform,
			// so rows and columns get explicit buttons.
			g = group();
			this.addButton(g, '', '+Row', 'Insert row below', () =>
				this.alterStructure('insert_row_below'),
			);
			this.addButton(g, '', '-Row', 'Delete selected rows', () =>
				this.alterStructure('remove_row'),
			);
			this.addButton(g, '', '+Col', 'Insert column after', () =>
				this.alterStructure('insert_col_end'),
			);
			this.addButton(g, '', '-Col', 'Delete selected columns', () =>
				this.alterStructure('remove_col'),
			);
		}

		g = group();
		this.addButton(g, '', '-', 'Zoom out', () =>
			this.applyZoom(this.zoom - 0.1),
		);
		this.ui.zoomLabel = this.addButton(
			g,
			'',
			'100%',
			'Reset zoom to 100%',
			() => this.applyZoom(1),
		);
		this.addButton(g, '', '+', 'Zoom in', () =>
			this.applyZoom(this.zoom + 0.1),
		);

		g = group();
		this.addButton(
			g,
			'arrow-left-right',
			'RTL/LTR',
			'Toggle RTL / LTR',
			() => this.toggleDirection(),
		);
	}

	private addButton(
		parent: HTMLElement,
		icon: string,
		label: string,
		title: string,
		onClick: () => void,
	): HTMLButtonElement {
		const btn = parent.createEl('button', {
			cls: 'rtl-sheet-btn',
			attr: { 'aria-label': title, title },
		});
		if (icon) setIcon(btn, icon);
		if (!btn.querySelector('svg')) btn.setText(label);
		// Keep the grid selection intact while clicking toolbar buttons.
		btn.addEventListener('mousedown', (e) => e.preventDefault());
		btn.addEventListener('click', () => {
			try {
				onClick();
			} catch (e) {
				this.plugin.reportError(title, e);
			}
		});
		return btn;
	}

	private addToggle(
		parent: HTMLElement,
		name: string,
		icon: string,
		label: string,
		title: string,
		onClick: () => void,
	): void {
		this.ui.toggles[name] = this.addButton(
			parent,
			icon,
			label,
			title,
			onClick,
		);
	}

	/**
	 * Remembers the selection when a control is first touched, so that opening a
	 * picker and clicking elsewhere in the grid cannot redirect the change to a
	 * different cell.
	 */
	private bindRangeControl(
		el: HTMLElement,
		eventName: 'input' | 'change',
		handler: (ranges: SelRange[]) => void,
	): void {
		let captured: SelRange[] = [];
		const capture = (): void => {
			const current = this.getSelectedRanges();
			if (current.length > 0) captured = current;
		};
		el.addEventListener('pointerdown', capture);
		el.addEventListener('focus', capture);
		el.addEventListener(eventName, () => {
			const ranges =
				captured.length > 0 ? captured : this.getSelectedRanges();
			if (ranges.length === 0) {
				new Notice('Select one or more cells first.');
				return;
			}
			try {
				handler(ranges);
			} catch (e) {
				this.plugin.reportError('Could not apply formatting', e);
			}
		});
	}

	private addColorPicker(
		parent: HTMLElement,
		label: string,
		title: string,
		initial: string,
		kind: ColorKind,
		setColor: (f: CellFormat, color: string) => void,
	): void {
		const wrap = parent.createEl('label', {
			cls: 'rtl-sheet-color',
			attr: { title },
		});
		wrap.createSpan({ text: label });
		const input = wrap.createEl('input', { type: 'color' });
		input.value = initial;
		this.bindRangeControl(input, 'input', (ranges) =>
			this.applyFormat(ranges, (f) => setColor(f, input.value), kind),
		);
		if (kind === 'border') return;

		// Saved colors: one click applies a color without opening the picker.
		this.ui.palettes[kind] = {
			el: parent.createDiv({ cls: 'rtl-sheet-swatches' }),
			input,
			setColor,
		};
		const manageBtn: HTMLButtonElement = this.addButton(
			parent,
			'plus',
			'+',
			`Save the current ${label.toLowerCase()} color, or remove a saved one`,
			() => this.showPaletteMenu(kind, manageBtn),
		);
		this.renderPalette(kind);
	}

	private savedColors(kind: PaletteKind): string[] {
		const s = this.plugin.settings;
		return kind === 'fill' ? s.savedFillColors : s.savedTextColors;
	}

	/** Saved colors are a plugin setting, so every sheet shows the same ones. */
	private setSavedColors(kind: PaletteKind, colors: string[]): void {
		if (kind === 'fill') this.plugin.settings.savedFillColors = colors;
		else this.plugin.settings.savedTextColors = colors;
		void this.plugin
			.saveSettings()
			.catch((e) =>
				this.plugin.reportError('Could not save the colors', e),
			);
	}

	private renderPalette(kind: PaletteKind): void {
		const palette = this.ui.palettes[kind];
		if (!palette) return;
		palette.el.empty();
		const colors = this.savedColors(kind);
		for (const color of colors) {
			const hint = `Apply ${color} (right-click to remove it)`;
			const swatch = palette.el.createEl('button', {
				cls: 'rtl-sheet-swatch',
				attr: { 'aria-label': hint, title: hint },
			});
			swatch.setCssProps({ '--rtl-sheet-swatch': color });
			// Keep the grid selection intact while clicking.
			swatch.addEventListener('mousedown', (e) => e.preventDefault());
			swatch.addEventListener('click', () => {
				palette.input.value = color;
				this.applyFormat(this.requireSelection(), (f) =>
					palette.setColor(f, color),
				);
			});
			swatch.addEventListener('contextmenu', (e) => {
				e.preventDefault();
				this.setSavedColors(
					kind,
					colors.filter((c) => c !== color),
				);
			});
		}
	}

	private showPaletteMenu(kind: PaletteKind, anchor: HTMLElement): void {
		const palette = this.ui.palettes[kind];
		if (!palette) return;
		const colors = this.savedColors(kind);
		const current = palette.input.value.toLowerCase();
		this.showMenu(anchor, [
			{
				title: `Save ${current}`,
				icon: 'plus',
				onClick: () => {
					if (colors.includes(current))
						new Notice('This color is already saved.');
					else if (colors.length >= MAX_SAVED_COLORS)
						new Notice(
							`You can save up to ${MAX_SAVED_COLORS} colors. Remove one first.`,
						);
					else this.setSavedColors(kind, [...colors, current]);
				},
			},
			...colors.map((color) => ({
				title: `Remove ${color}`,
				icon: 'trash-2',
				onClick: () =>
					this.setSavedColors(
						kind,
						colors.filter((c) => c !== color),
					),
			})),
		]);
	}

	private addSelect(
		parent: HTMLElement,
		title: string,
		options: { value: string; label: string }[],
		onChange: (ranges: SelRange[], value: string) => void,
	): HTMLSelectElement {
		const select = parent.createEl('select', {
			cls: 'dropdown rtl-sheet-select',
			attr: { title, 'aria-label': title },
		});
		for (const o of options)
			select.createEl('option', { value: o.value, text: o.label });
		this.bindRangeControl(select, 'change', (ranges) =>
			onChange(ranges, select.value),
		);
		return select;
	}

	/** Shows a menu under a button, or where the pointer is (a right-click). */
	private showMenu(
		anchor: HTMLElement | MouseEvent,
		entries: MenuEntry[],
	): void {
		const menu = new Menu();
		for (const entry of entries) {
			menu.addItem((item) => {
				item.setTitle(entry.title)
					.setIcon(entry.icon)
					.onClick(entry.onClick);
				if (entry.checked !== undefined) item.setChecked(entry.checked);
			});
		}
		if (anchor instanceof MouseEvent) {
			menu.showAtMouseEvent(anchor);
			return;
		}
		const rect = anchor.getBoundingClientRect();
		menu.showAtPosition({ x: rect.left, y: rect.bottom });
	}

	/** Syncs button highlights, selects and the formula bar with the active cell. */
	private refreshToolbarState(forceFormulaBar = false): void {
		const cell = this.formulaCell;
		const f = cell
			? this.model.cellFormats[key(cell.row, cell.col)]
			: undefined;
		const setActive = (name: string, on: boolean): void => {
			this.ui.toggles[name]?.toggleClass('is-active', on);
		};
		setActive('bold', !!f?.bold);
		setActive('italic', !!f?.italic);
		setActive('underline', !!f?.underline);
		setActive('strike', !!f?.strike);
		setActive('align-left', f?.align === 'left');
		setActive('align-center', f?.align === 'center');
		setActive('align-right', f?.align === 'right');
		setActive('valign-top', f?.valign === 'top');
		setActive('valign-middle', f?.valign === 'middle');
		setActive('valign-bottom', f?.valign === 'bottom');
		setActive('wrap-expand', f?.wrap === 'expand');
		setActive('wrap-wrap', f?.wrap === 'wrap');
		setActive('wrap-overflow', (f?.wrap ?? 'overflow') === 'overflow');

		if (this.ui.fontSelect)
			this.ui.fontSelect.value = f?.fontSize ? String(f.fontSize) : '';
		if (this.ui.rotateSelect)
			this.ui.rotateSelect.value = String(f?.rotate ?? 0);

		this.ui.address?.setText(
			cell ? `${indexToColumnLabel(cell.col)}${cell.row + 1}` : '',
		);
		const input = this.ui.formulaInput;
		if (
			input &&
			(forceFormulaBar || input.ownerDocument.activeElement !== input)
		) {
			const hot = this.hot;
			const raw: unknown =
				cell && hot && !hot.isDestroyed
					? hot.getSourceDataAtCell(cell.row, cell.col)
					: null;
			input.value = sourceText(raw);
		}
	}

	/** True for a cell picked into the formula being typed, or used by the selected cell's formula. */
	private isOutlined(row: number, col: number): boolean {
		const inside = (rg: SelRange): boolean =>
			row >= rg.r1 && row <= rg.r2 && col >= rg.c1 && col <= rg.c2;
		const picked = this.refPick?.range;
		return (
			(!!picked && inside(picked)) ||
			this.formulaRefs.some(inside) ||
			this.chartPickRanges.some(inside)
		);
	}

	/** Reads the references of the selected cell's formula; true when they changed. */
	private readFormulaRefs(): boolean {
		const hot = this.hot;
		const cell = this.formulaCell;
		const raw: unknown =
			cell &&
			hot &&
			!hot.isDestroyed &&
			this.plugin.settings.enableFormulas
				? hot.getSourceDataAtCell(cell.row, cell.col)
				: null;
		const next = typeof raw === 'string' ? formulaReferences(raw) : [];
		const changed = JSON.stringify(next) !== JSON.stringify(this.formulaRefs);
		this.formulaRefs = next;
		return changed;
	}

	/**
	 * Outlines the cells that the selected cell's formula uses (the G30 of "=G30"). Only the
	 * cells on screen are touched, so moving the selection does not draw the grid again; the
	 * renderer outlines the rest as they scroll into view.
	 */
	private refreshFormulaRefs(): void {
		if (this.readFormulaRefs()) this.paintOutlines();
	}

	/** Updates the dashed outlines of the cells on screen (see isOutlined). */
	private paintOutlines(): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		this.gridEl
			.querySelectorAll<HTMLTableCellElement>('.ht_master td')
			.forEach((td) => {
				const coords = hot.getCoords(td);
				const row = coords?.row ?? -1;
				const col = coords?.col ?? -1;
				td.classList.toggle(
					'rtl-sheet-ref-pick',
					row >= 0 && col >= 0 && this.isOutlined(row, col),
				);
			});
	}

	private onSelectionChanged(): void {
		this.clearRefPick();
		// While the chart editor is open, selected cells go into its active range field.
		if (this.chartEditor) {
			const ranges = this.getSelectedRanges();
			const last = ranges[ranges.length - 1];
			if (last) this.chartEditor.pickRange(last);
		}
		const first = this.getSelectedRanges()[0];
		this.formulaCell = first ? { row: first.r1, col: first.c1 } : null;
		this.refreshFormulaRefs();
		this.markActiveImageCell();
		this.setActiveFloat(null);
		this.queueCellInfo(0);

		if (this.painterSource) {
			const ranges = this.getSelectedRanges();
			const source = this.painterSource;
			this.applyFormat(ranges, (f) => {
				for (const k of Object.keys(f))
					delete (f as Record<string, unknown>)[k];
				Object.assign(f, cloneFormat(source));
			});
			// One-shot on desktop; stays on for touch so several taps can be painted.
			if (!IS_TOUCH) this.setPainter(null);
		}
		this.refreshToolbarState();
	}

	/* ---------------- Formatting actions ---------------- */

	private getSelectedRanges(): SelRange[] {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return [];
		const sel = hot.getSelected();
		if (!sel) return [];
		return sel
			.map((sel0) => {
				const [a = 0, b = 0, c = 0, d = 0] = sel0;
				return {
					r1: Math.max(0, Math.min(a, c)),
					c1: Math.max(0, Math.min(b, d)),
					r2: Math.max(a, c),
					c2: Math.max(b, d),
				};
			})
			.filter((rg) => rg.r2 >= 0 && rg.c2 >= 0);
	}

	private requireSelection(): SelRange[] {
		const ranges = this.getSelectedRanges();
		if (ranges.length === 0) new Notice('Select one or more cells first.');
		return ranges;
	}

	private applyFormat(
		ranges: SelRange[],
		mutate: (f: CellFormat, r: number, c: number, range: SelRange) => void,
		coalesce = '',
	): void {
		if (ranges.length === 0 || !this.hot) return;
		// Stored formats are replaced, never changed in place: undo steps keep the old objects.
		for (const rg of ranges) {
			for (let r = rg.r1; r <= rg.r2; r++) {
				for (let c = rg.c1; c <= rg.c2; c++) {
					const k = key(r, c);
					const f: CellFormat = {
						...(this.model.cellFormats[k] ?? {}),
					};
					mutate(f, r, c, rg);
					if (isEmptyFormat(f)) delete this.model.cellFormats[k];
					else this.model.cellFormats[k] = f;
				}
			}
		}
		this.invalidateRowSizes();
		const cols = new Set<number>();
		for (const rg of ranges)
			for (let c = rg.c1; c <= rg.c2; c++) cols.add(c);
		this.fitExpandColumns(cols);
		this.hot.render();
		// A color picker reports every color the pointer passes over: keep one drag one undo step.
		if (coalesce)
			this.coalesce = {
				tag: `${coalesce}:${JSON.stringify(ranges)}`,
				windowMs: 1500,
			};
		this.markDirty();
		this.refreshToolbarState();
	}

	/* ---------------- Undo ---------------- */

	private undoPlugin(): UndoRedoPlugin | null {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return null;
		return hot.getPlugin('undoRedo');
	}

	private captureSideState(): SideState {
		const m = this.model;
		return {
			formats: { ...m.cellFormats },
			images: { ...m.images },
			imageSizes: { ...m.imageSizes },
			comments: { ...m.comments },
			links: { ...m.links },
			widenRecords: JSON.stringify(m.widenRecords ?? {}),
			charts: JSON.stringify(m.charts),
			floatingImages: JSON.stringify(m.floatingImages),
			colWidths: m.colWidths.slice(),
			rowHeights: m.rowHeights.slice(),
		};
	}

	/** Forgets what the side state was compared with (after a load or an undo). */
	private resetSideBaseline(): void {
		this.sideBaseline = this.captureSideState();
	}

	/** Checks for side-state changes once the current action has finished (see markDirty). */
	private queueSideCheck(): void {
		if (this.sideCheckQueued) return;
		this.sideCheckQueued = true;
		queueMicrotask(() => {
			this.sideCheckQueued = false;
			try {
				this.recordSideChanges();
			} catch (e) {
				console.error('[Free-Excel] Could not record an undo step', e);
			}
		});
	}

	/**
	 * Makes whatever just changed outside the cell values undoable. Handsontable only undoes
	 * values, merges and inserted or removed rows and columns; formats, images and sizes are
	 * compared with how they were after the previous action. A difference is attached to the
	 * steps Handsontable recorded for the same action, or becomes a step of its own, so one
	 * Ctrl+Z takes back the whole action (like Google Sheets).
	 */
	private recordSideChanges(): void {
		const plugin = this.undoPlugin();
		const coalesce = this.coalesce;
		this.coalesce = null;
		if (!plugin || this.sideBusy) return;
		const baseline = this.sideBaseline;
		const current = this.captureSideState();
		this.sideBaseline = current;

		// The steps on top of the stack that were added since the last check.
		const done = plugin.doneActions as UndoStep[];
		let firstNew = done.length;
		while (firstNew > 0) {
			const action = done[firstNew - 1];
			if (!action || this.knownSteps.has(action)) break;
			firstNew--;
		}
		const fresh = done.slice(firstNew);
		for (const action of fresh) this.knownSteps.add(action);
		if (
			plugin.ignoreNewActions || // an undo or redo is still running
			!baseline ||
			sameSideState(baseline, current)
		)
			return;

		const now = Date.now();
		const last = this.lastSideStep;
		if (
			fresh.length === 0 &&
			coalesce &&
			last &&
			last.tag === coalesce.tag &&
			now - last.time < coalesce.windowMs &&
			done[done.length - 1] === last.step
		) {
			last.state.after = current;
			last.time = now;
			return;
		}

		const state = { before: baseline, after: current };
		const step = this.sideStep(fresh, state);
		this.knownSteps.add(step);
		if (fresh.length > 0) done.splice(firstNew, fresh.length, step);
		else plugin.done(() => step);
		this.lastSideStep = coalesce
			? { step, tag: coalesce.tag, time: now, state }
			: null;
	}

	/** An undo step that runs Handsontable's own steps and then puts the side state back. */
	private sideStep(
		inner: UndoStep[],
		state: { before: SideState; after: SideState },
	): UndoStep {
		const run = (
			hot: Handsontable.Core,
			steps: UndoStep[],
			method: 'undo' | 'redo',
			target: () => SideState,
			done: () => void,
		): void => {
			this.sideBusy = true;
			let i = 0;
			const finish = (): void => {
				try {
					this.restoreSideState(target());
				} catch (e) {
					this.plugin.reportError('Could not undo or redo', e);
				} finally {
					this.sideBusy = false;
					this.resetSideBaseline();
					done();
				}
			};
			// Handsontable reports a finished step from inside its render: continue after it.
			const next = (): void => {
				const action = steps[i++];
				if (!action) {
					finish();
					return;
				}
				try {
					action[method](hot, () => queueMicrotask(next));
				} catch (e) {
					console.error('[Free-Excel] Undo step failed', e);
					finish();
				}
			};
			next();
		};
		// Handsontable deep-clones a step for its hooks: keep the state in the closures.
		return {
			actionType: 'sheet_change',
			undo: (hot, done) =>
				run(
					hot,
					inner.slice().reverse(),
					'undo',
					() => state.before,
					done,
				),
			redo: (hot, done) =>
				run(hot, inner, 'redo', () => state.after, done),
		};
	}

	private restoreSideState(state: SideState): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const m = this.model;
		m.cellFormats = { ...state.formats };
		m.images = { ...state.images };
		m.imageSizes = { ...state.imageSizes };
		m.comments = { ...state.comments };
		m.links = { ...state.links };
		m.floatingImages = JSON.parse(state.floatingImages) as FloatingImage[];
		this.refreshCellInfo();
		this.refreshSidePanel();
		m.widenRecords = JSON.parse(state.widenRecords) as NonNullable<
			SheetModel['widenRecords']
		>;
		m.charts = JSON.parse(state.charts) as ChartSpec[];
		// The editor shows the chart as the undo left it (or closes when it is gone).
		window.setTimeout(() => this.chartEditor?.refresh());
		// Every size is applied again: Handsontable moved its own sizes while undoing.
		const cols = hot.countCols();
		m.colWidths = state.colWidths.slice(0, cols);
		for (let c = 0; c < cols; c++) {
			const w = state.colWidths[c];
			if (w) this.setColumnWidth(c, w);
			else this.clearColumnWidth(c);
		}
		const rows = hot.countRows();
		m.rowHeights = state.rowHeights.slice(0, rows);
		for (let r = 0; r < rows; r++) {
			const h = state.rowHeights[r];
			if (h) this.setRowHeight(r, h);
			else this.clearRowHeight(r);
		}
		this.invalidateRowSizes();
		hot.render();
		this.markDirty();
		this.refreshToolbarState();
	}

	/**
	 * Ctrl+F opens the find bar. Ctrl+Z and Ctrl+Y keep working after the grid lost the
	 * keyboard (to a color picker or a menu); while the grid has it, Handsontable handles them.
	 * Keys are matched by their position, so this works in every keyboard layout.
	 */
	private onViewKeyDown(e: KeyboardEvent): void {
		const mod = Platform.isMacOS ? e.metaKey : e.ctrlKey;
		if (!mod || e.altKey) return;
		if (this.app.workspace.getActiveViewOfType(SheetView) !== this) return;
		const target = e.target instanceof HTMLElement ? e.target : null;
		if (target?.closest('.modal-container, .menu, .suggestion-container'))
			return;

		if (e.code === 'KeyF' && !e.shiftKey) {
			e.preventDefault();
			this.openFind();
			return;
		}
		// Ctrl+G instead of the graph view (the view's scope may have handled it already).
		if (
			e.code === 'KeyG' &&
			!e.shiftKey &&
			this.plugin.settings.ctrlGFind &&
			!e.defaultPrevented
		) {
			e.preventDefault();
			e.stopImmediatePropagation();
			this.openFind();
			return;
		}
		const undo = e.code === 'KeyZ' && !e.shiftKey;
		const redo = e.code === 'KeyY' || (e.code === 'KeyZ' && e.shiftKey);
		if (!undo && !redo) return;
		const hot = this.hot;
		if (!hot || hot.isDestroyed || hot.isListening()) return;
		if (target && isTextEntry(target)) return;
		e.preventDefault();
		this.undoRedo(undo ? 'undo' : 'redo');
	}

	/* ---------------- Sort, filter, find ---------------- */

	/** The text a cell shows: the result of a formula, not its source. */
	private cellText(hot: Handsontable.Core, row: number, col: number): string {
		const display = this.resolveRaw(
			hot.getSourceDataAtCell(row, col),
			0,
			false,
			{ row, col },
		);
		if (display === null) return '';
		if (typeof display === 'boolean') return display ? 'TRUE' : 'FALSE';
		return String(display);
	}

	/** The last row that holds a value or an image, or -1 for an empty sheet. */
	private lastDataRow(hot: Handsontable.Core): number {
		let last = -1;
		for (const k of Object.keys(this.model.images))
			last = Math.max(last, Number(k.split(',')[0] ?? -1));
		for (let r = hot.countRows() - 1; r > last; r--) {
			const row = hot.getSourceDataAtRow(r) as unknown[];
			if (row.some((v) => v !== null && v !== undefined && v !== ''))
				return r;
		}
		return last;
	}

	private activeColumnName(): string {
		const cell = this.formulaCell;
		return cell
			? `column ${indexToColumnLabel(cell.col)}`
			: 'the selected column';
	}

	private headerRowEntry(): MenuEntry {
		return {
			title: 'First row is a header',
			icon: 'heading',
			checked: this.hasHeaderRow,
			onClick: () => {
				this.hasHeaderRow = !this.hasHeaderRow;
				if (this.filters.size > 0) this.applyFilters();
			},
		};
	}

	/**
	 * Sorts whole rows by the column of the active cell: the selected rows when several are
	 * selected, otherwise every row down to the last one with data. Rows hidden by a filter
	 * stay where they are.
	 */
	private sortByColumn(descending: boolean): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const cell = this.formulaCell;
		if (!cell) {
			new Notice('Select a cell in the column to sort by first.');
			return;
		}
		const selected = this.getSelectedRanges()[0];
		const span = selected && selected.r2 > selected.r1 ? selected : null;
		const first = Math.max(span?.r1 ?? 0, this.hasHeaderRow ? 1 : 0);
		const last = Math.min(
			span?.r2 ?? Number.MAX_SAFE_INTEGER,
			this.lastDataRow(hot),
		);

		const hidden = hot.getPlugin('hiddenRows');
		const rows: number[] = [];
		for (let r = first; r <= last; r++)
			if (!hidden.isHidden(r)) rows.push(r);
		if (rows.length < 2) {
			new Notice('There is nothing to sort.');
			return;
		}
		if (
			this.liveMerges(hot).some(
				(m) => m.row <= last && m.row + m.rowspan - 1 >= first,
			)
		) {
			new Notice('Unmerge the cells in the rows to sort first.');
			return;
		}

		const keys = rows.map((r) =>
			this.resolveRaw(hot.getSourceDataAtCell(r, cell.col), 0, true, {
				row: r,
				col: cell.col,
			}),
		);
		const order = rows
			.map((_r, i) => i)
			.sort((a, b) =>
				compareForSort(keys[a] ?? null, keys[b] ?? null, descending),
			);
		if (!order.every((from, i) => from === i)) {
			this.writeRows(
				this.reorderRows(this.snapshotRows(hot, rows), order),
			);
		}
		hot.listen();
	}

	private snapshotRows(hot: Handsontable.Core, rows: number[]): RowsSnapshot {
		const wanted = new Set(rows);
		const pick = <T>(source: Record<string, T>): Record<string, T> => {
			const out: Record<string, T> = {};
			for (const [k, value] of Object.entries(source))
				if (wanted.has(Number(k.split(',')[0]))) out[k] = value;
			return out;
		};
		return {
			rows,
			data: rows.map((r) =>
				(hot.getSourceDataAtRow(r) as Scalar[]).slice(),
			),
			formats: pick(this.model.cellFormats),
			images: pick(this.model.images),
			imageSizes: pick(this.model.imageSizes),
			comments: pick(this.model.comments),
			links: pick(this.model.links),
			heights: rows.map((r) => this.model.rowHeights[r] ?? null),
		};
	}

	/** The snapshot's rows in a new order: position `i` receives the row that was at `order[i]`. */
	private reorderRows(snap: RowsSnapshot, order: number[]): RowsSnapshot {
		const target = new Map<number, number>(); // row before -> row after
		order.forEach((from, i) => {
			const source = snap.rows[from];
			const dest = snap.rows[i];
			if (source !== undefined && dest !== undefined)
				target.set(source, dest);
		});
		const move = <T>(source: Record<string, T>): Record<string, T> => {
			const out: Record<string, T> = {};
			for (const [k, value] of Object.entries(source)) {
				const parts = k.split(',').map(Number);
				const r = parts[0] ?? 0;
				out[key(target.get(r) ?? r, parts[1] ?? 0)] = value;
			}
			return out;
		};
		const formulas = this.plugin.settings.enableFormulas;
		return {
			rows: snap.rows,
			data: order.map((from, i) => {
				const delta = (snap.rows[i] ?? 0) - (snap.rows[from] ?? 0);
				return (snap.data[from] ?? []).map((v) =>
					formulas && typeof v === 'string'
						? shiftFormulaRows(v, delta)
						: v,
				);
			}),
			formats: move(snap.formats),
			images: move(snap.images),
			imageSizes: move(snap.imageSizes),
			comments: move(snap.comments),
			links: move(snap.links),
			heights: order.map((from) => snap.heights[from] ?? null),
		};
	}

	/** Writes a snapshot into the sheet: values, formats, images and heights of its rows. */
	private writeRows(snap: RowsSnapshot): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const wanted = new Set(snap.rows);
		const replace = <T>(
			live: Record<string, T>,
			saved: Record<string, T>,
		): void => {
			for (const k of Object.keys(live))
				if (wanted.has(Number(k.split(',')[0]))) delete live[k];
			Object.assign(live, saved);
		};
		replace(this.model.cellFormats, snap.formats);
		replace(this.model.images, snap.images);
		replace(this.model.imageSizes, snap.imageSizes);
		replace(this.model.comments, snap.comments);
		replace(this.model.links, snap.links);

		const changes: [number, number, Scalar][] = [];
		snap.rows.forEach((r, i) => {
			const h = snap.heights[i] ?? null;
			if ((this.model.rowHeights[r] ?? null) !== h) {
				if (h) this.setRowHeight(r, h);
				else this.clearRowHeight(r);
			}
			(snap.data[i] ?? []).forEach((value, c) => {
				const current: unknown = hot.getSourceDataAtCell(r, c);
				if ((current ?? null) !== value) changes.push([r, c, value]);
			});
		});
		// Handsontable records the values; the formats and sizes join that undo step (markDirty).
		if (changes.length > 0) hot.setDataAtCell(changes, 'sort');
		this.invalidateRowSizes();
		hot.render();
		this.markDirty();
		this.refreshToolbarState();
	}

	private openColumnFilter(): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const cell = this.formulaCell;
		if (!cell) {
			new Notice('Select a cell in the column to filter first.');
			return;
		}
		const col = cell.col;
		const first = this.hasHeaderRow ? 1 : 0;
		const last = this.lastDataRow(hot);
		if (last < first) {
			new Notice('There is no data to filter.');
			return;
		}
		const values = new Set<string>();
		for (let r = first; r <= last; r++)
			values.add(this.cellText(hot, r, col));
		const header = this.hasHeaderRow ? this.cellText(hot, 0, col) : '';
		new ColumnFilterModal(
			this.app,
			`Filter column ${indexToColumnLabel(col)}${header ? ` (${header})` : ''}`,
			[...values].sort(compareText),
			this.filters.get(col) ?? null,
			(filter) => this.setColumnFilter(col, filter),
		).open();
	}

	private setColumnFilter(col: number, filter: ColumnFilter | null): void {
		if (filter) this.filters.set(col, filter);
		else this.filters.delete(col);
		this.applyFilters();
		if (this.hot && !this.hot.isDestroyed) this.hot.listen();
	}

	/** Hides the rows that do not pass every column filter. Rows below the data stay visible. */
	private applyFilters(): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const hide: number[] = [];
		if (this.filters.size > 0) {
			const last = this.lastDataRow(hot);
			for (let r = this.hasHeaderRow ? 1 : 0; r <= last; r++) {
				for (const [col, filter] of this.filters) {
					if (filterAccepts(filter, this.cellText(hot, r, col)))
						continue;
					hide.push(r);
					break;
				}
			}
		}
		const plugin = hot.getPlugin('hiddenRows');
		plugin.showRows(plugin.getHiddenRows());
		plugin.hideRows(hide);
		this.findStale = true;
		hot.render();
		this.ui.filterBtn?.toggleClass('is-active', this.filters.size > 0);
	}

	/** Keeps each filter on its column when columns are inserted or removed. */
	private remapFilters(
		index: number,
		amount: number,
		mode: 'insert' | 'remove',
	): void {
		if (this.filters.size === 0) return;
		const next = new Map<number, ColumnFilter>();
		let removed = false;
		for (const [col, filter] of this.filters) {
			if (mode === 'insert')
				next.set(col >= index ? col + amount : col, filter);
			else if (col < index) next.set(col, filter);
			else if (col >= index + amount) next.set(col - amount, filter);
			else removed = true;
		}
		this.filters = next;
		// The filtered column is gone: show its rows again, once Handsontable is done altering.
		if (removed) window.setTimeout(() => this.applyFilters());
	}

	private buildFindBar(): void {
		const bar = this.findBarEl;
		bar.empty();
		bar.hide();
		const input = bar.createEl('input', {
			type: 'text',
			cls: 'rtl-sheet-find-input',
			attr: {
				placeholder: 'Find in sheet',
				dir: 'auto',
				spellcheck: 'false',
				autocomplete: 'off',
			},
		});
		this.ui.findInput = input;
		this.ui.findCount = bar.createSpan({ cls: 'rtl-sheet-find-count' });
		this.addButton(bar, 'chevron-up', 'Previous', 'Previous match', () =>
			this.stepFind(-1),
		);
		this.addButton(bar, 'chevron-down', 'Next', 'Next match', () =>
			this.stepFind(1),
		);
		this.addButton(bar, 'x', 'Close', 'Close', () => this.closeFind());

		// Like the formula bar: typing here must not reach the grid.
		input.addEventListener('focus', () => this.hot?.unlisten());
		input.addEventListener('input', () => {
			this.runFind();
			this.showMatch(0);
		});
		input.addEventListener('keydown', (e) => {
			if (e.key === 'Enter') {
				e.preventDefault();
				this.stepFind(e.shiftKey ? -1 : 1);
			} else if (e.key === 'Escape') {
				e.preventDefault();
				e.stopPropagation();
				this.closeFind();
			}
		});
	}

	openFind(): void {
		const input = this.ui.findInput;
		if (!input) return;
		const wasShown = this.findBarEl.isShown();
		this.findBarEl.show();
		if (!wasShown) this.onResize();
		input.focus();
		input.select();
		if (!wasShown && input.value) {
			this.runFind();
			this.showMatch(0);
		}
	}

	/** Closes the find bar and leaves the selection on the match that was shown. */
	private closeFind(): void {
		const match = this.findMatches[this.findIndex];
		this.findBarEl.hide();
		this.findMatches = [];
		this.findHits.clear();
		this.findIndex = -1;
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		hot.render();
		this.onResize();
		if (match) hot.selectCell(match.row, match.col);
		hot.listen();
	}

	/** Collects the cells whose shown text contains the search text (case-insensitive). */
	private runFind(): void {
		const hot = this.hot;
		const query = (this.ui.findInput?.value ?? '').toLowerCase();
		this.findMatches = [];
		this.findHits.clear();
		this.findIndex = -1;
		this.findStale = false;
		if (!hot || hot.isDestroyed) return;
		if (query) {
			const hidden = hot.getPlugin('hiddenRows');
			const rows = hot.countRows();
			const cols = hot.countCols();
			for (let r = 0; r < rows; r++) {
				if (hidden.isHidden(r)) continue;
				for (let c = 0; c < cols; c++) {
					const raw: unknown = hot.getSourceDataAtCell(r, c);
					if (raw === null || raw === undefined || raw === '')
						continue;
					if (!this.cellText(hot, r, c).toLowerCase().includes(query))
						continue;
					this.findMatches.push({ row: r, col: c });
					this.findHits.add(key(r, c));
				}
			}
		}
		hot.render();
		this.updateFindCount();
	}

	private stepFind(delta: number): void {
		const current = this.findMatches[this.findIndex];
		if (this.findStale) {
			this.runFind();
			// Continue from the match that was shown, if it still matches.
			if (current)
				this.findIndex = this.findMatches.findIndex(
					(m) => m.row === current.row && m.col === current.col,
				);
		}
		const count = this.findMatches.length;
		if (count === 0) return;
		const next =
			this.findIndex < 0
				? delta > 0
					? 0
					: count - 1
				: (this.findIndex + delta + count) % count;
		this.showMatch(next);
	}

	/**
	 * Marks a match as the current one and scrolls to it. The cell is not selected while the
	 * find bar is in use: selecting would move the keyboard focus into the grid.
	 */
	private showMatch(index: number): void {
		const hot = this.hot;
		const match = this.findMatches[index];
		this.findIndex = match ? index : -1;
		if (hot && !hot.isDestroyed && match) {
			hot.scrollViewportTo({ row: match.row, col: match.col });
			hot.render();
		}
		this.updateFindCount();
	}

	private updateFindCount(): void {
		const count = this.findMatches.length;
		const hasQuery = !!this.ui.findInput?.value;
		this.ui.findCount?.setText(
			!hasQuery
				? ''
				: count === 0
					? 'No results'
					: `${this.findIndex + 1} / ${count}`,
		);
	}

	private toggleFlag(flag: 'bold' | 'italic' | 'underline' | 'strike'): void {
		const ranges = this.requireSelection();
		const firstRange = ranges[0];
		if (!firstRange) return;
		const turnOn =
			!this.model.cellFormats[key(firstRange.r1, firstRange.c1)]?.[flag];
		this.applyFormat(ranges, (f) => {
			if (turnOn) f[flag] = true;
			else delete f[flag];
		});
	}

	/** Sets a format property; pressing the same option again clears it. */
	private setFormatProp<K extends 'align' | 'valign' | 'wrap'>(
		prop: K,
		value: NonNullable<CellFormat[K]>,
	): void {
		const ranges = this.requireSelection();
		const firstRange = ranges[0];
		if (!firstRange) return;
		const same =
			this.model.cellFormats[key(firstRange.r1, firstRange.c1)]?.[
				prop
			] === value;
		this.applyFormat(ranges, (f) => {
			if (same) delete f[prop];
			else (f as Record<string, unknown>)[prop] = value;
		});
	}

	/**
	 * Left/right are visual (physical) sides. In an RTL sheet the column with the
	 * smaller index is the visually right-most one, so the two are swapped there.
	 * Single-edge modes add to the existing borders; "outer" and "all" replace them.
	 */
	private applyBorders(
		mode: BorderMode,
		ranges: SelRange[] = this.requireSelection(),
	): void {
		const rtl = this.model.rtl;
		this.applyFormat(ranges, (f, r, c, rg) => {
			const prev = f.borders ?? {};
			const withColor = (b: CellBorders): CellBorders =>
				prev.color ? { ...b, color: prev.color } : b;
			const visualLeft = rtl ? rg.c2 : rg.c1;
			const visualRight = rtl ? rg.c1 : rg.c2;
			switch (mode) {
				case 'none':
					delete f.borders;
					break;
				case 'all':
					f.borders = withColor({
						top: true,
						right: true,
						bottom: true,
						left: true,
					});
					break;
				case 'outer':
					f.borders = withColor({
						top: r === rg.r1,
						bottom: r === rg.r2,
						left: c === visualLeft,
						right: c === visualRight,
					});
					break;
				case 'top':
					if (r === rg.r1)
						f.borders = withColor({ ...prev, top: true });
					break;
				case 'bottom':
					if (r === rg.r2)
						f.borders = withColor({ ...prev, bottom: true });
					break;
				case 'left':
					if (c === visualLeft)
						f.borders = withColor({ ...prev, left: true });
					break;
				case 'right':
					if (c === visualRight)
						f.borders = withColor({ ...prev, right: true });
					break;
			}
		});
	}

	private toggleFormatPainter(): void {
		if (this.painterSource) {
			this.setPainter(null);
			return;
		}
		const cell = this.formulaCell;
		if (!cell) {
			new Notice('Select the cell to copy the formatting from first.');
			return;
		}
		this.setPainter(
			cloneFormat(this.model.cellFormats[key(cell.row, cell.col)] ?? {}),
		);
		new Notice(
			IS_TOUCH
				? 'Format painter on: tap the cells to format. Tap the brush again to turn it off.'
				: 'Format painter on: select the cells to format.',
		);
	}

	private setPainter(format: CellFormat | null): void {
		this.painterSource = format;
		this.ui.brush?.toggleClass('is-active', format !== null);
	}

	private mergeSelection(): void {
		const hot = this.hot;
		if (!hot) return;
		const ranges = this.requireSelection();
		if (ranges.length === 0) return;
		const plugin = hot.getPlugin('mergeCells');
		let merged = false;
		for (const rg of ranges) {
			if (rg.r1 === rg.r2 && rg.c1 === rg.c2) continue;
			this.keepMergeContent(hot, rg);
			plugin.merge(rg.r1, rg.c1, rg.r2, rg.c2);
			merged = true;
		}
		if (!merged) {
			new Notice('Select at least two cells to merge.');
			return;
		}
		this.invalidateRowSizes();
		hot.render();
		this.markDirty();
	}

	/**
	 * Handsontable keeps only the top-left value of a merge. Like Excel, keep the
	 * first value in the range instead (moving it to the top-left cell) and
	 * center the merged cell unless it already has an alignment.
	 */
	private keepMergeContent(hot: Handsontable.Core, rg: SelRange): void {
		const isBlank = (v: unknown): boolean =>
			v === null || v === undefined || v === '';
		const anchor = key(rg.r1, rg.c1);
		if (isBlank(hot.getSourceDataAtCell(rg.r1, rg.c1))) {
			search: for (let r = rg.r1; r <= rg.r2; r++) {
				for (let c = rg.c1; c <= rg.c2; c++) {
					const value: unknown = hot.getSourceDataAtCell(r, c);
					if (isBlank(value)) continue;
					hot.setDataAtCell(rg.r1, rg.c1, value);
					const f = this.model.cellFormats[key(r, c)];
					if (f && !this.model.cellFormats[anchor])
						this.model.cellFormats[anchor] = cloneFormat(f);
					break search;
				}
			}
		}
		const f: CellFormat = { ...(this.model.cellFormats[anchor] ?? {}) };
		f.align ??= 'center';
		f.valign ??= 'middle';
		this.model.cellFormats[anchor] = f;
	}

	private unmergeSelection(): void {
		const hot = this.hot;
		if (!hot) return;
		const ranges = this.requireSelection();
		if (ranges.length === 0) return;
		const plugin = hot.getPlugin('mergeCells');
		for (const rg of ranges) plugin.unmerge(rg.r1, rg.c1, rg.r2, rg.c2);
		this.invalidateRowSizes();
		hot.render();
	}

	private undoRedo(action: 'undo' | 'redo'): void {
		const plugin = this.undoPlugin();
		if (!plugin) return;
		if (action === 'undo') plugin.undo();
		else plugin.redo();
	}

	private alterStructure(
		action:
			| 'insert_row_below'
			| 'remove_row'
			| 'insert_col_end'
			| 'remove_col',
	): void {
		const hot = this.hot;
		if (!hot) return;
		const ranges = this.requireSelection();
		const rg = ranges[0];
		if (!rg) return;
		switch (action) {
			case 'insert_row_below':
				hot.alter('insert_row_below', rg.r2, 1);
				break;
			case 'insert_col_end':
				hot.alter('insert_col_end', rg.c2, 1);
				break;
			case 'remove_row':
				hot.alter('remove_row', rg.r1, rg.r2 - rg.r1 + 1);
				break;
			case 'remove_col':
				hot.alter('remove_col', rg.c1, rg.c2 - rg.c1 + 1);
				break;
		}
	}

	private toggleDirection(): void {
		this.setDirection(!this.model.rtl);
	}

	/** Changes the sheet direction (which rebuilds the grid) as an undoable step. */
	private setDirection(rtl: boolean, record = true): void {
		if (rtl === this.model.rtl) return;
		this.captureGridStateIfLive();
		this.model.rtl = rtl;
		this.rebuildGrid(true);
		this.markDirty();
		const plugin = this.undoPlugin();
		if (!record || !plugin) return;
		// The grid cannot be rebuilt while Handsontable is still running the undo: do it right after.
		const later =
			(value: boolean): UndoStep['undo'] =>
			(_hot, done) => {
				done();
				window.setTimeout(() => this.setDirection(value, false));
			};
		const step: UndoStep = {
			actionType: 'sheet_direction',
			undo: later(!rtl),
			redo: later(rtl),
		};
		this.knownSteps.add(step);
		plugin.done(() => step);
	}

	/* ---------------- Keyboard ---------------- */

	/** Registers Ctrl+G (Cmd+G) in the view's scope while the setting is on. */
	private applyShortcutSettings(): void {
		const scope = this.scope;
		if (!scope) return;
		if (this.findHotkey) {
			scope.unregister(this.findHotkey);
			this.findHotkey = null;
		}
		if (this.plugin.settings.ctrlGFind)
			this.findHotkey = scope.register(['Mod'], 'g', () => {
				this.openFind();
				return false; // handled: Obsidian does not open the graph view
			});
	}

	/* ---------------- Sheets (tabs) ---------------- */

	/** The tab bar under the grid: one tab per sheet, and a button that adds a sheet. */
	private renderTabs(): void {
		const bar = this.tabBarEl;
		bar.empty();
		bar.setAttr('dir', this.model.rtl ? 'rtl' : 'ltr');
		const add = bar.createEl('button', {
			cls: 'rtl-sheet-btn rtl-sheet-tab-add',
			attr: { 'aria-label': 'Add sheet', title: 'Add sheet' },
		});
		setIcon(add, 'plus');
		add.addEventListener('click', () => this.addSheet());

		const list = bar.createDiv({
			cls: 'rtl-sheet-tab-list',
			attr: { role: 'tablist' },
		});
		this.sheets.forEach((sheet, i) => {
			const active = i === this.activeSheet;
			const tab = list.createDiv({
				cls: `rtl-sheet-tab${active ? ' is-active' : ''}`,
				attr: {
					role: 'tab',
					'aria-selected': String(active),
					title: sheet.name,
				},
			});
			tab.createSpan({ cls: 'rtl-sheet-tab-name', text: sheet.name });
			tab.addEventListener('click', () => this.showSheet(i));
			tab.addEventListener('dblclick', () => this.renameSheet(i));
			tab.addEventListener('contextmenu', (e) => {
				e.preventDefault();
				this.showSheetMenu(i, tab);
			});
			if (active) {
				// Touch screens have no reliable right-click: the active tab has a menu button.
				const more = tab.createSpan({
					cls: 'rtl-sheet-tab-menu',
					attr: { 'aria-label': 'Sheet options' },
				});
				setIcon(more, 'chevron-down');
				more.addEventListener('click', (e) => {
					e.stopPropagation();
					this.showSheetMenu(i, tab);
				});
				window.requestAnimationFrame(() =>
					tab.scrollIntoView({ block: 'nearest', inline: 'nearest' }),
				);
			}
		});
	}

	private showSheetMenu(index: number, anchor: HTMLElement): void {
		const entries: MenuEntry[] = [
			{
				title: 'Rename…',
				icon: 'pencil',
				onClick: () => this.renameSheet(index),
			},
			{
				title: 'Duplicate',
				icon: 'copy',
				onClick: () => this.duplicateSheet(index),
			},
		];
		if (index > 0)
			entries.push({
				title: 'Move toward the start',
				icon: this.model.rtl ? 'arrow-right' : 'arrow-left',
				onClick: () => this.moveSheet(index, -1),
			});
		if (index < this.sheets.length - 1)
			entries.push({
				title: 'Move toward the end',
				icon: this.model.rtl ? 'arrow-left' : 'arrow-right',
				onClick: () => this.moveSheet(index, 1),
			});
		if (this.sheets.length > 1)
			entries.push({
				title: 'Delete',
				icon: 'trash-2',
				onClick: () => this.deleteSheet(index),
			});
		this.showMenu(anchor, entries);
	}

	/** Closes the cell editor and the formula bar, keeping what was typed. */
	private finishEditing(): void {
		const input = this.ui.formulaInput;
		if (input && input.ownerDocument.activeElement === input) input.blur();
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const editor = hot.getActiveEditor() as unknown as
			| { isOpened?: () => boolean; finishEditing?: () => void }
			| undefined;
		if (editor?.isOpened?.()) editor.finishEditing?.();
	}

	/**
	 * Shows another sheet. The grid is built again for it, so the undo history and the
	 * filters start over (they belong to the sheet that was shown).
	 */
	private showSheet(index: number, force = false): void {
		const next = this.sheets[index];
		if (!next || (index === this.activeSheet && !force)) return;
		this.finishEditing();
		this.captureGridStateIfLive();
		this.destroyGrid();
		this.filters.clear();
		this.formulaRefs = [];
		this.activeSheet = index;
		this.model = next;
		this.rebuildGrid();
	}

	private sheetNameError(name: string, except: number): string | null {
		const clean = cleanSheetName(name);
		if (!clean) return 'Enter a name.';
		const taken = this.sheets.some(
			(s, i) =>
				i !== except && s.name.toLowerCase() === clean.toLowerCase(),
		);
		return taken ? 'Another sheet already has this name.' : null;
	}

	private addSheet(): void {
		this.finishEditing();
		this.captureGridStateIfLive();
		const name = uniqueSheetName(
			this.sheets.map((s) => s.name),
			`Sheet${this.sheets.length + 1}`,
		);
		const index = this.activeSheet + 1;
		this.sheets.splice(
			index,
			0,
			createDefaultModel(this.plugin.settings, this.model.rtl, name),
		);
		this.showSheet(index, true);
		this.markDirty();
	}

	private duplicateSheet(index: number): void {
		const source = this.sheets[index];
		if (!source) return;
		this.finishEditing();
		this.captureGridStateIfLive();
		const copy = JSON.parse(JSON.stringify(source)) as SheetModel;
		copy.name = uniqueSheetName(
			this.sheets.map((s) => s.name),
			cleanSheetName(`${source.name} copy`),
		);
		copy.charts = copy.charts.map((c) => ({ ...c, id: newChartId() }));
		copy.floatingImages = copy.floatingImages.map((f) => ({
			...f,
			id: newChartId('img'),
		}));
		this.sheets.splice(index + 1, 0, copy);
		this.showSheet(index + 1, true);
		this.markDirty();
	}

	private moveSheet(index: number, delta: number): void {
		const target = index + delta;
		const sheet = this.sheets[index];
		const other = this.sheets[target];
		if (!sheet || !other) return;
		this.sheets[index] = other;
		this.sheets[target] = sheet;
		this.activeSheet = this.sheets.indexOf(this.model);
		this.renderTabs();
		this.markDirty();
	}

	private deleteSheet(index: number): void {
		const sheet = this.sheets[index];
		if (!sheet || this.sheets.length < 2) return;
		new ConfirmModal(
			this.app,
			'Delete sheet',
			`Delete "${sheet.name}" and everything on it? This cannot be undone.`,
			'Delete',
			() => {
				const at = this.sheets.indexOf(sheet);
				if (at < 0 || this.sheets.length < 2) return;
				if (at === this.activeSheet) {
					this.finishEditing();
					this.destroyGrid();
					this.sheets.splice(at, 1);
					const next = Math.min(at, this.sheets.length - 1);
					this.activeSheet = -1; // so showSheet does not skip it
					this.showSheet(next, true);
				} else {
					this.sheets.splice(at, 1);
					this.activeSheet = this.sheets.indexOf(this.model);
					this.renderTabs();
				}
				this.invalidateRowSizes(); // formulas that used the sheet now show #REF!
				this.hot?.render();
				this.markDirty();
			},
		).open();
	}

	private renameSheet(index: number): void {
		const sheet = this.sheets[index];
		if (!sheet) return;
		new TextPromptModal(
			this.app,
			'Rename sheet',
			sheet.name,
			(value) => this.sheetNameError(value, index),
			(value) => {
				const name = cleanSheetName(value);
				const old = sheet.name;
				if (name === old) return;
				sheet.name = name;
				this.renameSheetReferences(old, name);
				this.renderTabs();
				this.markDirty();
			},
		).open();
	}

	/** Formulas that name the renamed sheet ("Old!A1", "'Old name'!A1") follow the new name. */
	private renameSheetReferences(oldName: string, newName: string): void {
		const escaped = oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
		const pattern = new RegExp(
			`(^|[^A-Za-z0-9_.'])(?:'${escaped}'|${escaped})!`,
			'gi',
		);
		const replacement = `$1${sheetRefName(newName)}!`;
		const rewrite = (v: Scalar): Scalar =>
			typeof v === 'string' && v.startsWith('=')
				? v.replace(pattern, replacement)
				: v;

		this.captureGridStateIfLive();
		const changes: [number, number, Scalar][] = [];
		this.sheets.forEach((sheet, i) => {
			sheet.data.forEach((row, r) =>
				row.forEach((v, c) => {
					const next = rewrite(v);
					if (next === v) return;
					// The shown sheet changes through the grid, so the change can be undone.
					if (i === this.activeSheet && this.hot) changes.push([r, c, next]);
					else row[c] = next;
				}),
			);
		});
		const hot = this.hot;
		if (changes.length > 0 && hot && !hot.isDestroyed)
			hot.setDataAtCell(changes, 'rename');
		this.invalidateRowSizes();
		hot?.render();
	}

	/* ---------------- Charts ---------------- */

	/** Draws the charts, floating images and the copy outline again once the current work is done. */
	queueCharts(): void {
		if (this.chartFrame) return;
		this.chartFrame = window.requestAnimationFrame(() => {
			this.chartFrame = 0;
			try {
				this.renderCharts();
			} catch (e) {
				console.error('[Free-Excel] Could not draw the charts', e);
			}
		});
	}

	/** Where cell A1 starts inside the scrolled part of the grid (after the headers), in px. */
	private chartOrigin(hot: Handsontable.Core): { x: number; y: number } {
		const view = (
			hot as unknown as {
				view?: {
					getRowHeaderWidth?: () => number;
					getColumnHeaderHeight?: () => number;
				};
			}
		).view;
		return {
			x: view?.getRowHeaderWidth?.() ?? 0,
			y: view?.getColumnHeaderHeight?.() ?? 0,
		};
	}

	private renderCharts(): void {
		const hot = this.hot;
		const layer = this.chartLayer;
		if (!hot || hot.isDestroyed || !layer || !layer.isConnected) return;
		layer.empty();
		const origin = this.chartOrigin(hot);
		this.renderCopyMark(hot, layer, origin);
		const read = this.chartReader();
		// From the bottom up: later boxes cover earlier ones.
		for (const entry of this.layerOrder()) {
			if (entry.kind === 'image')
				this.renderFloatingImage(layer, entry.item, origin);
			else this.renderChartBox(layer, entry.item, origin, read);
		}
	}

	private renderChartBox(
		layer: HTMLElement,
		chart: ChartSpec,
		origin: { x: number; y: number },
		read: (row: number, col: number) => Scalar,
	): void {
		const box = layer.createDiv({ cls: 'rtl-sheet-chart' });
		box.toggleClass('is-editing', chart.id === this.chartEditorId);
		this.placeChart(box, chart, origin);
		// Presses on a chart must not select cells or start editing them.
		for (const type of [
			'mousedown',
			'pointerdown',
			'touchstart',
			'dblclick',
			'contextmenu',
		])
			box.addEventListener(type, (e) => e.stopPropagation(), {
				passive: true,
			});

		const head = box.createDiv({ cls: 'rtl-sheet-chart-head' });
		head.style.height = `${CHART_HEAD_H * this.zoomY}px`;
		head.style.fontSize = `${Math.round(13 * this.zoomY * 10) / 10}px`;
		const title = head.createSpan({
			cls: 'rtl-sheet-chart-title',
			text: chart.title || 'Chart',
			attr: { dir: 'auto' },
		});
		if (chart.titleColor) title.style.color = chart.titleColor;
		const editBtn = head.createEl('button', {
			cls: 'rtl-sheet-chart-btn',
			attr: { 'aria-label': 'Edit chart', title: 'Edit chart' },
		});
		setIcon(editBtn, 'pencil');
		editBtn.addEventListener('click', () => this.editChart(chart.id));
		const layersBtn = head.createEl('button', {
			cls: 'rtl-sheet-chart-btn',
			attr: { 'aria-label': 'Arrange', title: 'Arrange' },
		});
		setIcon(layersBtn, 'layers');
		layersBtn.addEventListener('click', () =>
			this.showMenu(layersBtn, this.layerMenuEntries(chart.id)),
		);
		box.addEventListener('contextmenu', (e) => {
			e.preventDefault();
			this.showMenu(e, [
				{
					title: 'Edit chart',
					icon: 'pencil',
					onClick: () => this.editChart(chart.id),
				},
				...this.layerMenuEntries(chart.id),
				{
					title: 'Delete chart',
					icon: 'trash-2',
					onClick: () => this.deleteChart(chart.id),
				},
			]);
		});
		const deleteBtn = head.createEl('button', {
			cls: 'rtl-sheet-chart-btn',
			attr: { 'aria-label': 'Delete chart', title: 'Delete chart' },
		});
		setIcon(deleteBtn, 'trash-2');
		deleteBtn.addEventListener('click', () => this.deleteChart(chart.id));
		head.addEventListener('pointerdown', (e) =>
			this.onChartDrag(e, box, chart.id, 'move'),
		);

		const body = box.createDiv({ cls: 'rtl-sheet-chart-body' });
		body.addEventListener('dblclick', () => this.editChart(chart.id));
		renderChart(
			body,
			chart,
			readChartData(chart, read),
			chart.w,
			chart.h - CHART_HEAD_H,
		);
		const grip = box.createDiv({
			cls: 'rtl-sheet-chart-grip',
			attr: { 'aria-label': 'Resize chart' },
		});
		grip.addEventListener('pointerdown', (e) =>
			this.onChartDrag(e, box, chart.id, 'resize'),
		);
	}

	/** Charts and floating images from the bottom up: by `z`, images first when it is the same. */
	private layerOrder(): (
		| { kind: 'image'; item: FloatingImage }
		| { kind: 'chart'; item: ChartSpec }
	)[] {
		const list = [
			...this.model.floatingImages.map((item) => ({
				kind: 'image' as const,
				item,
			})),
			...this.model.charts.map((item) => ({
				kind: 'chart' as const,
				item,
			})),
		];
		return list
			.map((entry, i) => ({ entry, i }))
			.sort(
				(a, b) =>
					(a.entry.item.z ?? 0) - (b.entry.item.z ?? 0) || a.i - b.i,
			)
			.map((e) => e.entry);
	}

	/** The `z` that puts a new chart or floating image on top of the others. */
	private topZ(): number {
		let top = 0;
		for (const { item } of this.layerOrder())
			top = Math.max(top, (item.z ?? 0) + 1);
		return top;
	}

	/** Moves a chart or floating image up or down the stack (one undo step). */
	private restack(
		id: string,
		move: 'front' | 'forward' | 'backward' | 'back',
	): void {
		const order = this.layerOrder();
		const i = order.findIndex((e) => e.item.id === id);
		const entry = order[i];
		if (!entry) return;
		order.splice(i, 1);
		const j =
			move === 'front'
				? order.length
				: move === 'back'
					? 0
					: move === 'forward'
						? Math.min(order.length, i + 1)
						: Math.max(0, i - 1);
		if (j === i) return; // already there
		order.splice(j, 0, entry);
		const z = new Map(order.map((e, k) => [e.item.id, k]));
		// Replaced, never changed in place (undo compares them).
		this.model.charts = this.model.charts.map((c) =>
			c.z === z.get(c.id) ? c : { ...c, z: z.get(c.id) },
		);
		this.model.floatingImages = this.model.floatingImages.map((f) =>
			f.z === z.get(f.id) ? f : { ...f, z: z.get(f.id) },
		);
		this.markDirty();
		this.queueCharts();
	}

	private layerMenuEntries(id: string): MenuEntry[] {
		return [
			{
				title: 'Bring to front',
				icon: 'bring-to-front',
				onClick: () => this.restack(id, 'front'),
			},
			{
				title: 'Bring forward',
				icon: 'arrow-up',
				onClick: () => this.restack(id, 'forward'),
			},
			{
				title: 'Send backward',
				icon: 'arrow-down',
				onClick: () => this.restack(id, 'backward'),
			},
			{
				title: 'Send to back',
				icon: 'send-to-back',
				onClick: () => this.restack(id, 'back'),
			},
		];
	}

	private placeChart(
		box: HTMLElement,
		chart: Pick<ChartSpec, 'x' | 'y' | 'w' | 'h'>,
		origin: { x: number; y: number },
	): void {
		// The position follows the columns and rows; the size follows the text, so a chart keeps
		// its shape when only the horizontal axis zooms.
		const z = this.zoom;
		const zy = this.zoomY;
		const st = box.style;
		st.top = `${origin.y + chart.y * zy}px`;
		// The first column is on the right in a right-to-left sheet.
		st.left = this.model.rtl ? '' : `${origin.x + chart.x * z}px`;
		st.right = this.model.rtl ? `${origin.x + chart.x * z}px` : '';
		st.width = `${chart.w * zy}px`;
		st.height = `${chart.h * zy}px`;
	}

	/** Dragging the title bar moves a chart, dragging its corner resizes it. One undo step each. */
	private onChartDrag(
		e: PointerEvent,
		box: HTMLElement,
		id: string,
		mode: 'move' | 'resize',
	): void {
		const chart = this.model.charts.find((c) => c.id === id);
		if (!chart) return;
		this.dragBox(
			e,
			box,
			chart,
			mode,
			{ w: CHART_MIN_W, h: CHART_MIN_H },
			false,
			(next) => this.replaceChart(id, (c) => ({ ...c, ...next })),
		);
	}

	/**
	 * Moves or resizes a box on the chart layer (a chart or a floating image) while the pointer
	 * is dragged, and reports where it ended up. `keepRatio` keeps the shape while resizing
	 * (Shift resizes freely).
	 */
	private dragBox(
		e: PointerEvent,
		box: HTMLElement,
		chart: { x: number; y: number; w: number; h: number },
		mode: 'move' | 'resize',
		min: { w: number; h: number },
		keepRatio: boolean,
		done: (next: { x: number; y: number; w: number; h: number }) => void,
	): void {
		if (e.pointerType === 'mouse' && e.button !== 0) return;
		if (e.target instanceof HTMLElement && e.target.closest('button')) return;
		const hot = this.hot;
		if (!hot) return;
		e.preventDefault();
		const handle = e.currentTarget as HTMLElement;
		handle.setPointerCapture(e.pointerId);
		const origin = this.chartOrigin(hot);
		// In a right-to-left sheet X grows to the left.
		const dir = this.model.rtl ? -1 : 1;
		const start = { x: e.clientX, y: e.clientY };
		const next = { x: chart.x, y: chart.y, w: chart.w, h: chart.h };
		let moved = false;
		box.addClass('is-dragging');

		const onMove = (ev: PointerEvent): void => {
			if (ev.pointerId !== e.pointerId) return;
			const px = (ev.clientX - start.x) * dir;
			const dy = (ev.clientY - start.y) / this.zoomY;
			if (mode === 'move') {
				next.x = Math.max(0, Math.round(chart.x + px / this.zoom));
				next.y = Math.max(0, Math.round(chart.y + dy));
			} else {
				next.w = Math.max(min.w, Math.round(chart.w + px / this.zoomY));
				next.h = Math.max(min.h, Math.round(chart.h + dy));
				if (keepRatio && !ev.shiftKey && chart.h > 0) {
					const ratio = chart.w / chart.h;
					next.h = Math.max(min.h, Math.round(next.w / ratio));
					next.w = Math.round(next.h * ratio);
				}
			}
			moved = true;
			this.placeChart(box, next, origin);
		};
		const onEnd = (ev: PointerEvent): void => {
			if (ev.pointerId !== e.pointerId) return;
			handle.removeEventListener('pointermove', onMove);
			handle.removeEventListener('pointerup', onEnd);
			handle.removeEventListener('pointercancel', onEnd);
			box.removeClass('is-dragging');
			if (!moved) return;
			done(next);
		};
		handle.addEventListener('pointermove', onMove);
		handle.addEventListener('pointerup', onEnd);
		handle.addEventListener('pointercancel', onEnd);
	}

	/** Charts are replaced, never changed in place (undo compares them). */
	private replaceChart(
		id: string,
		change: (chart: ChartSpec) => ChartSpec | null,
	): void {
		const charts: ChartSpec[] = [];
		for (const c of this.model.charts) {
			if (c.id !== id) {
				charts.push(c);
				continue;
			}
			const next = change(c);
			if (next) charts.push(next);
		}
		this.model.charts = charts;
		this.markDirty();
		this.queueCharts();
	}

	private editChart(id: string): void {
		if (this.chartEditorId === id) return;
		this.openChartEditor(id, false);
	}

	/** Reads what cells show on the shown sheet (formulas worked out), for charts. */
	private chartReader(): (row: number, col: number) => Scalar {
		return (row, col) =>
			this.resolveRaw(this.rawCell(this.activeSheet, row, col), 0, true, {
				row,
				col,
			});
	}

	/**
	 * Opens the chart editor next to the grid. It changes the chart live (every change is an
	 * undo step), and the sheet stays usable for picking ranges.
	 */
	private openChartEditor(id: string, isNew: boolean): void {
		this.closeChartEditor();
		this.chartEditorId = id;
		this.chartPanelEl.show();
		const get = (): ChartSpec | null =>
			this.model.charts.find((c) => c.id === id) ?? null;
		this.chartEditor = new ChartEditor(
			{
				el: this.chartPanelEl,
				get,
				update: (change, tag) => {
					const current = get();
					if (!current) return;
					const next = JSON.parse(JSON.stringify(current)) as ChartSpec;
					change(next);
					// Typing a title or dragging a color picker is one undo step.
					this.coalesce = { tag: `chart:${id}:${tag}`, windowMs: 2000 };
					this.replaceChart(id, () => next);
				},
				close: () => this.closeChartEditor(),
				highlight: (ranges) => {
					this.chartPickRanges = ranges;
					this.paintOutlines();
				},
				seriesNames: (spec) => {
					const data = readChartData(spec, this.chartReader());
					return spec.type === 'pie'
						? data.categories
						: data.series.map((s) => s.name);
				},
			},
			isNew,
		);
		this.onResize();
		this.queueCharts(); // marks the chart that is being edited
	}

	private closeChartEditor(): void {
		const editor = this.chartEditor;
		if (!editor) return;
		this.chartEditor = null;
		this.chartEditorId = null;
		this.queueCharts();
		editor.destroy();
		this.chartPanelEl.hide();
		this.chartPickRanges = [];
		this.paintOutlines();
		this.onResize();
		this.hot?.listen();
	}

	private deleteChart(id: string): void {
		if (this.chartEditorId === id) this.closeChartEditor();
		this.replaceChart(id, () => null);
		new Notice('Chart deleted. Undo (Ctrl+Z) brings it back.');
	}

	/**
	 * The block of filled cells around a cell (like Ctrl+A in Excel), so a chart can be made
	 * from one selected cell of a table.
	 */
	private dataRegion(hot: Handsontable.Core, row: number, col: number): SelRange {
		const filled = (r: number, c: number): boolean => {
			const v: unknown = hot.getSourceDataAtCell(r, c);
			return v !== null && v !== undefined && v !== '';
		};
		const rg: SelRange = { r1: row, r2: row, c1: col, c2: col };
		const rows = hot.countRows();
		const cols = hot.countCols();
		const anyIn = (r1: number, r2: number, c1: number, c2: number): boolean => {
			for (let r = r1; r <= r2; r++)
				for (let c = c1; c <= c2; c++) if (filled(r, c)) return true;
			return false;
		};
		for (let grew = true, guard = 0; grew && guard < 10000; guard++) {
			grew = false;
			if (rg.r1 > 0 && anyIn(rg.r1 - 1, rg.r1 - 1, rg.c1, rg.c2)) {
				rg.r1--;
				grew = true;
			}
			if (rg.r2 < rows - 1 && anyIn(rg.r2 + 1, rg.r2 + 1, rg.c1, rg.c2)) {
				rg.r2++;
				grew = true;
			}
			if (rg.c1 > 0 && anyIn(rg.r1, rg.r2, rg.c1 - 1, rg.c1 - 1)) {
				rg.c1--;
				grew = true;
			}
			if (rg.c2 < cols - 1 && anyIn(rg.r1, rg.r2, rg.c2 + 1, rg.c2 + 1)) {
				rg.c2++;
				grew = true;
			}
		}
		return rg;
	}

	private insertChart(type: ChartSpec['type']): void {
		const hot = this.hot;
		const selected = this.getSelectedRanges()[0];
		if (!hot || hot.isDestroyed || !selected) {
			new Notice('Select the cells to show in the chart first.');
			return;
		}
		const single = selected.r1 === selected.r2 && selected.c1 === selected.c2;
		const area = single
			? this.dataRegion(hot, selected.r1, selected.c1)
			: { ...selected };
		// The first column holds the labels (X values), the others the values.
		const split = area.c2 > area.c1;
		const range = split ? { ...area, c1: area.c1 + 1 } : area;
		const xRange = split ? { ...area, c2: area.c1 } : undefined;
		const read = this.chartReader();
		// Next to the data: right of its last column, level with its first row.
		let x = 0;
		for (let c = 0; c <= area.c2; c++)
			x += this.model.colWidths[c] || DEFAULT_COL_WIDTH;
		let y = 0;
		for (let r = 0; r < area.r1; r++)
			y += this.model.rowHeights[r] || this.computeRowHeight(r);
		const spec: ChartSpec = {
			id: newChartId(),
			type,
			title: '',
			range,
			xRange,
			headerRow: guessHeaderRow(area, read),
			trendline: false,
			x: x + 16,
			y,
			w: CHART_DEFAULT_W,
			h: CHART_DEFAULT_H,
			z: this.topZ(),
		};
		// The chart appears right away; the editor changes it from there.
		this.model.charts = [...this.model.charts, spec];
		this.markDirty();
		this.queueCharts();
		this.openChartEditor(spec.id, true);
	}

	/* ---------------- Zoom ---------------- */

	/** Zoom scales widths, and heights and fonts unless it is horizontal only (see zoomY); model values stay at 100%. */
	private applyZoom(next: number): void {
		const z = Math.min(
			MAX_ZOOM,
			Math.max(MIN_ZOOM, Math.round(next * 20) / 20),
		);
		if (z === this.zoom) return;
		this.zoom = z;
		this.applyZoomSizes();
		this.ui.zoomLabel?.setText(`${Math.round(z * 100)}%`);
	}

	/**
	 * The zoom of the vertical axis: row heights, the column of row numbers and the text. It is
	 * the zoom, or 1 when the zoom setting zooms the horizontal axis only (column widths and
	 * the column letters).
	 */
	private get zoomY(): number {
		return this.plugin.settings.zoomAxes === 'horizontal' ? 1 : this.zoom;
	}

	private rowHeaderWidth(): number {
		return Math.round(ROW_HEADER_WIDTH * this.zoomY);
	}

	/** Hands the zoomed sizes to the grid, after a zoom or a change of the zoom setting. */
	private applyZoomSizes(): void {
		this.applyZoomStyles();
		const hot = this.hot;
		if (hot && !hot.isDestroyed) {
			const cols = hot.getPlugin('manualColumnResize') as unknown as {
				setManualSize(c: number, w: number): void;
			};
			const rows = hot.getPlugin('manualRowResize') as unknown as {
				setManualSize(r: number, h: number): void;
			};
			this.model.colWidths.forEach((w, i) => {
				if (w) cols.setManualSize(i, Math.round(w * this.zoom));
			});
			this.model.rowHeights.forEach((h, i) => {
				if (h) rows.setManualSize(i, Math.round(h * this.zoomY));
			});
			// Handsontable keeps the theme's sizes, and its smallest row height comes from them.
			(
				hot as unknown as { stylesHandler?: { clearCache?(): void } }
			).stylesHandler?.clearCache?.();
			this.invalidateRowSizes();
			hot.updateSettings({ rowHeaderWidth: this.rowHeaderWidth() });
			hot.refreshDimensions();
		}
		this.queueCharts();
	}

	/* ---------------- Manual sizes ---------------- */

	/**
	 * "Widen column" cells: the column is as wide as the widest of those cells needs, and goes
	 * back to the width it had before as soon as no cell needs the extra room (a different
	 * mode, shorter text or a smaller font). A width the user set by dragging is never undone.
	 */
	private fitExpandColumns(cols: Iterable<number>): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const records = (this.model.widenRecords ??= {});
		const rows = hot.countRows();
		let changed = false;

		for (const col of new Set(cols)) {
			let needed = 0;
			let any = false;
			for (let r = 0; r < rows; r++) {
				const f = this.model.cellFormats[key(r, col)];
				if (
					f?.wrap !== 'expand' ||
					f.rotate ||
					this.model.images[key(r, col)]
				)
					continue;
				const mg = this.mergeAt(hot, r, col);
				// Cells hidden inside a merge keep their format but show nothing.
				if (mg && (mg.row !== r || mg.col !== col)) continue;
				any = true;
				const display = this.resolveRaw(
					hot.getSourceDataAtCell(r, col),
					0,
					false,
					{ row: r, col },
				);
				if (display === null || display === '') continue;
				const text =
					typeof display === 'boolean'
						? display
							? 'TRUE'
							: 'FALSE'
						: String(display);
				let width = this.measureTextWidth(text, f) + EXPAND_PADDING;
				// A merged cell already has the width of the other columns it covers.
				for (let c = col + 1; c < col + (mg?.colspan ?? 1); c++)
					width -= this.model.colWidths[c] || DEFAULT_COL_WIDTH;
				needed = Math.max(needed, width);
			}
			needed = Math.min(Math.ceil(needed), MAX_EXPAND_WIDTH);

			const colKey = String(col);
			const stored = this.model.colWidths[col] ?? null;
			const current = stored || DEFAULT_COL_WIDTH;
			let rec = records[colKey];
			if (rec && stored !== rec.w) {
				delete records[colKey]; // the user resized the column: forget the old width
				rec = undefined;
			}

			if (rec) {
				const baseWidth = rec.prevW || DEFAULT_COL_WIDTH;
				const target = any ? Math.max(baseWidth, needed) : baseWidth;
				if (target === baseWidth) {
					if (rec.prevW) this.setColumnWidth(col, rec.prevW);
					else this.clearColumnWidth(col);
					delete records[colKey];
					changed = true;
				} else if (target !== current) {
					this.setColumnWidth(col, target);
					rec.w = target;
					changed = true;
				}
			} else if (any && needed > current) {
				records[colKey] = { prevW: stored, w: needed };
				this.setColumnWidth(col, needed);
				changed = true;
			}
		}
		if (changed) hot.render();
	}

	private setColumnWidth(col: number, width: number): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const plugin = hot.getPlugin('manualColumnResize') as unknown as {
			setManualSize(column: number, width: number): void;
		};
		plugin.setManualSize(col, Math.round(width * this.zoom));
		while (this.model.colWidths.length <= col)
			this.model.colWidths.push(null);
		this.model.colWidths[col] = width;
	}

	private setRowHeight(row: number, height: number): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const plugin = hot.getPlugin('manualRowResize') as unknown as {
			setManualSize(row: number, height: number): void;
		};
		plugin.setManualSize(row, Math.round(height * this.zoomY));
		while (this.model.rowHeights.length <= row)
			this.model.rowHeights.push(null);
		this.model.rowHeights[row] = height;
	}

	/* ---------------- Images ---------------- */

	private pickLocalFile(accept: string): Promise<File | null> {
		return new Promise((resolve) => {
			const input = activeDocument.body.createEl('input', {
				type: 'file',
				cls: 'rtl-sheet-hidden',
			});
			input.accept = accept;
			const finish = (file: File | null): void => {
				input.remove();
				resolve(file);
			};
			input.addEventListener('change', () =>
				finish(input.files?.[0] ?? null),
			);
			input.addEventListener('cancel', () => finish(null));
			input.click();
		});
	}

	private pickImageFromVault(): void {
		if (!this.requireSelection()[0]) return;
		new VaultFilePickerModal(
			this.app,
			IMAGE_EXTENSIONS,
			'Choose an image from your vault…',
			(file) => {
				void this.setCellImage(file.path);
			},
		).open();
	}

	private async pickImageFromDevice(): Promise<void> {
		if (!this.requireSelection()[0]) return;
		const file = await this.pickLocalFile('image/*');
		if (!file) return;
		await this.addImageBlob(file, file.name);
	}

	private async pasteImageFromClipboard(): Promise<void> {
		if (!this.requireSelection()[0]) return;
		try {
			const items = await navigator.clipboard.read();
			for (const item of items) {
				const type = item.types.find((t) => t.startsWith('image/'));
				if (type) {
					await this.addImageBlob(await item.getType(type), null);
					return;
				}
			}
			new Notice('The clipboard does not contain an image.');
		} catch (e) {
			this.plugin.reportError(
				'Could not read an image from the clipboard',
				e,
			);
		}
	}

	/** Saves image data as an attachment in the vault and places it in a cell. */
	private async addImageBlob(
		blob: Blob,
		name: string | null,
		target?: { row: number; col: number },
	): Promise<void> {
		const selected = this.getSelectedRanges()[0];
		const cell =
			target ??
			(selected ? { row: selected.r1, col: selected.c1 } : null);
		if (!cell) {
			new Notice('Select a cell first.');
			return;
		}
		const path = await this.saveImageBlob(blob, name);
		if (path) await this.setCellImage(path, cell);
	}

	/** Saves image data as an attachment of this sheet; returns its vault path (null on failure). */
	private async saveImageBlob(
		blob: Blob,
		name: string | null,
	): Promise<string | null> {
		try {
			const subtype = (blob.type.split('/')[1] ?? 'png')
				.replace('jpeg', 'jpg')
				.replace('svg+xml', 'svg');
			const stamp = new Date().toISOString().replace(/[:.]/g, '-');
			const fileName =
				name && /\.[a-z0-9]+$/i.test(name)
					? name
					: `Pasted image ${stamp}.${subtype}`;
			const path =
				await this.app.fileManager.getAvailablePathForAttachment(
					fileName,
					this.file?.path ?? '',
				);
			await this.app.vault.createBinary(path, await blob.arrayBuffer());
			return path;
		} catch (e) {
			this.plugin.reportError('Could not add the image', e);
			return null;
		}
	}

	/** Ctrl+V / long-press paste of an image straight into the selected cell. */
	private onPaste(e: ClipboardEvent): void {
		// Any paste ends the copy outline, also one into a cell being edited or an image.
		this.clearCopyMark();
		const items = e.clipboardData?.items;
		if (!items) return;
		for (const item of Array.from(items)) {
			if (item.kind === 'file' && item.type.startsWith('image/')) {
				const blob = item.getAsFile();
				if (!blob) continue;
				e.preventDefault();
				e.stopPropagation();
				void this.addImageBlob(blob, null);
				return;
			}
		}
	}

	private hasImageFile(e: DragEvent): boolean {
		return Array.from(e.dataTransfer?.items ?? []).some(
			(i) => i.kind === 'file' && i.type.startsWith('image/'),
		);
	}

	/** Drop an image file from the file manager onto a cell. */
	private onDrop(e: DragEvent): void {
		const file = Array.from(e.dataTransfer?.files ?? []).find((f) =>
			f.type.startsWith('image/'),
		);
		if (!file) return;
		e.preventDefault();
		const hot = this.hot;
		const td =
			e.target instanceof HTMLElement ? e.target.closest('td') : null;
		let target: { row: number; col: number } | undefined;
		if (hot && td) {
			const coords = hot.getCoords(td);
			const row = Number(coords?.row ?? -1);
			const col = Number(coords?.col ?? -1);
			if (row >= 0 && col >= 0) {
				hot.selectCell(row, col);
				target = { row, col };
			}
		}
		void this.addImageBlob(file, file.name, target);
	}

	private async setCellImage(
		path: string,
		target?: { row: number; col: number },
	): Promise<void> {
		const hot = this.hot;
		const selected = this.getSelectedRanges()[0];
		const cell =
			target ??
			(selected ? { row: selected.r1, col: selected.c1 } : null);
		if (!hot || !cell) {
			new Notice('Select a cell first.');
			return;
		}
		const k = key(cell.row, cell.col);
		// The cell is resized once the image has loaded: that belongs to the same undo step.
		const coalesce = { tag: `image:${k}:${Date.now()}`, windowMs: 30000 };
		this.model.images[k] = path;
		hot.render();
		hot.selectCell(cell.row, cell.col);
		hot.listen();
		this.coalesce = coalesce;
		this.markDirty();

		await this.fitCellToImage(cell.row, cell.col, path);
		if (this.model.images[k] !== path || hot.isDestroyed) return; // removed in the meantime
		this.invalidateRowSizes();
		hot.render();
		this.coalesce = coalesce;
		this.markDirty();
	}

	/** Resizes the cell's row and column so the image is shown at a comfortable size. */
	private async fitCellToImage(
		row: number,
		col: number,
		path: string,
	): Promise<void> {
		const dims = await this.loadImageSize(path);
		if (!dims) return;
		if (this.model.images[key(row, col)] !== path) return; // image was removed while loading
		const scale = Math.min(
			1,
			MAX_IMAGE_CELL_PX / dims.w,
			MAX_IMAGE_CELL_PX / dims.h,
		);
		const w = Math.max(60, Math.round(dims.w * scale));
		const h = Math.max(30, Math.round(dims.h * scale));

		// Remember the original size (only the first time) so deleting the image can restore it.
		const k = key(row, col);
		const existing = this.model.imageSizes[k];
		this.model.imageSizes[k] = {
			prevW: existing
				? existing.prevW
				: (this.model.colWidths[col] ?? null),
			prevH: existing
				? existing.prevH
				: (this.model.rowHeights[row] ?? null),
			w,
			h,
		};
		this.setColumnWidth(col, w);
		this.setRowHeight(row, h);
	}

	/** The natural size of an image in the vault, or null when it cannot be loaded. */
	private loadImageSize(path: string): Promise<{ w: number; h: number } | null> {
		const file = this.app.vault.getAbstractFileByPath(path);
		if (!(file instanceof TFile)) return Promise.resolve(null);
		const url = this.app.vault.getResourcePath(file);
		return new Promise((resolve) => {
			const img = new Image();
			img.onload = () =>
				resolve(
					img.naturalWidth > 0 && img.naturalHeight > 0
						? { w: img.naturalWidth, h: img.naturalHeight }
						: null,
				);
			img.onerror = () => resolve(null);
			img.src = url;
		});
	}

	private clearColumnWidth(col: number): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const plugin = hot.getPlugin('manualColumnResize') as unknown as {
			clearManualSize(column: number): void;
		};
		plugin.clearManualSize(col);
		if (col < this.model.colWidths.length) this.model.colWidths[col] = null;
	}

	private clearRowHeight(row: number): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		// Unlike ManualColumnResize, ManualRowResize has no clearManualSize(); clear the
		// stored value in its index map instead, and fall back to a normal row's height.
		const plugin = hot.getPlugin('manualRowResize') as unknown as {
			rowHeightsMap?: { clearValue?: (physicalIndex: number) => void };
		};
		const map = plugin.rowHeightsMap;
		if (map && typeof map.clearValue === 'function') {
			map.clearValue(hot.toPhysicalRow(row));
		} else {
			this.setRowHeight(row, this.normalRowHeight(row));
		}
		if (row < this.model.rowHeights.length)
			this.model.rowHeights[row] = null;
	}

	/** Height of an ordinary row (no manual size, no image), used as a fallback default. */
	private normalRowHeight(excludeRow: number): number {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return 28;
		const limit = Math.min(hot.countRows(), 60);
		for (let r = 0; r < limit; r++) {
			if (r === excludeRow || this.model.rowHeights[r]) continue;
			if (
				Object.keys(this.model.images).some((k) =>
					k.startsWith(`${r},`),
				)
			)
				continue;
			const h = hot.getRowHeight(r);
			if (typeof h === 'number' && h > 0) return h;
		}
		return 28;
	}

	/** After an image is removed, give its row and column their old size back. */
	private restoreImageCellSize(row: number, col: number): void {
		const k = key(row, col);
		const rec = this.model.imageSizes[k];
		if (!rec) return;
		delete this.model.imageSizes[k];

		const others = Object.keys(this.model.images).map((imgKey) =>
			imgKey.split(',').map(Number),
		);
		const columnStillHasImage = others.some((p) => p[1] === col);
		const rowStillHasImage = others.some((p) => p[0] === row);

		// Only restore a size that the image set; if the user resized the cell since, keep their size.
		if (!columnStillHasImage && this.model.colWidths[col] === rec.w) {
			if (rec.prevW) this.setColumnWidth(col, rec.prevW);
			else this.clearColumnWidth(col);
		}
		if (!rowStillHasImage && this.model.rowHeights[row] === rec.h) {
			if (rec.prevH) this.setRowHeight(row, rec.prevH);
			else this.clearRowHeight(row);
		}
	}

	/** Removes the images in the given ranges and returns how many were removed. */
	private removeImagesInRanges(ranges: SelRange[]): number {
		let removed = 0;
		for (const rg of ranges) {
			for (let r = rg.r1; r <= rg.r2; r++) {
				for (let c = rg.c1; c <= rg.c2; c++) {
					if (this.model.images[key(r, c)]) {
						delete this.model.images[key(r, c)];
						this.restoreImageCellSize(r, c);
						removed++;
					}
				}
			}
		}
		return removed;
	}

	private removeImages(): void {
		const ranges = this.requireSelection();
		const hot = this.hot;
		if (ranges.length === 0 || !hot) return;
		if (this.removeImagesInRanges(ranges) === 0) {
			new Notice('There is no image in the selected cells.');
			return;
		}
		this.invalidateRowSizes();
		hot.render();
		this.markDirty();
	}

	/** Delete / Backspace on a selected cell also removes its image. */
	private onBeforeKeyDown(e: KeyboardEvent): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const editor = hot.getActiveEditor() as unknown as
			| { isOpened?: () => boolean }
			| undefined;
		if (editor?.isOpened?.()) return; // the user is editing text, not clearing cells
		if (e.key === 'Escape') {
			this.clearCopyMark();
			return;
		}
		if (e.key === ' ' && !e.ctrlKey && !e.metaKey && !e.altKey) {
			// Space checks or unchecks the selected check boxes (like Google Sheets).
			const cells = this.selectedCheckboxCells();
			if (cells.length > 0) {
				e.preventDefault();
				e.stopImmediatePropagation();
				this.toggleCheckboxCells(cells);
			}
			return;
		}
		if (e.key !== 'Delete' && e.key !== 'Backspace') return;
		if (this.activeFloatId) {
			e.preventDefault();
			e.stopImmediatePropagation();
			this.deleteFloatingImage(this.activeFloatId);
			return;
		}
		try {
			if (this.removeImagesInRanges(this.getSelectedRanges()) > 0) {
				this.invalidateRowSizes();
				hot.render();
				this.markDirty();
			}
		} catch (err) {
			this.plugin.reportError('Could not remove the image', err);
		}
	}

	/* ---------------- Floating images ---------------- */

	private pickFloatingImageFromVault(): void {
		new VaultFilePickerModal(
			this.app,
			IMAGE_EXTENSIONS,
			'Choose an image to float over the sheet…',
			(file) => {
				void this.addFloatingImage(file.path);
			},
		).open();
	}

	private async insertFloatingImageFrom(
		source: 'device' | 'clipboard',
	): Promise<void> {
		try {
			let blob: Blob | null = null;
			let name: string | null = null;
			if (source === 'device') {
				const file = await this.pickLocalFile('image/*');
				blob = file;
				name = file?.name ?? null;
			} else {
				for (const item of await navigator.clipboard.read()) {
					const type = item.types.find((t) => t.startsWith('image/'));
					if (type) {
						blob = await item.getType(type);
						break;
					}
				}
				if (!blob) {
					new Notice('The clipboard does not contain an image.');
					return;
				}
			}
			if (!blob) return;
			const path = await this.saveImageBlob(blob, name);
			if (path) await this.addFloatingImage(path);
		} catch (e) {
			this.plugin.reportError('Could not add the floating image', e);
		}
	}

	/** Puts an image over the sheet at a cell (the selected one or A1 when missing), at a comfortable size. */
	private async addFloatingImage(
		path: string,
		at?: { row: number; col: number },
	): Promise<void> {
		const dims = await this.loadImageSize(path);
		const scale = dims ? Math.min(1, FLOAT_DEFAULT_MAX_PX / dims.w) : 1;
		const w = Math.max(FLOAT_MIN_PX, Math.round((dims?.w ?? 200) * scale));
		const h = Math.max(FLOAT_MIN_PX, Math.round((dims?.h ?? 150) * scale));
		const cell = at ?? this.formulaCell ?? { row: 0, col: 0 };
		let x = 0;
		for (let c = 0; c < cell.col; c++)
			x += this.model.colWidths[c] || DEFAULT_COL_WIDTH;
		let y = 0;
		for (let r = 0; r < cell.row; r++)
			y += this.model.rowHeights[r] || this.computeRowHeight(r);
		const img: FloatingImage = {
			id: newChartId('img'),
			path,
			x,
			y,
			w,
			h,
			z: this.topZ(),
		};
		this.model.floatingImages = [...this.model.floatingImages, img];
		this.activeFloatId = img.id;
		this.markDirty();
		this.queueCharts();
	}

	private renderFloatingImage(
		layer: HTMLElement,
		img: FloatingImage,
		origin: { x: number; y: number },
	): void {
		const box = layer.createDiv({
			cls: 'rtl-sheet-float-img',
			attr: { 'data-id': img.id },
		});
		box.toggleClass('is-active', img.id === this.activeFloatId);
		this.placeChart(box, img, origin);
		// Presses on the image must not select cells or start editing them.
		for (const type of ['mousedown', 'touchstart', 'dblclick', 'contextmenu'])
			box.addEventListener(type, (e) => e.stopPropagation(), {
				passive: true,
			});
		const file = this.app.vault.getAbstractFileByPath(img.path);
		if (file instanceof TFile)
			box.createEl('img', {
				attr: {
					src: this.app.vault.getResourcePath(file),
					alt: file.name,
					draggable: 'false',
				},
			});
		else box.createDiv({ cls: 'rtl-sheet-float-missing', text: 'Missing image' });

		const deleteBtn = box.createEl('button', {
			cls: 'rtl-sheet-float-btn',
			attr: { 'aria-label': 'Delete image', title: 'Delete image' },
		});
		setIcon(deleteBtn, 'trash-2');
		deleteBtn.addEventListener('click', () => this.deleteFloatingImage(img.id));
		const layersBtn = box.createEl('button', {
			cls: 'rtl-sheet-float-btn is-layers',
			attr: { 'aria-label': 'Arrange', title: 'Arrange' },
		});
		setIcon(layersBtn, 'layers');
		layersBtn.addEventListener('click', () =>
			this.showMenu(layersBtn, this.layerMenuEntries(img.id)),
		);
		box.addEventListener('contextmenu', (e) => {
			e.preventDefault();
			this.setActiveFloat(img.id);
			this.showMenu(e, [
				...this.layerMenuEntries(img.id),
				{
					title: 'Delete image',
					icon: 'trash-2',
					onClick: () => this.deleteFloatingImage(img.id),
				},
			]);
		});
		const grip = box.createDiv({
			cls: 'rtl-sheet-chart-grip',
			attr: { 'aria-label': 'Resize image' },
		});
		const drag = (e: PointerEvent, mode: 'move' | 'resize'): void => {
			e.stopPropagation();
			this.setActiveFloat(img.id);
			this.hot?.listen(); // Delete removes the image
			this.dragBox(
				e,
				box,
				img,
				mode,
				{ w: FLOAT_MIN_PX, h: FLOAT_MIN_PX },
				true,
				(next) => this.replaceFloatingImage(img.id, (f) => ({ ...f, ...next })),
			);
		};
		grip.addEventListener('pointerdown', (e) => drag(e, 'resize'));
		box.addEventListener('pointerdown', (e) => {
			if (e.target === grip) return;
			drag(e, 'move');
		});
	}

	/** Marks the floating image that Delete removes (null: none) without drawing every chart again. */
	private setActiveFloat(id: string | null): void {
		if (this.activeFloatId === id) return;
		this.activeFloatId = id;
		this.chartLayer
			?.querySelectorAll<HTMLElement>('.rtl-sheet-float-img')
			.forEach((el) =>
				el.toggleClass('is-active', id !== null && el.dataset.id === id),
			);
	}

	/** Floating images are replaced, never changed in place (undo compares them). */
	private replaceFloatingImage(
		id: string,
		change: (img: FloatingImage) => FloatingImage | null,
	): void {
		const images: FloatingImage[] = [];
		for (const img of this.model.floatingImages) {
			if (img.id !== id) {
				images.push(img);
				continue;
			}
			const next = change(img);
			if (next) images.push(next);
		}
		this.model.floatingImages = images;
		this.markDirty();
		this.queueCharts();
	}

	private deleteFloatingImage(id: string): void {
		if (this.activeFloatId === id) this.activeFloatId = null;
		this.replaceFloatingImage(id, () => null);
		new Notice('Image deleted. Undo (Ctrl+Z) brings it back.');
	}

	/* ---------------- Emoji ---------------- */

	/** Inserts emojis where the user is typing (cell editor or formula bar), else at the end of the selected cell. */
	private openEmojiPicker(anchor: HTMLElement): void {
		const active = this.activeFormulaField();
		const field = active?.field ?? null;
		const caret = field
			? {
					from: field.selectionStart ?? field.value.length,
					to: field.selectionEnd ?? field.value.length,
				}
			: null;
		const cell = active?.cell ?? this.formulaCell;
		if (!field && !cell) {
			new Notice('Select a cell first.');
			return;
		}
		// The search box gets the keys; the grid gets them back when the picker closes.
		this.hot?.unlisten();
		this.emojiPicker.open(
			anchor,
			this.plugin.settings.recentEmojis,
			(emoji) => {
				this.rememberEmoji(emoji);
				if (field && caret && this.isStillEditing(field)) {
					field.focus();
					field.value =
						field.value.slice(0, caret.from) +
						emoji +
						field.value.slice(caret.to);
					caret.from = caret.to = caret.from + emoji.length;
					field.setSelectionRange(caret.from, caret.to);
					field.dispatchEvent(new Event('input', { bubbles: true }));
				} else if (cell) {
					this.appendToCell(cell, emoji);
				}
			},
			() => {
				if (field && this.isStillEditing(field)) field.focus();
				if (field !== this.ui.formulaInput) this.hot?.listen();
			},
		);
	}

	/** True while `field` is the formula bar, or the cell editor and it is still open. */
	private isStillEditing(field: HTMLInputElement | HTMLTextAreaElement): boolean {
		if (field === this.ui.formulaInput) return field.isConnected;
		const editor = this.hot?.getActiveEditor() as unknown as
			| { isOpened?: () => boolean; TEXTAREA?: HTMLTextAreaElement }
			| undefined;
		return !!editor?.isOpened?.() && editor.TEXTAREA === field;
	}

	private appendToCell(cell: { row: number; col: number }, text: string): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const current: unknown = hot.getSourceDataAtCell(cell.row, cell.col);
		if (typeof current === 'string' && current.startsWith('=')) {
			new Notice('The cell holds a formula: edit it to add an emoji.');
			return;
		}
		hot.setDataAtCell(cell.row, cell.col, sourceText(current) + text);
	}

	private rememberEmoji(emoji: string): void {
		const s = this.plugin.settings;
		s.recentEmojis = [emoji, ...s.recentEmojis.filter((e) => e !== emoji)].slice(
			0,
			MAX_RECENT_EMOJIS,
		);
		void this.plugin
			.saveSettings()
			.catch((e) => console.warn('[Free-Excel] Could not save the emojis', e));
	}

	/* ---------------- Check boxes ---------------- */

	/** Turns the selected cells into check boxes, or back into plain cells when the first one is a check box. */
	private toggleCheckboxes(): void {
		const ranges = this.requireSelection();
		const hot = this.hot;
		const first = ranges[0];
		if (!first || !hot || hot.isDestroyed) return;
		const on = !this.model.cellFormats[key(first.r1, first.c1)]?.checkbox;
		if (on) {
			// Empty cells and TRUE/FALSE-like text become real TRUE/FALSE values.
			const changes: [number, number, Scalar][] = [];
			for (const rg of ranges)
				for (let r = rg.r1; r <= rg.r2; r++)
					for (let c = rg.c1; c <= rg.c2; c++) {
						const v: unknown = hot.getSourceDataAtCell(r, c);
						if (typeof v === 'boolean') continue;
						const text = sourceText(v).trim();
						if (text === '' || /^(true|false|1|0)$/i.test(text))
							changes.push([r, c, isChecked(text)]);
					}
			if (changes.length > 0) hot.setDataAtCell(changes, 'checkbox');
		}
		this.applyFormat(ranges, (f) => {
			if (on) f.checkbox = true;
			else delete f.checkbox;
		});
	}

	/** The selected cells that show a check box. */
	private selectedCheckboxCells(): { row: number; col: number }[] {
		const out: { row: number; col: number }[] = [];
		for (const rg of this.getSelectedRanges())
			for (let r = rg.r1; r <= rg.r2; r++)
				for (let c = rg.c1; c <= rg.c2; c++)
					if (this.model.cellFormats[key(r, c)]?.checkbox)
						out.push({ row: r, col: c });
		return out;
	}

	/** Checks the boxes, or unchecks them all when the first one is checked. */
	private toggleCheckboxCells(cells: { row: number; col: number }[]): void {
		const hot = this.hot;
		const first = cells[0];
		if (!hot || hot.isDestroyed || !first) return;
		const isFormula = (c: { row: number; col: number }): boolean => {
			const v: unknown = hot.getSourceDataAtCell(c.row, c.col);
			return typeof v === 'string' && v.startsWith('=');
		};
		const free = cells.filter((c) => !isFormula(c));
		if (free.length === 0) {
			new Notice('This check box is set by a formula.');
			return;
		}
		const next = !isChecked(
			this.resolveRaw(hot.getSourceDataAtCell(first.row, first.col), 0, false, first),
		);
		hot.setDataAtCell(
			free.map((c): [number, number, Scalar] => [c.row, c.col, next]),
			'checkbox',
		);
	}

	private onCheckboxPress(e: PointerEvent): void {
		this.checkboxPress = null;
		if (e.pointerType === 'mouse' && e.button !== 0) return;
		if (!(e.target instanceof HTMLElement) || !e.target.closest('.rtl-sheet-checkbox'))
			return;
		if (this.refPickTarget(e)) return; // the click puts a reference into a formula
		const cell = this.cellFromEvent(e);
		if (cell)
			this.checkboxPress = {
				...cell,
				x: e.clientX,
				y: e.clientY,
				pointerId: e.pointerId,
			};
	}

	private onCheckboxRelease(e: PointerEvent): void {
		const press = this.checkboxPress;
		this.checkboxPress = null;
		if (!press || press.pointerId !== e.pointerId) return;
		if (Math.hypot(e.clientX - press.x, e.clientY - press.y) > 8) return;
		this.toggleCheckboxCells([{ row: press.row, col: press.col }]);
	}

	/** Check boxes change through toggleCheckboxCells only; Ctrl+click opens a link. */
	private onGridClick(e: MouseEvent): void {
		const target = e.target instanceof HTMLElement ? e.target : null;
		if (target?.closest('.rtl-sheet-checkbox')) {
			e.preventDefault();
			return;
		}
		const mod = Platform.isMacOS ? e.metaKey : e.ctrlKey;
		if (!mod || !target?.closest('.rtl-sheet-link')) return;
		const cell = this.cellFromEvent(e);
		const link = cell ? this.model.links[key(cell.row, cell.col)] : undefined;
		if (!link) return;
		e.preventDefault();
		e.stopPropagation();
		this.openLink(link);
	}

	/* ---------------- Links ---------------- */

	private editLink(): void {
		const hot = this.hot;
		const cell = this.formulaCell;
		if (!hot || hot.isDestroyed || !cell) {
			new Notice('Select a cell first.');
			return;
		}
		const k = key(cell.row, cell.col);
		const current = this.model.links[k];
		const raw: unknown = hot.getSourceDataAtCell(cell.row, cell.col);
		const isFormula = typeof raw === 'string' && raw.startsWith('=');
		const text = isFormula ? '' : sourceText(raw);
		new LinkModal(
			this.app,
			`${indexToColumnLabel(cell.col)}${cell.row + 1}`,
			{ text, target: current ?? '' },
			!!current,
			(link) => {
				if (!this.hot || this.hot.isDestroyed) return;
				const links = { ...this.model.links };
				if (link) links[k] = link.target;
				else delete links[k];
				this.model.links = links;
				// A formula keeps working: its result becomes the text of the link.
				if (link && !isFormula && link.text !== text)
					this.hot.setDataAtCell(cell.row, cell.col, link.text);
				this.hot.render();
				this.markDirty();
				this.refreshCellInfo();
			},
		).open();
	}

	/** The address a link gets in an .xlsx file: notes open in Obsidian. */
	private excelLinkTarget(target: string): string {
		if (isExternalLink(target)) return target;
		return `obsidian://open?vault=${encodeURIComponent(this.app.vault.getName())}&file=${encodeURIComponent(target)}`;
	}

	private openLink(target: string): void {
		if (isExternalLink(target)) {
			window.open(target);
			return;
		}
		void this.app.workspace.openLinkText(target, this.file?.path ?? '', 'tab');
	}

	/* ---------------- Comments ---------------- */

	private editComment(cell = this.formulaCell): void {
		if (!cell) {
			new Notice('Select a cell first.');
			return;
		}
		new CommentModal(
			this.app,
			`${indexToColumnLabel(cell.col)}${cell.row + 1}`,
			this.model.comments[key(cell.row, cell.col)] ?? null,
			(comment) => this.setComment(cell, comment),
		).open();
	}

	/** Comments are replaced, never changed in place (undo compares them). */
	private setComment(
		cell: { row: number; col: number },
		comment: CellComment | null,
	): void {
		const comments = { ...this.model.comments };
		if (comment) comments[key(cell.row, cell.col)] = comment;
		else delete comments[key(cell.row, cell.col)];
		this.model.comments = comments;
		this.hot?.render();
		this.markDirty();
		this.refreshCellInfo();
		this.refreshSidePanel();
	}

	private hasCellInfo(cell: { row: number; col: number }): boolean {
		const k = key(cell.row, cell.col);
		return !!(this.model.comments[k] || this.model.links[k]);
	}

	/** Updates the popup after `delay` ms, so moving over cells without a comment does not make it flicker. */
	private queueCellInfo(delay = 200): void {
		window.clearTimeout(this.cellInfoTimer);
		this.cellInfoTimer = window.setTimeout(() => this.refreshCellInfo(), delay);
	}

	/**
	 * Shows the comment and link of the cell under the pointer, or else of the selected cell,
	 * in a small popup next to it (like Excel). Hidden when neither has one.
	 */
	private refreshCellInfo(): void {
		window.clearTimeout(this.cellInfoTimer);
		const el = this.cellInfoEl;
		const hot = this.hot;
		const pick = (c: { row: number; col: number } | null) =>
			c && this.hasCellInfo(c) ? c : null;
		// While the pointer is on the popup it stays on its cell (its content may have changed).
		const cell =
			!hot || hot.isDestroyed || this.isEditing()
				? null
				: this.cellInfoHovered && this.cellInfoCell
					? pick(this.cellInfoCell)
					: (pick(this.hoverCell) ?? pick(this.formulaCell));
		if (!cell) this.cellInfoHovered = false;
		this.cellInfoCell = cell;
		el.empty();
		if (!cell) {
			el.hide();
			return;
		}
		const k = key(cell.row, cell.col);
		const comment = this.model.comments[k];
		const link = this.model.links[k];
		if (comment) {
			const box = el.createDiv({ cls: 'rtl-sheet-info-comment' });
			box.setCssProps({ '--rtl-sheet-comment-color': comment.color });
			box.createDiv({
				cls: 'rtl-sheet-info-text',
				text: comment.text,
				attr: { dir: 'auto' },
			});
			const foot = box.createDiv({ cls: 'rtl-sheet-info-foot' });
			foot.createSpan({
				cls: 'rtl-sheet-info-time',
				text: comment.time ? new Date(comment.time).toLocaleString() : '',
			});
			this.infoButton(foot, 'pencil', 'Edit comment', () =>
				this.editComment(cell),
			);
			this.infoButton(foot, 'trash-2', 'Delete comment', () =>
				this.setComment(cell, null),
			);
		}
		if (link) {
			const row = el.createDiv({ cls: 'rtl-sheet-info-link' });
			setIcon(row.createSpan({ cls: 'rtl-sheet-info-icon' }), 'link');
			const a = row.createEl('a', {
				text: link,
				href: isExternalLink(link) ? link : '#',
				attr: { dir: 'auto', title: link },
			});
			a.addEventListener('click', (e) => {
				e.preventDefault();
				this.openLink(link);
			});
			this.infoButton(row, 'pencil', 'Edit link', () => {
				this.hot?.selectCell(cell.row, cell.col);
				this.editLink();
			});
		}
		el.show();
		this.positionCellInfo();
	}

	private infoButton(
		parent: HTMLElement,
		icon: string,
		label: string,
		onClick: () => void,
	): void {
		const btn = parent.createEl('button', {
			cls: 'rtl-sheet-info-btn',
			attr: { 'aria-label': label, title: label },
		});
		setIcon(btn, icon);
		btn.addEventListener('click', onClick);
	}

	/** Places the popup beside its cell (on the side away from the first column), inside the grid. */
	private positionCellInfo(): void {
		const el = this.cellInfoEl;
		const cell = this.cellInfoCell;
		const hot = this.hot;
		if (!cell || !el.isShown() || !hot || hot.isDestroyed) return;
		const td = hot.getCell(cell.row, cell.col, true);
		const grid = this.gridEl.getBoundingClientRect();
		const rect = td?.getBoundingClientRect();
		if (
			!rect ||
			rect.bottom < grid.top ||
			rect.top > grid.bottom ||
			rect.right < grid.left ||
			rect.left > grid.right
		) {
			el.addClass('is-offscreen');
			return;
		}
		el.removeClass('is-offscreen');
		const host = this.contentEl.getBoundingClientRect();
		const width = el.offsetWidth;
		const height = el.offsetHeight;
		const after = rect.right - host.left + 4;
		const before = rect.left - host.left - width - 4;
		let left = this.model.rtl ? before : after;
		if (left + width > host.width - 4) left = before;
		if (left < 4) left = Math.min(after, host.width - width - 4);
		let top = rect.top - host.top;
		top = Math.max(4, Math.min(top, host.height - height - 4));
		el.style.left = `${Math.max(4, left)}px`;
		el.style.top = `${top}px`;
	}

	/* ---------------- Side panel (comments, AI) ---------------- */

	private toggleSidePanel(kind: 'comments' | 'ai'): void {
		if (this.sidePanel === kind) this.closeSidePanel();
		else this.openSidePanel(kind);
	}

	private openSidePanel(kind: 'comments' | 'ai'): void {
		this.sidePanel = kind;
		this.sidePanelEl.show();
		this.commentsPanelEl.toggle(kind === 'comments');
		this.aiPanelEl.toggle(kind === 'ai');
		if (kind === 'comments') this.renderCommentsPanel();
		else {
			if (!this.aiPanel)
				this.aiPanel = new AiAssistantPanel(this.aiPanelEl, {
					app: this.app,
					component: this,
					sourcePath: () => this.file?.path ?? '',
					settings: () => this.plugin.settings,
					describeSheet: () => this.describeSheetForAi(),
					tools: AI_TOOLS,
					runTool: (name, input) => this.runAiTool(name, input),
					close: () => this.closeSidePanel(),
				});
			if (!IS_TOUCH) this.aiPanel.focus();
		}
		this.onResize();
	}

	private closeSidePanel(): void {
		if (!this.sidePanel) return;
		this.sidePanel = null;
		this.sidePanelEl.hide();
		this.onResize();
		this.hot?.listen();
	}

	private refreshSidePanel(): void {
		if (this.sidePanel === 'comments') this.renderCommentsPanel();
	}

	/** Every comment of the shown sheet, in row order; a click selects its cell. */
	private renderCommentsPanel(): void {
		const el = this.commentsPanelEl;
		el.empty();
		const entries = Object.entries(this.model.comments)
			.map(([k, comment]) => {
				const parts = k.split(',').map(Number);
				return { row: parts[0] ?? 0, col: parts[1] ?? 0, comment };
			})
			.sort((a, b) => a.row - b.row || a.col - b.col);
		const head = el.createDiv({ cls: 'rtl-sheet-side-head' });
		const title = head.createDiv({ cls: 'rtl-sheet-side-title' });
		setIcon(title.createSpan(), 'messages-square');
		title.createSpan({ text: `Comments (${entries.length})` });
		const closeBtn = head.createEl('button', {
			cls: 'rtl-sheet-side-btn',
			attr: { 'aria-label': 'Close', title: 'Close' },
		});
		setIcon(closeBtn, 'x');
		closeBtn.addEventListener('click', () => this.closeSidePanel());

		const list = el.createDiv({ cls: 'rtl-sheet-comment-list' });
		if (entries.length === 0) {
			list.createDiv({
				cls: 'rtl-sheet-side-empty',
				text: 'No comments in this sheet. Select a cell and choose "Add comment".',
			});
			return;
		}
		for (const { row, col, comment } of entries) {
			const item = list.createDiv({ cls: 'rtl-sheet-comment-item' });
			item.setCssProps({ '--rtl-sheet-comment-color': comment.color });
			const top = item.createDiv({ cls: 'rtl-sheet-comment-item-head' });
			top.createSpan({
				cls: 'rtl-sheet-comment-cell',
				text: `${indexToColumnLabel(col)}${row + 1}`,
			});
			top.createSpan({
				cls: 'rtl-sheet-info-time',
				text: comment.time ? new Date(comment.time).toLocaleDateString() : '',
			});
			this.infoButton(top, 'pencil', 'Edit comment', () =>
				this.editComment({ row, col }),
			);
			this.infoButton(top, 'trash-2', 'Delete comment', () =>
				this.setComment({ row, col }, null),
			);
			item.createDiv({
				cls: 'rtl-sheet-info-text',
				text: comment.text,
				attr: { dir: 'auto' },
			});
			item.addEventListener('click', (e) => {
				if (e.target instanceof HTMLElement && e.target.closest('button')) return;
				const hot = this.hot;
				if (!hot || hot.isDestroyed) return;
				hot.selectCell(row, col);
				hot.scrollViewportTo({ row, col });
				hot.listen();
			});
		}
	}

	/* ---------------- AI ---------------- */

	private openPluginSettings(): void {
		const setting = (
			this.app as unknown as {
				setting?: { open(): void; openTabById(id: string): void };
			}
		).setting;
		setting?.open();
		setting?.openTabById(this.plugin.manifest.id);
	}

	/** Starts a CLAUDE() or GEMINI() formula in the formula bar of the selected cell. */
	private insertAiFormula(name: 'CLAUDE' | 'GEMINI'): void {
		const fn = FORMULA_FUNCTIONS.find((f) => f.name === name);
		const provider = name === 'CLAUDE' ? 'claude' : 'gemini';
		if (!aiApiKey(this.app, this.plugin.settings, provider))
			new Notice(
				`Add a ${name === 'CLAUDE' ? 'Claude' : 'Gemini'} API key in the settings, or the formula shows ${AI_ERROR_TEXT}.`,
			);
		if (fn) this.insertFunction(fn);
	}

	/** Forgets the answers of the AI formulas in the selection, so they are asked again. */
	private refreshAiFormulas(): void {
		const ranges = this.requireSelection();
		const hot = this.hot;
		if (ranges.length === 0 || !hot || hot.isDestroyed) return;
		const cells: string[] = [];
		for (const rg of ranges)
			for (let r = rg.r1; r <= rg.r2; r++)
				for (let c = rg.c1; c <= rg.c2; c++)
					cells.push(`${this.activeSheet}!${key(r, c)}`);
		this.aiRunner.retryFailed();
		const count = this.aiRunner.refreshCells(cells);
		this.invalidateRowSizes();
		hot.render();
		this.markDirty();
		new Notice(
			count > 0
				? `Asking again for ${count} AI formula answer(s).`
				: 'There are no AI formula answers in the selection.',
		);
	}

	/** A cell for the AI: its text, and for a formula the formula and its result. */
	private aiCellText(sheet: number, row: number, col: number): string {
		const raw = this.rawCell(sheet, row, col);
		const shown = this.resolveRaw(raw, 0, false, { row, col }, sheet);
		const text = shown === null ? '' : String(shown);
		if (typeof raw === 'string' && raw.startsWith('=') && raw.length > 1)
			return `${raw} → ${text}`;
		return text;
	}

	/** The sheet as text for the assistant: names, size, selection and the cells in use. */
	private describeSheetForAi(): string {
		const hot = this.hot;
		const lines: string[] = [];
		lines.push(
			`Shown sheet: "${this.model.name}" (sheets in the file: ${this.sheets.map((s) => `"${s.name}"`).join(', ')})`,
		);
		if (this.model.rtl) lines.push('Layout: right-to-left (Hebrew/Arabic).');
		if (!hot || hot.isDestroyed) return lines.join('\n');
		lines.push(`Grid size: ${hot.countRows()} rows × ${hot.countCols()} columns.`);
		const selected = this.getSelectedRanges().map(rangeLabel);
		lines.push(`Selected: ${selected.length > 0 ? selected.join(', ') : 'nothing'}.`);

		const lastRow = this.lastDataRow(hot);
		let lastCol = -1;
		for (let r = 0; r <= lastRow; r++) {
			const row = hot.getSourceDataAtRow(r) as unknown[];
			for (let c = row.length - 1; c > lastCol; c--)
				if (row[c] !== null && row[c] !== undefined && row[c] !== '') {
					lastCol = c;
					break;
				}
		}
		if (lastRow < 0 || lastCol < 0) {
			lines.push('The sheet is empty.');
		} else {
			const rows = Math.min(lastRow, AI_CONTEXT_ROWS - 1);
			const cols = Math.min(lastCol, AI_CONTEXT_COLS - 1);
			lines.push(
				`Cells in use: ${rangeLabel({ r1: 0, c1: 0, r2: lastRow, c2: lastCol })}. A formula is shown as "=FORMULA → result". Columns are separated by tabs:`,
			);
			const header = ['row'];
			for (let c = 0; c <= cols; c++) header.push(indexToColumnLabel(c));
			lines.push(header.join('\t'));
			for (let r = 0; r <= rows; r++) {
				const cells = [String(r + 1)];
				for (let c = 0; c <= cols; c++)
					cells.push(
						this.aiCellText(this.activeSheet, r, c)
							.replace(/[\t\n\r]+/g, ' ')
							.slice(0, AI_CONTEXT_CELL_CHARS),
					);
				lines.push(cells.join('\t'));
			}
			if (rows < lastRow || cols < lastCol)
				lines.push('(Only part of the cells is shown: use read_range for the rest.)');
		}
		const formats = Object.entries(this.model.cellFormats)
			.filter(([, f]) => f.checkbox)
			.map(([k]) => {
				const parts = k.split(',').map(Number);
				return `${indexToColumnLabel(parts[1] ?? 0)}${(parts[0] ?? 0) + 1}`;
			});
		if (formats.length > 0)
			lines.push(`Check box cells (TRUE = checked): ${formats.slice(0, 200).join(', ')}`);
		const comments = Object.entries(this.model.comments).slice(0, 100);
		if (comments.length > 0) {
			lines.push('Comments:');
			for (const [k, c] of comments) {
				const parts = k.split(',').map(Number);
				lines.push(
					`${indexToColumnLabel(parts[1] ?? 0)}${(parts[0] ?? 0) + 1}: ${c.text.replace(/\s+/g, ' ')}`,
				);
			}
		}
		const cellImages = Object.entries(this.model.images).slice(0, 50);
		if (cellImages.length > 0)
			lines.push(
				`Images in cells: ${cellImages
					.map(([k, path]) => {
						const parts = k.split(',').map(Number);
						return `${indexToColumnLabel(parts[1] ?? 0)}${(parts[0] ?? 0) + 1} = ${path}`;
					})
					.join('; ')}`,
			);
		if (this.model.floatingImages.length > 0)
			lines.push(
				`Floating images: ${this.model.floatingImages.map((f) => f.path).slice(0, 50).join('; ')}`,
			);
		if (this.model.charts.length > 0)
			lines.push(
				`Charts: ${this.model.charts.map((c) => `${c.type} of ${rangeLabel(c.range)}${c.title ? ` "${c.title}"` : ''}`).join('; ')}`,
			);
		return lines.join('\n');
	}

	/** Runs a tool call of the assistant on the shown sheet. Errors are thrown with a message for the model. */
	private runAiTool(
		name: string,
		input: Record<string, unknown>,
	): AiToolResult | Promise<AiToolResult> {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) throw new Error('The sheet is not ready.');
		const text = (k: string): string => {
			const v = input[k];
			return typeof v === 'string' ? v : typeof v === 'number' ? String(v) : '';
		};
		const range = (k = 'range'): SelRange => {
			const rg = parseRangeRef(text(k));
			if (!rg) throw new Error(`"${text(k)}" is not a cell or range in A1 notation.`);
			return rg;
		};
		const cellOf = (k: string): { row: number; col: number } => {
			const rg = range(k);
			return { row: rg.r1, col: rg.c1 };
		};
		/** Adds rows and columns so the range fits in the grid. */
		const fit = (rg: SelRange): void => {
			const rows = rg.r2 + 1 - hot.countRows();
			if (rows > 0) hot.alter('insert_row_below', hot.countRows() - 1, rows);
			const cols = rg.c2 + 1 - hot.countCols();
			if (cols > 0) hot.alter('insert_col_end', hot.countCols() - 1, cols);
		};
		const flag = (k: string): boolean | undefined => {
			const v = input[k];
			return typeof v === 'boolean' ? v : undefined;
		};
		/** Where an image tool puts its image: the given cell, else the selected one, else A1. */
		const imageCell = (): { row: number; col: number } => {
			if (text('cell')) {
				const cell = cellOf('cell');
				fit({ r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });
				return cell;
			}
			const sel = this.getSelectedRanges()[0];
			return sel ? { row: sel.r1, col: sel.c1 } : { row: 0, col: 0 };
		};

		switch (name) {
			case 'read_range': {
				const sheetName = text('sheet');
				const sheet = sheetName ? this.sheetIndexByName(sheetName) : this.activeSheet;
				if (sheet < 0) throw new Error(`There is no sheet named "${sheetName}".`);
				const rg = range();
				const width = rg.c2 - rg.c1 + 1;
				const r2 = Math.min(rg.r2, rg.r1 + Math.max(0, Math.floor(AI_MAX_READ_CELLS / width) - 1));
				const rows: string[][] = [];
				for (let r = rg.r1; r <= r2; r++) {
					const row: string[] = [];
					for (let c = rg.c1; c <= rg.c2; c++) row.push(this.aiCellText(sheet, r, c));
					rows.push(row);
				}
				const read = { ...rg, r2 };
				return {
					result: {
						range: rangeLabel(read),
						rows,
						...(r2 < rg.r2 ? { note: 'Cut short: read the rest separately.' } : {}),
					},
					summary: `Read ${sheetName ? `${sheetName}!` : ''}${rangeLabel(read)}`,
				};
			}
			case 'write_cells': {
				const start = cellOf('start');
				const raw = input.rows;
				if (!Array.isArray(raw)) throw new Error('"rows" must be a list of rows.');
				const changes: [number, number, Scalar][] = [];
				let width = 0;
				raw.forEach((row: unknown, i) => {
					const cells = Array.isArray(row) ? row : [row];
					width = Math.max(width, cells.length);
					cells.forEach((v: unknown, j) => {
						const value =
							v === null || v === undefined || v === ''
								? null
								: typeof v === 'boolean' || typeof v === 'number'
									? v
									: typeof v === 'string'
										? v
										: JSON.stringify(v);
						changes.push([start.row + i, start.col + j, value]);
					});
				});
				if (changes.length === 0) throw new Error('Nothing to write.');
				const rg = {
					r1: start.row,
					c1: start.col,
					r2: start.row + raw.length - 1,
					c2: start.col + Math.max(1, width) - 1,
				};
				fit(rg);
				hot.setDataAtCell(changes, 'ai');
				return { result: { written: rangeLabel(rg) }, summary: `Wrote ${rangeLabel(rg)}` };
			}
			case 'clear_range': {
				const rg = range();
				const changes: [number, number, Scalar][] = [];
				for (let r = rg.r1; r <= Math.min(rg.r2, hot.countRows() - 1); r++)
					for (let c = rg.c1; c <= Math.min(rg.c2, hot.countCols() - 1); c++)
						changes.push([r, c, null]);
				if (changes.length > 0) hot.setDataAtCell(changes, 'ai');
				return { result: { cleared: rangeLabel(rg) }, summary: `Cleared ${rangeLabel(rg)}` };
			}
			case 'format_range': {
				const rg = range();
				fit(rg);
				const color = (k: string): string | null | undefined => {
					if (!(k in input)) return undefined;
					const v = text(k).trim().toLowerCase();
					if (v === '') return null;
					if (!/^#[0-9a-f]{6}$/.test(v)) throw new Error(`${k} must be "#rrggbb".`);
					return v;
				};
				const textColor = color('text_color');
				const fillColor = color('fill_color');
				const pick = <T extends string>(k: string, allowed: readonly T[]): T | '' | undefined => {
					if (!(k in input)) return undefined;
					const v = text(k);
					if (v === '') return '';
					if (!(allowed as readonly string[]).includes(v))
						throw new Error(`${k} must be one of ${allowed.join(', ')}.`);
					return v as T;
				};
				const align = pick('align', ['left', 'center', 'right'] as const);
				const valign = pick('vertical_align', ['top', 'middle', 'bottom'] as const);
				const wrap = pick('wrap', ['overflow', 'wrap', 'expand'] as const);
				const size = typeof input.font_size === 'number' ? Math.round(input.font_size) : undefined;
				this.applyFormat([rg], (f) => {
					const set = (k: 'bold' | 'italic' | 'underline' | 'strike' | 'checkbox', v: boolean | undefined): void => {
						if (v === true) f[k] = true;
						else if (v === false) delete f[k];
					};
					set('bold', flag('bold'));
					set('italic', flag('italic'));
					set('underline', flag('underline'));
					set('strike', flag('strikethrough'));
					set('checkbox', flag('checkbox'));
					if (textColor === null) delete f.textColor;
					else if (textColor) f.textColor = textColor;
					if (fillColor === null) delete f.bgColor;
					else if (fillColor) f.bgColor = fillColor;
					if (align === '') delete f.align;
					else if (align) f.align = align;
					if (valign === '') delete f.valign;
					else if (valign) f.valign = valign;
					if (wrap === '') delete f.wrap;
					else if (wrap) f.wrap = wrap;
					if (size !== undefined) {
						if (size > 0) f.fontSize = Math.min(96, Math.max(6, size));
						else delete f.fontSize;
					}
				});
				const borders = pick('borders', ['all', 'outer', 'none'] as const);
				if (borders) this.applyBorders(borders, [rg]);
				return { result: { formatted: rangeLabel(rg) }, summary: `Formatted ${rangeLabel(rg)}` };
			}
			case 'set_comment': {
				const cell = cellOf('cell');
				const body = text('text').trim();
				const color = text('color').toLowerCase();
				this.setComment(
					cell,
					body
						? {
								text: body,
								color: /^#[0-9a-f]{6}$/.test(color)
									? color
									: (this.model.comments[key(cell.row, cell.col)]?.color ??
										DEFAULT_COMMENT_COLOR),
								time: Date.now(),
							}
						: null,
				);
				const label = rangeLabel({ r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });
				return {
					result: { ok: true },
					summary: body ? `Comment on ${label}` : `Removed the comment on ${label}`,
				};
			}
			case 'set_link': {
				const cell = cellOf('cell');
				const target = text('url').trim();
				const links = { ...this.model.links };
				if (target) links[key(cell.row, cell.col)] = target;
				else delete links[key(cell.row, cell.col)];
				this.model.links = links;
				const shown = text('text').trim();
				if (target && shown) hot.setDataAtCell(cell.row, cell.col, shown, 'ai');
				else if (target && hot.getSourceDataAtCell(cell.row, cell.col) == null)
					hot.setDataAtCell(cell.row, cell.col, target, 'ai');
				hot.render();
				this.markDirty();
				const label = rangeLabel({ r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });
				return { result: { ok: true }, summary: target ? `Link in ${label}` : `Removed the link in ${label}` };
			}
			case 'insert_chart': {
				const type = text('type') as ChartType;
				if (!CHART_TYPES.some((t) => t.value === type))
					throw new Error(`type must be one of ${CHART_TYPES.map((t) => t.value).join(', ')}.`);
				const rg = range();
				hot.selectCell(rg.r1, rg.c1, rg.r2, rg.c2);
				this.insertChart(type);
				const id = this.chartEditorId;
				const title = text('title');
				if (id && title) this.replaceChart(id, (c) => ({ ...c, title }));
				return { result: { ok: true }, summary: `Inserted a ${type} chart of ${rangeLabel(rg)}` };
			}
			case 'find_images': {
				const words = text('query').toLowerCase().split(/\s+/).filter(Boolean);
				const files = this.app.vault
					.getFiles()
					.filter((f) => IMAGE_EXTENSIONS.includes(f.extension.toLowerCase()))
					.map((f) => {
						const path = f.path.toLowerCase();
						return { f, score: words.filter((w) => path.includes(w)).length };
					})
					.filter((x) => words.length === 0 || x.score > 0)
					.sort((a, b) => b.score - a.score || b.f.stat.mtime - a.f.stat.mtime)
					.slice(0, AI_MAX_FOUND_IMAGES)
					.map((x) => x.f.path);
				return {
					result: files.length > 0 ? { images: files } : { images: [], note: 'No image in the vault matches.' },
					summary: `Found ${files.length} image(s)`,
				};
			}
			case 'search_web_images': {
				const query = text('query').trim();
				if (!query) throw new Error('"query" is missing.');
				return this.searchWebImages(query);
			}
			case 'insert_image': {
				const source = text('source').trim();
				if (!source) throw new Error('"source" is missing.');
				return this.aiInsertImage(source, imageCell(), flag('floating') ?? false);
			}
			case 'generate_image': {
				const prompt = text('prompt').trim();
				if (!prompt) throw new Error('"prompt" is missing.');
				return this.aiGenerateImage(prompt, imageCell(), flag('floating') ?? false);
			}
			case 'select_range': {
				const rg = range();
				hot.selectCell(rg.r1, rg.c1, rg.r2, rg.c2);
				hot.scrollViewportTo({ row: rg.r1, col: rg.c1 });
				return { result: { ok: true }, summary: `Selected ${rangeLabel(rg)}` };
			}
			default:
				throw new Error(`Unknown tool "${name}".`);
		}
	}

	/** Pictures on Wikimedia Commons that match a query (the search_web_images tool). */
	private async searchWebImages(query: string): Promise<AiToolResult> {
		const params = new URLSearchParams({
			action: 'query',
			format: 'json',
			generator: 'search',
			gsrsearch: `filetype:bitmap ${query}`,
			gsrnamespace: '6',
			gsrlimit: String(AI_MAX_WEB_IMAGES),
			prop: 'imageinfo',
			iiprop: 'url|size|mime',
			iiurlwidth: String(AI_WEB_IMAGE_WIDTH),
		});
		const res = await requestUrl({
			url: `https://commons.wikimedia.org/w/api.php?${params.toString()}`,
			throw: false,
		});
		if (res.status >= 400) throw new Error(`The image search failed (${res.status}).`);
		const pages = (
			res.json as {
				query?: {
					pages?: Record<
						string,
						{
							title?: string;
							index?: number;
							imageinfo?: { thumburl?: string; url?: string; width?: number; height?: number }[];
						}
					>;
				};
			}
		).query?.pages;
		const images = Object.values(pages ?? {})
			.sort((a, b) => (a.index ?? 0) - (b.index ?? 0))
			.flatMap((p) => {
				const info = p.imageinfo?.[0];
				const url = info?.thumburl ?? info?.url;
				if (!url) return [];
				return [
					{
						title: (p.title ?? '').replace(/^File:/, ''),
						url,
						size: `${info?.width ?? '?'}×${info?.height ?? '?'}`,
					},
				];
			});
		return {
			result: images.length > 0 ? { images } : { images: [], note: 'Nothing found: try other (English) words.' },
			summary: `Searched images for "${query}" (${images.length})`,
		};
	}

	/** Puts a vault image or a downloaded one in a cell or over the sheet (the insert_image tool). */
	private async aiInsertImage(
		source: string,
		cell: { row: number; col: number },
		floating: boolean,
	): Promise<AiToolResult> {
		let path: string | null;
		if (/^https?:\/\//i.test(source)) {
			const res = await requestUrl({ url: source, throw: false });
			if (res.status >= 400)
				throw new Error(`Downloading the image failed (${res.status}).`);
			const header = Object.entries(res.headers).find(
				([name]) => name.toLowerCase() === 'content-type',
			)?.[1];
			const fileName = decodeURIComponent(
				source.split(/[?#]/)[0]?.split('/').pop() ?? '',
			);
			const ext = fileName.split('.').pop()?.toLowerCase() ?? '';
			const type = header?.split(';')[0]?.trim().toLowerCase() ?? '';
			if (!type.startsWith('image/') && !IMAGE_EXTENSIONS.includes(ext))
				throw new Error(
					`The address is not an image (${type || 'unknown type'}). Use the address of the image file itself.`,
				);
			const blob = new Blob([res.arrayBuffer], {
				type: type.startsWith('image/') ? type : `image/${ext.replace('jpg', 'jpeg')}`,
			});
			path = await this.saveImageBlob(
				blob,
				IMAGE_EXTENSIONS.includes(ext) ? fileName : null,
			);
			if (!path) throw new Error('Could not save the downloaded image in the vault.');
		} else {
			const file =
				this.app.vault.getAbstractFileByPath(normalizePath(source)) ??
				this.app.metadataCache.getFirstLinkpathDest(
					source.replace(/^!?\[\[|\]\]$/g, ''),
					this.file?.path ?? '',
				);
			if (!(file instanceof TFile) || !IMAGE_EXTENSIONS.includes(file.extension.toLowerCase()))
				throw new Error(`There is no image "${source}" in the vault. Use find_images to look for one.`);
			path = file.path;
		}
		return this.placeAiImage(path, cell, floating);
	}

	/** Makes a picture with Gemini and places it (the generate_image tool). */
	private async aiGenerateImage(
		prompt: string,
		cell: { row: number; col: number },
		floating: boolean,
	): Promise<AiToolResult> {
		const image = await generateGeminiImage(this.app, this.plugin.settings, prompt);
		const ext = (image.mimeType.split('/')[1] ?? 'png').replace('jpeg', 'jpg');
		const stamp = new Date().toISOString().replace(/[:.]/g, '-');
		const path = await this.saveImageBlob(
			new Blob([image.data], { type: image.mimeType }),
			`Generated image ${stamp}.${ext}`,
		);
		if (!path) throw new Error('Could not save the image in the vault.');
		return this.placeAiImage(path, cell, floating);
	}

	private async placeAiImage(
		path: string,
		cell: { row: number; col: number },
		floating: boolean,
	): Promise<AiToolResult> {
		const label = rangeLabel({ r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });
		if (floating) await this.addFloatingImage(path, cell);
		else await this.setCellImage(path, cell);
		return {
			result: { ok: true, path, placed: floating ? `floating at ${label}` : `in ${label}` },
			summary: floating ? `Floating image at ${label}` : `Image in ${label}`,
		};
	}

	/* ---------------- Import ---------------- */

	private async importFromDevice(): Promise<void> {
		const file = await this.pickLocalFile(
			'.xlsx,.xlsm,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		);
		if (!file) return;
		await this.importBuffer(file.name, await file.arrayBuffer());
	}

	private importFromVault(): void {
		new VaultFilePickerModal(
			this.app,
			['xlsx', 'xlsm', 'csv'],
			'Choose a spreadsheet to import…',
			(file) => {
				void (async () => {
					try {
						await this.importBuffer(
							file.name,
							await this.app.vault.readBinary(file),
						);
					} catch (e) {
						this.plugin.reportError('Import failed', e);
					}
				})();
			},
		).open();
	}

	private askImportChoice(prompt: ImportPrompt): Promise<ImportChoice | null> {
		return new Promise((resolve) => {
			new ImportDirectionModal(this.app, prompt, resolve).open();
		});
	}

	private async importBuffer(name: string, data: ArrayBuffer): Promise<void> {
		try {
			const ext = name.split('.').pop()?.toLowerCase() ?? '';
			const imported =
				ext === 'csv'
					? parseCsv(decodeText(data))
					: await parseXlsx(data);

			// The start cell is read now: the dialog and a direction change both drop the selection.
			const anchor = this.getSelectedRanges()[0];
			const ar = anchor?.r1 ?? 0;
			const ac = anchor?.c1 ?? 0;

			const choice = await this.askImportChoice({
				detectedRtl: imported.rtl,
				startCell: `${indexToColumnLabel(ac)}${ar + 1}`,
				hasFormulas: Object.keys(imported.results).length > 0,
			});
			if (choice === null) return; // dialog dismissed: cancel the import
			if (choice.direction !== 'keep')
				this.setDirection(choice.direction === 'rtl');
			this.applyImported(imported, ar, ac, choice.keepFormulas);
		} catch (e) {
			this.plugin.reportError('Import failed', e);
		}
	}

	/**
	 * The cells of an imported table as they should be written at row `ar`, column `ac`:
	 * formulas are moved so they keep pointing at the same imported cells, or replaced by
	 * the values the file stored for them.
	 */
	private importedCells(
		imp: ImportedSheet,
		ar: number,
		ac: number,
		keepFormulas: boolean,
	): Scalar[][] {
		return imp.data.map((row, r) =>
			row.map((v, c) => {
				if (typeof v !== 'string' || v.length < 2 || !v.startsWith('='))
					return v;
				const k = key(r, c);
				if (!keepFormulas && k in imp.results)
					return imp.results[k] ?? null;
				return moveFormula(v, ar, ac);
			}),
		);
	}

	/** Writes imported data into this sheet, starting at row `ar`, column `ac`. */
	private applyImported(
		imp: ImportedSheet,
		ar: number,
		ac: number,
		keepFormulas: boolean,
	): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const rows = imp.data.length;
		const cols = Math.max(0, ...imp.data.map((r) => r.length));
		if (rows === 0 || cols === 0) {
			new Notice('Nothing to import.');
			return;
		}

		// One render for the whole import: each merge would otherwise draw the grid again.
		hot.batch(() => {
			const missingRows = ar + rows - hot.countRows();
			if (missingRows > 0)
				hot.alter('insert_row_below', hot.countRows() - 1, missingRows);
			const missingCols = ac + cols - hot.countCols();
			if (missingCols > 0)
				hot.alter('insert_col_end', hot.countCols() - 1, missingCols);

			hot.populateFromArray(
				ar,
				ac,
				this.importedCells(imp, ar, ac, keepFormulas),
			);

			const shifted = (k: string): string => {
				const parts = k.split(',').map(Number);
				return key((parts[0] ?? 0) + ar, (parts[1] ?? 0) + ac);
			};
			for (const [k, f] of Object.entries(imp.formats))
				this.model.cellFormats[shifted(k)] = f;
			const comments = { ...this.model.comments };
			for (const [k, text] of Object.entries(imp.comments))
				comments[shifted(k)] = {
					text,
					color: DEFAULT_COMMENT_COLOR,
					time: Date.now(),
				};
			this.model.comments = comments;
			const links = { ...this.model.links };
			for (const [k, target] of Object.entries(imp.links))
				links[shifted(k)] = target;
			this.model.links = links;

			const mergePlugin = hot.getPlugin('mergeCells');
			for (const mg of imp.merges) {
				mergePlugin.merge(
					mg.row + ar,
					mg.col + ac,
					mg.row + mg.rowspan - 1 + ar,
					mg.col + mg.colspan - 1 + ac,
				);
			}
			imp.colWidths.forEach((w, i) => {
				if (w) this.setColumnWidth(i + ac, w);
			});
			imp.rowHeights.forEach((h, i) => {
				if (h) this.setRowHeight(i + ar, h);
			});
		});

		this.invalidateRowSizes();
		hot.render();
		this.markDirty();
		new Notice(
			`Imported ${rows} rows × ${cols} columns starting at ${indexToColumnLabel(ac)}${ar + 1}.`,
		);
	}

	/* ---------------- Export ---------------- */

	private toExcelValue(
		raw: unknown,
		row: number,
		col: number,
		sheet: number,
	): ExcelJS.CellValue {
		if (typeof raw === 'string') {
			if (
				this.plugin.settings.enableFormulas &&
				raw.length > 1 &&
				raw.startsWith('=')
			) {
				const result = this.resolveRaw(
					raw,
					0,
					false,
					{ row, col },
					sheet,
				);
				return {
					formula: normalizeSeparators(raw.slice(1)),
					result: (result ??
						undefined) as ExcelJS.CellFormulaValue['result'],
				};
			}
			if (NUMERIC_RE.test(raw)) return Number(raw);
			return raw;
		}
		if (typeof raw === 'number' || typeof raw === 'boolean') return raw;
		return String(raw);
	}

	async exportToXlsx(): Promise<void> {
		const file = this.file;
		if (!file) {
			new Notice('Open a sheet file before exporting.');
			return;
		}
		try {
			this.captureGridState();
			const wb = new ExcelJS.Workbook();
			let skippedImages = 0;
			let skippedCharts = 0;
			for (let i = 0; i < this.sheets.length; i++) {
				const m = this.sheets[i];
				if (!m) continue;
				skippedImages += await this.exportSheet(wb, m, i);
				skippedCharts += m.charts.length;
			}

			const buffer = await wb.xlsx.writeBuffer();
			const bytes = new Uint8Array(buffer);
			const data = bytes.buffer;

			const path = normalizePath(
				file.path.replace(/\.[^./\\]+$/, '') + '.xlsx',
			);
			const existing = this.app.vault.getAbstractFileByPath(path);
			if (existing instanceof TFile)
				await this.app.vault.modifyBinary(existing, data);
			else await this.app.vault.createBinary(path, data);

			const notes: string[] = [];
			if (skippedImages > 0)
				notes.push(
					`${skippedImages} image(s) skipped: only png, jpg and gif are supported in .xlsx`,
				);
			if (skippedCharts > 0)
				notes.push(`${skippedCharts} chart(s) are not exported`);
			new Notice(
				notes.length > 0
					? `Exported to ${path} (${notes.join('; ')})`
					: `Exported to ${path}`,
			);
		} catch (e) {
			this.plugin.reportError('Excel export failed', e);
		}
	}

	/** Adds one sheet to the workbook; returns how many images could not be exported. */
	private async exportSheet(
		wb: ExcelJS.Workbook,
		m: SheetModel,
		sheetIndex: number,
	): Promise<number> {
		const defaultSizePx = this.plugin.settings.defaultFontSize;
		const ws = wb.addWorksheet(sanitizeSheetName(m.name), {
			views: [{ rightToLeft: m.rtl }],
		});
		{
			m.data.forEach((rowArr, r) => {
				rowArr.forEach((raw, c) => {
					const f = m.cellFormats[key(r, c)];
					const hasValue =
						raw !== null && raw !== undefined && raw !== '';
					if (!hasValue && !f) return;

					const cell = ws.getCell(r + 1, c + 1);
					let isNumber = false;
					if (hasValue) {
						const value = this.toExcelValue(raw, r, c, sheetIndex);
						cell.value = value;
						const target = m.links[key(r, c)];
						if (target && typeof value !== 'object')
							cell.value = {
								text: String(value),
								hyperlink: this.excelLinkTarget(target),
							};
						isNumber =
							typeof value === 'number' ||
							(typeof value === 'object' &&
								value !== null &&
								'formula' in value &&
								typeof (value as ExcelJS.CellFormulaValue)
									.result === 'number');
					}

					cell.font = {
						size:
							Math.round(
								(f?.fontSize ?? defaultSizePx) * 0.75 * 2,
							) / 2,
						bold: !!f?.bold,
						italic: !!f?.italic,
						underline: !!f?.underline,
						strike: !!f?.strike,
						...(f?.textColor
							? { color: { argb: toArgb(f.textColor) } }
							: {}),
					};
					cell.alignment = {
						horizontal:
							f?.align ?? (m.rtl || isNumber ? 'right' : 'left'),
						vertical: f?.valign ?? 'middle',
						readingOrder: m.rtl ? 'rtl' : 'ltr',
						wrapText: f?.wrap === 'wrap',
						...(f?.rotate
							? { textRotation: rotationToExcel(f.rotate) }
							: {}),
					};
					if (f?.bgColor) {
						cell.fill = {
							type: 'pattern',
							pattern: 'solid',
							fgColor: { argb: toArgb(f.bgColor) },
						};
					}
					if (f?.borders) {
						const side: Partial<ExcelJS.Border> = {
							style: 'thin',
							color: {
								argb: f.borders.color
									? toArgb(f.borders.color)
									: 'FF000000',
							},
						};
						cell.border = {
							top: f.borders.top ? side : undefined,
							right: f.borders.right ? side : undefined,
							bottom: f.borders.bottom ? side : undefined,
							left: f.borders.left ? side : undefined,
						};
					}
				});
			});

			m.colWidths.forEach((w, i) => {
				if (w) ws.getColumn(i + 1).width = Math.max(2, w / 7);
			});
			m.rowHeights.forEach((h, i) => {
				if (h) ws.getRow(i + 1).height = h * 0.75;
			});
			for (const mg of m.merges) {
				ws.mergeCells(
					mg.row + 1,
					mg.col + 1,
					mg.row + mg.rowspan,
					mg.col + mg.colspan,
				);
			}

			for (const [k, comment] of Object.entries(m.comments)) {
				const parts = k.split(',').map(Number);
				ws.getCell((parts[0] ?? 0) + 1, (parts[1] ?? 0) + 1).note =
					comment.text;
			}

			let skippedImages = 0;
			for (const [k, path] of Object.entries(m.images)) {
				const imageFile = this.app.vault.getAbstractFileByPath(path);
				const excelExt = excelImageType(imageFile);
				if (!(imageFile instanceof TFile) || !excelExt) {
					skippedImages++; // webp/svg are not supported inside .xlsx files
					continue;
				}
				const parts = k.split(',').map(Number);
				const r = parts[0] ?? 0;
				const c = parts[1] ?? 0;
				const bytes = await this.app.vault.readBinary(imageFile);
				const imageId = wb.addImage({
					buffer: bytes,
					extension: excelExt,
				});
				ws.addImage(imageId, {
					tl: { col: c, row: r },
					br: { col: c + 1, row: r + 1 },
					editAs: 'oneCell',
				} as unknown as Parameters<ExcelJS.Worksheet['addImage']>[1]);
			}

			// Floating images keep their place: px are turned into (fractional) columns and rows.
			const colAt = (px: number): number => {
				let c = 0;
				let rest = px;
				for (;;) {
					const w = m.colWidths[c] || DEFAULT_COL_WIDTH;
					if (rest < w) return c + rest / w;
					rest -= w;
					c++;
				}
			};
			const rowAt = (px: number): number => {
				let r = 0;
				let rest = px;
				for (;;) {
					const h = m.rowHeights[r] || rowBaseHeight(defaultSizePx);
					if (rest < h) return r + rest / h;
					rest -= h;
					r++;
				}
			};
			for (const img of m.floatingImages) {
				const imageFile = this.app.vault.getAbstractFileByPath(img.path);
				const excelExt = excelImageType(imageFile);
				if (!(imageFile instanceof TFile) || !excelExt) {
					skippedImages++;
					continue;
				}
				const imageId = wb.addImage({
					buffer: await this.app.vault.readBinary(imageFile),
					extension: excelExt,
				});
				ws.addImage(imageId, {
					tl: { col: colAt(img.x), row: rowAt(img.y) },
					ext: { width: img.w, height: img.h },
					editAs: 'absolute',
				});
			}
			return skippedImages;
		}
	}
}

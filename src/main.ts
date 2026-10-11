import {
	debounce,
	Menu,
	Notice,
	Platform,
	Plugin,
	TFile,
	TextFileView,
	WorkspaceLeaf,
	normalizePath,
	setIcon,
} from 'obsidian';
import Handsontable from 'handsontable';
import 'handsontable/styles/handsontable.min.css';
import 'handsontable/styles/ht-theme-main.min.css';
import './plugin.css';
import ExcelJS from 'exceljs';
import {
	DEFAULT_SETTINGS,
	MAX_SAVED_COLORS,
	PluginSettings,
	SettingsTab,
	sanitizeSettings,
} from './settings';
import {
	CellBorders,
	CellFormat,
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
	rowBaseHeight,
	serializeModel,
} from './model';
import {
	FormulaFunctionInfo,
	buildFunctionCall,
	evaluateFormula,
	normalizeSeparators,
	referenceSlot,
	shiftFormulaRows,
} from './formulas';
import { ImportedSheet, parseCsv, parseXlsx } from './xlsx';
import {
	ColumnFilterModal,
	FormulaAutocomplete,
	FunctionPickerModal,
	ImportDirectionChoice,
	ImportDirectionModal,
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
const MAX_IMAGE_CELL_PX = 240;
/** Limits (unzoomed px) for a cell that is resized by dragging a corner of its image. */
const MIN_IMAGE_RESIZE_PX = 24;
const MAX_IMAGE_RESIZE_PX = 2000;
/** The corners of an image cell that have a resize handle (physical sides). */
const IMAGE_CORNERS = ['nw', 'ne', 'sw', 'se'] as const;
/** Extra px added to measured text when a column is widened to fit it. */
const EXPAND_PADDING = 18;
const MAX_EXPAND_WIDTH = 600;
/** Horizontal padding inside a cell, used when measuring text. */
const CELL_PADDING = 8;

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

/** Draws cell borders as background layers (reliable inside Handsontable's table cells). */
function applyBorderStyles(st: CSSStyleDeclaration, b?: CellBorders): void {
	const color = b?.color ?? 'var(--text-normal)';
	const layer = `linear-gradient(${color}, ${color})`;
	const images: string[] = [];
	const sizes: string[] = [];
	const positions: string[] = [];
	if (b?.top) {
		images.push(layer);
		sizes.push('100% 1px');
		positions.push('0 0');
	}
	if (b?.right) {
		images.push(layer);
		sizes.push('1px 100%');
		positions.push('100% 0');
	}
	if (b?.bottom) {
		images.push(layer);
		sizes.push('100% 1px');
		positions.push('0 100%');
	}
	if (b?.left) {
		images.push(layer);
		sizes.push('1px 100%');
		positions.push('0 0');
	}
	st.backgroundImage = images.join(', ');
	st.backgroundSize = sizes.join(', ');
	st.backgroundPosition = positions.join(', ');
	st.backgroundRepeat = images.length > 0 ? 'no-repeat' : '';
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

	async createNewSheet(): Promise<void> {
		try {
			const active = this.app.workspace.getActiveFile();
			const folder = this.app.fileManager.getNewFileParent(
				active?.path ?? '',
			);
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
				serializeModel(
					createDefaultModel(this.settings, defaultSheetRtl()),
				),
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
	/** As JSON: these records are changed in place. */
	widenRecords: string;
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
		a.widenRecords === b.widenRecords &&
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
	heights: (number | null)[];
}

export class SheetView extends TextFileView {
	private plugin: RtlSpreadsheetPlugin;
	private model: SheetModel;
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

	/** View zoom (0.5-3). Scales column widths, row heights and fonts; it is not saved in the file. */
	private zoom = 1;
	private pinch: { startDistance: number; startZoom: number } | null = null;
	private pendingZoom = 1;
	private zoomFrame = 0;
	/** row -> columns that have a format (rebuilt lazily; used to compute row heights). */
	private rowIndex: Map<number, number[]> | null = null;
	private fontFamilyCache: string | null = null;
	private measureCtx: CanvasRenderingContext2D | null = null;

	constructor(leaf: WorkspaceLeaf, plugin: RtlSpreadsheetPlugin) {
		super(leaf);
		this.plugin = plugin;
		this.model = createDefaultModel(plugin.settings, defaultSheetRtl());
		this.appliedSettings = { ...plugin.settings };

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
		this.gridEl = this.contentEl.createDiv({ cls: 'rtl-sheet-grid' });

		this.buildFormulaBar();
		this.buildToolbar();
		this.buildFindBar();
		// Capture phase, so the keys are seen before Handsontable and whatever has the focus.
		this.registerDomEvent(document, 'keydown', (e) => this.onViewKeyDown(e), {
			capture: true,
		});
		this.applyToolbarScale();
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
		return serializeModel(this.model);
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
			const next = this.normalizeModel(parsed);
			this.diskData = data;
			if (clear) this.filters.clear(); // another file: its columns mean something else
			// Same file, new content: update the live grid so selection and scroll stay put.
			if (clear || !this.applyModelInPlace(next)) {
				// Drop the old grid first, rebuildGrid() would copy its stale cells into the new model.
				this.destroyGrid();
				this.model = next;
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
		this.autocomplete.destroy();
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
		this.rowIndex = null;
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
		this.applyToolbarScale();
		this.renderPalette('fill');
		this.renderPalette('text');

		if (
			prev.defaultFontSize !== s.defaultFontSize ||
			prev.enableFormulas !== s.enableFormulas
		) {
			this.invalidateRowSizes();
			if (this.hot && !this.hot.isDestroyed) this.hot.render();
		}
	}

	/* ---------------- Model helpers ---------------- */

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
			version: typeof r.version === 'number' ? r.version : 2,
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
			merges: Array.isArray(r.merges) ? r.merges : [],
			colWidths: Array.isArray(r.colWidths) ? r.colWidths : [],
			rowHeights: Array.isArray(r.rowHeights) ? r.rowHeights : [],
			rtl: typeof r.rtl === 'boolean' ? r.rtl : defaultSheetRtl(),
		};
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
		this.rowIndex = null;

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

	/** `at` is the cell the value lives in (ROW() and COLUMN() without an argument report it). */
	private resolveRaw(
		raw: unknown,
		depth: number,
		coerce: boolean,
		at?: { row: number; col: number },
	): Scalar {
		if (raw === undefined || raw === null) return null;
		if (typeof raw === 'string') {
			if (
				this.plugin.settings.enableFormulas &&
				raw.length > 1 &&
				raw.startsWith('=')
			) {
				if (depth > 40) return '#CYCLE!';
				return evaluateFormula(
					raw,
					(r, c) => {
						const hot: Handsontable.Core | null =
							this.gridForCalc ?? this.hot;
						const cellRaw =
							hot && !hot.isDestroyed
								? hot.getSourceDataAtCell(r, c)
								: null;
						return this.resolveRaw(cellRaw, depth + 1, true, {
							row: r,
							col: c,
						});
					},
					at,
				);
			}
			if (coerce && NUMERIC_RE.test(raw)) return Number(raw);
			return raw;
		}
		if (typeof raw === 'number' || typeof raw === 'boolean') return raw;
		return null;
	}

	/* ---------------- Grid lifecycle ---------------- */

	destroyGrid(): void {
		this.clearRefPick();
		this.autocomplete.detach();
		if (this.hot && !this.hot.isDestroyed) this.hot.destroy();
		this.hot = null;
		this.gridForCalc = null;
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

		this.fontFamilyCache = null;
		this.rowIndex = null;
		const zoom = this.zoom;
		const toSizes = (arr: (number | null)[]): (number | undefined)[] =>
			arr.map((v) => (v ? Math.round(v * zoom) : undefined));

		try {
			const options: Handsontable.GridSettings = {
				data: m.data,
				licenseKey: 'non-commercial-and-evaluation',
				themeName: handsontableThemeName(),
				layoutDirection: m.rtl ? 'rtl' : 'ltr',
				rowHeaders: true,
				colHeaders: true,
				width: '100%',
				height: '100%',
				manualColumnResize: toSizes(m.colWidths) as number[],
				manualRowResize: toSizes(m.rowHeights) as number[],
				// Functions (instead of auto-sizing) so the user decides column widths, and row heights
				// always follow the font size; manual resizing still overrides both.
				colWidths: () => Math.round(DEFAULT_COL_WIDTH * this.zoom),
				rowHeights: (row: number) =>
					Math.round(this.computeRowHeight(row) * this.zoom),
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
				afterChange: (changes, source) => {
					if (source === 'loadData' || this.loading) return;
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
					if (this.plugin.settings.enableFormulas && this.hot)
						this.hot.render();
					this.markDirty();
					this.refreshToolbarState();
				},
				afterSelectionEnd: () => this.onSelectionChanged(),
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
					this.markDirty();
				},
				afterRowResize: (newSize, row) => {
					while (this.model.rowHeights.length <= row)
						this.model.rowHeights.push(null);
					this.model.rowHeights[row] = Math.round(
						newSize / this.zoom,
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
			};

			this.hot = new Handsontable(this.gridEl, options);
			// The first render ran before `this.hot` existed; render again so formulas see all cells.
			if (this.plugin.settings.enableFormulas) this.hot.render();
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

	/** Row heights are computed from the fonts; call after anything that can change them. */
	private invalidateRowSizes(): void {
		this.rowIndex = null;
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
		const defaultSize = this.plugin.settings.defaultFontSize;
		let height = rowBaseHeight(defaultSize);
		const cols = this.getRowIndex().get(row);
		if (!cols) return height;

		const hot: Handsontable.Core | null = this.gridForCalc ?? this.hot;
		for (const col of cols) {
			const f = this.model.cellFormats[key(row, col)];
			if (!f) continue;
			const mg = this.mergeAt(hot, row, col);
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

	private fontFamily(): string {
		if (!this.fontFamilyCache) {
			this.fontFamilyCache =
				getComputedStyle(this.gridEl).fontFamily || 'sans-serif';
		}
		return this.fontFamilyCache;
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
		);
		const picked = this.refPick?.range;
		td.classList.toggle(
			'rtl-sheet-ref-pick',
			!!picked &&
				row >= picked.r1 &&
				row <= picked.r2 &&
				col >= picked.c1 &&
				col <= picked.c2,
		);
		const found = this.findMatches[this.findIndex];
		td.classList.toggle(
			'rtl-sheet-find-hit',
			this.findHits.has(key(row, col)),
		);
		td.classList.toggle(
			'rtl-sheet-find-current',
			!!found && found.row === row && found.col === col,
		);
		td.querySelectorAll('.rtl-sheet-spill-eraser').forEach((el) =>
			el.remove(),
		);

		const f = this.model.cellFormats[key(row, col)];
		const imagePath = this.model.images[key(row, col)];
		const rtl = this.model.rtl;
		const wrap = f?.wrap ?? 'overflow';
		const st = td.style;

		st.fontSize = `${Math.round((f?.fontSize ?? this.plugin.settings.defaultFontSize) * this.zoom * 10) / 10}px`;
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
		applyBorderStyles(st, f?.borders);

		if (imagePath) {
			this.renderImage(td, imagePath);
			const active = this.formulaCell;
			td.classList.toggle(
				'rtl-sheet-img-active',
				!!active && active.row === row && active.col === col,
			);
		} else if (f?.rotate) {
			this.renderRotated(td, text, f.rotate);
		} else if (wrap === 'overflow') {
			const spill = this.spillInfo(instance, row, col, text, isNumber, f);
			if (spill) {
				st.overflow = 'visible';
				st.position = 'relative';
				st.zIndex = '1';
				// Paint the overflow only as far as the empty cells it spills into.
				st.clipPath = spill.visualRight
					? `inset(0 -${spill.extra}px 0 0)`
					: `inset(0 0 0 -${spill.extra}px)`;
				// Hide the vertical grid lines the text crosses (like Google Sheets). The erasers sit
				// behind the text (z-index -1 inside this cell) but above the cell borders.
				const eraserColor =
					f?.bgColor ??
					'var(--ht-cell-background-color, var(--background-primary))';
				for (const offset of spill.lines) {
					const eraser = td.createDiv({
						cls: 'rtl-sheet-spill-eraser',
					});
					eraser.style.background = eraserColor;
					if (spill.visualRight)
						eraser.style.left = `${offset - 2}px`;
					else eraser.style.right = `${offset - 2}px`;
				}
			}
		}
	};

	/**
	 * Text that is wider than its cell spills into the empty cells next to it (like Excel).
	 * Numbers and numeric formula results never spill: only text, and only as far as needed.
	 */
	private spillInfo(
		hot: Handsontable.Core,
		row: number,
		col: number,
		text: string,
		isNumber: boolean,
		f: CellFormat | undefined,
	): { visualRight: boolean; extra: number; lines: number[] } | null {
		if (!text || isNumber) return null;

		// A merged cell is as wide as all the columns it covers.
		const colspan = this.mergeAt(hot, row, col)?.colspan ?? 1;
		let own = 0;
		for (let c = col; c < col + colspan; c++) own += hot.getColWidth(c);
		const needed =
			this.measureTextWidth(text, f) * this.zoom + 2 * CELL_PADDING;
		if (needed <= own) return null; // it fits: nothing crosses the cell border

		const visualRight = !/[\u0590-\u08FF]/.test(text);
		// In an RTL sheet, column index + 1 is the visually-left neighbour.
		const delta = visualRight === !this.model.rtl ? 1 : -1;

		const lines: number[] = [own]; // the cell's own border is crossed
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
				this.model.images[key(row, c)]
			)
				break;
			const w = hot.getColWidth(c);
			extra += w;
			reach += w;
			if (needed <= reach) break;
			lines.push(reach);
		}
		return extra > 0 ? { visualRight, extra, lines } : null;
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
			(hot.getRowHeight(row) ?? 0) / this.zoom ||
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
			const dy = ((ev.clientY - startY) * signY) / this.zoom;
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

	/** Mouse and touch events that follow a press on an image handle must not reach the grid. */
	private keepHandlePress(e: Event): void {
		if (
			e.target instanceof HTMLElement &&
			e.target.closest('.rtl-sheet-img-handle')
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
						icon: 'grid-3x3',
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

	private showMenu(anchor: HTMLElement, entries: MenuEntry[]): void {
		const menu = new Menu();
		for (const entry of entries) {
			menu.addItem((item) => {
				item.setTitle(entry.title)
					.setIcon(entry.icon)
					.onClick(entry.onClick);
				if (entry.checked !== undefined) item.setChecked(entry.checked);
			});
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

	private onSelectionChanged(): void {
		this.clearRefPick();
		const first = this.getSelectedRanges()[0];
		this.formulaCell = first ? { row: first.r1, col: first.c1 } : null;
		this.markActiveImageCell();

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
			widenRecords: JSON.stringify(m.widenRecords ?? {}),
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
		m.widenRecords = JSON.parse(state.widenRecords) as NonNullable<
			SheetModel['widenRecords']
		>;
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

	private openFind(): void {
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
	private applyBorders(mode: BorderMode): void {
		const rtl = this.model.rtl;
		this.applyFormat(this.requireSelection(), (f, r, c, rg) => {
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

	/* ---------------- Zoom ---------------- */

	/** Zoom scales widths, heights and fonts together; model values stay at 100%. */
	private applyZoom(next: number): void {
		const z = Math.min(
			MAX_ZOOM,
			Math.max(MIN_ZOOM, Math.round(next * 20) / 20),
		);
		if (z === this.zoom) return;
		this.zoom = z;

		const hot = this.hot;
		if (hot && !hot.isDestroyed) {
			const cols = hot.getPlugin('manualColumnResize') as unknown as {
				setManualSize(c: number, w: number): void;
			};
			const rows = hot.getPlugin('manualRowResize') as unknown as {
				setManualSize(r: number, h: number): void;
			};
			this.model.colWidths.forEach((w, i) => {
				if (w) cols.setManualSize(i, Math.round(w * z));
			});
			this.model.rowHeights.forEach((h, i) => {
				if (h) rows.setManualSize(i, Math.round(h * z));
			});
			hot.render();
			hot.refreshDimensions();
		}
		this.ui.zoomLabel?.setText(`${Math.round(z * 100)}%`);
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
		plugin.setManualSize(row, Math.round(height * this.zoom));
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
			await this.setCellImage(path, cell);
		} catch (e) {
			this.plugin.reportError('Could not add the image', e);
		}
	}

	/** Ctrl+V / long-press paste of an image straight into the selected cell. */
	private onPaste(e: ClipboardEvent): void {
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
		const file = this.app.vault.getAbstractFileByPath(path);
		if (!(file instanceof TFile)) return;
		const url = this.app.vault.getResourcePath(file);
		const dims = await new Promise<{ w: number; h: number } | null>(
			(resolve) => {
				const img = new Image();
				img.onload = () =>
					resolve({ w: img.naturalWidth, h: img.naturalHeight });
				img.onerror = () => resolve(null);
				img.src = url;
			},
		);
		if (!dims || dims.w <= 0 || dims.h <= 0) return;
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
		if (e.key !== 'Delete' && e.key !== 'Backspace') return;
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const editor = hot.getActiveEditor() as unknown as
			| { isOpened?: () => boolean }
			| undefined;
		if (editor?.isOpened?.()) return; // the user is editing text, not clearing cells
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

	private askImportDirection(
		detectedRtl: boolean,
	): Promise<ImportDirectionChoice | null> {
		return new Promise((resolve) => {
			new ImportDirectionModal(this.app, detectedRtl, resolve).open();
		});
	}

	private async importBuffer(name: string, data: ArrayBuffer): Promise<void> {
		try {
			const ext = name.split('.').pop()?.toLowerCase() ?? '';
			const imported =
				ext === 'csv'
					? parseCsv(decodeText(data))
					: await parseXlsx(data);

			const choice = await this.askImportDirection(imported.rtl);
			if (choice === null) return; // dialog dismissed: cancel the import
			if (choice !== 'keep') this.setDirection(choice === 'rtl');
			this.applyImported(imported);
		} catch (e) {
			this.plugin.reportError('Import failed', e);
		}
	}

	/** Writes imported data into this sheet, starting at the selected cell (or A1). */
	private applyImported(imp: ImportedSheet): void {
		const hot = this.hot;
		if (!hot || hot.isDestroyed) return;
		const rows = imp.data.length;
		const cols = Math.max(0, ...imp.data.map((r) => r.length));
		if (rows === 0 || cols === 0) {
			new Notice('Nothing to import.');
			return;
		}

		const anchor = this.getSelectedRanges()[0];
		const ar = anchor?.r1 ?? 0;
		const ac = anchor?.c1 ?? 0;

		const missingRows = ar + rows - hot.countRows();
		if (missingRows > 0)
			hot.alter('insert_row_below', hot.countRows() - 1, missingRows);
		const missingCols = ac + cols - hot.countCols();
		if (missingCols > 0)
			hot.alter('insert_col_end', hot.countCols() - 1, missingCols);

		hot.populateFromArray(ar, ac, imp.data);

		for (const [k, f] of Object.entries(imp.formats)) {
			const parts = k.split(',').map(Number);
			const r = parts[0] ?? 0;
			const c = parts[1] ?? 0;
			this.model.cellFormats[key(r + ar, c + ac)] = f;
		}

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

		this.invalidateRowSizes();
		hot.render();
		this.markDirty();
		new Notice(
			`Imported ${rows} rows × ${cols} columns starting at ${indexToColumnLabel(ac)}${ar + 1}.`,
		);
	}

	/* ---------------- Export ---------------- */

	private toExcelValue(raw: unknown, row: number, col: number): ExcelJS.CellValue {
		if (typeof raw === 'string') {
			if (
				this.plugin.settings.enableFormulas &&
				raw.length > 1 &&
				raw.startsWith('=')
			) {
				const result = this.resolveRaw(raw, 0, false, { row, col });
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
			const m = this.model;
			const defaultSizePx = this.plugin.settings.defaultFontSize;

			const wb = new ExcelJS.Workbook();
			const ws = wb.addWorksheet(sanitizeSheetName(file.basename), {
				views: [{ rightToLeft: m.rtl }],
			});

			m.data.forEach((rowArr, r) => {
				rowArr.forEach((raw, c) => {
					const f = m.cellFormats[key(r, c)];
					const hasValue =
						raw !== null && raw !== undefined && raw !== '';
					if (!hasValue && !f) return;

					const cell = ws.getCell(r + 1, c + 1);
					let isNumber = false;
					if (hasValue) {
						const value = this.toExcelValue(raw, r, c);
						cell.value = value;
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

			let skippedImages = 0;
			for (const [k, path] of Object.entries(m.images)) {
				const imageFile = this.app.vault.getAbstractFileByPath(path);
				if (!(imageFile instanceof TFile)) {
					skippedImages++;
					continue;
				}
				const ext = imageFile.extension.toLowerCase();
				const excelExt =
					ext === 'png'
						? 'png'
						: ext === 'jpg' || ext === 'jpeg'
							? 'jpeg'
							: ext === 'gif'
								? 'gif'
								: null;
				if (!excelExt) {
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

			new Notice(
				skippedImages > 0
					? `Exported to ${path} (${skippedImages} image(s) skipped: only png, jpg and gif are supported in .xlsx)`
					: `Exported to ${path}`,
			);
		} catch (e) {
			this.plugin.reportError('Excel export failed', e);
		}
	}
}

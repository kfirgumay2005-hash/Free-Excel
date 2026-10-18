import {
	App,
	FuzzySuggestModal,
	Modal,
	Setting,
	SuggestModal,
	TFile,
	setIcon,
} from 'obsidian';
import { EMOJI_CATEGORIES, EmojiInfo, searchEmojis } from './emoji';
import {
	FORMULA_FUNCTIONS,
	FormulaFunctionInfo,
	buildFunctionCall,
} from './formulas';
import {
	CellComment,
	ColumnFilter,
	DEFAULT_COMMENT_COLOR,
	FILTER_CONDITIONS,
	FilterCondition,
} from './model';

type TextField = HTMLInputElement | HTMLTextAreaElement;

/** After one of these characters a function name is expected, so the full list is offered. */
const FUNCTION_EXPECTED_AFTER = /[=(+\-*/^&,;<>]\s*$/;
const MAX_ITEMS_EMPTY_PREFIX = 40;
const MAX_ITEMS = 12;

/**
 * Excel-style function suggestions for a text field (the cell editor or the
 * formula bar). The popup is touch friendly: tap an item to insert it.
 * Keyboard: ArrowUp/ArrowDown to move, Tab (or Enter while a function name is being
 * typed) to insert it with example arguments, Escape to close.
 */
export class FormulaAutocomplete {
	private field: TextField | null = null;
	private popup: HTMLElement | null = null;
	private items: FormulaFunctionInfo[] = [];
	private active = 0;
	private prefixLength = 0;
	private suppressNextInput = false;
	/** True after the user moved through the list with the arrow keys. */
	private navigated = false;

	attach(field: TextField): void {
		if (this.field === field) return;
		this.detach();
		this.field = field;
		// Typed as HTMLElement so the 'keydown' overload resolves (a union of input/textarea does not).
		const el: HTMLElement = field;
		el.addEventListener('input', this.onInput);
		el.addEventListener('keydown', this.onKeyDown, true);
		el.addEventListener('blur', this.onBlur);
		this.refresh();
	}

	detach(): void {
		const field = this.field;
		if (field) {
			const el: HTMLElement = field;
			el.removeEventListener('input', this.onInput);
			el.removeEventListener('keydown', this.onKeyDown, true);
			el.removeEventListener('blur', this.onBlur);
		}
		this.field = null;
		this.hide();
	}

	destroy(): void {
		this.detach();
	}

	private readonly onInput = (): void => {
		if (this.suppressNextInput) {
			this.suppressNextInput = false;
			this.hide();
			return;
		}
		this.refresh();
	};

	private readonly onBlur = (): void => {
		this.hide();
	};

	private readonly onKeyDown = (e: KeyboardEvent): void => {
		if (!this.popup || this.items.length === 0) return;
		if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
			e.preventDefault();
			e.stopImmediatePropagation();
			this.navigated = true;
			const n = this.items.length;
			this.active =
				(this.active + (e.key === 'ArrowDown' ? 1 : n - 1)) % n;
			this.render();
		} else if (
			e.key === 'Tab' ||
			(e.key === 'Enter' && !e.shiftKey && this.enterPicksItem())
		) {
			// Enter while a function name is half-typed (or right after a lone "=", or after
			// choosing with the arrow keys) inserts the function instead of committing an
			// incomplete formula such as "=SUM" or "=".
			e.preventDefault();
			e.stopImmediatePropagation();
			this.choose(this.active);
		} else if (e.key === 'Escape') {
			this.hide();
		}
	};

	private enterPicksItem(): boolean {
		return (
			this.prefixLength > 0 ||
			this.navigated ||
			this.field?.value.trim() === '='
		);
	}

	/** Recomputes the suggestions for the text before the caret. */
	private refresh(): void {
		const field = this.field;
		if (!field) return;
		const value = field.value;
		this.navigated = false;
		if (!value.startsWith('=')) {
			this.hide();
			return;
		}
		const caret = field.selectionStart ?? value.length;
		const before = value.slice(0, caret);

		// Inside a string literal: no suggestions.
		if ((before.match(/"/g) ?? []).length % 2 === 1) {
			this.hide();
			return;
		}

		const m = /([A-Za-z][A-Za-z0-9_.]*)$/.exec(before);
		if (m && m[1]) {
			const prefix = m[1];
			if (/^[A-Za-z]{1,3}[0-9]+$/.test(prefix)) {
				this.hide(); // looks like a cell reference such as A1
				return;
			}
			const upper = prefix.toUpperCase();
			this.items = FORMULA_FUNCTIONS.filter((f) =>
				f.name.startsWith(upper),
			)
				.sort(
					(a, b) =>
						Number(b.name === upper) - Number(a.name === upper),
				)
				.slice(0, MAX_ITEMS);
			this.prefixLength = prefix.length;
		} else if (FUNCTION_EXPECTED_AFTER.test(before)) {
			this.items = FORMULA_FUNCTIONS.slice(0, MAX_ITEMS_EMPTY_PREFIX);
			this.prefixLength = 0;
		} else {
			this.hide();
			return;
		}

		if (this.items.length === 0) {
			this.hide();
			return;
		}
		this.active = 0;
		this.render();
	}

	private render(): void {
		const field = this.field;
		if (!field) return;
		const doc = field.ownerDocument;

		if (!this.popup) {
			this.popup = doc.body.createDiv({ cls: 'rtl-sheet-autocomplete' });
		}
		const popup = this.popup;
		popup.empty();

		this.items.forEach((fn, index) => {
			const row = popup.createDiv({
				cls:
					index === this.active
						? 'rtl-sheet-ac-item is-active'
						: 'rtl-sheet-ac-item',
			});
			row.createSpan({ cls: 'rtl-sheet-ac-name', text: fn.syntax });
			row.createSpan({ cls: 'rtl-sheet-ac-desc', text: fn.description });
			// Hover moves the highlight, so what Enter would insert matches what is highlighted.
			row.addEventListener('mouseenter', () => this.setActive(index));
			// Handsontable closes (and saves) the cell editor when it sees a press outside the
			// grid. Stop these events here so a click on the list does not save a bare "=".
			for (const type of [
				'mousedown',
				'pointerdown',
				'touchstart',
				'mouseup',
				'pointerup',
				'touchend',
			]) {
				row.addEventListener(type, (e) => e.stopPropagation());
			}
			// Keep focus in the text field when tapping/clicking an item.
			row.addEventListener('mousedown', (e) => e.preventDefault());
			row.addEventListener('click', () => this.choose(index));
		});

		this.position(field, popup);
		popup.children[this.active]?.scrollIntoView({ block: 'nearest' });
	}

	private setActive(index: number): void {
		this.active = index;
		Array.from(this.popup?.children ?? []).forEach((child, i) => {
			child.classList.toggle('is-active', i === index);
		});
	}

	private position(field: TextField, popup: HTMLElement): void {
		const win = field.ownerDocument.defaultView ?? window;
		const vv = win.visualViewport;
		const viewTop = vv?.offsetTop ?? 0;
		const viewHeight = vv?.height ?? win.innerHeight;
		const viewWidth = vv?.width ?? win.innerWidth;

		const rect = field.getBoundingClientRect();
		popup.style.maxHeight = `${Math.max(120, Math.min(280, Math.round(viewHeight * 0.4)))}px`;
		popup.style.minWidth = `${Math.min(280, viewWidth - 8)}px`;

		const width = popup.offsetWidth;
		const height = popup.offsetHeight;
		const left = Math.max(4, Math.min(rect.left, viewWidth - width - 4));
		let top = rect.bottom + 4;
		if (top + height > viewTop + viewHeight - 4) {
			top = Math.max(viewTop + 4, rect.top - height - 4);
		}
		popup.style.left = `${left}px`;
		popup.style.top = `${top}px`;
	}

	private choose(index: number): void {
		const field = this.field;
		const fn = this.items[index];
		if (!field || !fn) return;

		const caret = field.selectionStart ?? field.value.length;
		const start = Math.max(0, caret - this.prefixLength);
		const call = buildFunctionCall(fn);
		field.value = `${field.value.slice(0, start)}${call.text}${field.value.slice(caret)}`;
		// The example arguments end up selected, so typing replaces them.
		field.setSelectionRange(start + call.selectFrom, start + call.selectTo);
		field.focus();

		this.suppressNextInput = true;
		field.dispatchEvent(new Event('input', { bubbles: true }));
		this.hide();
	}

	private hide(): void {
		this.popup?.remove();
		this.popup = null;
		this.items = [];
	}
}

/** Searchable list of all known functions (works well on phones). */
export class FunctionPickerModal extends SuggestModal<FormulaFunctionInfo> {
	private readonly onChoose: (fn: FormulaFunctionInfo) => void;

	constructor(app: App, onChoose: (fn: FormulaFunctionInfo) => void) {
		super(app);
		this.onChoose = onChoose;
		this.setPlaceholder('Search functions, e.g. SUM or "average"');
	}

	getSuggestions(query: string): FormulaFunctionInfo[] {
		const q = query.trim().toLowerCase();
		if (!q) return FORMULA_FUNCTIONS;
		const starts = FORMULA_FUNCTIONS.filter((f) =>
			f.name.toLowerCase().startsWith(q),
		);
		const contains = FORMULA_FUNCTIONS.filter(
			(f) =>
				!f.name.toLowerCase().startsWith(q) &&
				(f.name.toLowerCase().includes(q) ||
					f.description.toLowerCase().includes(q)),
		);
		return [...starts, ...contains];
	}

	renderSuggestion(fn: FormulaFunctionInfo, el: HTMLElement): void {
		el.createDiv({ cls: 'rtl-sheet-fn-syntax', text: fn.syntax });
		el.createEl('small', { text: fn.description });
	}

	onChooseSuggestion(fn: FormulaFunctionInfo): void {
		this.onChoose(fn);
	}
}

/** Pick a file from the vault by extension. */
export class VaultFilePickerModal extends FuzzySuggestModal<TFile> {
	private readonly extensions: string[];
	private readonly onChoose: (file: TFile) => void;

	constructor(
		app: App,
		extensions: string[],
		placeholder: string,
		onChoose: (file: TFile) => void,
	) {
		super(app);
		this.extensions = extensions.map((e) => e.toLowerCase());
		this.onChoose = onChoose;
		this.setPlaceholder(placeholder);
	}

	getItems(): TFile[] {
		return this.app.vault
			.getFiles()
			.filter((f) => this.extensions.includes(f.extension.toLowerCase()));
	}

	getItemText(file: TFile): string {
		return file.path;
	}

	onChooseItem(file: TFile): void {
		this.onChoose(file);
	}
}

export type ImportDirectionChoice = 'rtl' | 'ltr' | 'keep';

export interface ImportChoice {
	direction: ImportDirectionChoice;
	/** False to import the values the file stored for its formulas instead of the formulas. */
	keepFormulas: boolean;
}

export interface ImportPrompt {
	detectedRtl: boolean;
	/** Label of the cell the table starts at, e.g. "B3". */
	startCell: string;
	/** True when the file has formulas together with their values, so there is a choice to make. */
	hasFormulas: boolean;
}

/**
 * Asks how the imported table should be laid out (right-to-left, left-to-right, or as the
 * sheet already is) and whether its formulas are kept or replaced by their values.
 */
export class ImportDirectionModal extends Modal {
	private readonly prompt: ImportPrompt;
	private readonly onChoose: (choice: ImportChoice | null) => void;
	private keepFormulas = true;
	private chosen = false;

	constructor(
		app: App,
		prompt: ImportPrompt,
		onChoose: (choice: ImportChoice | null) => void,
	) {
		super(app);
		this.prompt = prompt;
		this.onChoose = onChoose;
	}

	onOpen(): void {
		const { contentEl } = this;
		const { detectedRtl, startCell, hasFormulas } = this.prompt;
		contentEl.empty();
		contentEl.createEl('h3', { text: 'Import table' });
		contentEl.createEl('p', {
			text: `The table is placed starting at the selected cell (${startCell}).`,
		});

		if (hasFormulas) {
			new Setting(contentEl)
				.setName('Formulas')
				.setDesc(
					`Formulas are adjusted so that A1 of the imported table becomes ${startCell}.`,
				)
				.addDropdown((dropdown) => {
					dropdown
						.addOption('formulas', 'Keep formulas')
						.addOption('values', 'Values only')
						.setValue('formulas')
						.onChange((value) => {
							this.keepFormulas = value === 'formulas';
						});
				});
		}

		contentEl.createEl('p', {
			text: `The imported table looks ${detectedRtl ? 'right-to-left' : 'left-to-right'}. How should the sheet be laid out?`,
		});

		const buttons = contentEl.createDiv({
			cls: 'rtl-sheet-import-choices',
		});
		const add = (
			label: string,
			choice: ImportDirectionChoice,
			primary: boolean,
		): void => {
			const btn = buttons.createEl('button', {
				text: label,
				cls: primary ? 'mod-cta' : '',
			});
			btn.addEventListener('click', () => {
				this.chosen = true;
				this.close();
				this.onChoose({
					direction: choice,
					keepFormulas: this.keepFormulas,
				});
			});
		};
		add('Right to left', 'rtl', detectedRtl);
		add('Left to right', 'ltr', !detectedRtl);
		add("Keep this sheet's direction", 'keep', false);
	}

	onClose(): void {
		this.contentEl.empty();
		if (!this.chosen) this.onChoose(null); // dismissed: cancel the import
	}
}

/** Most values listed at once in the filter dialog; the search box narrows longer lists. */
const MAX_FILTER_VALUES = 500;

/** Picks which rows stay visible by the values of one column: a condition, a list of values, or both. */
export class ColumnFilterModal extends Modal {
	private readonly heading: string;
	private readonly values: string[];
	private readonly onApply: (filter: ColumnFilter | null) => void;
	private readonly checked: Set<string>;
	private condition: FilterCondition;
	private operand: string;
	private search = '';
	private listEl: HTMLElement | null = null;

	constructor(
		app: App,
		heading: string,
		values: string[],
		current: ColumnFilter | null,
		onApply: (filter: ColumnFilter | null) => void,
	) {
		super(app);
		this.heading = heading;
		this.values = values;
		this.onApply = onApply;
		const allowed = current?.values;
		this.checked = new Set(
			allowed ? values.filter((v) => allowed.has(v)) : values,
		);
		this.condition = current?.condition ?? 'none';
		this.operand = current?.operand ?? '';
	}

	onOpen(): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.addClass('rtl-sheet-filter-modal');
		contentEl.createEl('h3', { text: this.heading });

		new Setting(contentEl)
			.setName('Condition')
			.addDropdown((dropdown) => {
				for (const c of FILTER_CONDITIONS)
					dropdown.addOption(c.value, c.label);
				dropdown.setValue(this.condition).onChange((value) => {
					this.condition = value as FilterCondition;
				});
			})
			.addText((text) => {
				text.inputEl.setAttr('dir', 'auto');
				text.setPlaceholder('Value')
					.setValue(this.operand)
					.onChange((value) => {
						this.operand = value;
					});
			});

		const tools = contentEl.createDiv({ cls: 'rtl-sheet-filter-tools' });
		const searchEl = tools.createEl('input', {
			type: 'text',
			attr: { placeholder: 'Search values', dir: 'auto' },
		});
		searchEl.addEventListener('input', () => {
			this.search = searchEl.value.toLowerCase();
			this.renderList();
		});
		const setAll = (on: boolean): void => {
			for (const v of this.matching()) {
				if (on) this.checked.add(v);
				else this.checked.delete(v);
			}
			this.renderList();
		};
		tools
			.createEl('button', { text: 'Select all' })
			.addEventListener('click', () => setAll(true));
		tools
			.createEl('button', { text: 'Clear' })
			.addEventListener('click', () => setAll(false));

		this.listEl = contentEl.createDiv({ cls: 'rtl-sheet-filter-values' });
		this.renderList();

		const buttons = contentEl.createDiv({
			cls: 'rtl-sheet-import-choices',
		});
		buttons
			.createEl('button', { text: 'Apply', cls: 'mod-cta' })
			.addEventListener('click', () => {
				const all = this.checked.size === this.values.length;
				const active = !all || this.condition !== 'none';
				this.close();
				this.onApply(
					active
						? {
								values: all ? null : new Set(this.checked),
								condition: this.condition,
								operand: this.operand,
							}
						: null,
				);
			});
		buttons
			.createEl('button', { text: 'Remove filter' })
			.addEventListener('click', () => {
				this.close();
				this.onApply(null);
			});
		buttons
			.createEl('button', { text: 'Cancel' })
			.addEventListener('click', () => this.close());
	}

	private matching(): string[] {
		return this.search
			? this.values.filter((v) => v.toLowerCase().includes(this.search))
			: this.values;
	}

	private renderList(): void {
		const list = this.listEl;
		if (!list) return;
		list.empty();
		const matching = this.matching();
		for (const value of matching.slice(0, MAX_FILTER_VALUES)) {
			const row = list.createEl('label', {
				cls: 'rtl-sheet-filter-value',
			});
			const box = row.createEl('input', { type: 'checkbox' });
			box.checked = this.checked.has(value);
			box.addEventListener('change', () => {
				if (box.checked) this.checked.add(value);
				else this.checked.delete(value);
			});
			row.createSpan({
				text: value === '' ? '(Blanks)' : value,
				attr: { dir: 'auto' },
			});
		}
		if (matching.length === 0)
			list.createDiv({
				cls: 'rtl-sheet-filter-note',
				text: 'No values.',
			});
		else if (matching.length > MAX_FILTER_VALUES)
			list.createDiv({
				cls: 'rtl-sheet-filter-note',
				text: `Showing the first ${MAX_FILTER_VALUES} of ${matching.length} values. Search to narrow the list.`,
			});
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

/** Asks for one line of text (such as a new sheet name). `validate` returns an error message or null. */
export class TextPromptModal extends Modal {
	private value: string;

	constructor(
		app: App,
		private readonly heading: string,
		initial: string,
		private readonly validate: (value: string) => string | null,
		private readonly onSubmit: (value: string) => void,
	) {
		super(app);
		this.value = initial;
	}

	onOpen(): void {
		this.setTitle(this.heading);
		const { contentEl } = this;
		contentEl.empty();
		const error = contentEl.createDiv({ cls: 'rtl-sheet-prompt-error' });
		const submit = (): void => {
			const message = this.validate(this.value);
			if (message) {
				error.setText(message);
				return;
			}
			this.close();
			this.onSubmit(this.value);
		};
		new Setting(contentEl).addText((text) => {
			text.setValue(this.value).onChange((value) => {
				this.value = value;
				error.setText('');
			});
			text.inputEl.addClass('rtl-sheet-prompt-input');
			text.inputEl.setAttr('dir', 'auto');
			text.inputEl.addEventListener('keydown', (e) => {
				if (e.key === 'Enter' && !e.isComposing) {
					e.preventDefault();
					submit();
				}
			});
			window.setTimeout(() => text.inputEl.select(), 0);
		});
		new Setting(contentEl)
			.addButton((btn) =>
				btn.setButtonText('Save').setCta().onClick(submit),
			)
			.addButton((btn) =>
				btn.setButtonText('Cancel').onClick(() => this.close()),
			);
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

/** Asks before something that cannot be undone. */
export class ConfirmModal extends Modal {
	constructor(
		app: App,
		private readonly heading: string,
		private readonly message: string,
		private readonly confirmText: string,
		private readonly onConfirm: () => void,
	) {
		super(app);
	}

	onOpen(): void {
		this.setTitle(this.heading);
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('p', { text: this.message });
		new Setting(contentEl)
			.addButton((btn) =>
				btn
					.setButtonText(this.confirmText)
					.setDestructive()
					.onClick(() => {
						this.close();
						this.onConfirm();
					}),
			)
			.addButton((btn) =>
				btn.setButtonText('Cancel').onClick(() => this.close()),
			);
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

/* ------------------------------------------------------------------ */
/* Emoji picker                                                        */
/* ------------------------------------------------------------------ */

/**
 * A popup with a search box and every emoji by category. It lives in <body> like the
 * autocomplete list, and keeps its presses away from Handsontable so a cell being edited
 * stays open.
 */
export class EmojiPicker {
	private popup: HTMLElement | null = null;
	private gridEl: HTMLElement | null = null;
	private searchEl: HTMLInputElement | null = null;
	/** Index into EMOJI_CATEGORIES; -1 = recently used. */
	private category = 0;
	private onPick: ((emoji: string) => void) | null = null;
	private onClose: (() => void) | null = null;
	private recent: string[] = [];

	open(
		anchor: HTMLElement,
		recent: string[],
		onPick: (emoji: string) => void,
		onClose: () => void,
	): void {
		this.close();
		this.recent = recent;
		this.onPick = onPick;
		this.onClose = onClose;
		this.category = recent.length > 0 ? -1 : 0;
		const doc = anchor.ownerDocument;
		const popup = doc.body.createDiv({ cls: 'rtl-sheet-emoji-picker' });
		this.popup = popup;
		for (const type of [
			'mousedown',
			'pointerdown',
			'touchstart',
			'mouseup',
			'pointerup',
			'touchend',
			'keydown',
		])
			popup.addEventListener(type, (e) => e.stopPropagation());

		const search = popup.createEl('input', {
			type: 'text',
			cls: 'rtl-sheet-emoji-search',
			attr: { placeholder: 'Search emojis', dir: 'auto' },
		});
		this.searchEl = search;
		search.addEventListener('input', () => this.renderGrid());
		search.addEventListener('keydown', (e) => {
			if (e.key === 'Escape') this.close();
			else if (e.key === 'Enter') {
				e.preventDefault();
				this.gridEl?.querySelector('button')?.click();
			}
		});

		const tabs = popup.createDiv({ cls: 'rtl-sheet-emoji-tabs' });
		const addTab = (index: number, icon: string, label: string): void => {
			const tab = tabs.createEl('button', {
				cls: 'rtl-sheet-emoji-tab',
				attr: {
					'aria-label': label,
					title: label,
					'data-index': String(index),
				},
			});
			setIcon(tab, icon);
			tab.addEventListener('mousedown', (e) => e.preventDefault());
			tab.addEventListener('click', () => {
				this.category = index;
				search.value = '';
				this.renderGrid();
			});
		};
		if (recent.length > 0) addTab(-1, 'clock', 'Recently used');
		EMOJI_CATEGORIES.forEach((c, i) => addTab(i, c.icon, c.name));

		this.gridEl = popup.createDiv({ cls: 'rtl-sheet-emoji-grid' });
		this.renderGrid();
		this.position(anchor);
		doc.addEventListener('pointerdown', this.onOutside, true);
		// On phones the keyboard would cover the list: search only when the box is tapped.
		if (!doc.body.hasClass('is-mobile')) search.focus();
	}

	close(): void {
		const popup = this.popup;
		if (!popup) return;
		popup.ownerDocument.removeEventListener(
			'pointerdown',
			this.onOutside,
			true,
		);
		popup.remove();
		this.popup = null;
		this.gridEl = null;
		this.searchEl = null;
		const onClose = this.onClose;
		this.onClose = null;
		this.onPick = null;
		onClose?.();
	}

	private readonly onOutside = (e: PointerEvent): void => {
		if (e.target instanceof Node && this.popup?.contains(e.target)) return;
		this.close();
	};

	private renderGrid(): void {
		const grid = this.gridEl;
		if (!grid) return;
		grid.empty();
		const query = this.searchEl?.value.trim() ?? '';
		let list: EmojiInfo[];
		if (query) list = searchEmojis(query);
		else if (this.category < 0)
			list = this.recent.map((char) => ({ char, keywords: '' }));
		else list = EMOJI_CATEGORIES[this.category]?.emojis ?? [];
		this.popup
			?.querySelectorAll<HTMLElement>('.rtl-sheet-emoji-tab')
			.forEach((tab) =>
				tab.toggleClass(
					'is-active',
					!query && tab.dataset.index === String(this.category),
				),
			);
		if (list.length === 0) {
			grid.createDiv({
				cls: 'rtl-sheet-emoji-empty',
				text: 'No emojis found.',
			});
			return;
		}
		for (const e of list) {
			const hint = e.keywords || e.char;
			const btn = grid.createEl('button', {
				cls: 'rtl-sheet-emoji',
				text: e.char,
				attr: { title: hint, 'aria-label': hint },
			});
			btn.addEventListener('mousedown', (ev) => ev.preventDefault());
			btn.addEventListener('click', (ev) => {
				const pick = this.onPick;
				// Shift+click adds several emojis without closing the picker.
				if (!ev.shiftKey) this.close();
				pick?.(e.char);
			});
		}
	}

	private position(anchor: HTMLElement): void {
		const popup = this.popup;
		if (!popup) return;
		const win = anchor.ownerDocument.defaultView ?? window;
		const rect = anchor.getBoundingClientRect();
		const width = Math.min(340, win.innerWidth - 16);
		popup.style.width = `${width}px`;
		const height = popup.offsetHeight;
		const left = Math.max(8, Math.min(rect.left, win.innerWidth - width - 8));
		let top = rect.bottom + 4;
		if (top + height > win.innerHeight - 8)
			top = Math.max(8, rect.top - height - 4);
		popup.style.left = `${left}px`;
		popup.style.top = `${top}px`;
	}
}

/* ------------------------------------------------------------------ */
/* Comments and links                                                  */
/* ------------------------------------------------------------------ */

/** Colors offered for the corner mark of a comment. */
const COMMENT_COLORS = [
	DEFAULT_COMMENT_COLOR,
	'#e03131',
	'#2f9e44',
	'#1971c2',
	'#9c36b5',
	'#868e96',
];

/** Writes or changes the comment of a cell. `onSave(null)` removes it. */
export class CommentModal extends Modal {
	private text: string;
	private color: string;

	constructor(
		app: App,
		private readonly cellLabel: string,
		private readonly current: CellComment | null,
		private readonly onSave: (comment: CellComment | null) => void,
	) {
		super(app);
		this.text = current?.text ?? '';
		this.color = current?.color ?? DEFAULT_COMMENT_COLOR;
	}

	onOpen(): void {
		this.setTitle(
			this.current
				? `Edit comment on ${this.cellLabel}`
				: `Add comment to ${this.cellLabel}`,
		);
		const { contentEl } = this;
		contentEl.empty();
		const area = contentEl.createEl('textarea', {
			cls: 'rtl-sheet-comment-input',
			attr: { placeholder: 'Comment', dir: 'auto', rows: '5' },
		});
		area.value = this.text;
		area.addEventListener('input', () => {
			this.text = area.value;
		});
		area.addEventListener('keydown', (e) => {
			if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
				e.preventDefault();
				this.save();
			}
		});

		const colorRow = new Setting(contentEl)
			.setName('Corner color')
			.setDesc('The mark in the top right corner of the cell.');
		const swatches = colorRow.controlEl.createDiv({
			cls: 'rtl-sheet-comment-colors',
		});
		const picker = colorRow.controlEl.createEl('input', { type: 'color' });
		picker.value = this.color;
		const paint = (): void => {
			swatches
				.querySelectorAll<HTMLElement>('button')
				.forEach((b) =>
					b.toggleClass('is-active', b.dataset.color === this.color),
				);
		};
		for (const color of COMMENT_COLORS) {
			const b = swatches.createEl('button', {
				cls: 'rtl-sheet-swatch',
				attr: { 'aria-label': color, title: color, 'data-color': color },
			});
			b.setCssProps({ '--rtl-sheet-swatch': color });
			b.addEventListener('click', () => {
				this.color = color;
				picker.value = color;
				paint();
			});
		}
		picker.addEventListener('input', () => {
			this.color = picker.value.toLowerCase();
			paint();
		});
		paint();

		const buttons = new Setting(contentEl).addButton((btn) =>
			btn
				.setButtonText('Save')
				.setCta()
				.onClick(() => this.save()),
		);
		if (this.current)
			buttons.addButton((btn) =>
				btn
					.setButtonText('Delete comment')
					.setDestructive()
					.onClick(() => {
						this.close();
						this.onSave(null);
					}),
			);
		buttons.addButton((btn) =>
			btn.setButtonText('Cancel').onClick(() => this.close()),
		);
		window.setTimeout(() => area.focus(), 0);
	}

	private save(): void {
		const text = this.text.trim();
		this.close();
		this.onSave(text ? { text, color: this.color, time: Date.now() } : null);
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

export interface LinkInput {
	/** What the cell shows. */
	text: string;
	/** A web address, or the name or path of a note in the vault. */
	target: string;
}

/** Puts a link in a cell: the text it shows and where it goes. `onSave(null)` removes the link. */
export class LinkModal extends Modal {
	private text: string;
	private target: string;

	constructor(
		app: App,
		private readonly cellLabel: string,
		current: LinkInput,
		private readonly hasLink: boolean,
		private readonly onSave: (link: LinkInput | null) => void,
	) {
		super(app);
		this.text = current.text;
		this.target = current.target;
	}

	onOpen(): void {
		this.setTitle(
			`${this.hasLink ? 'Edit link in' : 'Insert link in'} ${this.cellLabel}`,
		);
		const { contentEl } = this;
		contentEl.empty();
		const submitOnEnter = (input: HTMLInputElement): void => {
			input.setAttr('dir', 'auto');
			input.addEventListener('keydown', (e) => {
				if (e.key === 'Enter' && !e.isComposing) {
					e.preventDefault();
					this.save();
				}
			});
		};
		let textInput: HTMLInputElement | null = null;
		new Setting(contentEl)
			.setName('Text')
			.setDesc('What the cell shows. Empty = the link itself.')
			.addText((t) => {
				textInput = t.inputEl;
				t.setValue(this.text).onChange((v) => {
					this.text = v;
				});
				submitOnEnter(t.inputEl);
			});
		let targetInput: HTMLInputElement | null = null;
		new Setting(contentEl)
			.setName('Link')
			.setDesc('A web address or the name of a note in your vault.')
			.addText((t) => {
				targetInput = t.inputEl;
				t.setPlaceholder('https://example.com or My note')
					.setValue(this.target)
					.onChange((v) => {
						this.target = v;
					});
				submitOnEnter(t.inputEl);
				window.setTimeout(() => t.inputEl.focus(), 0);
			})
			.addExtraButton((btn) =>
				btn
					.setIcon('file-search')
					.setTooltip('Choose a note')
					.onClick(() => {
						new VaultFilePickerModal(
							this.app,
							['md', 'canvas', 'pdf', 'sheet', 'rtlsheet'],
							'Choose a note to link to…',
							(file) => {
								this.target =
									file.extension === 'md'
										? file.path.slice(0, -3)
										: file.path;
								if (targetInput) targetInput.value = this.target;
								if (!this.text.trim()) {
									this.text = file.basename;
									if (textInput) textInput.value = this.text;
								}
							},
						).open();
					}),
			);
		const buttons = new Setting(contentEl).addButton((btn) =>
			btn
				.setButtonText('Save')
				.setCta()
				.onClick(() => this.save()),
		);
		if (this.hasLink)
			buttons.addButton((btn) =>
				btn
					.setButtonText('Remove link')
					.setDestructive()
					.onClick(() => {
						this.close();
						this.onSave(null);
					}),
			);
		buttons.addButton((btn) =>
			btn.setButtonText('Cancel').onClick(() => this.close()),
		);
	}

	private save(): void {
		// [[Note]] works too.
		const target = this.target.trim().replace(/^\[\[(.*)\]\]$/, '$1');
		this.close();
		if (!target) {
			if (this.hasLink) this.onSave(null);
			return;
		}
		this.onSave({ text: this.text.trim() || target, target });
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

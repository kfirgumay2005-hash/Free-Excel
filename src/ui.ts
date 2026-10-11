import {
	App,
	FuzzySuggestModal,
	Modal,
	Setting,
	SuggestModal,
	TFile,
} from 'obsidian';
import {
	FORMULA_FUNCTIONS,
	FormulaFunctionInfo,
	buildFunctionCall,
} from './formulas';
import { ColumnFilter, FILTER_CONDITIONS, FilterCondition } from './model';

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

/** Asks how the imported table should be laid out: right-to-left, left-to-right, or as the sheet already is. */
export class ImportDirectionModal extends Modal {
	private readonly detectedRtl: boolean;
	private readonly onChoose: (choice: ImportDirectionChoice | null) => void;
	private chosen = false;

	constructor(
		app: App,
		detectedRtl: boolean,
		onChoose: (choice: ImportDirectionChoice | null) => void,
	) {
		super(app);
		this.detectedRtl = detectedRtl;
		this.onChoose = onChoose;
	}

	onOpen(): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h3', { text: 'Import direction' });
		contentEl.createEl('p', {
			text: `The imported table looks ${this.detectedRtl ? 'right-to-left' : 'left-to-right'}. How should the sheet be laid out?`,
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
				this.onChoose(choice);
			});
		};
		add('Right to left', 'rtl', this.detectedRtl);
		add('Left to right', 'ltr', !this.detectedRtl);
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

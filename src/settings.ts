import {
	App,
	PluginSettingTab,
	Setting,
	SettingDefinitionItem,
	SliderComponent,
} from 'obsidian';
import { AiProvider, aiKeySecretId } from './ai';
import type RtlSpreadsheetPlugin from './main';

/** Opens Settings > Hotkeys filtered to this plugin's commands (internal API, so every step is optional). */
function openHotkeySettings(app: App): void {
	const setting = (
		app as unknown as {
			setting?: { openTabById?: (id: string) => unknown };
		}
	).setting;
	const tab = setting?.openTabById?.('hotkeys') as
		| {
				searchComponent?: { setValue(value: string): void };
				updateHotkeyVisibility?: () => void;
		  }
		| undefined;
	tab?.searchComponent?.setValue('Free-Excel');
	tab?.updateHotkeyVisibility?.();
}

export interface PluginSettings {
	autosave: boolean;
	autosaveDelayMs: number;
	/** Reload open sheets when their file changes on disk (for example through a sync plugin). */
	autoRefresh: boolean;
	refreshIntervalSec: number;
	defaultRowCount: number;
	defaultColCount: number;
	defaultFontSize: number;
	/** What the zoom scales: both axes, or only the horizontal one (column widths and letters). */
	zoomAxes: 'both' | 'horizontal';
	enableFormulas: boolean;
	/** Toolbar size in percent, on desktop and on mobile (phones and tablets). */
	toolbarScaleDesktop: number;
	toolbarScaleMobile: number;
	/** Touch screens: size in percent of the selection handles (the dot and its touch area). */
	selectionHandleScale: number;
	/** Touch screens: a handle on all four corners of the selection, or on two opposite ones. */
	selectionHandles: 'four' | 'two';
	/** Colors saved in the toolbar for one-click use, as "#rrggbb" (managed from the toolbar). */
	savedFillColors: string[];
	savedTextColors: string[];
	/** Ctrl+G (Cmd+G) opens the find bar of the active sheet instead of Obsidian's graph view. */
	ctrlGFind: boolean;
	/** Show a cell's comment and link when the mouse is over it (off: only when the cell is selected). */
	commentsOnHover: boolean;
	/** Names of the API keys in Obsidian's secret storage (the keys themselves are not saved here). */
	claudeKeySecret: string;
	geminiKeySecret: string;
	claudeModel: string;
	geminiModel: string;
	/** Gemini model that makes pictures for the assistant's generate_image tool. */
	geminiImageModel: string;
	/** The provider the assistant panel starts with. */
	aiProvider: 'claude' | 'gemini';
	/** Recently inserted emojis, newest first (managed by the emoji picker). */
	recentEmojis: string[];
}

export const DEFAULT_CLAUDE_MODEL = 'claude-opus-5-5';
export const DEFAULT_GEMINI_MODEL = 'gemini-flash-latest';
export const DEFAULT_GEMINI_IMAGE_MODEL = 'gemini-2.5-flash-image';
export const MAX_RECENT_EMOJIS = 24;

/** How many colors can be saved for the fill, and how many for the text. */
export const MAX_SAVED_COLORS = 10;

export const DEFAULT_SETTINGS: PluginSettings = {
	autosave: true,
	autosaveDelayMs: 1000,
	autoRefresh: true,
	refreshIntervalSec: 5,
	defaultRowCount: 50,
	defaultColCount: 26,
	defaultFontSize: 14,
	zoomAxes: 'both',
	enableFormulas: true,
	toolbarScaleDesktop: 100,
	toolbarScaleMobile: 100,
	selectionHandleScale: 100,
	selectionHandles: 'four',
	savedFillColors: [],
	savedTextColors: [],
	ctrlGFind: false,
	commentsOnHover: false,
	claudeKeySecret: '',
	geminiKeySecret: '',
	claudeModel: DEFAULT_CLAUDE_MODEL,
	geminiModel: DEFAULT_GEMINI_MODEL,
	geminiImageModel: DEFAULT_GEMINI_IMAGE_MODEL,
	aiProvider: 'claude',
	recentEmojis: [],
};

function toColors(value: unknown): string[] {
	if (!Array.isArray(value)) return [];
	const colors = value
		.filter((v): v is string => typeof v === 'string')
		.map((v) => v.toLowerCase())
		.filter((v) => /^#[0-9a-f]{6}$/.test(v));
	return [...new Set(colors)].slice(0, MAX_SAVED_COLORS);
}

/** Trimmed text, or the fallback when it is empty or not text. */
function toText(value: unknown, fallback: string): string {
	return typeof value === 'string' && value.trim() !== ''
		? value.trim()
		: fallback;
}

function toInt(
	value: unknown,
	fallback: number,
	min: number,
	max: number,
): number {
	const n = Math.round(Number(value));
	if (!Number.isFinite(n)) return fallback;
	return Math.min(max, Math.max(min, n));
}

/** Makes sure whatever was loaded from disk is valid and complete. */
export function sanitizeSettings(
	raw: Partial<PluginSettings> | null | undefined,
): PluginSettings {
	const s = { ...DEFAULT_SETTINGS, ...(raw ?? {}) };
	return {
		autosave: !!s.autosave,
		autosaveDelayMs: toInt(
			s.autosaveDelayMs,
			DEFAULT_SETTINGS.autosaveDelayMs,
			200,
			10000,
		),
		autoRefresh: !!s.autoRefresh,
		refreshIntervalSec: toInt(
			s.refreshIntervalSec,
			DEFAULT_SETTINGS.refreshIntervalSec,
			1,
			120,
		),
		defaultRowCount: toInt(
			s.defaultRowCount,
			DEFAULT_SETTINGS.defaultRowCount,
			1,
			5000,
		),
		defaultColCount: toInt(
			s.defaultColCount,
			DEFAULT_SETTINGS.defaultColCount,
			1,
			702,
		),
		defaultFontSize: toInt(
			s.defaultFontSize,
			DEFAULT_SETTINGS.defaultFontSize,
			8,
			36,
		),
		zoomAxes: s.zoomAxes === 'horizontal' ? 'horizontal' : 'both',
		enableFormulas: !!s.enableFormulas,
		toolbarScaleDesktop: toInt(
			s.toolbarScaleDesktop,
			DEFAULT_SETTINGS.toolbarScaleDesktop,
			50,
			150,
		),
		toolbarScaleMobile: toInt(
			s.toolbarScaleMobile,
			DEFAULT_SETTINGS.toolbarScaleMobile,
			50,
			150,
		),
		selectionHandleScale: toInt(
			s.selectionHandleScale,
			DEFAULT_SETTINGS.selectionHandleScale,
			50,
			250,
		),
		selectionHandles: s.selectionHandles === 'two' ? 'two' : 'four',
		savedFillColors: toColors(s.savedFillColors),
		savedTextColors: toColors(s.savedTextColors),
		ctrlGFind: !!s.ctrlGFind,
		commentsOnHover: !!s.commentsOnHover,
		claudeKeySecret: toText(s.claudeKeySecret, ''),
		geminiKeySecret: toText(s.geminiKeySecret, ''),
		claudeModel: toText(s.claudeModel, DEFAULT_CLAUDE_MODEL),
		geminiModel: toText(s.geminiModel, DEFAULT_GEMINI_MODEL),
		geminiImageModel: toText(s.geminiImageModel, DEFAULT_GEMINI_IMAGE_MODEL),
		aiProvider: s.aiProvider === 'gemini' ? 'gemini' : 'claude',
		recentEmojis: Array.isArray(s.recentEmojis)
			? s.recentEmojis
					.filter((v): v is string => typeof v === 'string')
					.slice(0, MAX_RECENT_EMOJIS)
			: [],
	};
}

export class SettingsTab extends PluginSettingTab {
	plugin: RtlSpreadsheetPlugin;

	constructor(app: App, plugin: RtlSpreadsheetPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	getControlValue(key: string): unknown {
		return this.plugin.settings[key as keyof PluginSettings];
	}

	/** Goes through saveSettings so the value is sanitized and open sheets pick it up. */
	async setControlValue(key: string, value: unknown): Promise<void> {
		Object.assign(this.plugin.settings, { [key]: value });
		await this.plugin.saveSettings();
		this.refreshDomState();
	}

	/**
	 * One field to paste the key into, and the steps to get one. The key goes straight to
	 * Obsidian's secret storage; it is never written to the plugin settings.
	 */
	private renderKeyField(setting: Setting, provider: AiProvider): void {
		const id = aiKeySecretId(provider);
		const steps =
			provider === 'claude'
				? {
						url: 'https://console.anthropic.com/settings/keys',
						site: 'console.anthropic.com',
						list: [
							'Open the Anthropic Console and sign in (or create an account).',
							'Under Settings > Billing, add credits: the API is paid separately from a Claude.ai subscription.',
							'Under Settings > API keys, choose "Create key", give it a name and copy it (it starts with "sk-ant-").',
							'Paste the key here.',
						],
					}
				: {
						url: 'https://aistudio.google.com/apikey',
						site: 'aistudio.google.com/apikey',
						list: [
							'Open Google AI Studio and sign in with your Google account.',
							'Choose "Create API key" (pick or create a Google Cloud project when asked).',
							'Open the key and copy the long line under "API key". There is a free tier with daily limits.',
							'Paste the key here.',
						],
					};
		setting.descEl.empty();
		setting.descEl.createDiv({
			text: `Used by the assistant panel and the ${provider === 'claude' ? 'CLAUDE()' : 'GEMINI()'} formula. How to get a key:`,
		});
		const ol = setting.descEl.createEl('ol', { cls: 'rtl-sheet-key-steps' });
		steps.list.forEach((step, i) => {
			const li = ol.createEl('li', { text: step });
			if (i === 0) {
				li.appendText(' ');
				li.createEl('a', { text: steps.site, href: steps.url });
			}
		});
		setting.descEl.createDiv({
			text: 'The key is kept in the Obsidian secret storage on this device, not in the plugin settings or the vault.',
		});

		let stored = '';
		try {
			stored = this.app.secretStorage.getSecret(id) ?? '';
		} catch {
			stored = '';
		}
		setting.addText((text) => {
			text.inputEl.type = 'password';
			text.inputEl.autocomplete = 'off';
			text.inputEl.spellcheck = false;
			text.setPlaceholder('Paste the API key').setValue(stored).onChange((value) => {
				try {
					this.app.secretStorage.setSecret(id, value.trim());
				} catch (e) {
					console.error('[Free-Excel] Could not store the API key', e);
				}
			});
		});
	}

	getSettingDefinitions(): SettingDefinitionItem<keyof PluginSettings>[] {
		return [
			{
				type: 'group',
				heading: 'Saving',
				items: [
					{
						name: 'Autosave',
						desc: 'Save automatically after every change. When off, sheets are saved when closed or via the "Save active spreadsheet" command.',
						control: { type: 'toggle', key: 'autosave' },
					},
					{
						name: 'Autosave delay (ms)',
						desc: 'How long to wait after the last change before saving.',
						control: {
							type: 'slider',
							key: 'autosaveDelayMs',
							min: 200,
							max: 10000,
							step: 100,
							disabled: () => !this.plugin.settings.autosave,
						},
					},
				],
			},
			{
				type: 'group',
				heading: 'Sync',
				items: [
					{
						name: 'Refresh open sheets',
						desc: 'Reload an open sheet in place when its file changes outside this view, for example when a sync plugin brings in changes from another device.',
						control: { type: 'toggle', key: 'autoRefresh' },
					},
					{
						name: 'Refresh interval (seconds)',
						desc: 'How often an open sheet checks its file for changes.',
						control: {
							type: 'slider',
							key: 'refreshIntervalSec',
							min: 1,
							max: 120,
							step: 1,
							disabled: () => !this.plugin.settings.autoRefresh,
						},
					},
				],
			},
			{
				type: 'group',
				heading: 'New sheets',
				items: [
					{
						name: 'Default row count',
						desc: 'Number of rows in a newly created sheet (1-5000).',
						control: {
							type: 'number',
							key: 'defaultRowCount',
							defaultValue: DEFAULT_SETTINGS.defaultRowCount,
							min: 1,
							max: 5000,
							step: 1,
						},
					},
					{
						name: 'Default column count',
						desc: 'Number of columns in a newly created sheet (1-702).',
						control: {
							type: 'number',
							key: 'defaultColCount',
							defaultValue: DEFAULT_SETTINGS.defaultColCount,
							min: 1,
							max: 702,
							step: 1,
						},
					},
				],
			},
			{
				type: 'group',
				heading: 'Appearance and calculation',
				items: [
					{
						name: 'Default font size',
						desc: 'Cell font size in pixels. Applies to all open sheets immediately.',
						control: {
							type: 'slider',
							key: 'defaultFontSize',
							min: 8,
							max: 36,
							step: 1,
						},
					},
					{
						name: 'Zoom',
						desc: 'What zooming in and out changes. Both axes: column widths and letters, row heights and numbers, and the text. Horizontal only: the column widths and letters; rows, row numbers and text keep their size.',
						control: {
							type: 'dropdown',
							key: 'zoomAxes',
							options: {
								both: 'Both axes',
								horizontal: 'Horizontal only (columns)',
							},
						},
					},
					{
						name: 'Enable formulas',
						desc: 'Evaluate cells that start with "=" (SUM, AVERAGE, COUNT, MIN, MAX, IF and many more). Applies to all open sheets immediately.',
						control: { type: 'toggle', key: 'enableFormulas' },
					},
					{
						name: 'Show comments on hover',
						desc: 'Show the comment and link of a cell as soon as the mouse is over it. When off, they show only when you click the cell.',
						control: { type: 'toggle', key: 'commentsOnHover' },
					},
				],
			},
			{
				type: 'group',
				heading: 'Keyboard',
				items: [
					{
						name: 'Ctrl+G finds in the sheet',
						desc: 'While a spreadsheet is active, Ctrl+G (Cmd+G on macOS) opens the find bar instead of the graph view. Everywhere else Ctrl+G keeps opening the graph view.',
						control: { type: 'toggle', key: 'ctrlGFind' },
					},
					{
						name: 'Custom find shortcut',
						desc: 'Prefer another key? Assign any shortcut to the "Find in spreadsheet" command in Settings > Hotkeys.',
						render: (setting) => {
							setting.addButton((btn) =>
								btn
									.setButtonText('Open hotkeys')
									.onClick(() => openHotkeySettings(this.app)),
							);
						},
					},
				],
			},
			{
				type: 'group',
				heading: 'AI (Claude and Gemini)',
				items: [
					{
						name: 'Claude API key',
						desc: 'Used by the assistant panel and the CLAUDE() formula.',
						render: (setting) => this.renderKeyField(setting, 'claude'),
					},
					{
						name: 'Claude model',
						desc: `Model ID, for example ${DEFAULT_CLAUDE_MODEL} or claude-haiku-5-5 (faster and cheaper for many formulas).`,
						control: {
							type: 'text',
							key: 'claudeModel',
							placeholder: DEFAULT_CLAUDE_MODEL,
						},
					},
					{
						name: 'Gemini API key',
						desc: 'Used by the assistant panel and the GEMINI() formula.',
						render: (setting) => this.renderKeyField(setting, 'gemini'),
					},
					{
						name: 'Gemini model',
						desc: `Model ID, for example ${DEFAULT_GEMINI_MODEL} or gemini-3.1-pro-preview.`,
						control: {
							type: 'text',
							key: 'geminiModel',
							placeholder: DEFAULT_GEMINI_MODEL,
						},
					},
					{
						name: 'Gemini image model',
						desc: 'Makes pictures when you ask the assistant (with Claude or Gemini) to create an image. Needs the Gemini API key; image models may need a paid plan.',
						control: {
							type: 'text',
							key: 'geminiImageModel',
							placeholder: DEFAULT_GEMINI_IMAGE_MODEL,
						},
					},
					{
						name: 'Assistant provider',
						desc: 'Which AI the assistant panel uses when it opens. You can switch in the panel.',
						control: {
							type: 'dropdown',
							key: 'aiProvider',
							options: { claude: 'Claude', gemini: 'Gemini' },
						},
					},
				],
			},
			{
				type: 'group',
				heading: 'Toolbar',
				items: [
					{
						name: 'Toolbar size on desktop (%)',
						desc: 'Size of the toolbar buttons on computers and laptops. Applies to all open sheets immediately.',
						control: {
							type: 'slider',
							key: 'toolbarScaleDesktop',
							min: 50,
							max: 150,
							step: 5,
						},
					},
					{
						name: 'Toolbar size on mobile (%)',
						desc: 'Size of the toolbar buttons on phones and tablets. A smaller toolbar shows more buttons without scrolling.',
						control: {
							type: 'slider',
							key: 'toolbarScaleMobile',
							min: 50,
							max: 150,
							step: 5,
						},
					},
				],
			},
			{
				type: 'group',
				heading: 'Touch selection',
				items: [
					{
						name: 'Selection handle size (%)',
						desc: 'Size of the dots on the corners of the selection on touch screens, together with their touch area. Applies to all open sheets immediately.',
						render: (setting) => {
							let slider: SliderComponent | null = null;
							setting
								.addSlider((sl) => {
									slider = sl;
									sl.setLimits(50, 250, 10)
										.setValue(
											this.plugin.settings
												.selectionHandleScale,
										)
										.onChange((value) =>
											this.setControlValue(
												'selectionHandleScale',
												value,
											),
										);
								})
								.addExtraButton((btn) =>
									btn
										.setIcon('rotate-ccw')
										.setTooltip('Reset to default size')
										.onClick(() => {
											const value =
												DEFAULT_SETTINGS.selectionHandleScale;
											slider?.setValue(value);
											void this.setControlValue(
												'selectionHandleScale',
												value,
											);
										}),
								);
						},
					},
					{
						name: 'Selection handles',
						desc: 'Show a handle on every corner of the selection, or only on the top right and bottom left corners.',
						control: {
							type: 'dropdown',
							key: 'selectionHandles',
							options: {
								four: 'All four corners',
								two: 'Top right and bottom left',
							},
						},
					},
				],
			},
		];
	}
}

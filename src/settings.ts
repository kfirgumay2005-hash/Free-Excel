import {
	App,
	PluginSettingTab,
	SettingDefinitionItem,
	SliderComponent,
} from 'obsidian';
import type RtlSpreadsheetPlugin from './main';

export interface PluginSettings {
	autosave: boolean;
	autosaveDelayMs: number;
	/** Reload open sheets when their file changes on disk (for example through a sync plugin). */
	autoRefresh: boolean;
	refreshIntervalSec: number;
	defaultRowCount: number;
	defaultColCount: number;
	defaultFontSize: number;
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
}

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
	enableFormulas: true,
	toolbarScaleDesktop: 100,
	toolbarScaleMobile: 100,
	selectionHandleScale: 100,
	selectionHandles: 'four',
	savedFillColors: [],
	savedTextColors: [],
};

function toColors(value: unknown): string[] {
	if (!Array.isArray(value)) return [];
	const colors = value
		.filter((v): v is string => typeof v === 'string')
		.map((v) => v.toLowerCase())
		.filter((v) => /^#[0-9a-f]{6}$/.test(v));
	return [...new Set(colors)].slice(0, MAX_SAVED_COLORS);
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
						name: 'Enable formulas',
						desc: 'Evaluate cells that start with "=" (SUM, AVERAGE, COUNT, MIN, MAX, IF and many more). Applies to all open sheets immediately.',
						control: { type: 'toggle', key: 'enableFormulas' },
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
										.setDynamicTooltip()
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

import { App, PluginSettingTab, SettingDefinitionItem } from 'obsidian';
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
}

export const DEFAULT_SETTINGS: PluginSettings = {
	autosave: true,
	autosaveDelayMs: 1000,
	autoRefresh: true,
	refreshIntervalSec: 5,
	defaultRowCount: 50,
	defaultColCount: 26,
	defaultFontSize: 14,
	enableFormulas: true,
};

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
		];
	}
}

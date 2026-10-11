import { App, PluginSettingTab, Setting } from 'obsidian';
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

	private async savePatch(patch: Partial<PluginSettings>): Promise<void> {
		Object.assign(this.plugin.settings, patch);
		await this.plugin.saveSettings();
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();

		const s = this.plugin.settings;

		new Setting(containerEl).setName('Saving').setHeading();

		let delaySetting: Setting | null = null;

		new Setting(containerEl)
			.setName('Autosave')
			.setDesc(
				'Save automatically after every change. When off, sheets are saved when closed or via the "Save active spreadsheet" command.',
			)
			.addToggle((toggle) =>
				toggle.setValue(s.autosave).onChange(async (value) => {
					await this.savePatch({ autosave: value });
					delaySetting?.setDisabled(!value);
				}),
			);

		delaySetting = new Setting(containerEl)
			.setName('Autosave delay (ms)')
			.setDesc('How long to wait after the last change before saving.')
			.addSlider((slider) =>
				slider
					.setLimits(200, 10000, 100)
					.setValue(s.autosaveDelayMs)
					.setDynamicTooltip()
					.onChange(async (value) => {
						await this.savePatch({ autosaveDelayMs: value });
					}),
			);
		delaySetting.setDisabled(!s.autosave);

		new Setting(containerEl).setName('Sync').setHeading();

		let intervalSetting: Setting | null = null;

		new Setting(containerEl)
			.setName('Refresh open sheets')
			.setDesc(
				'Reload an open sheet in place when its file changes outside this view, for example when a sync plugin brings in changes from another device.',
			)
			.addToggle((toggle) =>
				toggle.setValue(s.autoRefresh).onChange(async (value) => {
					await this.savePatch({ autoRefresh: value });
					intervalSetting?.setDisabled(!value);
				}),
			);

		intervalSetting = new Setting(containerEl)
			.setName('Refresh interval (seconds)')
			.setDesc('How often an open sheet checks its file for changes.')
			.addSlider((slider) =>
				slider
					.setLimits(1, 120, 1)
					.setValue(s.refreshIntervalSec)
					.setDynamicTooltip()
					.onChange(async (value) => {
						await this.savePatch({ refreshIntervalSec: value });
					}),
			);
		intervalSetting.setDisabled(!s.autoRefresh);

		new Setting(containerEl).setName('New sheets').setHeading();

		new Setting(containerEl)
			.setName('Default row count')
			.setDesc('Number of rows in a newly created sheet (1-5000).')
			.addText((text) => {
				text.inputEl.type = 'number';
				text.inputEl.min = '1';
				text.inputEl.max = '5000';
				text.setValue(String(s.defaultRowCount)).onChange(
					async (value) => {
						const n = parseInt(value, 10);
						if (!Number.isFinite(n)) return;
						await this.savePatch({
							defaultRowCount: toInt(
								n,
								DEFAULT_SETTINGS.defaultRowCount,
								1,
								5000,
							),
						});
					},
				);
			});

		new Setting(containerEl)
			.setName('Default column count')
			.setDesc('Number of columns in a newly created sheet (1-702).')
			.addText((text) => {
				text.inputEl.type = 'number';
				text.inputEl.min = '1';
				text.inputEl.max = '702';
				text.setValue(String(s.defaultColCount)).onChange(
					async (value) => {
						const n = parseInt(value, 10);
						if (!Number.isFinite(n)) return;
						await this.savePatch({
							defaultColCount: toInt(
								n,
								DEFAULT_SETTINGS.defaultColCount,
								1,
								702,
							),
						});
					},
				);
			});

		new Setting(containerEl)
			.setName('Appearance and calculation')
			.setHeading();

		new Setting(containerEl)
			.setName('Default font size')
			.setDesc(
				'Cell font size in pixels. Applies to all open sheets immediately.',
			)
			.addSlider((slider) =>
				slider
					.setLimits(8, 36, 1)
					.setValue(s.defaultFontSize)
					.setDynamicTooltip()
					.onChange(async (value) => {
						await this.savePatch({ defaultFontSize: value });
					}),
			);

		new Setting(containerEl)
			.setName('Enable formulas')
			.setDesc(
				'Evaluate cells that start with "=" (SUM, AVERAGE, COUNT, MIN, MAX, IF and many more). Applies to all open sheets immediately.',
			)
			.addToggle((toggle) =>
				toggle.setValue(s.enableFormulas).onChange(async (value) => {
					await this.savePatch({ enableFormulas: value });
				}),
			);
	}
}

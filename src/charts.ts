import { setIcon } from 'obsidian';
import {
	AxisStyle,
	ChartSpec,
	ChartType,
	LegendPosition,
	Scalar,
	SelRange,
	SeriesStyle,
	indexToColumnLabel,
} from './model';

export const CHART_TYPES: { value: ChartType; label: string }[] = [
	{ value: 'column', label: 'Column' },
	{ value: 'bar', label: 'Bar' },
	{ value: 'line', label: 'Line' },
	{ value: 'pie', label: 'Pie' },
	{ value: 'scatter', label: 'Scatter (XY)' },
];

/** Size in px (at 100% zoom) of a new chart, and the smallest size it can be resized to. */
export const CHART_DEFAULT_W = 480;
export const CHART_DEFAULT_H = 320;
export const CHART_MIN_W = 180;
export const CHART_MIN_H = 140;
/** Height of the chart's title bar (at 100% zoom). */
export const CHART_HEAD_H = 28;
/** Most rows a chart reads from its range. */
const MAX_POINTS = 2000;

/** Obsidian's accent colors: they follow the light and dark themes. */
const PALETTE = [
	'var(--color-blue)',
	'var(--color-orange)',
	'var(--color-green)',
	'var(--color-red)',
	'var(--color-purple)',
	'var(--color-cyan)',
	'var(--color-yellow)',
	'var(--color-pink)',
];

const seriesColor = (i: number): string =>
	PALETTE[i % PALETTE.length] ?? 'var(--color-blue)';

/** Reads the value a cell shows (formulas already worked out). */
export type CellReader = (row: number, col: number) => Scalar;

export interface ChartSeries {
	name: string;
	values: (number | null)[];
}

export interface ChartData {
	/** One label per data row (the first column of the range, or 1, 2, 3...). */
	categories: string[];
	/** The first column as numbers, used as X by scatter charts. */
	xValues: (number | null)[];
	series: ChartSeries[];
}

export interface Trend {
	slope: number;
	intercept: number;
	r2: number;
}

function toNumber(v: Scalar): number | null {
	if (typeof v === 'number') return Number.isFinite(v) ? v : null;
	if (typeof v !== 'string') return null;
	const t = v.trim().replace(/,/g, '');
	if (t === '') return null;
	const n = Number(t);
	return Number.isFinite(n) ? n : null;
}

function toText(v: Scalar): string {
	if (v === null) return '';
	if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
	return String(v);
}

/* ------------------------------------------------------------------ */
/* Ranges                                                               */
/* ------------------------------------------------------------------ */

function columnIndex(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase())
		n = n * 26 + (ch.charCodeAt(0) - 64);
	return n - 1;
}

/** "A1:C10" (or a single cell such as "B2") to a range; null when it is not a range. */
export function parseRangeText(text: string): SelRange | null {
	const m =
		/^\s*\$?([A-Za-z]{1,3})\$?(\d+)\s*(?::\s*\$?([A-Za-z]{1,3})\$?(\d+))?\s*$/.exec(
			text,
		);
	if (!m) return null;
	const c1 = columnIndex(m[1] ?? 'A');
	const r1 = Number(m[2]) - 1;
	const c2 = m[3] ? columnIndex(m[3]) : c1;
	const r2 = m[4] ? Number(m[4]) - 1 : r1;
	if (r1 < 0 || r2 < 0) return null;
	return {
		r1: Math.min(r1, r2),
		r2: Math.max(r1, r2),
		c1: Math.min(c1, c2),
		c2: Math.max(c1, c2),
	};
}

export function formatRange(rg: SelRange): string {
	const first = `${indexToColumnLabel(rg.c1)}${rg.r1 + 1}`;
	const last = `${indexToColumnLabel(rg.c2)}${rg.r2 + 1}`;
	return first === last ? first : `${first}:${last}`;
}

/** True when the first row looks like titles: some text that is not a number above the values. */
export function guessHeaderRow(range: SelRange, read: CellReader): boolean {
	if (range.r2 <= range.r1) return false;
	const first = range.c2 > range.c1 ? range.c1 + 1 : range.c1;
	for (let c = first; c <= range.c2; c++) {
		const v = read(range.r1, c);
		if (v !== null && v !== '' && toNumber(v) === null) return true;
	}
	return false;
}

/**
 * The values a chart shows. With two or more columns the first one holds the labels (or the
 * X values of a scatter chart) and every other column is a series; a single column is one
 * series numbered 1, 2, 3...
 */
export function readChartData(spec: ChartSpec, read: CellReader): ChartData {
	if (spec.type === 'scatter' && spec.xRange)
		return readXYData(spec, spec.xRange, read);
	const { r1, c1, c2 } = spec.range;
	const r2 = Math.min(spec.range.r2, r1 + MAX_POINTS);
	const firstRow = spec.headerRow ? r1 + 1 : r1;
	const hasLabels = c2 > c1;
	const cols: number[] = [];
	for (let c = hasLabels ? c1 + 1 : c1; c <= c2; c++) cols.push(c);

	const categories: string[] = [];
	const xValues: (number | null)[] = [];
	const series: ChartSeries[] = cols.map((c, i) => ({
		name:
			(spec.headerRow ? toText(read(r1, c)).trim() : '') ||
			(cols.length > 1
				? `Column ${indexToColumnLabel(c)}`
				: `Series ${i + 1}`),
		values: [],
	}));
	for (let r = firstRow; r <= r2; r++) {
		const label = hasLabels ? read(r, c1) : null;
		const values = cols.map((c) => toNumber(read(r, c)));
		// Rows that are completely empty (often the end of the range) are left out.
		if (toText(label).trim() === '' && values.every((v) => v === null))
			continue;
		const n = categories.length + 1;
		categories.push(hasLabels ? toText(label) : String(n));
		xValues.push(hasLabels ? toNumber(label) : n);
		values.forEach((v, i) => series[i]?.values.push(v));
	}
	return { categories, xValues, series };
}

/** A range as lines of cells: each of its columns, or its only row when it is one row high. */
function rangeLines(rg: SelRange): [number, number][][] {
	if (rg.r2 === rg.r1 && rg.c2 > rg.c1) {
		const line: [number, number][] = [];
		for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + MAX_POINTS); c++)
			line.push([rg.r1, c]);
		return [line];
	}
	const lines: [number, number][][] = [];
	for (let c = rg.c1; c <= rg.c2; c++) {
		const line: [number, number][] = [];
		for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + MAX_POINTS); r++)
			line.push([r, c]);
		lines.push(line);
	}
	return lines;
}

/**
 * Scatter chart with its own X range: the n-th X value goes with the n-th Y value of every
 * series (each column of the Y range, or its only row).
 */
function readXYData(
	spec: ChartSpec,
	xRange: SelRange,
	read: CellReader,
): ChartData {
	const skip = spec.headerRow ? 1 : 0;
	const xLine = (rangeLines(xRange)[0] ?? []).slice(skip);
	const yLines = rangeLines(spec.range);
	const ySeries = yLines.map((line, i) => {
		const first = line[0];
		const header =
			spec.headerRow && first
				? toText(read(first[0], first[1])).trim()
				: '';
		return {
			name:
				header ||
				(yLines.length > 1 && first
					? `Column ${indexToColumnLabel(first[1])}`
					: `Series ${i + 1}`),
			values: line.slice(skip).map(([r, c]) => toNumber(read(r, c))),
		};
	});
	const count = Math.max(
		xLine.length,
		...ySeries.map((s) => s.values.length),
	);
	const categories: string[] = [];
	const xValues: (number | null)[] = [];
	const series: ChartSeries[] = ySeries.map((s) => ({
		name: s.name,
		values: [],
	}));
	for (let i = 0; i < count; i++) {
		const cell = xLine[i];
		const x = cell ? read(cell[0], cell[1]) : null;
		const ys = ySeries.map((s) => s.values[i] ?? null);
		if (toNumber(x) === null && ys.every((v) => v === null)) continue;
		categories.push(toText(x));
		xValues.push(toNumber(x));
		ys.forEach((v, j) => series[j]?.values.push(v));
	}
	return { categories, xValues, series };
}

/** Least-squares line through the points; null with fewer than two distinct X values. */
export function linearTrend(points: { x: number; y: number }[]): Trend | null {
	const n = points.length;
	if (n < 2) return null;
	let sx = 0;
	let sy = 0;
	for (const p of points) {
		sx += p.x;
		sy += p.y;
	}
	const mx = sx / n;
	const my = sy / n;
	let sxx = 0;
	let sxy = 0;
	let syy = 0;
	for (const p of points) {
		sxx += (p.x - mx) ** 2;
		sxy += (p.x - mx) * (p.y - my);
		syy += (p.y - my) ** 2;
	}
	if (sxx === 0) return null;
	const slope = sxy / sxx;
	const intercept = my - slope * mx;
	const r2 = syy === 0 ? 1 : (sxy * sxy) / (sxx * syy);
	return { slope, intercept, r2 };
}

/* ------------------------------------------------------------------ */
/* Drawing                                                              */
/* ------------------------------------------------------------------ */

const NUMBER_FORMAT = new Intl.NumberFormat(undefined, {
	maximumFractionDigits: 2,
});
const COMPACT_FORMAT = new Intl.NumberFormat(undefined, {
	notation: 'compact',
	maximumFractionDigits: 1,
});

function formatNumber(n: number): string {
	return Math.abs(n) >= 1e6
		? COMPACT_FORMAT.format(n)
		: NUMBER_FORMAT.format(n);
}

function formatCoefficient(n: number): string {
	const abs = Math.abs(n);
	if (abs !== 0 && (abs < 0.01 || abs >= 1e6)) return n.toExponential(2);
	return NUMBER_FORMAT.format(n);
}

/** "y = 2x + 1, R² = 0.950", or only one of the two. */
export function trendLabel(t: Trend, equation = true, r2 = true): string {
	const parts: string[] = [];
	if (equation) {
		const sign = t.intercept < 0 ? '−' : '+';
		parts.push(
			`y = ${formatCoefficient(t.slope)}x ${sign} ${formatCoefficient(Math.abs(t.intercept))}`,
		);
	}
	if (r2) parts.push(`R² = ${t.r2.toFixed(3)}`);
	return parts.join(', ');
}

/** Color of a series, or of a slice of a pie chart. */
export function colorOf(spec: ChartSpec, i: number): string {
	return spec.series?.[i]?.color ?? seriesColor(i);
}

function hasTrend(spec: ChartSpec, i: number): boolean {
	return (
		spec.type === 'scatter' &&
		(spec.series?.[i]?.trendline ?? spec.trendline)
	);
}

function trendColorOf(spec: ChartSpec, i: number): string {
	return spec.series?.[i]?.trendColor ?? spec.trendColor ?? colorOf(spec, i);
}

interface Scale {
	min: number;
	max: number;
	step: number;
}

/** Round steps (1, 2, 5 x 10^n) over lo..hi; the axis' own min and max win when set. */
function axisScale(
	lo: number,
	hi: number,
	count: number,
	axis?: AxisStyle,
): Scale {
	let fixedMin = typeof axis?.min === 'number' ? axis.min : undefined;
	let fixedMax = typeof axis?.max === 'number' ? axis.max : undefined;
	if (
		fixedMin !== undefined &&
		fixedMax !== undefined &&
		fixedMax <= fixedMin
	) {
		fixedMin = undefined;
		fixedMax = undefined;
	}
	let min = fixedMin ?? lo;
	let max = fixedMax ?? hi;
	if (max < min) [min, max] = [max, min];
	if (min === max) {
		const pad = min === 0 ? 1 : Math.abs(min) * 0.5;
		if (fixedMin === undefined) min -= pad;
		else max += pad;
		if (fixedMax === undefined) max += pad;
	}
	const rough = (max - min) / Math.max(1, count);
	const power = 10 ** Math.floor(Math.log10(rough));
	const unit = rough / power;
	const step = (unit <= 1 ? 1 : unit <= 2 ? 2 : unit <= 5 ? 5 : 10) * power;
	return {
		min: fixedMin ?? Math.floor(min / step) * step,
		max: fixedMax ?? Math.max(max, Math.ceil(max / step) * step),
		step,
	};
}

function ticks(scale: Scale): number[] {
	const out: number[] = [];
	const first = Math.ceil(scale.min / scale.step - 1e-9) * scale.step;
	for (let v = first; v <= scale.max + scale.step * 1e-6; v += scale.step)
		out.push(Math.abs(v) < scale.step / 1e6 ? 0 : v);
	return out;
}

function truncate(text: string, max: number): string {
	if (max < 2) return '';
	return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

type SvgAttrs = Record<string, string | number>;

function svg<K extends keyof SVGElementTagNameMap>(
	parent: Element,
	tag: K,
	attrs: SvgAttrs = {},
	style: Partial<Record<'fill' | 'stroke', string | undefined>> = {},
): SVGElementTagNameMap[K] {
	const el = parent.createSvg(tag);
	for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
	// CSS variables only work through style, not through presentation attributes.
	if (style.fill) el.style.fill = style.fill;
	if (style.stroke) el.style.stroke = style.stroke;
	return el;
}

function tooltip(el: SVGElement, text: string): void {
	svg(el, 'title').textContent = text;
}

function label(
	parent: Element,
	x: number,
	y: number,
	text: string,
	anchor: 'start' | 'middle' | 'end',
	fontSize: number,
	cls = 'rtl-chart-label',
): SVGTextElement {
	const t = svg(parent, 'text', {
		x,
		y,
		'text-anchor': anchor,
		'dominant-baseline': 'middle',
		class: cls,
	});
	t.style.fontSize = `${fontSize}px`;
	t.textContent = text;
	return t;
}

interface Box {
	x: number;
	y: number;
	w: number;
	h: number;
}

interface LegendItem {
	name: string;
	color: string;
	/** Drawn as a dashed line (a trendline) instead of a square. */
	line?: boolean;
}

/** Draws the legend on one side of `area` and returns what is left for the plot. */
function placeLegend(
	root: Element,
	items: LegendItem[],
	pos: LegendPosition,
	area: Box,
	fs: number,
): Box {
	if (pos === 'none' || items.length === 0) return area;
	const charW = fs * 0.6;
	const itemW = (it: LegendItem): number =>
		fs + 8 + Math.min(it.name.length, 60) * charW + 10;
	const rowH = fs + 8;
	const mark = (it: LegendItem, x: number, y: number): void => {
		if (it.line)
			svg(
				root,
				'line',
				{ x1: x, y1: y, x2: x + fs, y2: y, class: 'rtl-chart-trend' },
				{ stroke: it.color },
			);
		else
			svg(
				root,
				'rect',
				{ x, y: y - fs / 2 + 1, width: fs - 2, height: fs - 2, rx: 2 },
				{ fill: it.color },
			);
	};

	if (pos === 'right') {
		const w = Math.min(area.w * 0.45, Math.max(...items.map(itemW)) + 8);
		const x0 = area.x + area.w - w + 4;
		const maxChars = Math.floor((w - fs - 14) / charW);
		let y = area.y + 8 + rowH / 2;
		for (const it of items) {
			if (y > area.y + area.h - rowH / 2) break;
			mark(it, x0, y);
			label(
				root,
				x0 + fs + 6,
				y,
				truncate(it.name, maxChars),
				'start',
				fs,
			);
			y += rowH;
		}
		return { ...area, w: area.w - w };
	}

	const rows: LegendItem[][] = [];
	let row: LegendItem[] = [];
	let used = 0;
	for (const it of items) {
		if (row.length > 0 && used + itemW(it) > area.w - 16) {
			rows.push(row);
			row = [];
			used = 0;
		}
		row.push(it);
		used += itemW(it);
	}
	if (row.length > 0) rows.push(row);
	const h = Math.min(rows.length * rowH + 6, area.h * 0.4);
	const top = pos === 'top' ? area.y : area.y + area.h - h;
	rows.forEach((r, i) => {
		const y = top + 3 + i * rowH + rowH / 2;
		if (y > top + h) return;
		const total = r.reduce((sum, it) => sum + itemW(it), 0);
		let x = area.x + Math.max(8, (area.w - total) / 2);
		for (const it of r) {
			mark(it, x, y);
			label(
				root,
				x + fs + 6,
				y,
				truncate(it.name, Math.floor((area.w - 24) / charW)),
				'start',
				fs,
			);
			x += itemW(it);
		}
	});
	return pos === 'top'
		? { ...area, y: area.y + h, h: area.h - h }
		: { ...area, h: area.h - h };
}

let clipCounter = 0;

/**
 * Draws a chart into `el` as an SVG that fills it. Sizes are in px at 100% zoom: the SVG
 * scales with the element, so zooming the sheet zooms the chart.
 */
export function renderChart(
	el: HTMLElement,
	spec: ChartSpec,
	data: ChartData,
	width: number,
	height: number,
): void {
	el.empty();
	const root = svg(el, 'svg', {
		viewBox: `0 0 ${width} ${height}`,
		width: '100%',
		height: '100%',
		class: 'rtl-chart-svg',
	});
	// Charts read left to right in every sheet, like in Excel.
	root.setAttribute('direction', 'ltr');
	const fs = Math.min(32, Math.max(6, spec.fontSize ?? 11));
	if (spec.bg)
		svg(root, 'rect', { x: 0, y: 0, width, height }, { fill: spec.bg });

	const hasData = data.series.some((s) => s.values.some((v) => v !== null));
	if (!hasData || data.categories.length === 0) {
		label(
			root,
			width / 2,
			height / 2,
			'No numbers in the data range',
			'middle',
			fs + 2,
			'rtl-chart-empty',
		);
		return;
	}
	const area: Box = { x: 0, y: 0, w: width, h: height };
	if (spec.type === 'pie') drawPie(root, spec, data, area, fs);
	else drawAxes(root, spec, data, area, fs);
}

function drawPie(
	root: SVGSVGElement,
	spec: ChartSpec,
	data: ChartData,
	area: Box,
	fs: number,
): void {
	const series = data.series[0];
	if (!series) return;
	const slices = data.categories
		.map((name, i) => ({ name, value: series.values[i] ?? 0, i }))
		.filter((s) => s.value > 0);
	const total = slices.reduce((sum, s) => sum + s.value, 0);
	if (total <= 0) {
		label(
			root,
			area.w / 2,
			area.h / 2,
			'No positive values to show',
			'middle',
			fs + 2,
			'rtl-chart-empty',
		);
		return;
	}
	const plot = placeLegend(
		root,
		slices.map((s) => ({
			name: s.name || `Item ${s.i + 1}`,
			color: colorOf(spec, s.i),
		})),
		spec.legend ?? 'bottom',
		area,
		fs,
	);
	const cx = plot.x + plot.w / 2;
	const cy = plot.y + plot.h / 2;
	const radius = Math.max(10, Math.min(plot.w, plot.h) / 2 - 10);
	const hole =
		(Math.min(90, Math.max(0, spec.donutHole ?? 0)) / 100) * radius;
	const pt = (r: number, a: number): string =>
		`${cx + r * Math.cos(a)} ${cy + r * Math.sin(a)}`;
	let angle = -Math.PI / 2;
	for (const s of slices) {
		const share = s.value / total;
		const sweep = share * Math.PI * 2;
		const end = angle + sweep;
		const large = sweep > Math.PI ? 1 : 0;
		let d: string;
		if (share >= 0.9999) {
			// A full ring: two circles, the inner one cut out.
			d = `M ${cx - radius} ${cy} A ${radius} ${radius} 0 1 0 ${cx + radius} ${cy} A ${radius} ${radius} 0 1 0 ${cx - radius} ${cy} Z`;
			if (hole > 0)
				d += ` M ${cx - hole} ${cy} A ${hole} ${hole} 0 1 0 ${cx + hole} ${cy} A ${hole} ${hole} 0 1 0 ${cx - hole} ${cy} Z`;
		} else if (hole > 0) {
			d = `M ${pt(radius, angle)} A ${radius} ${radius} 0 ${large} 1 ${pt(radius, end)} L ${pt(hole, end)} A ${hole} ${hole} 0 ${large} 0 ${pt(hole, angle)} Z`;
		} else {
			d = `M ${cx} ${cy} L ${pt(radius, angle)} A ${radius} ${radius} 0 ${large} 1 ${pt(radius, end)} Z`;
		}
		const shape = svg(
			root,
			'path',
			{ d, class: 'rtl-chart-slice', 'fill-rule': 'evenodd' },
			{ fill: colorOf(spec, s.i) },
		);
		tooltip(
			shape,
			`${s.name}: ${formatNumber(s.value)} (${(share * 100).toFixed(1)}%)`,
		);
		if (share >= 0.05) {
			const mid = angle + sweep / 2;
			const r = hole > 0 ? (radius + hole) / 2 : radius * 0.65;
			label(
				root,
				cx + r * Math.cos(mid),
				cy + r * Math.sin(mid),
				`${Math.round(share * 100)}%`,
				'middle',
				fs,
				'rtl-chart-slice-label',
			);
		}
		angle = end;
	}
}

function drawAxes(
	root: SVGSVGElement,
	spec: ChartSpec,
	data: ChartData,
	area: Box,
	fs: number,
): void {
	const horizontal = spec.type === 'bar';
	const scatter = spec.type === 'scatter';
	const n = data.categories.length;
	const charW = fs * 0.6;
	// Points: X is the X values for scatter charts, the position (1, 2, 3...) otherwise.
	// Without a single number among the X values, scatter points are numbered 1, 2, 3...
	const numericX = data.xValues.some((v) => v !== null);
	const xOf = (i: number): number | null =>
		scatter && numericX ? (data.xValues[i] ?? null) : i + 1;

	const trends = data.series.map((s, si) =>
		hasTrend(spec, si)
			? linearTrend(
					s.values.flatMap((y, i) => {
						const x = xOf(i);
						return y === null || x === null ? [] : [{ x, y }];
					}),
				)
			: null,
	);
	const legend: LegendItem[] = data.series.map((s, i) => ({
		name: s.name,
		color: colorOf(spec, i),
	}));
	trends.forEach((t, i) => {
		if (!t) return;
		const of =
			data.series.length > 1 ? ` (${data.series[i]?.name ?? ''})` : '';
		const text = trendLabel(
			t,
			spec.showEquation !== false,
			spec.showR2 !== false,
		);
		legend.push({
			name: `Trendline${of}${text ? `: ${text}` : ''}`,
			color: trendColorOf(spec, i),
			line: true,
		});
	});
	const showLegend = spec.legend
		? spec.legend !== 'none'
		: data.series.length > 1 || trends.some((t) => t !== null);
	const box = showLegend
		? placeLegend(root, legend, spec.legend ?? 'bottom', area, fs)
		: area;

	const valueAxis = horizontal ? spec.xAxis : spec.yAxis;
	const catAxis = horizontal ? spec.yAxis : spec.xAxis;
	const values = data.series.flatMap((s) =>
		s.values.filter((v): v is number => v !== null),
	);
	let vMin = Math.min(...values);
	let vMax = Math.max(...values);
	if (spec.type === 'column' || spec.type === 'bar') {
		vMin = Math.min(0, vMin);
		vMax = Math.max(0, vMax);
	}
	const vScale = axisScale(vMin, vMax, 5, valueAxis);
	const vTicks = ticks(vScale);
	const xs = data.categories.flatMap((_c, i) => {
		const x = xOf(i);
		return x === null ? [] : [x];
	});
	const xScale =
		scatter && xs.length > 0
			? axisScale(Math.min(...xs), Math.max(...xs), 6, spec.xAxis)
			: null;

	const showV = valueAxis?.showLabels !== false;
	const showC = catAxis?.showLabels !== false;
	const xTitle = spec.xAxis?.title?.trim() ?? '';
	const yTitle = spec.yAxis?.title?.trim() ?? '';
	const tickW =
		Math.max(...vTicks.map((t) => formatNumber(t).length)) * charW + 8;
	const catW = Math.min(
		box.w * 0.35,
		Math.max(...data.categories.map((c) => Math.min(c.length, 16))) *
			charW +
			10,
	);
	const leftLabels = horizontal ? (showC ? catW : 4) : showV ? tickW : 4;
	const bottomLabels = (horizontal ? showV : showC) ? fs + 10 : 4;
	const left = box.x + leftLabels + (yTitle ? fs + 8 : 0) + 2;
	const top = box.y + 10;
	const plotW = Math.max(10, box.x + box.w - 12 - left);
	const plotH = Math.max(
		10,
		box.y + box.h - bottomLabels - (xTitle ? fs + 8 : 0) - top,
	);

	if (yTitle) {
		const x = box.x + fs / 2 + 4;
		const y = top + plotH / 2;
		label(
			root,
			x,
			y,
			truncate(yTitle, Math.floor(plotH / charW)),
			'middle',
			fs,
			'rtl-chart-axis-title',
		).setAttribute('transform', `rotate(-90 ${x} ${y})`);
	}
	if (xTitle)
		label(
			root,
			left + plotW / 2,
			box.y + box.h - fs / 2 - 2,
			truncate(xTitle, Math.floor(plotW / charW)),
			'middle',
			fs,
			'rtl-chart-axis-title',
		);

	const vSpan = vScale.max - vScale.min || 1;
	const valuePos = (v: number): number =>
		horizontal
			? left + ((v - vScale.min) / vSpan) * plotW
			: top + plotH - ((v - vScale.min) / vSpan) * plotH;
	const gridStyle = { stroke: spec.gridColor };

	for (const t of vTicks) {
		const p = valuePos(t);
		if (horizontal) {
			if (spec.gridlines !== false)
				svg(
					root,
					'line',
					{
						x1: p,
						y1: top,
						x2: p,
						y2: top + plotH,
						class: 'rtl-chart-grid',
					},
					gridStyle,
				);
			if (showV)
				label(
					root,
					p,
					top + plotH + fs / 2 + 4,
					formatNumber(t),
					'middle',
					fs,
				);
		} else {
			if (spec.gridlines !== false)
				svg(
					root,
					'line',
					{
						x1: left,
						y1: p,
						x2: left + plotW,
						y2: p,
						class: 'rtl-chart-grid',
					},
					gridStyle,
				);
			if (showV) label(root, left - 6, p, formatNumber(t), 'end', fs);
		}
	}
	// Axis line at zero (or at the edge when zero is out of range).
	const base = valuePos(Math.min(Math.max(0, vScale.min), vScale.max));
	if (horizontal)
		svg(root, 'line', {
			x1: base,
			y1: top,
			x2: base,
			y2: top + plotH,
			class: 'rtl-chart-axis',
		});
	else
		svg(root, 'line', {
			x1: left,
			y1: base,
			x2: left + plotW,
			y2: base,
			class: 'rtl-chart-axis',
		});

	const band = (horizontal ? plotH : plotW) / Math.max(1, n);
	const catPos = (i: number): number =>
		horizontal ? top + band * (i + 0.5) : left + band * (i + 0.5);
	const xSpan = xScale ? xScale.max - xScale.min || 1 : 1;
	const xPos = (x: number): number =>
		xScale ? left + ((x - xScale.min) / xSpan) * plotW : left;
	/** Where position X (1, 2, 3... or a scatter X value) is drawn along the category axis. */
	const alongPos = (x: number): number =>
		scatter
			? xPos(x)
			: horizontal
				? top + band * (x - 0.5)
				: left + band * (x - 0.5);

	if (xScale) {
		for (const t of ticks(xScale)) {
			if (spec.xGridlines)
				svg(
					root,
					'line',
					{
						x1: xPos(t),
						y1: top,
						x2: xPos(t),
						y2: top + plotH,
						class: 'rtl-chart-grid',
					},
					gridStyle,
				);
			if (showC)
				label(
					root,
					xPos(t),
					top + plotH + fs / 2 + 4,
					formatNumber(t),
					'middle',
					fs,
				);
		}
	} else if (showC) {
		const room = horizontal ? fs + 4 : charW * 10;
		const every = Math.max(1, Math.ceil(room / Math.max(1, band)));
		data.categories.forEach((c, i) => {
			if (i % every !== 0) return;
			if (horizontal)
				label(
					root,
					left - 6,
					catPos(i),
					truncate(c, Math.floor((catW - 10) / charW)),
					'end',
					fs,
				);
			else
				label(
					root,
					catPos(i),
					top + plotH + fs / 2 + 4,
					truncate(
						c,
						Math.max(3, Math.floor((band * every) / charW)),
					),
					'middle',
					fs,
				);
		});
	}

	// Marks are clipped to the plot, so a fixed axis min or max cuts them off cleanly.
	const clipId = `rtl-chart-clip-${++clipCounter}`;
	const clip = svg(svg(root, 'defs'), 'clipPath', { id: clipId });
	svg(clip, 'rect', {
		x: left - 6,
		y: top - 6,
		width: plotW + 12,
		height: plotH + 12,
	});
	const marks = svg(root, 'g', { 'clip-path': `url(#${clipId})` });

	const seriesCount = Math.max(1, data.series.length);
	const pointSize = Math.min(
		20,
		Math.max(1, spec.pointSize ?? (scatter ? 4 : 3)),
	);
	const lineWidth = Math.min(10, Math.max(1, spec.lineWidth ?? 2));
	data.series.forEach((s, si) => {
		const color = colorOf(spec, si);
		if (spec.type === 'column' || spec.type === 'bar') {
			const inner = band * 0.8;
			const barSize = inner / seriesCount;
			const a = valuePos(Math.max(Math.min(0, vScale.max), vScale.min));
			s.values.forEach((v, i) => {
				if (v === null) return;
				const start = catPos(i) - inner / 2 + si * barSize;
				const b = valuePos(v);
				const rect = horizontal
					? svg(
							marks,
							'rect',
							{
								x: Math.min(a, b),
								y: start,
								width: Math.max(1, Math.abs(b - a)),
								height: Math.max(1, barSize - 1),
							},
							{ fill: color },
						)
					: svg(
							marks,
							'rect',
							{
								x: start,
								y: Math.min(a, b),
								width: Math.max(1, barSize - 1),
								height: Math.max(1, Math.abs(b - a)),
							},
							{ fill: color },
						);
				tooltip(
					rect,
					`${s.name} – ${data.categories[i] ?? ''}: ${formatNumber(v)}`,
				);
			});
		} else {
			const points: { x: number; y: number; i: number }[] = [];
			s.values.forEach((v, i) => {
				const x = xOf(i);
				if (v === null || x === null) return;
				points.push({ x: alongPos(x), y: valuePos(v), i });
			});
			if (spec.type === 'line' && points.length > 1) {
				const line = svg(
					marks,
					'polyline',
					{
						points: points.map((p) => `${p.x},${p.y}`).join(' '),
						class: 'rtl-chart-line',
					},
					{ stroke: color },
				);
				line.style.strokeWidth = String(lineWidth);
			}
			for (const p of points) {
				const dot = svg(
					marks,
					'circle',
					{ cx: p.x, cy: p.y, r: pointSize },
					{ fill: color },
				);
				const v = s.values[p.i] ?? 0;
				tooltip(
					dot,
					scatter
						? `${s.name}: (${formatNumber(xOf(p.i) ?? 0)}, ${formatNumber(v)})`
						: `${s.name} – ${data.categories[p.i] ?? ''}: ${formatNumber(v)}`,
				);
			}
		}

		const t = trends[si];
		if (t) {
			const xsAll = s.values.flatMap((y, i) => {
				const x = xOf(i);
				return y === null || x === null ? [] : [x];
			});
			const x1 = Math.min(...xsAll);
			const x2 = Math.max(...xsAll);
			const line = svg(
				marks,
				'line',
				{
					x1: alongPos(x1),
					y1: valuePos(t.slope * x1 + t.intercept),
					x2: alongPos(x2),
					y2: valuePos(t.slope * x2 + t.intercept),
					class: 'rtl-chart-trend',
				},
				{ stroke: trendColorOf(spec, si) },
			);
			tooltip(line, `${s.name}: ${trendLabel(t)}`);
		}
	});
}

/* ------------------------------------------------------------------ */
/* Chart editor                                                         */
/* ------------------------------------------------------------------ */

/** Trendline color shown in the editor before one is picked. */
export const DEFAULT_TREND_COLOR = '#e8590c';

export interface ChartEditorHost {
	/** The side panel the editor is drawn in (the sheet stays usable next to it). */
	el: HTMLElement;
	/** The chart as it is now, or null when it was deleted. */
	get(): ChartSpec | null;
	/**
	 * Changes the chart right away. Changes with the same tag that follow each other quickly
	 * (typing a title) are one undo step.
	 */
	update(change: (spec: ChartSpec) => void, tag: string): void;
	close(): void;
	/** The ranges of the chart, outlined in the sheet while the editor is open. */
	highlight(ranges: SelRange[]): void;
	/** The series names (the slice names of a pie chart) as the chart shows them now. */
	seriesNames(spec: ChartSpec): string[];
}

type RangeField = 'x' | 'y';
type Section =
	'style' | 'titles' | 'series' | 'legend' | 'haxis' | 'vaxis' | 'grid';

/** "rgb(1, 2, 3)" or "#abc" (what a CSS color resolves to) as "#rrggbb", for color inputs. */
function cssColorToHex(probe: HTMLElement, css: string): string {
	if (/^#[0-9a-f]{6}$/i.test(css)) return css;
	probe.setCssProps({ color: css });
	const rgb = /(\d+)\D+(\d+)\D+(\d+)/.exec(getComputedStyle(probe).color);
	probe.setCssProps({ color: '' });
	if (!rgb) return '#888888';
	return `#${[rgb[1], rgb[2], rgb[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('')}`;
}

/**
 * The chart editor (like Google Sheets): a Setup tab for the type and the data, and a
 * Customize tab with the look. Every change shows on the chart at once; there is nothing to
 * confirm. While a range field is active, selecting cells in the sheet fills it.
 */
export class ChartEditor {
	private tab: 'setup' | 'customize' = 'setup';
	private readonly open = new Set<Section>(['style']);
	private readonly text: Record<RangeField, string> = { x: '', y: '' };
	/** The range field that selecting cells fills, if any. */
	private armed: RangeField | null = null;
	private inputs: Partial<Record<RangeField, HTMLInputElement>> = {};
	private errorEl: HTMLElement | null = null;
	/** Customize > Series: -1 for all series, or the index of one. */
	private seriesPick = -1;

	constructor(
		private readonly host: ChartEditorHost,
		isNew: boolean,
	) {
		const spec = host.get();
		if (spec) {
			this.syncText(spec);
			// A new chart is waiting for its data; an existing one must not change by accident.
			if (isNew)
				this.armed =
					spec.type === 'scatter' && !spec.xRange ? 'x' : 'y';
		}
		host.el.addEventListener('keydown', this.onKeyDown);
		this.render();
	}

	destroy(): void {
		this.host.el.removeEventListener('keydown', this.onKeyDown);
		this.host.el.empty();
	}

	/** Shows the chart again as it is now (after an undo, or a change from outside). */
	refresh(): void {
		const spec = this.host.get();
		if (!spec) {
			this.host.close();
			return;
		}
		this.syncText(spec);
		this.render();
	}

	/** Called when cells are selected in the sheet; true when the range went into a field. */
	pickRange(rg: SelRange): boolean {
		const field = this.armed;
		if (!field || this.tab !== 'setup') return false;
		this.text[field] = formatRange(rg);
		const input = this.inputs[field];
		if (input) input.value = this.text[field];
		this.applyRange(field, rg);
		// X first, then Y.
		if (field === 'x') this.arm('y');
		return true;
	}

	private syncText(spec: ChartSpec): void {
		this.text.x = spec.xRange ? formatRange(spec.xRange) : '';
		this.text.y = formatRange(spec.range);
	}

	private readonly onKeyDown = (e: KeyboardEvent): void => {
		if (e.key !== 'Escape') return;
		e.preventDefault();
		e.stopPropagation();
		this.host.close();
	};

	private set(tag: string, change: (spec: ChartSpec) => void): void {
		this.host.update(change, tag);
	}

	private arm(field: RangeField | null): void {
		this.armed = field;
		for (const [name, input] of Object.entries(this.inputs))
			input?.toggleClass('is-armed', name === field);
	}

	private applyRange(field: RangeField, rg: SelRange): void {
		this.set(`range-${field}`, (s) => {
			if (field === 'x') s.xRange = rg;
			else s.range = rg;
		});
		this.errorEl?.setText('');
		this.emitHighlight();
	}

	private emitHighlight(): void {
		const spec = this.host.get();
		const fields: RangeField[] =
			spec?.type === 'scatter' ? ['x', 'y'] : ['y'];
		this.host.highlight(
			fields.flatMap((f) => {
				const rg = parseRangeText(this.text[f]);
				return rg ? [rg] : [];
			}),
		);
	}

	private render(): void {
		const el = this.host.el;
		el.empty();
		this.inputs = {};
		const spec = this.host.get();
		if (!spec) return;

		const head = el.createDiv({ cls: 'rtl-chart-ed-head' });
		setIcon(head.createSpan({ cls: 'rtl-chart-ed-icon' }), 'chart-column');
		head.createSpan({ cls: 'rtl-chart-ed-heading', text: 'Chart editor' });
		const close = head.createEl('button', {
			cls: 'clickable-icon rtl-chart-ed-close',
			attr: { 'aria-label': 'Close' },
		});
		setIcon(close, 'x');
		close.addEventListener('click', () => this.host.close());

		const tabs = el.createDiv({ cls: 'rtl-chart-ed-tabs' });
		for (const [id, name] of [
			['setup', 'Setup'],
			['customize', 'Customize'],
		] as const) {
			const tab = tabs.createEl('button', {
				cls: `rtl-chart-ed-tab${this.tab === id ? ' is-active' : ''}`,
				text: name,
			});
			tab.addEventListener('click', () => {
				this.tab = id;
				this.render();
			});
		}

		const body = el.createDiv({ cls: 'rtl-chart-ed-body' });
		if (this.tab === 'setup') this.renderSetup(body, spec);
		else this.renderCustomize(body, spec);
		this.emitHighlight();
	}

	/* ---------------- Controls ---------------- */

	private row(parent: HTMLElement, name: string): HTMLElement {
		const row = parent.createDiv({ cls: 'rtl-chart-ed-row' });
		row.createDiv({ cls: 'rtl-chart-ed-label', text: name });
		return row.createDiv({ cls: 'rtl-chart-ed-control' });
	}

	private select<T extends string>(
		parent: HTMLElement,
		name: string,
		options: [T, string][],
		value: T,
		onChange: (value: T) => void,
	): void {
		const sel = this.row(parent, name).createEl('select', {
			cls: 'dropdown',
		});
		for (const [v, text] of options)
			sel.createEl('option', { value: v, text });
		sel.value = value;
		sel.addEventListener('change', () => onChange(sel.value as T));
	}

	private textInput(
		parent: HTMLElement,
		name: string,
		value: string,
		tag: string,
		onChange: (s: ChartSpec, value: string) => void,
	): void {
		const input = this.row(parent, name).createEl('input', {
			type: 'text',
			attr: { dir: 'auto' },
		});
		input.value = value;
		input.addEventListener('input', () =>
			this.set(tag, (s) => onChange(s, input.value)),
		);
	}

	private check(
		parent: HTMLElement,
		name: string,
		value: boolean,
		onChange: (value: boolean) => void,
	): void {
		const label = parent.createEl('label', { cls: 'rtl-chart-ed-check' });
		const box = label.createEl('input', { type: 'checkbox' });
		box.checked = value;
		label.createSpan({ text: name });
		box.addEventListener('change', () => onChange(box.checked));
	}

	private numberInput(
		parent: HTMLElement,
		name: string,
		value: number | undefined,
		tag: string,
		limits: { min?: number; max?: number; placeholder?: string },
		onChange: (s: ChartSpec, value: number | undefined) => void,
	): void {
		const input = this.row(parent, name).createEl('input', {
			type: 'number',
			attr: { placeholder: limits.placeholder ?? 'Auto' },
		});
		if (limits.min !== undefined) input.min = String(limits.min);
		if (limits.max !== undefined) input.max = String(limits.max);
		input.value = value === undefined ? '' : String(value);
		input.addEventListener('input', () => {
			const raw = input.value.trim();
			const n = Number(raw);
			if (raw !== '' && !Number.isFinite(n)) return;
			this.set(tag, (s) => onChange(s, raw === '' ? undefined : n));
		});
	}

	/** A color picker with a button that goes back to the automatic color. */
	private color(
		parent: HTMLElement,
		name: string,
		value: string | undefined,
		fallback: string,
		tag: string,
		onChange: (s: ChartSpec, value: string | undefined) => void,
	): void {
		const control = this.row(parent, name);
		const input = control.createEl('input', { type: 'color' });
		input.value = value ?? cssColorToHex(control, fallback);
		input.addEventListener('input', () =>
			this.set(tag, (s) => onChange(s, input.value)),
		);
		const reset = control.createEl('button', {
			cls: 'rtl-chart-ed-reset',
			text: 'Auto',
			attr: { title: 'Use the automatic color' },
		});
		reset.toggleClass('is-active', value === undefined);
		reset.addEventListener('click', () => {
			this.set(tag, (s) => onChange(s, undefined));
			this.render();
		});
	}

	private section(
		parent: HTMLElement,
		id: Section,
		name: string,
		build: (content: HTMLElement) => void,
	): void {
		const wrap = parent.createDiv({ cls: 'rtl-chart-ed-section' });
		const header = wrap.createEl('button', {
			cls: 'rtl-chart-ed-section-head',
		});
		const chevron = header.createSpan({ cls: 'rtl-chart-ed-chevron' });
		header.createSpan({ text: name });
		const content = wrap.createDiv({ cls: 'rtl-chart-ed-section-body' });
		const apply = (): void => {
			const isOpen = this.open.has(id);
			wrap.toggleClass('is-open', isOpen);
			setIcon(chevron, isOpen ? 'chevron-down' : 'chevron-right');
			content.toggle(isOpen);
		};
		header.addEventListener('click', () => {
			if (this.open.has(id)) this.open.delete(id);
			else this.open.add(id);
			apply();
		});
		build(content);
		apply();
	}

	/* ---------------- Setup ---------------- */

	private renderSetup(body: HTMLElement, spec: ChartSpec): void {
		const scatter = spec.type === 'scatter';
		this.select(
			body,
			'Chart type',
			CHART_TYPES.map((t) => [t.value, t.label]),
			spec.type,
			(value) => this.changeType(value),
		);
		if (scatter) {
			this.rangeField(body, 'x', 'X axis values');
			this.rangeField(body, 'y', 'Y axis values (a series per column)');
		} else {
			this.rangeField(
				body,
				'y',
				spec.type === 'pie'
					? 'Data range (labels, then values)'
					: 'Data range (labels, then a series per column)',
			);
		}
		body.createDiv({
			cls: 'rtl-chart-ed-hint',
			text: 'Click a range field, then select the cells in the sheet, or type the range.',
		});
		this.errorEl = body.createDiv({ cls: 'rtl-chart-ed-error' });
		this.check(body, 'Use the first row as names', spec.headerRow, (v) =>
			this.set('header', (s) => {
				s.headerRow = v;
			}),
		);
	}

	private rangeField(parent: HTMLElement, f: RangeField, name: string): void {
		const input = this.row(parent, name).createEl('input', {
			type: 'text',
			cls: 'rtl-chart-ed-range',
			attr: {
				placeholder: f === 'x' ? 'A2:A20' : 'B2:B20',
				dir: 'ltr',
				spellcheck: 'false',
				autocomplete: 'off',
				autocapitalize: 'off',
			},
		});
		input.value = this.text[f];
		input.toggleClass('is-armed', this.armed === f);
		input.addEventListener('focus', () => this.arm(f));
		input.addEventListener('input', () => {
			this.text[f] = input.value;
			const rg = parseRangeText(input.value);
			if (rg) this.applyRange(f, rg);
			else
				this.errorEl?.setText(
					input.value.trim()
						? 'Not a range yet: use the form A1:B10.'
						: '',
				);
		});
		this.inputs[f] = input;
	}

	/** Switching to or from a scatter chart splits the first column off as X, or joins it back. */
	private changeType(next: ChartType): void {
		const spec = this.host.get();
		if (!spec) return;
		const prev = spec.type;
		const x = spec.xRange;
		const y = spec.range;
		this.set('type', (s) => {
			s.type = next;
			if (next === 'scatter' && prev !== 'scatter') {
				if (!x && y.c2 > y.c1) {
					s.xRange = { ...y, c2: y.c1 };
					s.range = { ...y, c1: y.c1 + 1 };
				}
			} else if (next !== 'scatter' && prev === 'scatter') {
				if (
					x &&
					x.c1 === x.c2 &&
					x.c2 + 1 === y.c1 &&
					x.r1 === y.r1 &&
					x.r2 === y.r2
				)
					s.range = { ...y, c1: x.c1 };
				delete s.xRange;
			}
		});
		const updated = this.host.get();
		if (updated) this.syncText(updated);
		this.render();
	}

	/* ---------------- Customize ---------------- */

	private renderCustomize(body: HTMLElement, spec: ChartSpec): void {
		const pie = spec.type === 'pie';
		const scatter = spec.type === 'scatter';
		const horizontal = spec.type === 'bar';

		this.section(body, 'style', 'Chart style', (c) => {
			this.color(
				c,
				'Background color',
				spec.bg,
				'var(--background-primary)',
				'bg',
				(s, v) => {
					s.bg = v;
				},
			);
			this.numberInput(
				c,
				'Font size',
				spec.fontSize,
				'font',
				{ min: 6, max: 32, placeholder: '11' },
				(s, v) => {
					s.fontSize = v;
				},
			);
			if (spec.type === 'line' || scatter)
				this.numberInput(
					c,
					'Point size',
					spec.pointSize,
					'point',
					{ min: 1, max: 20, placeholder: scatter ? '4' : '3' },
					(s, v) => {
						s.pointSize = v;
					},
				);
			if (spec.type === 'line')
				this.numberInput(
					c,
					'Line thickness',
					spec.lineWidth,
					'line',
					{ min: 1, max: 10, placeholder: '2' },
					(s, v) => {
						s.lineWidth = v;
					},
				);
			if (pie)
				this.numberInput(
					c,
					'Donut hole (%)',
					spec.donutHole,
					'donut',
					{ min: 0, max: 90, placeholder: '0' },
					(s, v) => {
						s.donutHole = v;
					},
				);
		});

		this.section(body, 'titles', 'Chart and axis titles', (c) => {
			this.textInput(c, 'Chart title', spec.title, 'title', (s, v) => {
				s.title = v;
			});
			this.color(
				c,
				'Title color',
				spec.titleColor,
				'var(--text-normal)',
				'title-color',
				(s, v) => {
					s.titleColor = v;
				},
			);
			if (pie) return;
			this.textInput(
				c,
				'Horizontal axis title',
				spec.xAxis?.title ?? '',
				'x-title',
				(s, v) => {
					s.xAxis = { ...(s.xAxis ?? {}), title: v || undefined };
				},
			);
			this.textInput(
				c,
				'Vertical axis title',
				spec.yAxis?.title ?? '',
				'y-title',
				(s, v) => {
					s.yAxis = { ...(s.yAxis ?? {}), title: v || undefined };
				},
			);
		});

		this.section(body, 'series', pie ? 'Slices' : 'Series', (c) => {
			const names = this.host.seriesNames(spec);
			if (this.seriesPick >= names.length) this.seriesPick = -1;
			const pick = this.seriesPick;
			this.select(
				c,
				'Apply to',
				[
					['-1', pie ? 'All slices' : 'All series'],
					...names.map((n, i): [string, string] => [
						String(i),
						n || `${pie ? 'Slice' : 'Series'} ${i + 1}`,
					]),
				],
				String(pick),
				(v) => {
					this.seriesPick = Number(v);
					this.render();
				},
			);
			const setStyle = (
				s: ChartSpec,
				i: number,
				change: (st: SeriesStyle) => SeriesStyle,
			): void => {
				const list = (s.series ?? []).slice();
				while (list.length <= i) list.push({});
				list[i] = change({ ...(list[i] ?? {}) });
				s.series = list;
			};
			if (pick >= 0)
				this.color(
					c,
					'Color',
					spec.series?.[pick]?.color,
					seriesColor(pick),
					`color-${pick}`,
					(s, v) => setStyle(s, pick, (st) => ({ ...st, color: v })),
				);
			else
				c.createDiv({
					cls: 'rtl-chart-ed-hint',
					text: `Pick ${pie ? 'a slice' : 'a series'} above to change its color.`,
				});
			if (!scatter) return;

			const own = pick >= 0 ? spec.series?.[pick] : undefined;
			const trendOn = own?.trendline ?? spec.trendline;
			this.check(c, 'Trendline', trendOn, (v) => {
				this.set(`trend-${pick}`, (s) => {
					if (pick >= 0)
						setStyle(s, pick, (st) => ({ ...st, trendline: v }));
					else {
						s.trendline = v;
						// "All series" overrides what was set for single series.
						s.series = s.series?.map((st) => ({
							...st,
							trendline: undefined,
						}));
					}
				});
				this.render();
			});
			if (!trendOn) return;
			this.color(
				c,
				'Trendline color',
				pick >= 0 ? own?.trendColor : spec.trendColor,
				pick >= 0
					? (spec.trendColor ?? colorOf(spec, pick))
					: DEFAULT_TREND_COLOR,
				`trend-color-${pick}`,
				(s, v) => {
					if (pick >= 0)
						setStyle(s, pick, (st) => ({ ...st, trendColor: v }));
					else {
						s.trendColor = v;
						s.series = s.series?.map((st) => ({
							...st,
							trendColor: undefined,
						}));
					}
				},
			);
			this.check(
				c,
				'Show the equation',
				spec.showEquation !== false,
				(v) =>
					this.set('equation', (s) => {
						s.showEquation = v ? undefined : false;
					}),
			);
			this.check(c, 'Show R²', spec.showR2 !== false, (v) =>
				this.set('r2', (s) => {
					s.showR2 = v ? undefined : false;
				}),
			);
		});

		this.section(body, 'legend', 'Legend', (c) => {
			this.select<LegendPosition | ''>(
				c,
				'Position',
				[
					['', 'Auto'],
					['bottom', 'Bottom'],
					['top', 'Top'],
					['right', 'Right'],
					['none', 'None'],
				],
				spec.legend ?? '',
				(v) =>
					this.set('legend', (s) => {
						s.legend = v || undefined;
					}),
			);
		});
		if (pie) return;

		const axisSection = (
			id: 'haxis' | 'vaxis',
			name: string,
			which: 'xAxis' | 'yAxis',
			numeric: boolean,
		): void => {
			this.section(body, id, name, (c) => {
				const axis = spec[which];
				this.check(c, 'Show labels', axis?.showLabels !== false, (v) =>
					this.set(`${which}-labels`, (s) => {
						s[which] = {
							...(s[which] ?? {}),
							showLabels: v ? undefined : false,
						};
					}),
				);
				if (!numeric) return;
				this.numberInput(
					c,
					'Minimum',
					axis?.min,
					`${which}-min`,
					{},
					(s, v) => {
						s[which] = { ...(s[which] ?? {}), min: v };
					},
				);
				this.numberInput(
					c,
					'Maximum',
					axis?.max,
					`${which}-max`,
					{},
					(s, v) => {
						s[which] = { ...(s[which] ?? {}), max: v };
					},
				);
			});
		};
		axisSection('haxis', 'Horizontal axis', 'xAxis', horizontal || scatter);
		axisSection('vaxis', 'Vertical axis', 'yAxis', !horizontal);

		this.section(body, 'grid', 'Gridlines', (c) => {
			this.check(
				c,
				horizontal ? 'Vertical gridlines' : 'Horizontal gridlines',
				spec.gridlines !== false,
				(v) =>
					this.set('grid', (s) => {
						s.gridlines = v ? undefined : false;
					}),
			);
			if (scatter)
				this.check(c, 'Vertical gridlines', !!spec.xGridlines, (v) =>
					this.set('x-grid', (s) => {
						s.xGridlines = v || undefined;
					}),
				);
			this.color(
				c,
				'Gridline color',
				spec.gridColor,
				'var(--background-modifier-border)',
				'grid-color',
				(s, v) => {
					s.gridColor = v;
				},
			);
		});
	}
}

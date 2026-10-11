import ExcelJS from 'exceljs';
import { CellFormat, MergeInfo, Scalar, key } from './model';

/** Data read from an external file, positioned relative to its own top-left cell. */
export interface ImportedSheet {
	/** Cell contents; formula cells hold the formula text ("=A1+B1") as written in the file. */
	data: Scalar[][];
	/** The values the file stored for its formula cells, keyed "row,col" (empty when it stores none, like CSV). */
	results: Record<string, Scalar>;
	formats: Record<string, CellFormat>;
	merges: MergeInfo[];
	colWidths: (number | null)[];
	rowHeights: (number | null)[];
	/** True when the content (or the source worksheet) is right-to-left. */
	rtl: boolean;
}

const MAX_ROWS = 5000;
const MAX_COLS = 702;
/** Excel's default font size; treated as "use the plugin default". */
const EXCEL_DEFAULT_FONT_SIZE = 11;

/** Guesses the direction from the letters in the first cells (Hebrew/Arabic vs Latin). */
export function detectRtl(data: Scalar[][]): boolean {
	let rtl = 0;
	let ltr = 0;
	for (const row of data.slice(0, 200)) {
		for (const v of row) {
			if (typeof v !== 'string') continue;
			rtl += (v.match(/[\u0590-\u08FF]/g) ?? []).length;
			ltr += (v.match(/[A-Za-z]/g) ?? []).length;
		}
	}
	return rtl > 0 && rtl >= ltr;
}

/* ------------------------------------------------------------------ */
/* Dates and times (so imported cells look the way they do in Excel)    */
/* ------------------------------------------------------------------ */

interface FmtToken {
	kind: 'tok' | 'lit';
	text: string;
}

function tokenizeNumFmt(fmt: string): FmtToken[] {
	const section = fmt.split(';')[0] ?? '';
	const out: FmtToken[] = [];
	let i = 0;
	while (i < section.length) {
		const ch = section.charAt(i);
		if (ch === '"') {
			const end = section.indexOf('"', i + 1);
			out.push({
				kind: 'lit',
				text: section.slice(i + 1, end < 0 ? undefined : end),
			});
			i = end < 0 ? section.length : end + 1;
			continue;
		}
		if (ch === '\\') {
			out.push({ kind: 'lit', text: section.charAt(i + 1) });
			i += 2;
			continue;
		}
		if (ch === '_' || ch === '*') {
			i += 2; // padding / fill character
			continue;
		}
		if (ch === '[') {
			const end = section.indexOf(']', i);
			const inner = section.slice(i + 1, end < 0 ? undefined : end);
			i = end < 0 ? section.length : end + 1;
			// [h] / [mm] / [ss] (elapsed time) are shown like h / mm / ss; other brackets are ignored.
			if (/^(h+|m+|s+)$/i.test(inner))
				out.push({ kind: 'tok', text: inner.toLowerCase() });
			continue;
		}
		const m =
			/^(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|A\/P)/i.exec(
				section.slice(i),
			);
		if (m) {
			out.push({ kind: 'tok', text: m[0] });
			i += m[0].length;
			continue;
		}
		out.push({ kind: 'lit', text: ch });
		i++;
	}
	return out;
}

function isDateFormat(fmt: string | undefined): boolean {
	if (!fmt) return false;
	const stripped = (fmt.split(';')[0] ?? '').replace(
		/"[^"]*"|\\.|\[[^\]]*\]/g,
		'',
	);
	return /[ydhs]/i.test(stripped) && !/^[#0,.?%\s]+$/.test(stripped);
}

function usesUsDateOrder(): boolean {
	const lang =
		typeof navigator !== 'undefined' ? navigator.language : 'en-US';
	return /^en-US/i.test(lang);
}

function formatDateWithNumFmt(d: Date, fmtRaw: string | undefined): string {
	let fmt = fmtRaw ?? '';
	// Built-in "short date" formats depend on the system locale in Excel.
	if (fmt === 'mm-dd-yy') fmt = usesUsDateOrder() ? 'm/d/yy' : 'd/m/yy';
	if (fmt === 'm/d/yy h:mm')
		fmt = usesUsDateOrder() ? 'm/d/yy h:mm' : 'd/m/yy h:mm';
	if (!fmt || fmt === 'General') {
		const hasTime =
			d.getUTCHours() + d.getUTCMinutes() + d.getUTCSeconds() > 0;
		fmt = hasTime ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd';
	}

	const tokens = tokenizeNumFmt(fmt);
	const hasAmPm = tokens.some((t) => t.kind === 'tok' && /^a/i.test(t.text));
	const lang = typeof navigator !== 'undefined' ? navigator.language : 'en';
	const pad = (n: number, width = 2): string =>
		String(n).padStart(width, '0');

	const year = d.getUTCFullYear();
	const month = d.getUTCMonth() + 1;
	const day = d.getUTCDate();
	const hours = d.getUTCHours();
	const minutes = d.getUTCMinutes();
	const seconds = d.getUTCSeconds();
	const name = (opts: Intl.DateTimeFormatOptions): string =>
		new Intl.DateTimeFormat(lang, { ...opts, timeZone: 'UTC' }).format(d);

	// "m" / "mm" means minutes next to hours or seconds, and month otherwise.
	const toks = tokens
		.map((t, i) => ({ t, i }))
		.filter((x) => x.t.kind === 'tok');
	const minuteAt = new Set<number>();
	toks.forEach((x, k) => {
		if (!/^mm?$/i.test(x.t.text)) return;
		const prev = toks[k - 1]?.t.text ?? '';
		const next = toks[k + 1]?.t.text ?? '';
		if (/^h/i.test(prev) || /^s/i.test(next)) minuteAt.add(x.i);
	});

	return tokens
		.map((t, i) => {
			if (t.kind === 'lit') return t.text;
			const k = t.text.toLowerCase();
			switch (k) {
				case 'yyyy':
					return String(year);
				case 'yy':
					return pad(year % 100);
				case 'mmmm':
					return name({ month: 'long' });
				case 'mmm':
					return name({ month: 'short' });
				case 'mm':
					return minuteAt.has(i) ? pad(minutes) : pad(month);
				case 'm':
					return minuteAt.has(i) ? String(minutes) : String(month);
				case 'dddd':
					return name({ weekday: 'long' });
				case 'ddd':
					return name({ weekday: 'short' });
				case 'dd':
					return pad(day);
				case 'd':
					return String(day);
				case 'hh':
					return pad(hasAmPm ? hours % 12 || 12 : hours);
				case 'h':
					return String(hasAmPm ? hours % 12 || 12 : hours);
				case 'ss':
					return pad(seconds);
				case 's':
					return String(seconds);
				case 'am/pm':
					return hours < 12 ? 'AM' : 'PM';
				case 'a/p':
					return hours < 12 ? 'A' : 'P';
				default:
					return t.text;
			}
		})
		.join('');
}

/* ------------------------------------------------------------------ */
/* CSV                                                                  */
/* ------------------------------------------------------------------ */

export function parseCsv(text: string): ImportedSheet {
	const src = text.replace(/^\uFEFF/, '');
	const firstLine = src.split(/\r?\n/, 1)[0] ?? '';
	const count = (ch: string): number => firstLine.split(ch).length - 1;
	let delimiter = ',';
	if (count(';') > count(',') && count(';') >= count('\t')) delimiter = ';';
	else if (count('\t') > count(',')) delimiter = '\t';

	const rows: string[][] = [];
	let row: string[] = [];
	let field = '';
	let inQuotes = false;

	for (let i = 0; i < src.length; i++) {
		const ch = src.charAt(i);
		if (inQuotes) {
			if (ch === '"') {
				if (src.charAt(i + 1) === '"') {
					field += '"';
					i++;
				} else {
					inQuotes = false;
				}
			} else {
				field += ch;
			}
			continue;
		}
		if (ch === '"') {
			inQuotes = true;
		} else if (ch === delimiter) {
			row.push(field);
			field = '';
		} else if (ch === '\n') {
			row.push(field);
			rows.push(row);
			row = [];
			field = '';
		} else if (ch !== '\r') {
			field += ch;
		}
	}
	if (field !== '' || row.length > 0) {
		row.push(field);
		rows.push(row);
	}

	while (
		rows.length > 0 &&
		(rows[rows.length - 1] ?? []).every((c) => c === '')
	)
		rows.pop();
	if (rows.length > MAX_ROWS) rows.length = MAX_ROWS;

	const width = Math.min(MAX_COLS, Math.max(1, ...rows.map((r) => r.length)));
	const data: Scalar[][] = rows.map((r) => {
		const out: Scalar[] = [];
		for (let c = 0; c < width; c++) {
			const v = r[c];
			out.push(v === undefined || v === '' ? null : v);
		}
		return out;
	});

	return {
		data,
		results: {},
		formats: {},
		merges: [],
		colWidths: [],
		rowHeights: [],
		rtl: detectRtl(data),
	};
}

/* ------------------------------------------------------------------ */
/* XLSX                                                                 */
/* ------------------------------------------------------------------ */

interface LooseStyle {
	font?: {
		bold?: boolean;
		italic?: boolean;
		underline?: boolean | string;
		strike?: boolean;
		size?: number;
		color?: { argb?: string };
	};
	fill?: { type?: string; pattern?: string; fgColor?: { argb?: string } };
	alignment?: {
		horizontal?: string;
		vertical?: string;
		wrapText?: boolean;
		textRotation?: number | string;
	};
	border?: Partial<
		Record<
			'top' | 'right' | 'bottom' | 'left',
			{ style?: string; color?: { argb?: string } }
		>
	>;
}

function argbToHex(argb: string | undefined): string | undefined {
	if (!argb) return undefined;
	if (argb.length === 8) return `#${argb.slice(2).toLowerCase()}`;
	if (argb.length === 6) return `#${argb.toLowerCase()}`;
	return undefined;
}

/** The formula of a cell as "=…" (shared formulas come already adjusted to the cell), or null. */
function cellFormula(cell: ExcelJS.Cell): string | null {
	const formula = (cell as unknown as { formula?: unknown }).formula;
	return typeof formula === 'string' && formula.length > 0
		? `=${formula}`
		: null;
}

/** The value of a cell; for a formula cell, the result the file stored for it. */
function cellValueToScalar(cell: ExcelJS.Cell): Scalar {
	const v: unknown = cell.value;
	if (v === null || v === undefined) return null;
	if (v instanceof Date) return formatDateWithNumFmt(v, cell.numFmt);
	// A number with a date/time format is a date serial (days since 1899-12-30).
	if (typeof v === 'number' && isDateFormat(cell.numFmt)) {
		return formatDateWithNumFmt(
			new Date(Math.round((v - 25569) * 86400 * 1000)),
			cell.numFmt,
		);
	}
	if (
		typeof v === 'string' ||
		typeof v === 'number' ||
		typeof v === 'boolean'
	)
		return v;
	if (typeof v === 'object') {
		const o = v as Record<string, unknown>;
		if ('result' in o) {
			const r = o.result;
			if (typeof r === 'number' && isDateFormat(cell.numFmt)) {
				return formatDateWithNumFmt(
					new Date(Math.round((r - 25569) * 86400 * 1000)),
					cell.numFmt,
				);
			}
			if (
				typeof r === 'string' ||
				typeof r === 'number' ||
				typeof r === 'boolean'
			)
				return r;
			if (r instanceof Date) return formatDateWithNumFmt(r, cell.numFmt);
			if (typeof r === 'object' && r !== null) {
				const error = (r as { error?: unknown }).error;
				if (typeof error === 'string') return error;
			}
			return null;
		}
		if (Array.isArray(o.richText)) {
			return (o.richText as { text?: string }[])
				.map((t) => t.text ?? '')
				.join('');
		}
		if (typeof o.text === 'string') return o.text;
		if (typeof o.error === 'string') return o.error;
	}
	return null;
}

function rotationFromExcel(
	value: number | string | undefined,
): number | undefined {
	if (typeof value !== 'number' || value === 0) return undefined;
	if (value >= 1 && value <= 67) return 45;
	if (value >= 68 && value <= 90) return 90;
	if (value >= 91 && value <= 157) return -45;
	if (value >= 158 && value <= 180) return -90;
	return undefined;
}

function extractFormat(cell: ExcelJS.Cell): CellFormat | null {
	const st = cell as unknown as LooseStyle;
	const f: CellFormat = {};

	const font = st.font;
	if (font) {
		if (font.bold) f.bold = true;
		if (font.italic) f.italic = true;
		if (font.underline) f.underline = true;
		if (font.strike) f.strike = true;
		if (font.size && Math.round(font.size) !== EXCEL_DEFAULT_FONT_SIZE) {
			f.fontSize = Math.round((font.size * 4) / 3);
		}
		const color = argbToHex(font.color?.argb);
		if (color && color !== '#000000') f.textColor = color;
	}

	const fill = st.fill;
	if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
		const bg = argbToHex(fill.fgColor?.argb);
		if (bg) f.bgColor = bg;
	}

	const al = st.alignment;
	if (al) {
		if (
			al.horizontal === 'left' ||
			al.horizontal === 'center' ||
			al.horizontal === 'right'
		) {
			f.align = al.horizontal;
		}
		if (
			al.vertical === 'top' ||
			al.vertical === 'middle' ||
			al.vertical === 'bottom'
		) {
			f.valign = al.vertical;
		}
		if (al.wrapText) f.wrap = 'wrap';
		const rot = rotationFromExcel(al.textRotation);
		if (rot) f.rotate = rot;
	}

	const b = st.border;
	if (b) {
		const borders = {
			top: !!b.top?.style,
			right: !!b.right?.style,
			bottom: !!b.bottom?.style,
			left: !!b.left?.style,
		};
		if (borders.top || borders.right || borders.bottom || borders.left)
			f.borders = borders;
	}

	return Object.keys(f).length > 0 ? f : null;
}

function columnLettersToIndex(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase())
		n = n * 26 + (ch.charCodeAt(0) - 64);
	return n - 1;
}

/** Reads the first worksheet of an .xlsx file. */
export async function parseXlsx(buffer: ArrayBuffer): Promise<ImportedSheet> {
	const wb = new ExcelJS.Workbook();
	await wb.xlsx.load(buffer);
	const ws = wb.worksheets[0];
	if (!ws) throw new Error('The workbook has no worksheets.');

	const rows = Math.min(MAX_ROWS, Math.max(1, ws.rowCount));
	const cols = Math.min(MAX_COLS, Math.max(1, ws.columnCount));

	const data: Scalar[][] = Array.from({ length: rows }, () =>
		Array<Scalar>(cols).fill(null),
	);
	const formats: Record<string, CellFormat> = {};
	const results: Record<string, Scalar> = {};

	ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
		const r = rowNumber - 1;
		if (r >= rows) return;
		row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
			const c = colNumber - 1;
			if (c >= cols) return;

			const isSlaveOfMerge = cell.isMerged && cell.master !== cell;
			if (!isSlaveOfMerge) {
				const rowData = data[r];
				const formula = cellFormula(cell);
				if (rowData) rowData[c] = formula ?? cellValueToScalar(cell);
				if (formula) results[key(r, c)] = cellValueToScalar(cell);
			}
			const f = extractFormat(cell);
			if (f) formats[key(r, c)] = f;
		});
	});

	const merges: MergeInfo[] = [];
	const rawMerges =
		(ws as unknown as { model?: { merges?: string[] } }).model?.merges ??
		[];
	for (const ref of rawMerges) {
		const m = /^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/.exec(ref);
		if (!m) continue;
		const c1 = columnLettersToIndex(m[1] ?? '');
		const r1 = parseInt(m[2] ?? '0', 10) - 1;
		const c2 = columnLettersToIndex(m[3] ?? '');
		const r2 = parseInt(m[4] ?? '0', 10) - 1;
		if (r1 < 0 || c1 < 0 || r2 >= rows || c2 >= cols) continue;
		if (r2 > r1 || c2 > c1)
			merges.push({
				row: r1,
				col: c1,
				rowspan: r2 - r1 + 1,
				colspan: c2 - c1 + 1,
			});
	}

	const colWidths: (number | null)[] = [];
	for (let c = 1; c <= cols; c++) {
		const w = ws.getColumn(c).width;
		colWidths.push(
			typeof w === 'number' && w > 0 ? Math.round(w * 7) : null,
		);
	}
	const rowHeights: (number | null)[] = [];
	for (let r = 1; r <= rows; r++) {
		const h = ws.getRow(r).height;
		rowHeights.push(
			typeof h === 'number' && h > 0 ? Math.round(h / 0.75) : null,
		);
	}

	const view = (ws as unknown as { views?: { rightToLeft?: boolean }[] })
		.views?.[0];
	const rtl = view?.rightToLeft === true || detectRtl(data);

	return { data, results, formats, merges, colWidths, rowHeights, rtl };
}

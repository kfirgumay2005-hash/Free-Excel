import FormulaParser from 'fast-formula-parser';
import { Infix, Postfix } from 'fast-formula-parser/formulas/operators';
import { indexToColumnLabel } from './model';
import type { Scalar } from './model';

/* ------------------------------------------------------------------ */
/* Formula evaluation                                                  */
/* ------------------------------------------------------------------ */
/*
 * evaluateFormula is the single entry point for formulas. It receives the
 * raw formula text (e.g. "=SUM(A1:A3)") and a getCell(row, col) callback
 * (0-based, already evaluated and number-coerced).
 *
 * Engine order:
 *   1. fast-formula-parser (MIT) - ~280 Excel functions.
 *   2. Built-in mini evaluator (SUM, AVERAGE, COUNT, MIN, MAX, IF, basic
 *      arithmetic and comparisons) - used only if the library throws, so a
 *      library problem never breaks the sheet.
 */

type Val = Scalar | Scalar[];
type Token = { t: 'num' | 'str' | 'id' | 'op'; v: string };
/** `sheet` is the name of another sheet (as written in the formula), or undefined for the formula's own sheet. */
export type GetCell = (row: number, col: number, sheet?: string) => Scalar;

/** Sheet name given to the parser for the formula's own sheet, so "Sheet1!A1" still means another sheet. */
const OWN_SHEET = '#own-sheet#';
const otherSheet = (sheet: unknown): string | undefined =>
	typeof sheet === 'string' && sheet !== OWN_SHEET ? sheet : undefined;

const ERR_RE =
	/^#(?:DIV\/0!|N\/A|NAME\?|VALUE!|REF!|NUM!|NULL!|CYCLE!|ERROR!)$/;

/** Accepts ";" as argument separator (common in some locales) by converting it to ",". */
export function normalizeSeparators(src: string): string {
	let out = '';
	let inString = false;
	for (const ch of src) {
		if (ch === '"') inString = !inString;
		out += !inString && ch === ';' ? ',' : ch;
	}
	return out;
}

function roundResult(v: Scalar): Scalar {
	return typeof v === 'number' && Number.isFinite(v)
		? Number(v.toPrecision(12))
		: v;
}

/* ------------------------------------------------------------------ */
/* Fixes and additions to fast-formula-parser                           */
/* ------------------------------------------------------------------ */
/*
 * The library leaves out some everyday functions (they are added in EXTRA_FUNCTIONS) and
 * mishandles empty cells: an empty cell passed to a text function became the text
 * "[object Object]" (so LEN of an empty cell was 15), IF(A1, …) was always true, and an
 * empty cell was never equal to "". patchLibrary fixes those in place, once.
 */

/** An argument as the library hands it to a function. */
interface Arg {
	value?: unknown;
	isArray?: boolean;
	isRangeRef?: boolean;
	isCellRef?: boolean;
}
type LibFunction = (...args: (Arg | undefined)[]) => unknown;

const LibFormulaError = FormulaParser.FormulaError;

/** Cells hold errors as text such as "#DIV/0!"; the library needs its own error objects to pass them on. */
function toLibValue(v: Scalar): unknown {
	if (v === null) return undefined;
	return typeof v === 'string' && ERR_RE.test(v) ? new LibFormulaError(v) : v;
}

/** All values of an argument. An error in any of them is thrown, unless `keepErrors` is set. */
function argValues(arg: Arg | undefined, keepErrors = false): unknown[] {
	if (!arg) throw LibFormulaError.VALUE;
	const v = arg.value;
	const values: unknown[] = Array.isArray(v) ? v.flat(Infinity) : [v];
	if (!keepErrors)
		for (const x of values) if (x instanceof LibFormulaError) throw x;
	return values;
}

/** The numbers at the given positions; an error there is thrown, anything else is skipped. */
function numbersAt(values: unknown[], indexes: number[]): number[] {
	const out: number[] = [];
	for (const index of indexes) {
		const v = values[index];
		if (v instanceof LibFormulaError) throw v;
		if (typeof v === 'number') out.push(v);
	}
	return out;
}

const isBlank = (v: unknown): boolean =>
	v === null || v === undefined || v === '';

/** First value of an argument (a range gives its top-left cell). */
function argScalar(arg: Arg | undefined): unknown {
	return argValues(arg)[0];
}

function argText(arg: Arg | undefined): string {
	const v = argScalar(arg);
	if (typeof v === 'string') return v;
	if (typeof v === 'number') return String(v);
	if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
	return '';
}

function argNumber(arg: Arg | undefined): number {
	const v = argScalar(arg);
	if (isBlank(v)) return 0;
	if (typeof v === 'number') return v;
	if (typeof v === 'boolean') return Number(v);
	const n = Number(v);
	if (Number.isNaN(n)) throw LibFormulaError.VALUE;
	return n;
}

/** The numbers in the arguments. Like Excel, text, booleans and empty cells inside references are skipped. */
function argNumbers(args: (Arg | undefined)[]): number[] {
	const out: number[] = [];
	for (const arg of args) {
		const fromCells = !!(arg?.isRangeRef || arg?.isCellRef || arg?.isArray);
		for (const v of argValues(arg)) {
			if (typeof v === 'number') out.push(v);
			else if (fromCells || isBlank(v)) continue;
			else if (typeof v === 'boolean') out.push(Number(v));
			else if (!Number.isNaN(Number(v))) out.push(Number(v));
			else throw LibFormulaError.VALUE;
		}
	}
	return out;
}

const sameValue = (a: unknown, b: unknown): boolean =>
	typeof a === 'string' && typeof b === 'string'
		? a.toLowerCase() === b.toLowerCase()
		: a === b;

/** Turns "a*" or "?b" into a case-insensitive test of the whole text ("*" and "?" are wildcards). */
function wildcardRegExp(pattern: string): RegExp {
	const source = pattern
		.replace(/[.+^${}()|[\]\\]/g, '\\$&')
		.replace(/\*/g, '[^]*')
		.replace(/\?/g, '[^]');
	return new RegExp(`^${source}$`, 'i');
}

/** A SUMIFS / COUNTIFS condition such as ">5", "<>done", "a*" or 7. */
function criteriaTest(criteria: unknown): (v: unknown) => boolean {
	if (typeof criteria !== 'string')
		return (v) => (isBlank(criteria) ? isBlank(v) : v === criteria);
	const m = /^(<=|>=|<>|<|>|=)?([^]*)$/.exec(criteria);
	const op = m?.[1] ?? '=';
	const rhs = m?.[2] ?? '';
	const compare = (a: number | string, b: number | string): boolean =>
		op === '<' ? a < b : op === '>' ? a > b : op === '<=' ? a <= b : a >= b;

	if (rhs.trim() !== '' && !Number.isNaN(Number(rhs))) {
		const n = Number(rhs);
		if (op === '=') return (v) => v === n;
		if (op === '<>') return (v) => v !== n;
		return (v) => typeof v === 'number' && compare(v, n);
	}
	if (op === '=' || op === '<>') {
		const re = wildcardRegExp(rhs);
		return (v) =>
			(isBlank(v) ? rhs === '' : re.test(String(v))) ===
			(op === '=');
	}
	const text = rhs.toLowerCase();
	return (v) => typeof v === 'string' && compare(v.toLowerCase(), text);
}

/** Rows (as indexes into the flattened ranges) that satisfy every range/criteria pair. */
function matchingIndexes(pairs: (Arg | undefined)[]): number[] {
	if (pairs.length === 0 || pairs.length % 2 !== 0) throw LibFormulaError.VALUE;
	let indexes: number[] | null = null;
	for (let p = 0; p < pairs.length; p += 2) {
		const values = argValues(pairs[p], true);
		const test = criteriaTest(argScalar(pairs[p + 1]));
		const all: number[] = indexes ?? values.map((_, index) => index);
		indexes = all.filter((index) => test(values[index]));
	}
	return indexes ?? [];
}

function variance(args: (Arg | undefined)[]): number {
	const n = argNumbers(args);
	if (n.length < 2) throw LibFormulaError.DIV0;
	const mean = n.reduce((a, b) => a + b, 0) / n.length;
	return n.reduce((a, b) => a + (b - mean) ** 2, 0) / (n.length - 1);
}

/** The k-th value (1-based) of the sorted numbers. */
function kth(args: (Arg | undefined)[], descending: boolean): number {
	const n = argNumbers([args[0]]).sort((a, b) => a - b);
	if (descending) n.reverse();
	const picked = n[Math.trunc(argNumber(args[1])) - 1];
	if (picked === undefined) throw LibFormulaError.NUM;
	return picked;
}

/** Rounds to a number of decimal digits: half away from zero, or always away from / toward zero. */
function roundTo(
	number: Arg | undefined,
	digits: Arg | undefined,
	mode: 'half' | 'up' | 'down',
): number {
	const x = argNumber(number);
	const factor = 10 ** Math.trunc(argNumber(digits));
	// toPrecision removes binary noise, so 1.005 * 100 is 100.5 and not 100.49999999999999.
	const scaled = Number((Math.abs(x) * factor).toPrecision(15));
	const rounded =
		mode === 'up'
			? Math.ceil(scaled)
			: mode === 'down'
				? Math.floor(scaled)
				: Math.round(scaled);
	return (Math.sign(x) * rounded) / factor;
}

/**
 * VLOOKUP (vertical) and HLOOKUP. The library's versions returned #REF! when the first
 * result cell was empty, and the lookup key instead of the result past the end of the table.
 */
function lookup(args: (Arg | undefined)[], vertical: boolean): unknown {
	const [value, table, indexArg, approximateArg] = args;
	const raw = argScalar(value);
	const wanted: unknown = isBlank(raw) ? 0 : raw;
	const rows = table?.value;
	if (!Array.isArray(rows) || !Array.isArray(rows[0]))
		throw LibFormulaError.NA;
	const grid = rows as unknown[][];
	const index = Math.trunc(argNumber(indexArg));
	if (index < 1) throw LibFormulaError.VALUE;
	if (index > (vertical ? (grid[0]?.length ?? 0) : grid.length))
		throw LibFormulaError.REF;
	const approximate = approximateArg
		? isTrue(argScalar(approximateArg))
		: true;

	const keys: unknown[] = vertical
		? grid.map((row) => row[0])
		: (grid[0] ?? []);
	const fold = (v: unknown): unknown =>
		typeof v === 'string' ? v.toLowerCase() : v;
	const target = fold(wanted);
	let found = -1;
	if (approximate) {
		// The keys are sorted: take the last one that is not above the value.
		for (let i = 0; i < keys.length; i++) {
			const k = fold(keys[i]);
			if (isBlank(k) || typeof k !== typeof target) continue;
			if ((k as number | string) > (target as number | string)) break;
			found = i;
		}
	} else if (typeof wanted === 'string' && /[*?]/.test(wanted)) {
		const re = wildcardRegExp(wanted);
		found = keys.findIndex((k) => typeof k === 'string' && re.test(k));
	} else {
		found = keys.findIndex((k) => !isBlank(k) && fold(k) === target);
	}
	if (found < 0) throw LibFormulaError.NA;
	const result: unknown = vertical
		? grid[found]?.[index - 1]
		: grid[index - 1]?.[found];
	if (result instanceof LibFormulaError) throw result;
	return result ?? 0;
}

/** The library's own functions, for the wrappers below that only correct a special case. */
const LIBRARY_FUNCTIONS = new FormulaParser().functions;

/**
 * YEAR, MONTH and DAY of an empty cell (or 0): Excel calls that day 0 January 1900, the
 * library answered 31 December 1899.
 */
function datePart(name: string, ofDayZero: number): LibFunction {
	return (date) => {
		const v = argScalar(date);
		if (isBlank(v) || (typeof v === 'number' && Math.trunc(v) === 0))
			return ofDayZero;
		return LIBRARY_FUNCTIONS[name]?.(date);
	};
}

function isTrue(v: unknown): boolean {
	if (isBlank(v)) return false;
	if (typeof v === 'boolean') return v;
	if (typeof v === 'number') return v !== 0;
	const text = String(v).toUpperCase();
	if (text === 'TRUE') return true;
	if (text === 'FALSE') return false;
	throw LibFormulaError.VALUE;
}

/** Which AI a CLAUDE() or GEMINI() formula asks. */
export type AiFormulaProvider = 'claude' | 'gemini';

/**
 * Answers an AI formula: the cached answer, or a placeholder while the request runs. Set by
 * the view, which owns the requests and the cache; without it the functions are unknown.
 */
export type AiFormulaHandler = (
	provider: AiFormulaProvider,
	prompt: string,
	context: string,
) => Scalar;

let aiFormulaHandler: AiFormulaHandler | null = null;

/** Sets the handler for the formulas evaluated next; returns the one it replaces, to put back. */
export function setAiFormulaHandler(
	handler: AiFormulaHandler | null,
): AiFormulaHandler | null {
	const previous = aiFormulaHandler;
	aiFormulaHandler = handler;
	return previous;
}

/** An argument as text for an AI prompt: a range becomes tab-separated rows. */
function argAsPromptText(arg: Arg | undefined): string {
	const cell = (v: unknown): string => {
		if (isBlank(v)) return '';
		if (v instanceof LibFormulaError) {
			const error = (v as { error?: unknown }).error;
			return typeof error === 'string' ? error : '#VALUE!';
		}
		if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
		return String(v);
	};
	const v = arg?.value;
	if (!Array.isArray(v)) return cell(v);
	return v
		.map((row: unknown) =>
			Array.isArray(row) ? row.map(cell).join('\t') : cell(row),
		)
		.join('\n');
}

function aiFormula(
	provider: AiFormulaProvider,
	args: (Arg | undefined)[],
): unknown {
	if (!aiFormulaHandler) throw new LibFormulaError('#NAME?');
	const prompt = argAsPromptText(args[0]).trim();
	if (!prompt) throw LibFormulaError.VALUE;
	const context = args
		.slice(1)
		.map(argAsPromptText)
		.filter((t) => t !== '')
		.join('\n\n');
	return aiFormulaHandler(provider, prompt, context);
}

/** Functions the library lacks, or gets wrong for empty cells. */
const EXTRA_FUNCTIONS: Record<string, LibFunction> = {
	CLAUDE: (...args) => aiFormula('claude', args),
	GEMINI: (...args) => aiFormula('gemini', args),
	// The library's SUM skips an error in a cell instead of showing it.
	SUM: (...args) => argNumbers(args).reduce((a, b) => a + b, 0),
	// Excel gives 0 (the library gave 1) when there is no number to multiply.
	PRODUCT: (...args) => {
		const n = argNumbers(args);
		return n.length ? n.reduce((a, b) => a * b, 1) : 0;
	},
	VLOOKUP: (...args) => lookup(args, true),
	HLOOKUP: (...args) => lookup(args, false),
	YEAR: datePart('YEAR', 1900),
	MONTH: datePart('MONTH', 1),
	DAY: datePart('DAY', 0),
	// The library's ROUNDUP adds a step to numbers that are already round (5 became 6).
	ROUND: (number, digits) => roundTo(number, digits, 'half'),
	ROUNDUP: (number, digits) => roundTo(number, digits, 'up'),
	ROUNDDOWN: (number, digits) => roundTo(number, digits, 'down'),
	// Excel's TRIM also collapses repeated spaces between words.
	TRIM: (text) => argText(text).trim().replace(/ {2,}/g, ' '),
	// The library's SEARCH is case sensitive, which is what FIND is for.
	SEARCH: (find, text, start) => {
		const from = start ? Math.trunc(argNumber(start)) : 1;
		if (from < 1) throw LibFormulaError.VALUE;
		const at = argText(text)
			.toLowerCase()
			.indexOf(argText(find).toLowerCase(), from - 1);
		if (at < 0) throw LibFormulaError.VALUE;
		return at + 1;
	},
	AVERAGE: (...args) => {
		const n = argNumbers(args);
		if (n.length === 0) throw LibFormulaError.DIV0;
		return n.reduce((a, b) => a + b, 0) / n.length;
	},
	COUNTA: (...args) =>
		args.reduce(
			(count, arg) =>
				count + argValues(arg).filter((v) => !isBlank(v)).length,
			0,
		),
	MIN: (...args) => {
		const n = argNumbers(args);
		return n.length ? Math.min(...n) : 0;
	},
	MAX: (...args) => {
		const n = argNumbers(args);
		return n.length ? Math.max(...n) : 0;
	},
	MEDIAN: (...args) => {
		const n = argNumbers(args).sort((a, b) => a - b);
		const mid = n.length >> 1;
		const upper = n[mid];
		if (upper === undefined) throw LibFormulaError.NUM;
		return n.length % 2 ? upper : ((n[mid - 1] ?? upper) + upper) / 2;
	},
	LARGE: (...args) => kth(args, true),
	SMALL: (...args) => kth(args, false),
	VAR: (...args) => variance(args),
	STDEV: (...args) => Math.sqrt(variance(args)),
	SUMIFS: (sumRange, ...pairs) =>
		numbersAt(argValues(sumRange, true), matchingIndexes(pairs)).reduce(
			(a, b) => a + b,
			0,
		),
	COUNTIFS: (...pairs) => matchingIndexes(pairs).length,
	// The library's criteria are case sensitive ("b*" did not match "Banana").
	COUNTIF: (range, criteria) => matchingIndexes([range, criteria]).length,
	MATCH: (value, range, type) => {
		const lookup = argScalar(value);
		const values = argValues(range);
		const mode = type ? argNumber(type) : 1;
		let found = -1;
		if (mode === 0) {
			found = values.findIndex((v) => sameValue(v, lookup));
		} else {
			// 1: the last value that is not above the lookup; -1: the last that is not below it.
			const key = typeof lookup === 'string' ? lookup.toLowerCase() : lookup;
			values.forEach((v, index) => {
				const x = typeof v === 'string' ? v.toLowerCase() : v;
				if (typeof x !== typeof key || isBlank(x)) return;
				const a = x as number | string;
				const b = key as number | string;
				if (mode > 0 ? a <= b : a >= b) found = index;
			});
		}
		if (found < 0) throw LibFormulaError.NA;
		return found + 1;
	},
	// The library calls CHOOSE and IF with the parser as the first argument.
	CHOOSE: (_parser, index, ...values) => {
		const i = Math.trunc(argNumber(index));
		if (i < 1 || i > values.length) throw LibFormulaError.VALUE;
		return argScalar(values[i - 1]) ?? 0;
	},
	UPPER: (text) => argText(text).toUpperCase(),
	SUBSTITUTE: (text, oldText, newText, instance) => {
		const source = argText(text);
		const from = argText(oldText);
		const to = argText(newText);
		if (from === '') return source;
		if (!instance) return source.split(from).join(to);
		const nth = Math.trunc(argNumber(instance));
		if (nth < 1) throw LibFormulaError.VALUE;
		let at = -1;
		for (let n = 0; n < nth; n++) {
			at = source.indexOf(from, at + 1);
			if (at < 0) return source;
		}
		return source.slice(0, at) + to + source.slice(at + from.length);
	},
	VALUE: (text) => {
		const v = argScalar(text);
		if (typeof v === 'number') return v;
		const s = argText(text).trim().replace(/,/g, '');
		const percent = s.endsWith('%');
		const n = Number(percent ? s.slice(0, -1) : s);
		if (Number.isNaN(n)) throw LibFormulaError.VALUE;
		return percent ? n / 100 : n;
	},
	PMT: (rate, periods, presentValue, futureValue, type) => {
		const r = argNumber(rate);
		const n = argNumber(periods);
		const pv = argNumber(presentValue);
		const fv = futureValue ? argNumber(futureValue) : 0;
		const atStart = type ? argNumber(type) !== 0 : false;
		if (n === 0) throw LibFormulaError.NUM;
		if (r === 0) return -(pv + fv) / n;
		const growth = (1 + r) ** n;
		return (
			-(r * (pv * growth + fv)) / ((atStart ? 1 + r : 1) * (growth - 1))
		);
	},
};

/** The parser, as the library passes it to the functions that receive their arguments unread. */
interface LibContext {
	utils: { extractRefValue(v: unknown): { val: unknown; isArray: boolean } };
}
interface LibCell {
	row: number;
	col: number;
}
/** An unread argument: a reference to a cell or a range, or a plain value. */
type RawArg =
	{ ref?: Partial<LibCell> & { from?: LibCell; to?: LibCell } } | null | undefined;

function readArg(context: LibContext, raw: unknown): Arg | undefined {
	if (raw === undefined) return undefined;
	const read = context.utils.extractRefValue(raw);
	return { value: read.val, isArray: read.isArray, isRangeRef: true };
}

/**
 * SUMIF and AVERAGEIF: the numbers to add are the cells of the sum range (which starts at
 * its top-left cell and takes the shape of the criteria range) whose criteria cell matches.
 */
function conditionalNumbers(
	context: LibContext,
	range: unknown,
	criteria: unknown,
	sumRange: unknown,
): number[] {
	const tested = readArg(context, range);
	const indexes = matchingIndexes([tested, readArg(context, criteria)]);
	let source = tested;
	const from = (range as RawArg)?.ref;
	const to = (sumRange as RawArg)?.ref;
	if (to) {
		const start: Partial<LibCell> = to.from ?? to;
		const height = from?.from && from.to ? from.to.row - from.from.row : 0;
		const width = from?.from && from.to ? from.to.col - from.from.col : 0;
		if (start.row === undefined || start.col === undefined)
			throw LibFormulaError.VALUE;
		source = readArg(context, {
			ref: {
				from: { row: start.row, col: start.col },
				to: { row: start.row + height, col: start.col + width },
			},
		});
	}
	return numbersAt(argValues(source, true), indexes);
}

/** Functions the library calls with the parser first and the arguments unread. */
const CONTEXT_FUNCTIONS: Record<
	string,
	(context: LibContext, ...args: unknown[]) => unknown
> = {
	// The library treated every reference in the condition as TRUE.
	IF: (context, test, whenTrue, whenFalse) => {
		if (test === undefined || whenTrue === undefined)
			throw LibFormulaError.VALUE;
		if (isTrue(argScalar(readArg(context, test)))) return whenTrue ?? 0;
		return whenFalse === undefined ? false : (whenFalse ?? 0);
	},
	// The library's criteria are case sensitive ("b*" did not match "Banana").
	SUMIF: (context, range, criteria, sumRange) =>
		conditionalNumbers(context, range, criteria, sumRange).reduce(
			(a, b) => a + b,
			0,
		),
	AVERAGEIF: (context, range, criteria, averageRange) => {
		const n = conditionalNumbers(context, range, criteria, averageRange);
		if (n.length === 0) throw LibFormulaError.DIV0;
		return n.reduce((a, b) => a + b, 0) / n.length;
	},
};

let libraryPatched = false;

function patchLibrary(): void {
	if (libraryPatched) return;
	libraryPatched = true;

	// An empty cell is 0, "" or FALSE, depending on what the function asks for.
	const helpers = FormulaParser.FormulaHelpers;
	const types = FormulaParser.Types;
	const accept = helpers.accept.bind(helpers);
	helpers.accept = (param: unknown, type?: unknown, ...rest: unknown[]) => {
		const kind = Array.isArray(type) ? (type[0] as unknown) : type;
		if (
			typeof param === 'object' &&
			param !== null &&
			!Array.isArray(param) &&
			'isArray' in param &&
			isBlank((param as Arg).value)
		) {
			if (kind === undefined || kind === null) return 0;
			if (kind === types.STRING) return '';
			if (kind === types.BOOLEAN) return false;
			if (kind === types.NUMBER || kind === types.NUMBER_NO_BOOLEAN)
				return 0;
		}
		return accept(param, type, ...rest);
	};

	// Comparisons: an empty cell equals "" and FALSE (it already equals 0), and text ignores case.
	const compareOp = Infix.compareOp;
	Infix.compareOp = (a, op, b, aIsArray, bIsArray) => {
		if (!aIsArray && !bIsArray) {
			const blankLike = (other: unknown): unknown =>
				typeof other === 'string'
					? ''
					: typeof other === 'boolean'
						? false
						: 0;
			if (a === null || a === undefined) a = blankLike(b);
			if (b === null || b === undefined) b = blankLike(a);
			if (typeof a === 'string' && typeof b === 'string') {
				a = a.toLowerCase();
				b = b.toLowerCase();
			}
		}
		return compareOp(a, op, b, aIsArray, bIsArray);
	};

	// "=A1%" with an empty A1.
	const percentOp = Postfix.percentOp;
	Postfix.percentOp = (value, postfix, isArray) =>
		percentOp(value ?? 0, postfix, isArray);
}

/** Where the formula lives (0-based); ROW() and COLUMN() without an argument report it. */
export interface CellPosition {
	row: number;
	col: number;
}

function evaluateWithLibrary(
	formula: string,
	getCell: GetCell,
	at?: CellPosition,
): Scalar {
	patchLibrary();
	const parser = new FormulaParser({
		functions: EXTRA_FUNCTIONS,
		functionsNeedContext: CONTEXT_FUNCTIONS,
		onCell: (ref: { sheet?: string; row: number; col: number }) =>
			toLibValue(getCell(ref.row - 1, ref.col - 1, otherSheet(ref.sheet))),
		onRange: (ref: {
			sheet?: string;
			from: { row: number; col: number };
			to: { row: number; col: number };
		}) => {
			const sheet = otherSheet(ref.sheet);
			const rows: unknown[][] = [];
			for (let r = ref.from.row; r <= ref.to.row; r++) {
				const row: unknown[] = [];
				for (let c = ref.from.col; c <= ref.to.col; c++)
					row.push(toLibValue(getCell(r - 1, c - 1, sheet)));
				rows.push(row);
			}
			return rows;
		},
	});

	const parsed: unknown = parser.parse(
		normalizeSeparators(formula.replace(/^\s*=/, '')),
		{
			sheet: OWN_SHEET,
			row: (at?.row ?? 0) + 1,
			col: (at?.col ?? 0) + 1,
		},
	);

	let out: unknown = parsed;
	if (Array.isArray(out)) out = Array.isArray(out[0]) ? out[0][0] : out[0];
	if (out === undefined || out === null) return null;
	if (typeof out === 'object') {
		const err = (out as { error?: unknown }).error;
		return typeof err === 'string' ? err : '#VALUE!';
	}
	if (
		typeof out === 'number' ||
		typeof out === 'string' ||
		typeof out === 'boolean'
	)
		return roundResult(out);
	return '#VALUE!';
}

function colToIndex(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase())
		n = n * 26 + (ch.charCodeAt(0) - 64);
	return n - 1;
}

function toNum(v: Val): number {
	const x: Scalar | undefined = Array.isArray(v) ? v[0] : v;
	if (x === undefined || x === null || x === '') return 0;
	if (typeof x === 'number') return x;
	if (typeof x === 'boolean') return x ? 1 : 0;
	if (ERR_RE.test(x)) throw new Error(x);
	const n = Number(x);
	if (!Number.isNaN(n)) return n;
	throw new Error('#VALUE!');
}

function toBool(v: Val): boolean {
	const x = Array.isArray(v) ? v[0] : v;
	if (typeof x === 'string') {
		if (ERR_RE.test(x)) throw new Error(x);
		return x !== '' && x.toUpperCase() !== 'FALSE' && Number(x) !== 0;
	}
	return !!x;
}

function toText(v: Val): string {
	const x = Array.isArray(v) ? v[0] : v;
	return x === null || x === undefined ? '' : String(x);
}

function tokenize(src: string): Token[] {
	const re =
		/\s*(?:(\d+\.?\d*)|"((?:[^"]|"")*)"|([A-Za-z_][A-Za-z0-9_]*)|(<>|<=|>=|[-+*/^=<>(),:;]))/y;
	const tokens: Token[] = [];
	let pos = 0;
	while (pos < src.length) {
		if (/^\s*$/.test(src.slice(pos))) break;
		re.lastIndex = pos;
		const m = re.exec(src);
		if (!m) throw new Error('#ERROR!');
		if (m[1] !== undefined) tokens.push({ t: 'num', v: m[1] });
		else if (m[2] !== undefined)
			tokens.push({ t: 'str', v: m[2].replace(/""/g, '"') });
		else if (m[3] !== undefined) tokens.push({ t: 'id', v: m[3] });
		else tokens.push({ t: 'op', v: m[4] ?? '' });
		pos = re.lastIndex;
	}
	return tokens;
}

function numericValues(args: Val[]): number[] {
	const out: number[] = [];
	for (const a of args) {
		if (Array.isArray(a)) {
			for (const x of a) {
				if (typeof x === 'string' && ERR_RE.test(x)) throw new Error(x);
				if (typeof x === 'number') out.push(x);
			}
		} else {
			out.push(toNum(a));
		}
	}
	return out;
}

function callFunction(name: string, args: Val[]): Scalar {
	switch (name) {
		case 'SUM':
			return numericValues(args).reduce((a, b) => a + b, 0);
		case 'AVERAGE': {
			const n = numericValues(args);
			if (n.length === 0) throw new Error('#DIV/0!');
			return n.reduce((a, b) => a + b, 0) / n.length;
		}
		case 'COUNT': {
			let count = 0;
			for (const a of args) {
				if (Array.isArray(a))
					count += a.filter((x) => typeof x === 'number').length;
				else if (typeof a === 'number') count++;
			}
			return count;
		}
		case 'MIN': {
			const n = numericValues(args);
			return n.length ? Math.min(...n) : 0;
		}
		case 'MAX': {
			const n = numericValues(args);
			return n.length ? Math.max(...n) : 0;
		}
		case 'IF': {
			if (args.length < 2) throw new Error('#ERROR!');
			const cond: Val = args[0] ?? null;
			const picked: Val = toBool(cond)
				? (args[1] ?? null)
				: args.length > 2
					? (args[2] ?? null)
					: false;
			return Array.isArray(picked) ? (picked[0] ?? null) : picked;
		}
		default:
			throw new Error('#NAME?');
	}
}

function evaluateBuiltin(formula: string, getCell: GetCell): Scalar {
	try {
		const tokens = tokenize(formula.replace(/^\s*=/, ''));
		let i = 0;
		const isOp = (v: string): boolean => {
			const tk = tokens[i];
			return tk !== undefined && tk.t === 'op' && tk.v === v;
		};
		const expectOp = (v: string): void => {
			if (!isOp(v)) throw new Error('#ERROR!');
			i++;
		};
		const isRef = (s: string): boolean => /^[A-Za-z]{1,3}[0-9]+$/.test(s);
		const refToCoords = (s: string): [number, number] => {
			const m = /^([A-Za-z]+)([0-9]+)$/.exec(s);
			const letters = m?.[1];
			const digits = m?.[2];
			if (!letters || !digits) throw new Error('#REF!');
			return [parseInt(digits, 10) - 1, colToIndex(letters)];
		};

		const parseComparison = (): Val => {
			const left = parseAdditive();
			const cmp = ['<>', '<=', '>=', '=', '<', '>'].find((op) =>
				isOp(op),
			);
			if (!cmp) return left;
			i++;
			const right = parseAdditive();
			const l = Array.isArray(left) ? (left[0] ?? null) : left;
			const r = Array.isArray(right) ? (right[0] ?? null) : right;
			const bothNumeric = typeof l !== 'string' && typeof r !== 'string';
			const a: number | string = bothNumeric
				? toNum(l)
				: toText(l).toLowerCase();
			const b: number | string = bothNumeric
				? toNum(r)
				: toText(r).toLowerCase();
			switch (cmp) {
				case '=':
					return a === b;
				case '<>':
					return a !== b;
				case '<':
					return a < b;
				case '>':
					return a > b;
				case '<=':
					return a <= b;
				default:
					return a >= b;
			}
		};

		const parseAdditive = (): Val => {
			let left = parseTerm();
			while (isOp('+') || isOp('-')) {
				const op = tokens[i++]?.v ?? '';
				const right = parseTerm();
				left =
					op === '+'
						? toNum(left) + toNum(right)
						: toNum(left) - toNum(right);
			}
			return left;
		};

		const parseTerm = (): Val => {
			let left = parsePower();
			while (isOp('*') || isOp('/')) {
				const op = tokens[i++]?.v ?? '';
				const right = parsePower();
				if (op === '*') left = toNum(left) * toNum(right);
				else {
					const d = toNum(right);
					if (d === 0) throw new Error('#DIV/0!');
					left = toNum(left) / d;
				}
			}
			return left;
		};

		const parsePower = (): Val => {
			let base = parseUnary();
			while (isOp('^')) {
				i++;
				base = Math.pow(toNum(base), toNum(parseUnary()));
			}
			return base;
		};

		const parseUnary = (): Val => {
			if (isOp('-')) {
				i++;
				return -toNum(parseUnary());
			}
			if (isOp('+')) {
				i++;
				return toNum(parseUnary());
			}
			return parsePrimary();
		};

		const parsePrimary = (): Val => {
			const tok = tokens[i++];
			if (!tok) throw new Error('#ERROR!');
			if (tok.t === 'num') return parseFloat(tok.v);
			if (tok.t === 'str') return tok.v;
			if (tok.t === 'op') {
				if (tok.v !== '(') throw new Error('#ERROR!');
				const inner = parseComparison();
				expectOp(')');
				return inner;
			}
			const upper = tok.v.toUpperCase();
			if (isOp('(')) {
				i++;
				const args: Val[] = [];
				if (!isOp(')')) {
					for (;;) {
						args.push(parseComparison());
						if (isOp(',') || isOp(';')) {
							i++;
							continue;
						}
						break;
					}
				}
				expectOp(')');
				return callFunction(upper, args);
			}
			if (upper === 'TRUE') return true;
			if (upper === 'FALSE') return false;
			if (isRef(tok.v)) {
				const [r1, c1] = refToCoords(tok.v);
				if (isOp(':')) {
					i++;
					const end = tokens[i++];
					if (!end || end.t !== 'id' || !isRef(end.v))
						throw new Error('#REF!');
					const [r2, c2] = refToCoords(end.v);
					const out: Scalar[] = [];
					for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
						for (
							let c = Math.min(c1, c2);
							c <= Math.max(c1, c2);
							c++
						)
							out.push(getCell(r, c));
					}
					return out;
				}
				return getCell(r1, c1);
			}
			throw new Error('#NAME?');
		};

		const result = parseComparison();
		if (i < tokens.length) throw new Error('#ERROR!');
		const out = Array.isArray(result) ? (result[0] ?? null) : result;
		return roundResult(out);
	} catch (e) {
		return e instanceof Error && ERR_RE.test(e.message)
			? e.message
			: '#ERROR!';
	}
}

export function evaluateFormula(
	formula: string,
	getCell: GetCell,
	at?: CellPosition,
): Scalar {
	try {
		return evaluateWithLibrary(formula, getCell, at);
	} catch (e) {
		console.debug(
			'[Free-Excel] fast-formula-parser failed, using built-in evaluator',
			e,
		);
		return evaluateBuiltin(formula, getCell);
	}
}

/* ------------------------------------------------------------------ */
/* Function catalog (used by the function picker and autocomplete)      */
/* ------------------------------------------------------------------ */

export interface FormulaFunctionInfo {
	name: string;
	syntax: string;
	description: string;
}

const fn = (
	name: string,
	syntax: string,
	description: string,
): FormulaFunctionInfo => ({
	name,
	syntax,
	description,
});

/** Most commonly used first. Anything the formula engine does not know shows #NAME? in the cell. */
export const FORMULA_FUNCTIONS: FormulaFunctionInfo[] = [
	fn('SUM', 'SUM(number1, [number2], …)', 'Adds numbers or ranges'),
	fn('AVERAGE', 'AVERAGE(number1, [number2], …)', 'Average of numbers'),
	fn(
		'COUNT',
		'COUNT(value1, [value2], …)',
		'Counts cells that contain numbers',
	),
	fn(
		'COUNTA',
		'COUNTA(value1, [value2], …)',
		'Counts cells that are not empty',
	),
	fn('MIN', 'MIN(number1, [number2], …)', 'Smallest value'),
	fn('MAX', 'MAX(number1, [number2], …)', 'Largest value'),
	fn(
		'IF',
		'IF(condition, value_if_true, [value_if_false])',
		'Returns one value if true, another if false',
	),
	fn(
		'IFERROR',
		'IFERROR(value, value_if_error)',
		'Replaces an error with a value',
	),
	fn(
		'AND',
		'AND(condition1, [condition2], …)',
		'True if all conditions are true',
	),
	fn(
		'OR',
		'OR(condition1, [condition2], …)',
		'True if any condition is true',
	),
	fn('NOT', 'NOT(condition)', 'Reverses true and false'),
	fn('ROUND', 'ROUND(number, digits)', 'Rounds a number to given digits'),
	fn('ROUNDUP', 'ROUNDUP(number, digits)', 'Rounds up, away from zero'),
	fn('ROUNDDOWN', 'ROUNDDOWN(number, digits)', 'Rounds down, toward zero'),
	fn('ABS', 'ABS(number)', 'Absolute value'),
	fn('SQRT', 'SQRT(number)', 'Square root'),
	fn('POWER', 'POWER(number, power)', 'Raises a number to a power'),
	fn('MOD', 'MOD(number, divisor)', 'Remainder after division'),
	fn('INT', 'INT(number)', 'Rounds down to the nearest integer'),
	fn('PRODUCT', 'PRODUCT(number1, [number2], …)', 'Multiplies numbers'),
	fn('MEDIAN', 'MEDIAN(number1, [number2], …)', 'Median value'),
	fn(
		'SUMIF',
		'SUMIF(range, criteria, [sum_range])',
		'Sums cells that match a condition',
	),
	fn(
		'SUMIFS',
		'SUMIFS(sum_range, range1, criteria1, …)',
		'Sums cells that match several conditions',
	),
	fn(
		'COUNTIF',
		'COUNTIF(range, criteria)',
		'Counts cells that match a condition',
	),
	fn(
		'COUNTIFS',
		'COUNTIFS(range1, criteria1, …)',
		'Counts cells that match several conditions',
	),
	fn(
		'AVERAGEIF',
		'AVERAGEIF(range, criteria, [average_range])',
		'Average of cells that match a condition',
	),
	fn(
		'SUMPRODUCT',
		'SUMPRODUCT(array1, [array2], …)',
		'Sum of products of ranges',
	),
	fn(
		'VLOOKUP',
		'VLOOKUP(value, table, column, [exact])',
		'Looks up a value in the first column of a table',
	),
	fn(
		'HLOOKUP',
		'HLOOKUP(value, table, row, [exact])',
		'Looks up a value in the first row of a table',
	),
	fn(
		'INDEX',
		'INDEX(range, row, [column])',
		'Returns the value at a position in a range',
	),
	fn(
		'MATCH',
		'MATCH(value, range, [type])',
		'Position of a value in a range',
	),
	fn(
		'CHOOSE',
		'CHOOSE(index, value1, [value2], …)',
		'Picks a value by position',
	),
	fn('CONCATENATE', 'CONCATENATE(text1, [text2], …)', 'Joins text together'),
	fn('LEFT', 'LEFT(text, [count])', 'First characters of text'),
	fn('RIGHT', 'RIGHT(text, [count])', 'Last characters of text'),
	fn('MID', 'MID(text, start, count)', 'Characters from the middle of text'),
	fn('LEN', 'LEN(text)', 'Length of text'),
	fn('UPPER', 'UPPER(text)', 'Converts to upper case'),
	fn('LOWER', 'LOWER(text)', 'Converts to lower case'),
	fn('PROPER', 'PROPER(text)', 'Capitalizes each word'),
	fn('TRIM', 'TRIM(text)', 'Removes extra spaces'),
	fn(
		'SUBSTITUTE',
		'SUBSTITUTE(text, old, new, [instance])',
		'Replaces text inside text',
	),
	fn(
		'FIND',
		'FIND(find, text, [start])',
		'Position of text (case sensitive)',
	),
	fn(
		'SEARCH',
		'SEARCH(find, text, [start])',
		'Position of text (not case sensitive)',
	),
	fn('REPT', 'REPT(text, times)', 'Repeats text'),
	fn('TEXT', 'TEXT(value, format)', 'Formats a number as text'),
	fn('VALUE', 'VALUE(text)', 'Converts text to a number'),
	fn('TODAY', 'TODAY()', "Today's date"),
	fn('NOW', 'NOW()', 'Current date and time'),
	fn('DATE', 'DATE(year, month, day)', 'Builds a date'),
	fn('YEAR', 'YEAR(date)', 'Year of a date'),
	fn('MONTH', 'MONTH(date)', 'Month of a date'),
	fn('DAY', 'DAY(date)', 'Day of a date'),
	fn('WEEKDAY', 'WEEKDAY(date, [type])', 'Day of the week'),
	fn('LARGE', 'LARGE(range, k)', 'The k-th largest value'),
	fn('SMALL', 'SMALL(range, k)', 'The k-th smallest value'),
	fn(
		'STDEV',
		'STDEV(number1, [number2], …)',
		'Standard deviation of a sample',
	),
	fn('VAR', 'VAR(number1, [number2], …)', 'Variance of a sample'),
	fn('RAND', 'RAND()', 'Random number between 0 and 1'),
	fn(
		'RANDBETWEEN',
		'RANDBETWEEN(bottom, top)',
		'Random integer between two numbers',
	),
	fn('PI', 'PI()', 'The number pi'),
	fn('CEILING', 'CEILING(number, significance)', 'Rounds up to a multiple'),
	fn('FLOOR', 'FLOOR(number, significance)', 'Rounds down to a multiple'),
	fn('ISBLANK', 'ISBLANK(value)', 'True if the cell is empty'),
	fn('ISNUMBER', 'ISNUMBER(value)', 'True if the value is a number'),
	fn('ISTEXT', 'ISTEXT(value)', 'True if the value is text'),
	fn('ISERROR', 'ISERROR(value)', 'True if the value is an error'),
	fn('ROW', 'ROW([reference])', 'Row number of a reference'),
	fn('COLUMN', 'COLUMN([reference])', 'Column number of a reference'),
	fn('PMT', 'PMT(rate, periods, present_value)', 'Loan payment per period'),
	fn(
		'CLAUDE',
		'CLAUDE(prompt, [cells], …)',
		'Asks Claude; the answer goes in the cell (needs an API key)',
	),
	fn(
		'GEMINI',
		'GEMINI(prompt, [cells], …)',
		'Asks Gemini; the answer goes in the cell (needs an API key)',
	),
];

/* ------------------------------------------------------------------ */
/* Picking a reference by clicking a cell                               */
/* ------------------------------------------------------------------ */

const REFERENCE_SRC = '\\$?[A-Za-z]{1,3}\\$?\\d+(?::\\$?[A-Za-z]{1,3}\\$?\\d+)?';
/**
 * Strings, references (group 1), names (functions, TRUE/FALSE) and numbers. Names come
 * before numbers, and a reference is not followed by "(", so LOG10(…) stays a name.
 */
const OPERAND_RE = new RegExp(
	`"[^"]*"?|(${REFERENCE_SRC})(?![\\w.(])|[A-Za-z_][\\w.]*|\\d+(?:\\.\\d+)?`,
	'g',
);
const LITERAL_RE = /^(?:"[^]*|\d+(?:\.\d+)?|TRUE|FALSE)$/i;
const REFERENCE_EXPECTED_AFTER = /[=(,;+\-*/^&<>:]\s*$/;

/**
 * The part of a formula that a clicked cell should replace, or null when a click should
 * not insert a reference. That is the selected text, the reference under the caret, a
 * literal argument of a function under the caret (such as the 1 in "IF(A1>0, 1, 0)"), or
 * the empty spot after "=", "(", "," or an operator.
 */
export function referenceSlot(
	value: string,
	start: number,
	end: number,
): { from: number; to: number } | null {
	if (!value.startsWith('=') || start < 1) return null;
	if (end > start) return { from: start, to: end };

	let depth = 0;
	let inString = false;
	for (const ch of value.slice(0, start)) {
		if (ch === '"') inString = !inString;
		else if (!inString && ch === '(') depth++;
		else if (!inString && ch === ')') depth--;
	}

	// Right after a separator the caret belongs to the argument that follows the spaces.
	const afterSeparator = REFERENCE_EXPECTED_AFTER.test(value.slice(0, start));
	let next = start;
	if (afterSeparator) while (value[next] === ' ') next++;

	for (const m of value.matchAll(OPERAND_RE)) {
		const from = m.index ?? 0;
		const to = from + m[0].length;
		if (from > next) break;
		if (to < start) continue;
		if (m[1] !== undefined || (depth > 0 && LITERAL_RE.test(m[0])))
			return { from, to };
		if (from > start) break; // a name after the spaces: insert before it
		return null;
	}
	return afterSeparator ? { from: start, to: start } : null;
}

/* ------------------------------------------------------------------ */
/* Inserting a function with example arguments                          */
/* ------------------------------------------------------------------ */

/** A sample value for a parameter, based on its name in the syntax text. */
function sampleFor(param: string, nextCell: () => string): string {
	const p = param.toLowerCase().replace(/[[\]]/g, '');
	if (/range|array|sum_range|average_range/.test(p)) return 'A1:A10';
	if (p === 'table') return 'A1:B10';
	if (p === 'criteria' || /^criteria\d*$/.test(p)) return '">0"';
	if (p === 'condition' || /^condition\d*$/.test(p)) return 'A1>0';
	if (p === 'value_if_true') return '1';
	if (p === 'value_if_false' || p === 'value_if_error') return '0';
	if (
		p === 'digits' ||
		p === 'power' ||
		p === 'divisor' ||
		p === 'times' ||
		p === 'column' ||
		p === 'row' ||
		p === 'index'
	)
		return '2';
	if (
		p === 'significance' ||
		p === 'k' ||
		p === 'count' ||
		p === 'start' ||
		p === 'bottom'
	)
		return '1';
	if (p === 'top') return '10';
	if (p === 'exact') return 'FALSE';
	if (p === 'type') return '0';
	if (p === 'find' || p === 'old') return '"a"';
	if (p === 'new') return '"b"';
	if (p === 'format') return '"0.00"';
	if (p === 'year') return '2026';
	if (p === 'month' || p === 'day') return '1';
	if (p === 'rate') return '0.05';
	if (p === 'periods') return '12';
	if (p === 'present_value') return '1000';
	if (p === 'prompt') return '"Summarize"';
	if (p === 'cells') return 'A1';
	return nextCell();
}

/** Example arguments for a function, e.g. "A1, A2" for SUM or "A1>0, 1, 0" for IF. */
export function exampleArguments(fn: FormulaFunctionInfo): string {
	const open = fn.syntax.indexOf('(');
	const close = fn.syntax.lastIndexOf(')');
	if (open < 0 || close <= open) return '';
	const inner = fn.syntax.slice(open + 1, close).trim();
	if (inner === '') return '';

	let n = 0;
	const nextCell = (): string => `A${++n}`;
	const out: string[] = [];
	for (const raw of inner.split(',')) {
		const token = raw.trim();
		if (token === '' || token.includes('…')) continue;
		const optional = token.startsWith('[');
		// Optional parameters are skipped, except the second item of a list (number2) and IF's else value.
		if (
			optional &&
			!/2\]?$/.test(token) &&
			!token.includes('value_if_false')
		)
			continue;
		out.push(sampleFor(token, nextCell));
	}
	return out.join(', ');
}

export interface FunctionCall {
	/** The text to insert, e.g. "SUM(A1, A2)". */
	text: string;
	/** Selection (relative to the start of text) that covers the example arguments. */
	selectFrom: number;
	selectTo: number;
}

/** Builds "NAME(example args)" and the selection that lets the user overtype the arguments. */
export function buildFunctionCall(fn: FormulaFunctionInfo): FunctionCall {
	const args = exampleArguments(fn);
	const text = `${fn.name}(${args})`;
	if (args === '')
		return { text, selectFrom: text.length, selectTo: text.length };
	return {
		text,
		selectFrom: fn.name.length + 1,
		selectTo: fn.name.length + 1 + args.length,
	};
}

/**
 * Moves the relative row references of a formula by `delta` rows, like Excel does when a
 * row is moved by sorting. Absolute rows (A$1) and text in quotes are left alone.
 */
export function shiftFormulaRows(formula: string, delta: number): string {
	if (delta === 0 || !formula.startsWith('=')) return formula;
	return formula
		.split(/("(?:[^"]|"")*")/)
		.map((part, i) =>
			i % 2 === 1
				? part
				: part.replace(
						/(^|[^A-Za-z0-9_.$])(\$?[A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(.])/g,
						(
							match: string,
							before: string,
							column: string,
							absolute: string,
							rowText: string,
						) => {
							if (absolute) return match;
							const row = Number(rowText) + delta;
							return row < 1
								? `${before}#REF!`
								: `${before}${column}${row}`;
						},
					),
		)
		.join('');
}

export interface ReferencedRange {
	r1: number;
	c1: number;
	r2: number;
	c2: number;
}

const REFERENCED_RE =
	/(^|[^A-Za-z0-9_.$!])\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?(?![A-Za-z0-9_(.!])/g;

/**
 * The cells and ranges of this sheet that a formula refers to (0-based), e.g. G30 and
 * A1:B5 for "=G30+SUM(A1:B5)". References to other sheets and text in quotes are skipped.
 */
export function formulaReferences(formula: string): ReferencedRange[] {
	if (!formula.startsWith('=')) return [];
	const out: ReferencedRange[] = [];
	formula.split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/).forEach((part, i) => {
		if (i % 2 === 1) return;
		for (const m of part.matchAll(REFERENCED_RE)) {
			const c1 = colToIndex(m[2] ?? '');
			const r1 = Number(m[3]) - 1;
			const c2 = m[4] ? colToIndex(m[4]) : c1;
			const r2 = m[5] ? Number(m[5]) - 1 : r1;
			if (r1 < 0 || r2 < 0) continue;
			out.push({
				r1: Math.min(r1, r2),
				c1: Math.min(c1, c2),
				r2: Math.max(r1, r2),
				c2: Math.max(c1, c2),
			});
		}
	});
	return out;
}

const MOVE_CELL_RE =
	/(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(.])/g;
const MOVE_COLUMNS_RE =
	/(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(.$])/g;
const MOVE_ROWS_RE =
	/(^|[^A-Za-z0-9_.$])(\$?)(\d+):(\$?)(\d+)(?![A-Za-z0-9_(.])/g;

/**
 * Rewrites a formula for a table that was moved as a whole by `rowDelta` rows and
 * `colDelta` columns (an import that does not start at A1). Every reference moves with
 * the table, absolute ones ($A$1) too, because the cells they point at moved as well.
 * Whole columns (A:C) and whole rows (1:5) move the same way; text in quotes is left alone.
 */
export function moveFormula(
	formula: string,
	rowDelta: number,
	colDelta: number,
): string {
	if ((rowDelta === 0 && colDelta === 0) || !formula.startsWith('=')) {
		return formula;
	}
	const column = (letters: string): string =>
		indexToColumnLabel(colToIndex(letters) + colDelta);
	const row = (digits: string): string => String(Number(digits) + rowDelta);
	return formula
		.split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/)
		.map((part, i) =>
			i % 2 === 1
				? part
				: part
						.replace(
							MOVE_CELL_RE,
							(
								_match: string,
								before: string,
								absCol: string,
								letters: string,
								absRow: string,
								digits: string,
							) =>
								`${before}${absCol}${column(letters)}${absRow}${row(digits)}`,
						)
						.replace(
							MOVE_COLUMNS_RE,
							(
								_match: string,
								before: string,
								abs1: string,
								from: string,
								abs2: string,
								to: string,
							) =>
								`${before}${abs1}${column(from)}:${abs2}${column(to)}`,
						)
						.replace(
							MOVE_ROWS_RE,
							(
								_match: string,
								before: string,
								abs1: string,
								from: string,
								abs2: string,
								to: string,
							) => `${before}${abs1}${row(from)}:${abs2}${row(to)}`,
						),
		)
		.join('');
}

import FormulaParser from 'fast-formula-parser';
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
export type GetCell = (row: number, col: number) => Scalar;

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

function evaluateWithLibrary(formula: string, getCell: GetCell): Scalar {
	const FP = FormulaParser as unknown as new (options: unknown) => {
		parse(formula: string, position: unknown): unknown;
	};
	const parser = new FP({
		onCell: (ref: { row: number; col: number }) => {
			const v = getCell(ref.row - 1, ref.col - 1);
			return v === null ? undefined : v;
		},
		onRange: (ref: {
			from: { row: number; col: number };
			to: { row: number; col: number };
		}) => {
			const rows: (Scalar | undefined)[][] = [];
			for (let r = ref.from.row; r <= ref.to.row; r++) {
				const row: (Scalar | undefined)[] = [];
				for (let c = ref.from.col; c <= ref.to.col; c++) {
					const v = getCell(r - 1, c - 1);
					row.push(v === null ? undefined : v);
				}
				rows.push(row);
			}
			return rows;
		},
	});

	const parsed: unknown = parser.parse(
		normalizeSeparators(formula.replace(/^\s*=/, '')),
		{
			sheet: 'Sheet1',
			row: 1,
			col: 1,
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
	return String(out);
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

export function evaluateFormula(formula: string, getCell: GetCell): Scalar {
	try {
		return evaluateWithLibrary(formula, getCell);
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
];

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

// Ambient declarations for modules that ship without TypeScript types.
// This file must live inside src/ (next to main.ts) and must NOT contain
// any top-level import/export, otherwise the declarations stop being ambient.

// CSS files are bundled by esbuild (emitted as main.css, then moved to styles.css).
declare module '*.css';

// fast-formula-parser ships plain JavaScript without type declarations.
declare module 'fast-formula-parser' {
	interface FormulaPosition {
		sheet?: string;
		row: number;
		col: number;
	}

	class FormulaParser {
		constructor(config?: unknown);
		parse(formula: string, position: FormulaPosition): unknown;
		/** Every function the parser knows, by upper-case name. */
		functions: Record<string, (...args: unknown[]) => unknown>;
		/** Error class; `new FormulaError('#DIV/0!')` returns the shared instance for that error. */
		static FormulaError: unknown;
		static FormulaHelpers: {
			accept(param: unknown, type?: unknown, ...rest: unknown[]): unknown;
		};
		static Types: Record<string, number>;
	}

	export default FormulaParser;
}

// The operator implementations, patched in formulas.ts for empty cells.
declare module 'fast-formula-parser/formulas/operators' {
	export const Infix: {
		compareOp(
			a: unknown,
			op: string,
			b: unknown,
			aIsArray: boolean,
			bIsArray: boolean,
		): unknown;
	};
	export const Postfix: {
		percentOp(value: unknown, postfix: string, isArray: boolean): unknown;
	};
}

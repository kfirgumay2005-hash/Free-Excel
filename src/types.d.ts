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
	}

	export default FormulaParser;
}

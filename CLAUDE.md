# Free-Excel (Obsidian plugin)

Excel-style spreadsheet for Obsidian with full RTL support, built on Handsontable. Works on desktop and mobile.

## Language

- Talk to the user in Hebrew. Code, comments, identifiers and UI text are in English.
- The plugin UI is English-only (the plugin is global); sentence case for labels.

## Commands

- `npm run build` runs `tsc -noEmit -skipLibCheck` and then the esbuild production bundle. Run it after every change and fix type errors first.
- Never edit `main.js` or `styles.css` in the repo root: both are generated.

## Layout (`src/`)

- `main.ts`: plugin class and `SheetView` (a `TextFileView`): grid, toolbar, formula bar, zoom, images, import and export.
- `model.ts`: file model and shared helpers (`SheetModel`, `CellFormat`, row-height and zoom constants).
- `formulas.ts`: formula engine (fast-formula-parser with a small built-in fallback) and the function catalog.
- `xlsx.ts`: import of .xlsx and .csv (formats, merges, dates and times via the number format).
- `ui.ts`: formula autocomplete popup, the emoji picker and the modals (comment, link, filter, import).
- `ai.ts`: Claude and Gemini: requests, the runner for `CLAUDE()`/`GEMINI()` formulas (answers cached in the file by request, `WorkbookModel.aiResults`, not part of undo) and the assistant panel with its tool loop. The tools themselves (`AI_TOOLS`, `runAiTool`) live in `main.ts`; a tool may return a promise (the image tools download or generate files), and tool calls run one after another. Busy responses (429/5xx) are retried in `postJson`.
- `emoji.ts`: emoji list with search keywords.
- `settings.ts`: settings tab. `types.d.ts`: ambient declarations. `plugin.css`: plugin styles.
- CSS from Handsontable and `plugin.css` is bundled by esbuild into `main.css`, and the build renames it to `styles.css`.

## Rules that are easy to break

- `tsconfig` uses `strict` and `noUncheckedIndexedAccess`: every `arr[i]` may be `undefined`, so guard it or give it a default.
- The model stores sizes at 100% zoom. Zoom is view-only and is applied when sizes are handed to Handsontable. Column widths use `zoom`; row heights, the column of row numbers (`rowHeaderWidth`), fonts and chart sizes use `zoomY`, which is 1 when the "Zoom" setting is "Horizontal only". Handsontable never makes a row lower than its theme's line height plus padding, so `applyZoomStyles` zooms those CSS variables too.
- Per-cell data is keyed `"row,col"`. When rows or columns are inserted or removed, remap every keyed map (`cellFormats`, `images`, `imageSizes`, `comments`, `links`, `widenRecords`) in `remapCellMaps`, and add new ones to `RowsSnapshot` (sorting) and `SideState` (undo).
- Visual and physical row/column indexes must stay equal: filters hide rows with the `hiddenRows` plugin (view-only, not saved), and sorting rewrites the rows themselves (`sortByColumn`). Never enable Handsontable's `filters`, `columnSorting` or trimming plugins.
- Undo: Handsontable undoes cell values, merges and inserted or removed rows and columns. Everything else in the model (`SideState`: formats, images, sizes) is recorded by `recordSideChanges`, which runs after every `markDirty()` and attaches the difference to the same undo step. So call `markDirty()` after every model change, and replace format and image-size objects instead of changing them in place (they are compared by identity).
- Row heights come from the font size (`computeRowHeight`), not from Handsontable auto-sizing. Column widths are never automatic.
- Handsontable cell renderers reuse `<td>` elements: reset every style you set, and remove any child elements you added.
- Handsontable v16+: styles are imported from `handsontable/styles/...`, the theme is set with `themeName`, and `layoutDirection` can only be set at creation (a direction change rebuilds the grid).
- `ManualRowResize` has no `clearManualSize` (only `ManualColumnResize` does).
- Sheet files use the `.sheet` extension and fall back to `.rtlsheet` when another plugin owns it.
- API keys are kept in Obsidian's secret storage (`app.secretStorage`, Obsidian 1.13.2+) under fixed IDs (`aiKeySecretId`); the settings tab shows one password field per key. `claudeKeySecret`/`geminiKeySecret` are only read as a fallback from the earlier secret picker. Requests go through `requestUrl` (no CORS, works on mobile).
- Floating images and charts are drawn on the chart layer and positioned in px at 100% zoom from A1; they are replaced, never changed in place.
- Floating images and charts share one stacking order: `z` on both, drawn bottom-up by `layerOrder()`; new ones get `topZ()`. Find a box on the layer by its `data-id`, never by its index.
- `normalizeCharts` rebuilds every chart field by field when a file loads: add each new optional `ChartSpec` field to `normalizeChartExtras` (`charts.ts`), or it is lost on reopen.
- Every chart type reads two ranges: `xRange` (labels, slice names or X values) and `range` (a series per column). Without `xRange` the first column of `range` holds the labels, unless `noLabels` is set.
- Handsontable's Ctrl+C/Ctrl+V carries text only. Formats (check boxes too) go through the module-level `copiedCells`: kept in `beforeCopy`/`beforeCut` and put back in `afterPaste` when the pasted text is the same.
- The chart layer is emptied and drawn again after every forced render (`queueCharts`), so a CSS animation on it restarts each time and looks frozen: give it a negative `animation-delay` from the clock (see `renderCopyMark`). When a setting controls motion, do not also turn it off for `prefers-reduced-motion` (Windows "Animation effects" off sets it).
- The Handsontable license key is non-commercial: a commercial release needs its own license.

## Verification

- There is no automated test setup. Behavior that depends on rendering (grid lines, spill, touch gestures, RTL) has to be checked in Obsidian on desktop and on a phone. Say so when something is untested.

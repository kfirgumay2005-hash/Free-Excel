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
- `ui.ts`: formula autocomplete popup and the modals.
- `settings.ts`: settings tab. `types.d.ts`: ambient declarations. `plugin.css`: plugin styles.
- CSS from Handsontable and `plugin.css` is bundled by esbuild into `main.css`, and the build renames it to `styles.css`.

## Rules that are easy to break

- `tsconfig` uses `strict` and `noUncheckedIndexedAccess`: every `arr[i]` may be `undefined`, so guard it or give it a default.
- The model stores sizes at 100% zoom. Zoom is view-only and is applied when sizes are handed to Handsontable.
- Per-cell data is keyed `"row,col"`. When rows or columns are inserted or removed, remap every keyed map (`cellFormats`, `images`, `imageSizes`, `widenRecords`) in `remapCellMaps`.
- Visual and physical row/column indexes must stay equal: filters hide rows with the `hiddenRows` plugin (view-only, not saved), and sorting rewrites the rows themselves (`sortByColumn`). Never enable Handsontable's `filters`, `columnSorting` or trimming plugins.
- Undo: Handsontable undoes cell values, merges and inserted or removed rows and columns. Everything else in the model (`SideState`: formats, images, sizes) is recorded by `recordSideChanges`, which runs after every `markDirty()` and attaches the difference to the same undo step. So call `markDirty()` after every model change, and replace format and image-size objects instead of changing them in place (they are compared by identity).
- Row heights come from the font size (`computeRowHeight`), not from Handsontable auto-sizing. Column widths are never automatic.
- Handsontable cell renderers reuse `<td>` elements: reset every style you set, and remove any child elements you added.
- Handsontable v16+: styles are imported from `handsontable/styles/...`, the theme is set with `themeName`, and `layoutDirection` can only be set at creation (a direction change rebuilds the grid).
- `ManualRowResize` has no `clearManualSize` (only `ManualColumnResize` does).
- Sheet files use the `.sheet` extension and fall back to `.rtlsheet` when another plugin owns it.
- The Handsontable license key is non-commercial: a commercial release needs its own license.

## Verification

- There is no automated test setup. Behavior that depends on rendering (grid lines, spill, touch gestures, RTL) has to be checked in Obsidian on desktop and on a phone. Say so when something is untested.

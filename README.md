# Free-Excel

> [!NOTE]
> **About the "Caution" review score.** The plugin page shows a Caution rating because the automated scan reports a few dozen issues. Almost all of them are one CSS lint warning repeated: "Avoid `!important`".
>
> None of those come from this plugin's own styles. They come from the stylesheet of [Handsontable](https://handsontable.com/), the grid library the spreadsheet is built on, which is bundled into `styles.css`. Handsontable positions and colors parts of the grid (selection borders, the fill handle, resize handles) with inline styles set from JavaScript, and `!important` is the only thing in CSS that can override an inline style. Higher selector specificity or CSS variables, which the lint suggests, cannot do it, so removing these rules would break the grid.
>
> It is a style warning, not a security or stability problem. The plugin works offline, and its styles only apply inside the spreadsheet view.

Create and edit spreadsheets with full right-to-left (RTL) support, formulas, formatting and export to .xlsx.

A real spreadsheet inside your Obsidian vault. Free, offline, on desktop and mobile.

## A few words

can't belive this op plugin got shadowbanned because of one warning pretty shameful.

When I heard the following lines I felt like I had to give the people MORE! :

“It's about drive, it's about power.
We stay hungry, we devour.
Put in the work, put in the hours,
And take what's ours"

## Features

- **True RTL.** Supports right to left, and one button switches a sheet to LTR.
- **Formulas.** Around 280 Excel functions such as `SUM`, `IF` and `VLOOKUP`, with autocomplete and a function picker. While typing a formula, click a cell (or drag over several) to insert its reference.
- **Formatting.** Fonts, colors, borders, alignment, merged cells, text wrap and rotation, plus a format painter.
- **Images in cells.** Paste, drop or pick one from your vault.
- **Excel in and out.** Import `.xlsx` and `.csv`, export to `.xlsx`.
- **Sync friendly.** Open sheets refresh on their own when a sync plugin brings in changes from another device.
- **Touch ready.** Pinch to zoom and a toolbar made for phones.

## Get started

1. Select the table icon in the ribbon, or run **Create new spreadsheet** from the command palette.
2. Type. A cell that starts with `=` is a formula.
3. Select **Export** in the toolbar to get an `.xlsx` file next to your sheet.

Sheets are saved as `.sheet` files in your vault and autosave as you work.

## Install

Copy `main.js`, `manifest.json` and `styles.css` into `<Vault>/.obsidian/plugins/freeexcel/`, then enable **Free-Excel** in **Settings → Community plugins**.

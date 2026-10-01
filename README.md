# Sheetian

Create and edit spreadsheet files inside **desktop Obsidian**. Sheetian is based on [Divam Gupta's plugin](https://github.com/divamgupta/obsidian-spreadsheets), uses [FortuneSheet](https://github.com/ruilisi/fortune-sheet) 1.0.4 and retains the existing JSON `.sheet` format.

Sheetian repairs data-loss and save lifecycle defects, preserves images, filters, formula chains and other workbook metadata, isolates its stylesheet, fixes paste ownership across multiple open workbooks, and resolves spreadsheet shortcuts intercepted by Obsidian. The full [upstream issue review](ISSUE_TRIAGE.md) documents what is fixed, already available, unconfirmed, or a separate feature request.

## Use

1. Click the table ribbon icon or run **Sheetian: New spreadsheet**.
2. Right-click a folder or file to create a spreadsheet in that folder.
3. Edit cells, use formulas, format cells, insert images, and filter/sort with the spreadsheet toolbar.
4. Run **Sheetian: Export active sheet to CSV** to save displayed values beside the workbook. Existing CSVs receive a numbered suffix.

In **Settings → Sheetian**, choose the default folder, currency symbol, and light/dark/Obsidian theme. Settings apply when you open a spreadsheet. Dark display inverts canvas colors with hue rotation; stored cell formatting remains intact and images receive a compensating filter. It does not reproduce every custom Obsidian theme's palette.

- **Double-click a column header’s right edge** to fit its contents. Select several column headers first to fit them all. Widths account for displayed values, font sizes, bold/italic text, and explicit line breaks; horizontally merged cells are excluded.
- **Double-click a row header’s bottom edge** to fit its height. Select several row headers first to fit them all. Heights account for wrapped text at the current column widths, line breaks, and fonts/rich text. Cells merged across columns use their combined width; cells merged across rows are excluded. Heights support undo/redo and persist after reopening.
- **F2** opens the current cell editor with the caret at the end.
- **Cmd+B / Ctrl+B** toggles the selected cells' bold format.
- **Alt+Enter** inserts a line break while editing a cell. Use the toolbar's text-wrap setting to display wrapped content.
- The current engine also provides fill-down and edge/range navigation shortcuts.
- Enter `$300`, `€1,234.50`, `£25`, or the configured currency symbol to store numeric currency values. Cells explicitly formatted as Text stay text.
- For dates, use `YYYY-MM-DD`; Date-formatted cells also accept `M/D/YYYY`. Dates are stored as Excel-compatible serial numbers.

Files that cannot be parsed show an error and preserve their original contents. Opening and closing an untouched file does not rewrite it.

## Import from Google Sheets or Excel

1. In Google Sheets, choose **File → Download → Microsoft Excel (.xlsx)**. Existing Excel `.xlsx` workbooks can be imported directly.
2. In Obsidian, open a sheet's **⋯** menu at the top right and choose **Import Excel workbook (.xlsx)**, or run **Sheetian: Import Excel workbook (.xlsx)** from the command palette. Select the downloaded file. Alternatively, right-click an `.xlsx` file already in your vault and choose **Import into Sheetian**.
3. Sheetian creates a new `.sheet` workbook and opens it. The command uses the default folder from Settings; the context-menu action uses the source file's folder. Repeated imports receive numbered filenames.

The importer retains worksheet tabs, ordinary and shared formulas, cross-tab references, cached results, numeric/date formats, basic fonts/alignment/solid colors, rich text, hyperlinks, merged cells, row/column sizes, hidden rows/columns/tabs, and frozen panes. Supported formulas recalculate when their inputs change. Double-clicking and committing a formula retains its formula source, including formulas formatted across several lines. Formulas without cached results are calculated on import. Import uses [ExcelJS](https://github.com/exceljs/exceljs).

This is a one-time import; changes do not sync back to Google Sheets. CSV and normal clipboard paste cannot preserve the original formulas. Formula text and available cached results are retained for unsupported formulas, but Google-specific functions (such as `QUERY`, `IMPORTRANGE`, and `GOOGLEFINANCE`), newer Excel functions, named ranges, external/structured references, and array/spill expansion may not recalculate. Recognized limitations are shown after import. Named-range definitions are retained as metadata. Images, charts, and Excel-specific features such as table/filter definitions, conditional formatting, and validation are not imported; theme colors and borders are not translated. Imports are limited to 25 MB and 1,000,000 grid cells across all tabs, including the engine's minimum 100 × 26 grid per tab.

## Build and install

Use Node.js 18+ and npm:

```sh
npm ci
npm test
npm run build
```

Copy **main.js**, **manifest.json**, and **styles.css** into `<vault>/.obsidian/plugins/sheetian/`, then enable **Sheetian** in Obsidian's Community plugins settings. When migrating from Spreadsheets, disable that plugin first and copy its `data.json` into the Sheetian folder to preserve settings. Existing `.sheet` files remain compatible. Rebuild all three files together; `styles.css` is generated from the pinned engine and `spreadsheet.css`.

The bundled JavaScript is generated and intentionally ignored by Git. Build-time compatibility patches in `engine-compat.mjs` guard known engine defects and must be reviewed when changing the engine version. The lockfile makes installs reproducible.

## Scope

This is a desktop plugin. Mobile compatibility, live Google Sheets synchronization, native vault-wide content search, Markdown embeddings, custom executable functions, automatic numeric alignment, and native cell wikilinks remain separate features. CSV exports values, not workbook styling or editable formulas.

Use GitHub issues for support. The plugin is MIT-licensed; original author: Divam Gupta.

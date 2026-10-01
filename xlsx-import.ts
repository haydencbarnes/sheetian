import * as ExcelJS from "exceljs";
import { Buffer } from "buffer";
import { is_date, update } from "@fortune-sheet/core";
import type { Cell, Sheet } from "@fortune-sheet/core";

export interface ImportedWorkbook {
  sheets: Sheet[];
  formulaCount: number;
  warnings: string[];
  missingResults: { id: string; r: number; c: number }[];
}
const MAX_GRID_CELLS = 1_000_000;
const MAX_FILE_BYTES = 25 * 1024 * 1024;

function color(value: Partial<ExcelJS.Color> | undefined): string | undefined {
  return value?.argb && /^[0-9a-f]{6,8}$/i.test(value.argb) ? `#${value.argb.slice(-6)}` : undefined;
}
function font(value: Partial<ExcelJS.Font> | undefined): Partial<Cell> {
  if (!value) return {};
  return { ...(value.name ? { ff: value.name } : {}), ...(value.size ? { fs: value.size } : {}),
    ...(value.bold ? { bl: 1 } : {}), ...(value.italic ? { it: 1 } : {}),
    ...(value.strike ? { cl: 1 } : {}), ...(value.underline ? { un: 1 } : {}),
    ...(color(value.color) ? { fc: color(value.color) } : {}) };
}
function dateSerial(date: Date): number {
  const serial = date.getTime() / 86400000 + 25569;
  return serial < 61 ? serial - 1 : serial;
}
function scalar(value: ExcelJS.CellValue): string | number | boolean | undefined {
  if (value instanceof Date) return dateSerial(value);
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  if (value && typeof value === "object" && "error" in value) return value.error;
  return undefined;
}
function address(value: string): { r: number; c: number } {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(value);
  if (!match) throw new Error(`Unsupported cell address: ${value}`);
  let column = 0;
  for (const ch of match[1].toUpperCase()) column = column * 26 + ch.charCodeAt(0) - 64;
  return { r: Number(match[2]) - 1, c: column - 1 };
}

export async function importXlsx(bytes: ArrayBuffer | Uint8Array): Promise<ImportedWorkbook> {
  if (bytes.byteLength > MAX_FILE_BYTES) throw new Error("This workbook exceeds the 25 MB import limit.");
  const workbook = new ExcelJS.Workbook();
  try { await workbook.xlsx.load(Buffer.from(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes))); }
  catch { throw new Error("Could not read this XLSX file. Download an unencrypted Microsoft Excel (.xlsx) copy and try again."); }
  if (!workbook.worksheets.length) throw new Error("The workbook contains no worksheets.");
  const warnings = new Set<string>();
  const missingResults: ImportedWorkbook["missingResults"] = [];
  let formulaCount = 0, gridCells = 0;
  const sheets = workbook.worksheets.map((worksheet, index): Sheet => {
    const id = `xlsx-${index + 1}`;
    // FortuneSheet expands sparse cells into a dense grid when opening it.
    const rows = Math.max(100, worksheet.rowCount), columns = Math.max(26, worksheet.columnCount);
    gridCells += rows * columns;
    if (gridCells > MAX_GRID_CELLS) throw new Error("This workbook needs more than 1,000,000 grid cells. Export a smaller range or remove unused formatted rows/columns first.");
    const sheet: Sheet = { name: worksheet.name, id, order: index, status: 0,
      hide: worksheet.state === "visible" ? 0 : 1, row: rows, column: columns,
      celldata: [], calcChain: [], config: {}, hyperlink: {} };
    const config = sheet.config!;
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (source, columnNumber) => {
        if (source.isMerged && source.master.address !== source.address) return;
        const r = rowNumber - 1, c = columnNumber - 1;
        const cell: Cell = { ...font(source.font) };
        const fill = source.fill;
        if (fill?.type === "pattern" && fill.pattern === "solid" && color(fill.fgColor)) cell.bg = color(fill.fgColor);
        const align = source.alignment;
        if (align?.horizontal === "center") cell.ht = 0;
        else if (align?.horizontal === "left") cell.ht = 1;
        else if (align?.horizontal === "right") cell.ht = 2;
        if (align?.vertical) cell.vt = { middle: 0, top: 1, bottom: 2 }[align.vertical as "middle" | "top" | "bottom"];
        if (align?.wrapText) cell.tb = "2";
        if (typeof align?.textRotation === "number") cell.rt = align.textRotation < 0 ? 90 - align.textRotation : align.textRotation;
        if (align?.textRotation === "vertical") cell.tr = "3";
        const formula = source.formula;
        let value: ExcelJS.CellValue = formula ? source.result : source.value;
        if (formula) {
          cell.f = formula.startsWith("=") ? formula : `=${formula}`;
          formulaCount++;
          sheet.calcChain!.push({ r, c, id });
          if (value === undefined || value === null) missingResults.push({ id, r, c });
          const outsideStrings = formula.replace(/"(?:[^"]|"")*"/g, '""');
          if (/\b(?:_xlfn\.|_xlws\.|IMPORTRANGE\s*\(|GOOGLEFINANCE\s*\(|QUERY\s*\(|ARRAYFORMULA\s*\(|IMPORT(?:XML|HTML|DATA|FEED)\s*\(|SPARKLINE\s*\()/i.test(outsideStrings)) {
            warnings.add("Some formulas use Google-only or newer Excel functions. Their text and cached results are preserved, but Sheetian may not recalculate them.");
          }
          const array = source.value as ExcelJS.CellFormulaValue & { shareType?: string; ref?: string };
          if (array.shareType === "array") warnings.add("Array/spill formulas retain their anchor formula and cached cells; automatic spill expansion is not supported.");
          if (/\[[^\]]+\]/.test(outsideStrings)) warnings.add("External workbook or structured table references are preserved, but cannot be resolved by Sheetian.");
        }
        if (value && typeof value === "object" && "richText" in value) {
          cell.ct = { t: "inlineStr", fa: "General", s: value.richText.map(run => ({ v: run.text, ...font(run.font) })) };
        } else {
          if (value && typeof value === "object" && "hyperlink" in value) {
            sheet.hyperlink![`${r}_${c}`] = { linkType: "external", linkAddress: value.hyperlink };
            value = value.text;
          }
          const v = scalar(value);
          if (v !== undefined) cell.v = v;
          const fmt = source.numFmt || "General";
          const date = value instanceof Date || is_date(fmt);
          cell.ct = { fa: fmt, t: date ? "d" : typeof v === "string" ? "s" : typeof v === "boolean" ? "b" : "n" };
          if (v !== undefined) {
            try { cell.m = typeof v === "number" ? String(update(fmt, v)) : typeof v === "boolean" ? String(v).toUpperCase() : v; }
            catch { cell.m = String(v); }
          }
        }
        if (source.value !== null && source.value !== undefined || Object.keys(source.style).length) sheet.celldata!.push({ r, c, v: cell });
      });
      if (row.height) { (config.rowlen ??= {})[rowNumber - 1] = Math.max(1, row.height * 4 / 3 - 1); (config.customHeight ??= {})[rowNumber - 1] = 1; }
      if (row.hidden) (config.rowhidden ??= {})[rowNumber - 1] = 0;
    });
    worksheet.columns?.forEach((column, i) => {
      if (column.width) { (config.columnlen ??= {})[i] = Math.max(1, Math.round(column.width * 7 + 5) - 1); (config.customWidth ??= {})[i] = 1; }
      if (column.hidden) (config.colhidden ??= {})[i] = 0;
    });
    const cells = new Map(sheet.celldata!.map(cell => [`${cell.r}_${cell.c}`, cell]));
    for (const merge of worksheet.model.merges ?? []) {
      const [first, last = first] = String(merge).split(":");
      const start = address(first), end = address(last);
      const key = `${start.r}_${start.c}`, region = { ...start, rs: end.r - start.r + 1, cs: end.c - start.c + 1 };
      (config.merge ??= {})[key] = region;
      for (let r = start.r; r <= end.r; r++) for (let c = start.c; c <= end.c; c++) {
        const cellKey = `${r}_${c}`;
        const entry = cells.get(cellKey) ?? { r, c, v: {} };
        entry.v!.mc = cellKey === key ? region : start;
        cells.set(cellKey, entry);
      }
    }
    sheet.celldata = [...cells.values()];
    const frozen = worksheet.views?.find(view => view.state === "frozen");
    if (frozen?.state === "frozen" && (frozen.xSplit || frozen.ySplit)) sheet.frozen = {
      type: frozen.xSplit && frozen.ySplit ? "both" : frozen.xSplit ? "column" : "row",
      range: { row_focus: Math.max(0, (frozen.ySplit || 0) - 1), column_focus: Math.max(0, (frozen.xSplit || 0) - 1) },
    };
    if (worksheet.getImages().length) warnings.add("Embedded images are not imported.");
    return sheet;
  });
  const active = sheets.find(sheet => !sheet.hide) ?? sheets[0];
  active.status = 1; active.hide = 0;
  if (workbook.definedNames.model.length) warnings.add("Named ranges are retained in import metadata, but formulas that use their names may not recalculate.");
  // Preserve source metadata that the engine does not yet understand.
  Object.assign(sheets[0], { sheetianImport: { format: "xlsx", date1904: !!workbook.properties.date1904,
    definedNames: workbook.definedNames.model, warnings: [...warnings] } });
  return { sheets, formulaCount, warnings: [...warnings], missingResults };
}

import { clearMeasureTextCache, getCellTextInfo } from "@fortune-sheet/core";
import type { Cell, Context, Selection, Sheet } from "@fortune-sheet/core";

// Text layout reads these three context fields. Measure in saved pixels so
// sheet/display zoom cannot change the resulting row height.
const textContext = { lang: "en", defaultFontSize: 10, zoomRatio: 1 } as Context;

export function fitRowHeight(sheet: Sheet, row: number, canvas: CanvasRenderingContext2D): number {
  clearMeasureTextCache();
  let height = sheet.defaultRowHeight ?? 19;
  const include = (cell: Cell | null | undefined, column: number) => {
    if (!cell || sheet.config?.colhidden?.[column] === 0) return;
    // A cell spanning several rows cannot determine one row's height.
    if (cell.mc && (cell.mc.rs === undefined || cell.mc.rs > 1)) return;
    let width = 0;
    for (let c = column; c < column + (cell.mc?.cs ?? 1); c++) {
      if (sheet.config?.colhidden?.[c] !== 0) width += (sheet.config?.columnlen?.[c] ?? sheet.defaultColWidth ?? 73) + 1;
    }
    if (!width) return;
    const rich = cell.ct?.t === "inlineStr" && Array.isArray(cell.ct.s);
    const text = cell.m ?? cell.v;
    if (!rich && (text === undefined || text === "")) return;
    // The engine expects displayed values as strings and represents explicit
    // line breaks as rich text. Work on a copy without rewriting the cell.
    const measured: Cell = rich ? cell : { ...cell, m: String(text) };
    if (!rich && /\r|\n/.test(String(text))) {
      measured.ct = { ...cell.ct, t: "inlineStr", s: [{ ...cell, v: String(text) }] };
    }
    const info = getCellTextInfo(measured, canvas, textContext, {
      cellWidth: width, cellHeight: Number.POSITIVE_INFINITY, r: row, c: column,
    }, textContext);
    if (Number.isFinite(info?.textHeightAll)) height = Math.max(height, Math.ceil(info.textHeightAll + 8));
  };
  if (sheet.data) sheet.data[row]?.forEach(include);
  else sheet.celldata?.forEach(cell => { if (cell.r === row) include(cell.v, cell.c); });
  return height;
}

export function autoFitRows(sheet: Sheet, clicked: number, selection: Selection[] | undefined = sheet.luckysheet_select_save): number[] {
  const lastColumn = (sheet.data?.[0]?.length ?? sheet.column ?? 26) - 1;
  const ranges = selection?.filter(range => range.row_select ||
    (range.column[0] === 0 && range.column[1] === lastColumn));
  if (!ranges?.some(range => clicked >= range.row[0] && clicked <= range.row[1])) return [clicked];
  const rows = new Set<number>();
  for (const range of ranges) for (let r = range.row[0]; r <= range.row[1]; r++) {
    if (sheet.config?.rowhidden?.[r] !== 0) rows.add(r);
  }
  return [...rows];
}

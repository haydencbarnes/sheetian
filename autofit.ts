import type { Cell, CellStyle, Selection, Sheet } from "@fortune-sheet/core";

type Measure = (text: string, style: CellStyle) => number;

// Match FortuneSheet's English font list and point-based canvas font sizes.
export function cellFont(style: CellStyle): string {
  const fonts = ["Times New Roman", "Arial", "Tahoma", "Verdana"];
  const family = typeof style.ff === "string" && !/^\d+$/.test(style.ff)
    ? style.ff.replace(/["']/g, "") : fonts[Number(style.ff ?? 0)] ?? fonts[0];
  return `${style.it ? "italic" : "normal"} ${style.bl ? "bold" : "normal"} ${Math.ceil(style.fs || 10)}pt "${family}", "Helvetica Neue", Helvetica, Arial, sans-serif`;
}

function textWidth(cell: Cell, measure: Measure): number | null {
  // A horizontally merged cell does not belong to one column's width.
  if (cell.mc && (cell.mc.cs === undefined || cell.mc.cs > 1)) return null;
  const rich = cell.ct?.t === "inlineStr" && Array.isArray(cell.ct.s) ? cell.ct.s : null;
  const text = cell.m ?? cell.v;
  if (!rich && (text === undefined || text === "")) return null;
  const runs: (CellStyle & { v: unknown })[] = rich ?? [{ ...cell, v: text }];
  let lineWidth = 0, width = 0;
  for (const run of runs) {
    const style = { ...cell, ...run };
    const lines = String(run.v ?? "").split(/\r\n|\r|\n/);
    lines.forEach((line, i) => {
      if (i) { width = Math.max(width, lineWidth); lineWidth = 0; }
      lineWidth += measure(line, style);
    });
  }
  return Math.max(width, lineWidth);
}

export function fitColumnWidth(sheet: Sheet, column: number, measure: Measure): number {
  let width: number | null = null;
  const include = (cell: Cell | null | undefined) => {
    if (!cell) return;
    const measured = textWidth(cell, measure);
    if (measured !== null) width = Math.max(width ?? 0, measured);
  };
  if (sheet.data) sheet.data.forEach(row => include(row?.[column]));
  else sheet.celldata?.forEach(cell => { if (cell.c === column) include(cell.v); });
  return width === null ? sheet.defaultColWidth ?? 73 : Math.max(20, Math.ceil(width + 12));
}

export function autoFitColumns(sheet: Sheet, clicked: number, selection: Selection[] | undefined = sheet.luckysheet_select_save): number[] {
  // getSelection() is live, but omits the engine's column_select flag. A
  // whole-column selection covers every row; stored sheet selections can lag.
  const lastRow = (sheet.data?.length ?? sheet.row ?? 100) - 1;
  const ranges = selection?.filter(range => range.column_select ||
    (range.row[0] === 0 && range.row[1] === lastRow));
  if (!ranges?.some(range => clicked >= range.column[0] && clicked <= range.column[1])) return [clicked];
  const columns = new Set<number>();
  for (const range of ranges) {
    for (let c = range.column[0]; c <= range.column[1]; c++) {
      if (sheet.config?.colhidden?.[c] !== 0) columns.add(c);
    }
  }
  return [...columns];
}

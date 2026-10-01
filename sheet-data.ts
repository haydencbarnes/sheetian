import type { CellWithRowAndCol, Sheet } from "@fortune-sheet/core";

export function parseSheets(source: string): Sheet[] {
  if (!source.trim()) return [{ name: "Sheet1", id: "sheet-1", status: 1 }];
  const sheets: unknown = JSON.parse(source);
  if (!Array.isArray(sheets) || sheets.length === 0 || sheets.some(sheet =>
    !sheet || typeof sheet !== "object" || typeof sheet.name !== "string" ||
    (sheet.data !== undefined && (!Array.isArray(sheet.data) ||
      sheet.data.some((row: unknown) => !Array.isArray(row)))) ||
    (sheet.celldata !== undefined && (!Array.isArray(sheet.celldata) ||
      sheet.celldata.some((cell: CellWithRowAndCol) => !cell ||
        !Number.isInteger(cell.r) || cell.r < 0 || !Number.isInteger(cell.c) || cell.c < 0)))
  )) throw new Error("Expected an array of named sheets with valid cell coordinates.");
  // Preserve IDs: formulas and hyperlinks can refer to other sheets by ID.
  const ids = new Set(sheets.map(sheet => sheet.id).filter(Boolean));
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    if (!sheet.id) {
      let id = `sheet-${i + 1}`;
      while (ids.has(id)) id += "-new";
      sheet.id = id;
      ids.add(id);
    }
    if (!sheet.luckysheet_select_save?.length) {
      sheet.luckysheet_select_save = [{ row: [0, 0], column: [0, 0], row_focus: 0, column_focus: 0 }];
    }
    if (sheet.data) {
      sheet.celldata = cellsFromSheet(sheet);
      delete sheet.data;
    }
  }
  if (!sheets.some(sheet => sheet.status === 1)) sheets[0].status = 1;
  return sheets;
}

export function cellsFromSheet(sheet: Sheet): CellWithRowAndCol[] {
  if (!sheet.data) return sheet.celldata ?? [];
  const cells: CellWithRowAndCol[] = [];
  sheet.data.forEach((row, r) => row.forEach((v, c) => {
    if (v !== null && v !== undefined) cells.push({ r, c, v });
  }));
  return cells;
}

export function serializeSheets(sheets: Sheet[]): string {
  return JSON.stringify(sheets.map(sheet => {
    // Retain all metadata, including unknown future engine fields.
    const { data: _data, celldata: _cells, ...metadata } = sheet;
    return { ...metadata, celldata: cellsFromSheet(sheet) };
  }), null, 2);
}

export function sheetToCsv(sheet: Sheet): string {
  const cells = cellsFromSheet(sheet);
  if (!cells.length) return "";
  let rows = 0, columns = 0;
  for (const cell of cells) { rows = Math.max(rows, cell.r + 1); columns = Math.max(columns, cell.c + 1); }
  const values = new Map(cells.map(cell => [`${cell.r}:${cell.c}`, cell.v]));
  return Array.from({ length: rows }, (_, r) => Array.from({ length: columns }, (_, c) => {
    const cell = values.get(`${r}:${c}`);
    const value = String(cell?.m ?? cell?.v ?? cell?.ct?.s?.map((part: { v: string }) => part.v).join("") ?? "");
    return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
  }).join(",")).join("\r\n") + "\r\n";
}

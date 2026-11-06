import { getBorderInfoComputeRange, type Cell, type CellStyle, type Context, type Sheet } from "@fortune-sheet/core";
import { cellsFromSheet } from "./sheet-data";
import { cellFont } from "./autofit";

const escape = (value: unknown): string => String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const size = (value: number | undefined, fallback: number): number => Number.isFinite(value) && value! > 0 ? value! : fallback;
const color = (value: string | undefined, fallback: string): string => value && /^(#[\da-f]{3,8}|[a-z]+|(?:rgb|hsl)a?\([\d.,%\s+-]+\))$/i.test(value) ? value : fallback;

function fontStyle(cell: CellStyle): string {
  const decoration = [cell.cl ? "line-through" : "", cell.un ? "underline" : ""].filter(Boolean).join(" ") || "none";
  return `font:${cellFont(cell)};color:${color(cell.fc, "#000")};text-decoration:${decoration};`;
}

function cellText(cell: Cell): string {
  if (cell.ct?.t === "inlineStr" && Array.isArray(cell.ct.s)) {
    return cell.ct.s.map((run: CellStyle & { v?: string }) => `<span style="${escape(fontStyle({ ...cell, ...run }))}">${escape(run.v)}</span>`).join("");
  }
  // Use the displayed result, including number/date formats, rather than a formula.
  return escape(cell.m ?? cell.v);
}

function borderStyle(border: { style: number; color: string } | undefined): string {
  if (!border || !Number(border.style)) return "none";
  const kind = Number(border.style);
  const line = kind === 7 ? "double" : [3, 6, 11].includes(kind) ? "dotted" : [4, 5, 9, 10, 12].includes(kind) ? "dashed" : "solid";
  return `${kind === 7 || kind === 13 ? 3 : kind >= 8 ? 2 : 1}px ${line} ${color(border.color, "#000")}`;
}

export interface SheetPrintDocument { html: string; landscape: boolean; }

// Render saved sheet data, not the virtualized viewport. This also keeps display
// zoom, selection highlights and Obsidian's dark theme out of the exported PDF.
export function sheetPrintDocument(sheet: Sheet): SheetPrintDocument {
  const cells = cellsFromSheet(sheet);
  const config = sheet.config ?? {};
  const merges = { ...config.merge };
  let rowCount = 0, columnCount = 0;
  const include = (r: number, c: number) => {
    rowCount = Math.max(rowCount, r + 1); columnCount = Math.max(columnCount, c + 1);
  };
  for (const { r, c, v } of cells) {
    if (!v || !Object.keys(v).length) continue;
    include(r, c);
    if (v.mc?.rs && v.mc.cs) merges[`${v.mc.r}_${v.mc.c}`] = { ...v.mc, rs: v.mc.rs, cs: v.mc.cs };
  }
  for (const merge of Object.values(merges)) include(merge.r + merge.rs - 1, merge.c + merge.cs - 1);
  for (const border of config.borderInfo ?? []) {
    if (border.rangeType === "cell") include(border.value.row_index, border.value.col_index);
    else for (const range of border.range ?? []) include(range.row[1], range.column[1]);
  }
  // Bound expansion of sparse sheets before allocating a printable matrix.
  if (!Number.isSafeInteger(rowCount) || !Number.isSafeInteger(columnCount) || rowCount * columnCount > 1_000_000) {
    throw new Error("This sheet is too large to export to PDF. Reduce its used range to one million cells or fewer.");
  }
  const rows = Array.from({ length: rowCount }, (_, r) => r).filter(r => config.rowhidden?.[r] == null);
  const columns = Array.from({ length: columnCount }, (_, c) => c).filter(c => config.colhidden?.[c] == null);
  const data = Array.from({ length: rowCount }, () => Array<Cell | null>(columnCount).fill(null));
  for (const { r, c, v } of cells) if (r < rowCount && c < columnCount) data[r][c] = v;
  const id = sheet.id ?? "pdf";
  const context = { currentSheetId: id, luckysheetfile: [{ ...sheet, id, data }] } as Context;
  const borders = rowCount && columnCount ? getBorderInfoComputeRange(context, 0, rowCount - 1, 0, columnCount - 1, id) : {};
  const covered = new Set<string>();
  const anchors = new Map<string, { r: number; c: number; rs: number; cs: number; lastRow: number; lastColumn: number }>();
  for (const merge of Object.values(merges)) {
    const rr = rows.filter(r => r >= merge.r && r < merge.r + merge.rs);
    const cc = columns.filter(c => c >= merge.c && c < merge.c + merge.cs);
    if (!rr.length || !cc.length) continue;
    for (const r of rr) for (const c of cc) covered.add(`${r}_${c}`);
    const key = `${rr[0]}_${cc[0]}`;
    covered.delete(key);
    anchors.set(key, { r: merge.r, c: merge.c, rs: rr.length, cs: cc.length, lastRow: rr[rr.length - 1], lastColumn: cc[cc.length - 1] });
  }
  const widths = columns.map(c => size(config.columnlen?.[c], size(sheet.defaultColWidth, 73)) + 1);
  const tableWidth = widths.reduce((sum, width) => sum + width, 0);
  // Sheetian stores inserted/imported images as embedded data URLs. Never fetch
  // remote resources while exporting a workbook.
  const images = (sheet.images ?? []).filter(image => /^data:image\/(png|jpeg|gif|webp);base64,/i.test(image.src));
  const imageRight = Math.max(0, ...images.map(image => image.left + image.width));
  const imageBottom = Math.max(0, ...images.map(image => image.top + image.height));
  const width = Math.max(tableWidth, imageRight, (!rows.length || !columns.length) && !images.length ? 718 : 1);
  const landscape = width > 718; // A4 portrait minus two 10 mm margins, in pixels.
  const scale = Math.min(1, (landscape ? 1047 : 718) / width);
  const body = rows.map(r => {
    const contents = columns.map(c => {
      const key = `${r}_${c}`;
      if (covered.has(key)) return "";
      const merge = anchors.get(key);
      const cell = data[merge?.r ?? r]?.[merge?.c ?? c] ?? {};
      let style = fontStyle(cell) + `background:${color(cell.bg, "#fff")};text-align:${cell.ht === 0 ? "center" : cell.ht === 2 ? "right" : "left"};vertical-align:${cell.vt === 1 ? "top" : cell.vt === 2 ? "bottom" : "middle"};`;
      for (const [side, row, column, edge] of [
        ["top", r, c, "t"], ["left", r, c, "l"],
        ["bottom", merge?.lastRow ?? r, c, "b"], ["right", r, merge?.lastColumn ?? c, "r"],
      ] as const) {
        const border = borders[`${row}_${column}`]?.[edge];
        if (border != null) style += `border-${side}:${borderStyle(border)};`;
      }
      return `<td${merge ? ` rowspan="${merge.rs}" colspan="${merge.cs}"` : ""} style="${escape(style)}"><div>${cellText(cell)}</div></td>`;
    }).join("");
    return `<tr style="height:${size(config.rowlen?.[r], size(sheet.defaultRowHeight, 19)) + 1}px">${contents}</tr>`;
  }).join("");
  const gridlines = sheet.showGridLines === false || sheet.showGridLines === 0 ? "transparent" : "#d5d8dc";
  const imageHtml = images.map(image => `<img alt="" src="${escape(image.src)}" style="position:absolute;left:${Math.max(0, image.left)}px;top:${Math.max(0, image.top)}px;width:${size(image.width, 1)}px;height:${size(image.height, 1)}px">`).join("");
  return { landscape, html: `<!doctype html><html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data:">
<title>${escape(sheet.name)}</title><style>
* { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body { margin: 0; background: white; color: black; font: 10pt Arial, sans-serif; }
h1 { font: 600 14pt Arial, sans-serif; margin: 0 0 12px; overflow-wrap: anywhere; }
.sheet { position: relative; width: ${width}px; min-height: ${imageBottom}px; zoom: ${scale}; }
table { border-collapse: collapse; table-layout: fixed; width: ${tableWidth}px; }
tr { break-inside: avoid; }
td { padding: 2px 4px; border: 1px solid ${gridlines}; }
td > div { white-space: pre-wrap; overflow-wrap: anywhere; }
</style></head><body><h1>${escape(sheet.name)}</h1><div class="sheet">
${rows.length && columns.length ? `<table><colgroup>${widths.map(w => `<col style="width:${w}px">`).join("")}</colgroup><tbody>${body}</tbody></table>` : images.length ? "" : "<p>This sheet has no visible cells.</p>"}
${imageHtml}</div></body></html>` };
}

export async function downloadSheetPdf(sheet: Sheet, filename: string): Promise<string | null> {
  // Capture before opening the asynchronous dialog, even if the user switches tabs.
  const document = sheetPrintDocument(sheet);
  const remote = require("electron").remote ?? require("@electron/remote");
  const { canceled, filePath } = await remote.dialog.showSaveDialog(remote.getCurrentWindow(), {
    title: "Export active sheet to PDF", defaultPath: filename,
    filters: [{ name: "PDF document", extensions: ["pdf"] }],
  });
  if (canceled || !filePath) return null;
  const printWindow = new remote.BrowserWindow({ show: false, webPreferences: {
    nodeIntegration: false, contextIsolation: true, sandbox: true,
  } });
  const fs = require("fs").promises;
  const path = require("path");
  let temporaryDirectory: string | undefined;
  try {
    // Chromium rejects large data URLs. Load a private temporary HTML file so
    // large sheets and embedded images do not hit the URL-length limit.
    temporaryDirectory = await fs.mkdtemp(path.join(require("os").tmpdir(), "sheetian-pdf-"));
    const htmlPath = path.join(temporaryDirectory!, "sheet.html");
    await fs.writeFile(htmlPath, document.html, { encoding: "utf8", mode: 0o600 });
    await printWindow.loadFile(htmlPath);
    await printWindow.webContents.executeJavaScript("document.fonts.ready.then(() => Promise.all(Array.from(document.images, image => image.decode().catch(() => {}))))");
    const bytes = await printWindow.webContents.printToPDF({
      pageSize: "A4", landscape: document.landscape, printBackground: true,
      margins: { top: 10 / 25.4, bottom: 10 / 25.4, left: 10 / 25.4, right: 10 / 25.4 },
    });
    await fs.writeFile(filePath, bytes);
    return filePath;
  } finally {
    try {
      if (!printWindow.isDestroyed()) printWindow.destroy();
    } finally {
      if (temporaryDirectory) await fs.rm(temporaryDirectory, { recursive: true, force: true });
    }
  }
}

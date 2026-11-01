import {
  type Cell, type Context, type Selection, getSheetIndex, getFlowdata, handleCopy,
  handlePasteByClick, isAllowEdit, hasPartMC, getBorderInfoCompute, functionCopy,
  expandRowsAndColumns, delFunctionGroup, insertUpdateFunctionGroup, execFunctionGroup, groupValuesRefresh,
  genarate, update, setFormulaCellInfoMap, getQKBorder,
} from "@fortune-sheet/core";

export type PasteMode = "all" | "values" | "formulas" | "formatting";
export interface ClipboardContents { html: string; text: string; }
interface ClipboardCell { cell: Cell | null; row: number; column: number; border?: Record<string, unknown>; }
interface SheetClipboard {
  version: 1;
  id: string;
  sheetId: string;
  mode: PasteMode;
  cells: ClipboardCell[][];
}
type ClipboardContext = Context & { sheetianClipboardId?: string; sheetianPasteFormulaKeys?: string[]; };
const MAX_CELLS = 1_000_000;
const styles: (keyof Cell)[] = ["bg", "ff", "fc", "bl", "it", "fs", "cl", "un", "vt", "ht", "tr", "tb", "rt"];
const clone = <T,>(value: T): T => value == null ? value : JSON.parse(JSON.stringify(value));
const escape = (value: string): string => value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function cellValue(cell: Cell | null): string | number | boolean | undefined {
  return cell?.ct?.t === "inlineStr" ? cell.ct.s?.map((run: { v?: string }) => run.v ?? "").join("") : cell?.v;
}
function textValue(cell: Cell | null, mode: PasteMode): string {
  return String(mode === "formulas" ? cell?.f ?? "" : cell?.m ?? cellValue(cell) ?? "");
}
// TSV quoting keeps tabs, newlines and double quotes inside a single cell.
export function toTsv(rows: string[][]): string {
  return rows.map(row => row.map(value => /[\t\n\r"]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value).join("\t")).join("\n");
}
export function parseTsv(text: string): string[][] {
  const rows: string[][] = []; let row: string[] = [], value = "", quoted = false;
  text = text.replace(/\r\n?/g, "\n");
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (char === '"' && (quoted || value === "")) {
      if (quoted && text[i + 1] === '"') { value += '"'; i++; }
      else quoted = !quoted;
    } else if (!quoted && (char === "\t" || char === "\n")) {
      row.push(value); value = "";
      if (char === "\n") { rows.push(row); row = []; }
    } else value += char;
  }
  if (row.length || value !== "" || !rows.length || !text.endsWith("\n")) { row.push(value); rows.push(row); }
  const width = rows.reduce((maximum, r) => Math.max(maximum, r.length), 0);
  return rows.map(r => Array.from({ length: width }, (_, c) => r[c] ?? ""));
}

export function readSheetClipboard(html: string): SheetClipboard | null {
  const encoded = html.match(/data-sheetian-clipboard="([^"]+)"/)?.[1];
  if (!encoded) return null;
  try {
    const data: SheetClipboard = JSON.parse(decodeURIComponent(encoded));
    const width = data.cells?.[0]?.length;
    if (data.version !== 1 || typeof data.id !== "string" || typeof data.sheetId !== "string" ||
        !["all", "values", "formulas", "formatting"].includes(data.mode) || !width ||
        data.cells.length * width > MAX_CELLS || !data.cells.every(row => Array.isArray(row) && row.length === width && row.every(entry =>
          entry && Number.isInteger(entry.row) && entry.row >= 0 && Number.isInteger(entry.column) && entry.column >= 0 &&
          (entry.cell === null || (typeof entry.cell === "object" && !Array.isArray(entry.cell) &&
            (entry.cell.f == null || typeof entry.cell.f === "string")))))) return null;
    return data;
  } catch { return null; }
}

export function writeDesktopClipboard(contents: ClipboardContents): void { require("electron").clipboard.write(contents); }
export function readDesktopClipboard(): ClipboardContents {
  const clipboard = require("electron").clipboard;
  return { html: clipboard.readHTML(), text: clipboard.readText() };
}

export function copyCells(ctx: ClipboardContext, mode: PasteMode = "all", cut = false,
  writer: (contents: ClipboardContents) => void = writeDesktopClipboard): void {
  if (cut && !isAllowEdit(ctx, ctx.luckysheet_select_save)) return;
  if (cut && ctx.luckysheet_select_save?.length !== 1) {
    ctx.luckysheet_paste_iscut = false;
    ctx.warnDialog = "Select one contiguous range to cut."; return;
  }
  let copied = false;
  (handleCopy as (ctx: Context, writer: (html: string) => void) => void)(ctx, html => {
    const ranges = ctx.luckysheet_select_save;
    const sheetId = ctx.currentSheetId;
    const data = getFlowdata(ctx);
    if (!ranges?.length || !data || !sheetId) return;
    const borders = getBorderInfoCompute(ctx, sheetId);
    const blocks = ranges.map(range => {
      const rows: ClipboardCell[][] = [];
      for (let r = range.row[0]; r <= range.row[1]; r++) {
        if (ctx.config.rowhidden?.[r] != null) continue;
        const row: ClipboardCell[] = [];
        for (let c = range.column[0]; c <= range.column[1]; c++) {
          if (ctx.config.colhidden?.[c] != null) continue;
          row.push({ cell: clone(data[r]?.[c] ?? null), row: r, column: c, border: clone(borders[`${r}_${c}`]) });
        }
        if (row.length) rows.push(row);
      }
      return rows;
    });
    const cells = blocks.length > 1 && ranges[0].row[0] === ranges[1].row[0] && ranges[0].row[1] === ranges[1].row[1]
      ? blocks[0].map((_, r) => blocks.flatMap(block => block[r] ?? [])) : blocks.flat();
    if (!cells.length || !cells[0].length) return;
    const snapshot: SheetClipboard = { version: 1, id: `${Date.now()}-${Math.random()}`, sheetId, mode, cells };
    const textRows = cells.map(row => row.map(entry => textValue(entry.cell, mode)));
    if (mode !== "all") html = `<table>${textRows.map(row => `<tr>${row.map(value => `<td>${escape(value).replace(/\n/g, "<br>")}</td>`).join("")}</tr>`).join("")}</table>`;
    html = html.replace(/<table\b/i, `<table data-sheetian-clipboard="${encodeURIComponent(JSON.stringify(snapshot))}"`);
    writer({ html, text: toTsv(textRows) });
    ctx.sheetianClipboardId = snapshot.id;
    copied = true;
  });
  // Never cut a previous selection when the engine refuses this copy.
  ctx.luckysheet_paste_iscut = cut && copied;
}

function shiftFormula(ctx: Context, formula: string, row: number, column: number): string {
  if (row) formula = `=${functionCopy(ctx, formula, row > 0 ? "down" : "up", Math.abs(row))}`;
  if (column) formula = `=${functionCopy(ctx, formula, column > 0 ? "right" : "left", Math.abs(column))}`;
  return formula;
}

function parseTextCell(text: string): Cell | null {
  if (text === "") return null;
  if (text.startsWith("=")) return { f: text, v: text };
  const [m, ct, v] = genarate(text) ?? [text, { fa: "General", t: "s" }, text];
  return { m, ct, v };
}

// Parse in a detached document: clipboard HTML must never execute or load into
// the workbook. Inline styles and Excel's simple class rules are sufficient for
// ordinary spreadsheet tables; native Sheetian copies retain exact cell data.
function externalCells(contents: ClipboardContents, ownerDocument?: Document): ClipboardCell[][] {
  if (contents.html && ownerDocument) {
    const doc = new ownerDocument.defaultView!.DOMParser().parseFromString(contents.html, "text/html");
    const table = doc.querySelector("table");
    if (table) {
      const rows: ClipboardCell[][] = [];
      const rules = Array.from(doc.querySelectorAll("style"), el => el.textContent ?? "").join("\n");
      Array.from(table.rows).forEach((tr, r) => {
        rows[r] ??= []; let c = 0;
        for (const td of Array.from(tr.cells)) {
          while (rows[r][c]) c++;
          const plain = td.cloneNode(true) as HTMLElement;
          plain.querySelectorAll("br").forEach(br => br.replaceWith("\n"));
          const cell = parseTextCell(plain.textContent ?? "") ?? {};
          for (const match of rules.matchAll(/([^{}]+)\{([^{}]+)\}/g)) {
            try { if (td.matches(match[1].trim())) td.style.cssText = match[2] + ";" + td.style.cssText; } catch { /* Unsupported clipboard selector. */ }
          }
          const style = td.style;
          if (style.backgroundColor) cell.bg = style.backgroundColor;
          if (style.color) cell.fc = style.color;
          if (style.fontFamily) cell.ff = style.fontFamily.replace(/["']/g, "");
          if (style.fontSize) cell.fs = parseFloat(style.fontSize) * (style.fontSize.endsWith("px") ? 0.75 : 1);
          if (style.fontWeight) cell.bl = style.fontWeight === "bold" || Number(style.fontWeight) >= 600 ? 1 : 0;
          if (style.fontStyle) cell.it = style.fontStyle === "italic" ? 1 : 0;
          if (style.textDecoration.includes("underline")) cell.un = 1;
          if (style.textDecoration.includes("line-through")) cell.cl = 1;
          if (style.textAlign) cell.ht = style.textAlign === "center" ? 0 : style.textAlign === "right" ? 2 : 1;
          if (style.verticalAlign) cell.vt = style.verticalAlign === "middle" ? 0 : style.verticalAlign === "top" ? 1 : 2;
          if (style.whiteSpace === "normal" || style.whiteSpace === "pre-wrap") cell.tb = "2";
          const numberFormat = style.getPropertyValue("mso-number-format").replace(/^["']|["']$/g, "").replace(/\\/g, "");
          if (numberFormat) cell.ct = { ...cell.ct, fa: numberFormat };
          const rs = Math.max(1, td.rowSpan), cs = Math.max(1, td.colSpan);
          if ((r + rs) * (c + cs) > MAX_CELLS) throw new Error("The clipboard selection is too large.");
          if (rs > 1 || cs > 1) {
            cell.mc = { r, c, rs, cs };
            for (let mr = r; mr < r + rs; mr++) {
              rows[mr] ??= [];
              for (let mc = c; mc < c + cs; mc++) rows[mr][mc] = { cell: { mc: { r, c } }, row: mr, column: mc };
            }
          }
          rows[r][c] = { cell, row: r, column: c }; c += cs;
          const sides = [["top", "t"], ["bottom", "b"], ["left", "l"], ["right", "r"]];
          for (const [side, key] of sides) {
            const line = style.getPropertyValue(`border-${side}-style`), thickness = style.getPropertyValue(`border-${side}-width`);
            if (!line || line === "none" || line === "hidden" || !parseFloat(thickness)) continue;
            const [borderStyle, color] = getQKBorder(thickness, line, style.getPropertyValue(`border-${side}-color`));
            for (let mr = r; mr < r + rs; mr++) for (let mc = c - cs; mc < c; mc++) {
              if ((key === "t" && mr !== r) || (key === "b" && mr !== r + rs - 1) ||
                  (key === "l" && mc !== c - cs) || (key === "r" && mc !== c - 1)) continue;
              rows[mr][mc].border ??= {};
              rows[mr][mc].border![key] = { style: borderStyle, color };
            }
          }
        }
      });
      const width = rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
      return rows.map((row, r) => Array.from({ length: width }, (_, c) => row[c] ?? { cell: null, row: r, column: c }));
    }
  }
  if (!contents.text) return [];
  return parseTsv(contents.text).map((row, r) => row.map((text, c) => ({ cell: parseTextCell(text), row: r, column: c })));
}

function withValue(destination: Cell | null, source: Cell | null): Cell {
  const result = { ...clone(destination ?? {}) };
  delete result.f; delete result.spl; delete result.qp;
  const value = cellValue(source);
  result.ct = { fa: result.ct?.fa ?? "General", t: typeof value === "number" ? "n" : typeof value === "boolean" ? "b" : "s" };
  if (value == null) { delete result.v; delete result.m; }
  else {
    result.v = value;
    result.m = typeof value === "string" ? value : String(update(result.ct.fa!, value));
  }
  return result;
}

export function pasteCells(ctx: ClipboardContext, contents: ClipboardContents, requestedMode: PasteMode = "all", ownerDocument?: Document): void {
  const selection = ctx.luckysheet_select_save;
  if (selection?.length !== 1) { ctx.warnDialog = "Select one destination range to paste into."; return; }
  const snapshot = readSheetClipboard(contents.html);
  const mode = requestedMode === "all" ? snapshot?.mode ?? "all" : requestedMode;
  const cells = snapshot?.cells ?? externalCells(contents, ownerDocument);
  if (!cells.length || !cells[0].length) return;
  const target = selection[0], height = cells.length, width = cells[0].length;
  const selectedHeight = target.row[1] - target.row[0] + 1, selectedWidth = target.column[1] - target.column[0] + 1;
  const repeat = selectedHeight % height === 0 && selectedWidth % width === 0;
  const rows = repeat ? selectedHeight : height, columns = repeat ? selectedWidth : width;
  if (rows * columns > MAX_CELLS) { ctx.warnDialog = "The clipboard selection is too large."; return; }
  const range: Selection = { row: [target.row[0], target.row[0] + rows - 1], column: [target.column[0], target.column[0] + columns - 1] };
  if (!isAllowEdit(ctx, [range])) return;
  if (hasPartMC(ctx, ctx.config, range.row[0], range.row[1], range.column[0], range.column[1])) {
    ctx.warnDialog = "Select the whole merged cell before pasting."; return;
  }
  if (ctx.hooks.beforePaste?.([range], contents.html || contents.text) === false) return;
  // Use the engine's move operation for cuts to retain formula source and move
  // merges and metadata. Only the workbook owning this clipboard token may move it.
  if (snapshot && mode === "all" && ctx.sheetianClipboardId === snapshot.id && ctx.luckysheet_paste_iscut) {
    const sourceIndex = getSheetIndex(ctx, snapshot.sheetId);
    if (sourceIndex == null) return;
    const sourceContext = { ...ctx, currentSheetId: snapshot.sheetId, config: ctx.luckysheetfile[sourceIndex].config ?? {} };
    if (!isAllowEdit(sourceContext, ctx.luckysheet_copy_save?.copyRange)) return;
    ctx.luckysheet_select_save = [{ ...range, row_focus: range.row[0], column_focus: range.column[0] }];
    handlePasteByClick(ctx, contents.html, "sheetian-cut"); return;
  }
  const data = getFlowdata(ctx);
  const index = getSheetIndex(ctx, ctx.currentSheetId);
  if (!data || index == null) return;
  expandRowsAndColumns(data, Math.max(0, range.row[1] + 1 - data.length), Math.max(0, range.column[1] + 1 - data[0].length));
  const config = ctx.config;
  const destinationBorders = mode === "all" || mode === "formatting" ? getBorderInfoCompute(ctx, ctx.currentSheetId) : {};
  const changed: { r: number; c: number; i: string }[] = [];
  const formulas: string[] = [];
  if (mode === "all") {
    config.merge ??= {};
    for (const [key, merge] of Object.entries(config.merge)) {
      if (merge.r >= range.row[0] && merge.r <= range.row[1] && merge.c >= range.column[0] && merge.c <= range.column[1]) delete config.merge[key];
    }
  }
  for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) {
    const row = range.row[0] + r, column = range.column[0] + c;
    const entry = cells[r % height][c % width], source = clone(entry.cell), destination = data[row][column];
    if (mode === "formulas" && !source?.f) continue;
    let value: Cell | null;
    if (mode === "values") value = withValue(destination, source);
    else if (mode === "formatting") {
      value = clone(destination ?? {});
      for (const attr of styles) { delete value[attr]; if (source?.[attr] != null) Object.assign(value, { [attr]: source[attr] }); }
      value.ct = { ...value.ct, fa: source?.ct?.fa ?? "General" };
      if (value.v != null) value.m = typeof value.v === "string" ? value.v : String(update(value.ct.fa!, value.v));
    } else {
      value = mode === "formulas" ? withValue(destination, source) : source;
      if (source?.f && value) {
        value.f = snapshot ? shiftFormula(ctx, source.f, row - entry.row, column - entry.column) : source.f;
        delete value.spl;
      }
      if (mode === "all" && value?.mc) {
        const mc = value.mc;
        const mergeRow = row + mc.r - entry.row, mergeColumn = column + mc.c - entry.column;
        value.mc = { ...mc, r: mergeRow, c: mergeColumn };
        if (mc.rs && mc.cs) config.merge![`${row}_${column}`] = { r: row, c: column, rs: mc.rs, cs: mc.cs };
      }
    }
    if (mode !== "formatting") {
      delFunctionGroup(ctx, row, column);
      changed.push({ r: row, c: column, i: ctx.currentSheetId! });
      if (value?.f) {
        insertUpdateFunctionGroup(ctx, row, column);
        formulas.push(`r${row}c${column}i${ctx.currentSheetId}`);
      }
    }
    data[row][column] = value;
    if ((mode === "all" || mode === "formatting") && (entry.border || destinationBorders[`${row}_${column}`])) {
      config.borderInfo ??= [];
      config.borderInfo.push({ rangeType: "cell", value: { row_index: row, col_index: column,
        l: null, r: null, t: null, b: null, ...clone(entry.border ?? {}) } });
    }
  }
  ctx.luckysheetfile[index].config = config;
  ctx.luckysheet_select_save = [range];
  ctx.luckysheetCellUpdate = [];
  if (changed.length) {
    ctx.formulaCache.formulaCellInfoMap = {};
    for (const sheet of ctx.luckysheetfile) setFormulaCellInfoMap(ctx, sheet.calcChain, sheet.data);
    ctx.formulaCache.execFunctionGlobalData = null;
    ctx.formulaCache.execFunctionExist = changed;
    ctx.sheetianPasteFormulaKeys = formulas;
    execFunctionGroup(ctx, 0, 0, null, undefined, data);
    delete ctx.sheetianPasteFormulaKeys;
    groupValuesRefresh(ctx);
    ctx.formulaCache.execFunctionGlobalData = null;
  }
}

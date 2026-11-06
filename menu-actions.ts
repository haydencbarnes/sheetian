import {
  type Context, getSheetIndex, isAllowEdit, toolbarItemClickHandler,
  handleSort, handleFreeze, createFilter, clearFilter, selectAll,
  deleteSelectedCellText, updateCell, moveHighlightCell, selectionCache,
} from "@fortune-sheet/core";
import { copyCells, pasteCells, readDesktopClipboard } from "./clipboard";

export interface MenuPayload { clipboard?: ClipboardEvent; zoom?: number; editor?: HTMLDivElement; moveDown?: boolean; writeClipboard?: (html: string) => void; }

// Runs through FortuneSheet's normal context update/history mechanism. Every
// action uses the invoking workbook's context, including in split panes.
export function runMenuAction(ctx: Context, action: string, input: HTMLDivElement | null, payload?: MenuPayload): void {
  if (action === "commit-formula" && payload?.editor) {
    const ranges = ctx.luckysheet_select_save;
    const selection = ranges?.[ranges.length - 1];
    const [row, column] = ctx.luckysheetCellUpdate.length ? ctx.luckysheetCellUpdate : [selection?.row_focus, selection?.column_focus];
    if (row != null && column != null && isAllowEdit(ctx, ranges)) {
      updateCell(ctx, row, column, payload.editor);
      if (payload.moveDown) moveHighlightCell(ctx, "down", 1, "rangeOfSelect");
    }
    return;
  }
  if (action === "find" || action === "replace") {
    ctx.showSearch = true; ctx.showReplace = action === "replace"; return;
  }
  if (action === "select-all") { selectAll(ctx); return; }
  if (action === "copy" || action === "cut" || action === "copy-values" || action === "copy-formulas") {
    const mode = action === "copy-values" ? "values" : action === "copy-formulas" ? "formulas" : "all";
    copyCells(ctx, mode, action === "cut", payload?.writeClipboard ? contents => payload.writeClipboard!(contents.html) : undefined);
    return;
  }
  const sheet = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)!];
  if (!sheet) return;
  if (action === "gridlines") { sheet.showGridLines = sheet.showGridLines === false || sheet.showGridLines === 0 ? 1 : 0; return; }
  if (action === "zoom" && payload?.zoom != null) {
    sheet.zoomRatio = ctx.zoomRatio = payload.zoom; return;
  }
  if (!isAllowEdit(ctx, ctx.luckysheet_select_save)) return;
  if (action === "clear-contents") { deleteSelectedCellText(ctx); return; }
  if (["paste", "paste-values", "paste-formulas", "paste-formatting"].includes(action)) {
    selectionCache.isPasteAction = false;
    const data = payload?.clipboard?.clipboardData;
    const contents = data ? { html: data.getData("text/html"), text: data.getData("text/plain") } : readDesktopClipboard();
    const mode = action === "paste-values" ? "values" : action === "paste-formulas" ? "formulas" : action === "paste-formatting" ? "formatting" : "all";
    pasteCells(ctx, contents, mode, input?.ownerDocument); return;
  }
  if (action === "sort-asc" || action === "sort-desc") { handleSort(ctx, action === "sort-asc"); return; }
  if (action === "filter") { createFilter(ctx); return; }
  if (action === "clear-filter") { clearFilter(ctx); return; }
  if (action.startsWith("freeze-")) { handleFreeze(ctx, action); return; }
  const handler = toolbarItemClickHandler(action);
  if (handler && input) handler(ctx, input);
}

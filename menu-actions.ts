import {
  type Context, getSheetIndex, isAllowEdit, toolbarItemClickHandler,
  handleSort, handleFreeze, createFilter, clearFilter, selectAll,
  deleteSelectedCellText, handleCopy, handlePaste, selectionCache, updateCell, moveHighlightCell,
} from "@fortune-sheet/core";

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
  if (action === "copy" || action === "cut") {
    if (action === "cut" && !isAllowEdit(ctx, ctx.luckysheet_select_save)) return;
    let copied = false;
    const writer = payload?.writeClipboard;
    (handleCopy as (ctx: Context, writer?: (html: string) => void) => void)(ctx, writer ? html => { writer(html); copied = true; } : undefined);
    // A refused partial-merge/disjoint copy must not cut a previous range.
    if (action === "cut") ctx.luckysheet_paste_iscut = copied;
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
  if (action === "paste" && payload?.clipboard) {
    selectionCache.isPasteAction = true; handlePaste(ctx, payload.clipboard); return;
  }
  if (action === "sort-asc" || action === "sort-desc") { handleSort(ctx, action === "sort-asc"); return; }
  if (action === "filter") { createFilter(ctx); return; }
  if (action === "clear-filter") { clearFilter(ctx); return; }
  if (action.startsWith("freeze-")) { handleFreeze(ctx, action); return; }
  const handler = toolbarItemClickHandler(action);
  if (handler && input) handler(ctx, input);
}

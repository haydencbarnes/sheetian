import * as React from "react";
import { flushSync } from "react-dom";
import { getFlowdata, updateContextWithSheetData } from "@fortune-sheet/core";
import type { CellMatrix, Context } from "@fortune-sheet/core";

function activeSheet(ctx: Context) {
  return ctx.luckysheetfile.find(sheet => sheet.id === ctx.currentSheetId);
}

// Object identity alone is not a geometry change. Native resize events can
// produce fresh config maps with identical sizes; measuring those again from
// a layout effect creates a render loop. Compare the actual grid boundaries.
export function sheetGeometryMatches(ctx: Context, data: CellMatrix): boolean {
  const columns = data[0]?.length ?? 0;
  if (ctx.visibledatarow.length !== data.length || ctx.visibledatacolumn.length !== columns) return false;
  let height = 0;
  for (let r = 0; r < data.length; r++) {
    if (ctx.config.rowhidden?.[r] == null) height += Math.round(((ctx.config.rowlen?.[r] || ctx.defaultrowlen) + 1) * ctx.zoomRatio);
    if (ctx.visibledatarow[r] !== height) return false;
  }
  if (ctx.rh_height !== height + 80) return false;
  let width = 0;
  for (let c = 0; c < columns; c++) {
    if (ctx.config.colhidden?.[c] == null) {
      let size = ctx.config.columnlen?.[c] || ctx.defaultcollen;
      if (!ctx.config.columnlen?.[c] && data[0]?.[c] && size > 300) size = 300;
      width += Math.round((size + 1) * ctx.zoomRatio);
    }
    if (ctx.visibledatacolumn[c] !== width) return false;
  }
  return ctx.ch_width === width;
}

// Preview only the drawing config. The saved sheet and its undo history are
// updated once, by the engine's usual mouseup transaction.
export function previewColumnResize(ctx: Context, pageX: number): boolean {
  const sheet = activeSheet(ctx), start = ctx.luckysheet_cols_change_size_start;
  if (!ctx.luckysheet_cols_change_size || !sheet || !Number.isFinite(start[2])) return false;
  const original = sheet.config ?? {}, column = start[1] as number;
  const delta = pageX - start[2];
  const columnlen = { ...original.columnlen };
  if (Math.abs(delta) >= 3) {
    const width = Math.ceil(Math.max(10, (original.columnlen?.[column] ?? ctx.defaultcollen) + delta / ctx.zoomRatio));
    const selected = ctx.luckysheet_select_save?.filter(range => range.column_select) ?? [];
    const ranges = selected.some(range => column >= range.column[0] && column <= range.column[1])
      ? selected : [{ column: [column, column] }];
    for (const range of ranges) for (let c = range.column[0]; c <= range.column[1]; c++) columnlen[c] = width;
  }
  // Native pointer/scroll events and React updater replays can repeat the
  // same position. Replacing an identical map retriggers the geometry layout
  // effect, which can enter a nested render loop at the last-column edge.
  const current = ctx.config.columnlen ?? {};
  const keys = new Set([...Object.keys(current), ...Object.keys(columnlen)]);
  if ([...keys].every(key => current[key] === columnlen[key])) return false;
  ctx.config = { ...original, columnlen };
  return true;
}

export function followColumnResize(ctx: Context, column: number, viewportWidth: number): void {
  const edge = ctx.visibledatacolumn[column];
  if (!Number.isFinite(edge) || viewportWidth <= 0) return;
  const frozen = activeSheet(ctx)?.frozen;
  const frozenColumn = frozen && /column|both/i.test(frozen.type) ? frozen.range?.column_focus ?? 0 : -1;
  const maximum = Math.max(0, ctx.ch_width - viewportWidth);
  let next = Math.min(ctx.scrollLeft, maximum);
  if (column > frozenColumn) {
    const left = (ctx.visibledatacolumn[frozenColumn] ?? 0) + 6;
    if (edge > next + viewportWidth) next = edge - viewportWidth;
    else if (edge < next + left) next = edge - left;
  }
  ctx.scrollLeft = Math.max(0, Math.min(maximum, next));
}

export function prepareColumnResizeCommit(ctx: Context, extraPixels: number): number | undefined {
  if (!ctx.luckysheet_cols_change_size) return;
  ctx.config = activeSheet(ctx)?.config ?? {};
  ctx.luckysheet_cols_change_size_start[2] -= extraPixels;
  return ctx.luckysheet_cols_change_size_start[1];
}

export function resizeHistoryScrollTarget(ctx: Context & { sheetianFollowResizeEnd?: boolean }, viewportWidth: number): number | undefined {
  if (!ctx.sheetianFollowResizeEnd) return;
  const sheet = activeSheet(ctx);
  const columns = sheet?.data?.[0]?.length ?? 0;
  let width = 0;
  for (let c = 0; c < columns; c++) {
    if (ctx.config.colhidden?.[c] != null) continue;
    width += Math.round(((ctx.config.columnlen?.[c] || ctx.defaultcollen) + 1) * ctx.zoomRatio);
  }
  // Undo commits saved data before the geometry effect rebuilds its columns.
  // Wait for that geometry before restoring the rightmost scroll position.
  if (ctx.ch_width === width) return Math.max(0, width - viewportWidth);
}

type Refs = {
  workbookContainer: React.RefObject<HTMLDivElement>;
  cellArea: React.RefObject<HTMLDivElement>;
  scrollbarX: React.RefObject<HTMLDivElement>;
};
type SetContext = (recipe: (ctx: Context) => void, options?: { noHistory: boolean }) => void;

export function useColumnResizePreview(context: Context, setContext: SetContext, refs: Refs) {
  const latest = React.useRef(context);
  React.useLayoutEffect(() => { latest.current = context; });
  const controller = React.useMemo(() => {
    let drag: { start: number; column: number; sheetId: string; pageX: number; extra: number } | null = null;
    let frame: number | null = null, lastTime = 0;
    const ownerWindow = () => refs.workbookContainer.current?.ownerDocument.defaultView;
    const stop = () => {
      if (frame !== null) ownerWindow()?.cancelAnimationFrame(frame);
      frame = null; lastTime = 0;
    };
    const paint = () => {
      if (!drag) return;
      const pageX = drag.pageX + drag.extra;
      flushSync(() => setContext(ctx => {
        if (!ctx.luckysheet_cols_change_size) return;
        if (previewColumnResize(ctx, pageX)) {
          const data = getFlowdata(ctx);
          if (data) updateContextWithSheetData(ctx, data);
        }
        followColumnResize(ctx, ctx.luckysheet_cols_change_size_start[1], refs.scrollbarX.current?.clientWidth ?? 0);
      }, { noHistory: true }));
    };
    const atEdge = () => {
      const ctx = latest.current, area = refs.cellArea.current, win = ownerWindow();
      if (!drag || !area || !win || !ctx.luckysheet_cols_change_size || ctx.currentSheetId !== drag.sheetId) return false;
      // A simple click must not start resizing. Continue only an outward drag
      // of the last visible column while the pointer is held at the right edge.
      const last = ctx.visibledatacolumn.length - 1;
      return ctx.visibledatacolumn[drag.column] === ctx.visibledatacolumn[last]
        && drag.pageX - drag.start >= 3
        && drag.pageX - win.scrollX >= area.getBoundingClientRect().right - 8;
    };
    const tick = (time: number) => {
      frame = null;
      if (!atEdge() || !drag) { lastTime = 0; return; }
      if (lastTime) { drag.extra += Math.min(32, time - lastTime) * 0.18; paint(); }
      lastTime = time;
      frame = ownerWindow()!.requestAnimationFrame(tick);
    };
    return {
      move(event: MouseEvent): boolean {
        const ctx = latest.current, start = ctx.luckysheet_cols_change_size_start;
        if (!ctx.luckysheet_cols_change_size) return false;
        if (!drag || drag.start !== start[2] || drag.sheetId !== ctx.currentSheetId) {
          stop(); drag = { start: start[2], column: start[1], sheetId: ctx.currentSheetId, pageX: event.pageX, extra: 0 };
        }
        drag.pageX = event.pageX; paint();
        if (atEdge()) { if (frame === null) frame = ownerWindow()!.requestAnimationFrame(tick); }
        else stop();
        return true;
      },
      finish(): number { stop(); const extra = drag?.extra ?? 0; drag = null; return extra; },
      stop,
    };
  }, [setContext, refs]);
  React.useEffect(() => {
    const win = refs.workbookContainer.current?.ownerDocument.defaultView;
    win?.addEventListener("blur", controller.stop);
    return () => { controller.finish(); win?.removeEventListener("blur", controller.stop); };
  }, [controller, refs]);
  return controller;
}

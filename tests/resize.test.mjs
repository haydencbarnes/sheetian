import { test } from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import { readFile } from "node:fs/promises";
import produce, { applyPatches, enablePatches, produceWithPatches } from "immer";
import { patchCore, patchHeaderResizing, patchHistoryLayout } from "../engine-compat.mjs";

enablePatches();

test("header drags expand at the viewport edge, scale once, and ignore clicks", async () => {
  for (const file of ["index.js", "index.esm.js"]) {
    const original = await readFile(`node_modules/@fortune-sheet/core/dist/${file}`, "utf8");
    const source = patchHeaderResizing(original);
    const start = source.indexOf("  if (ctx.luckysheet_rows_change_size) {", source.indexOf("function handleOverlayMouseUp("));
    const end = source.indexOf("  if (ctx.luckysheet_cols_freeze_drag) {", start);
    const getSheetIndex = () => 0;
    const resize = vm.runInNewContext("(function(ctx, e, rect) {" + source.slice(start, end) + "})", {
      getSheetIndex, core: { getSheetIndex }, window: { scrollX: 0, scrollY: 0 },
    });
    for (const zoomRatio of [0.5, 1, 1.3, 2]) for (const axis of ["rows", "cols"]) for (const grouped of [false, true]) {
      for (const delta of [0, 2, 60, -1000]) {
        const row = axis === "rows", size = row ? 29 : 100, index = row ? 2 : 25;
        const length = row ? "rowlen" : "columnlen", custom = row ? "customHeight" : "customWidth";
        const start = 1000, scroll = 500, rect = { left: 20, top: 50, width: 600, height: 400 };
        const config = { [length]: { [index]: size } };
        const ctx = { zoomRatio, rowHeaderWidth: 46, columnHeaderHeight: 20, scrollLeft: scroll, scrollTop: scroll,
          config, luckysheetfile: [{ config }], currentSheetId: "sheet", defaultrowlen: 19, defaultcollen: 73,
          [`luckysheet_${axis}_change_size`]: true, [`luckysheet_${axis}_change_size_start`]: [start, index,
            start + (row ? rect.top + 20 : rect.left + 46) - scroll],
          luckysheet_select_save: grouped ? [{ [row ? "row_select" : "column_select"]: true,
            row: row ? [index - 1, index] : [0, 99], column: row ? [0, 25] : [index - 1, index] }] : [],
        };
        const e = row ? { pageY: start + delta + rect.top + ctx.columnHeaderHeight - scroll }
          : { pageX: start + delta + rect.left + ctx.rowHeaderWidth - scroll };
        resize(ctx, e, rect);
        const expected = Math.abs(delta) < 3 ? size : Math.ceil(Math.max(10, size + delta / zoomRatio));
        assert.equal(ctx.config[length][index], expected, `${file}, ${axis}, zoom ${zoomRatio}, delta ${delta}`);
        assert.equal(ctx[`luckysheet_${axis}_change_size`], false);
        if (Math.abs(delta) < 3) {
          assert.equal(ctx.config[custom], undefined, "clicking a boundary must not mark it resized");
        } else {
          assert.equal(ctx.luckysheetfile[0].config[length][index], expected);
          if (grouped) assert.equal(ctx.config[length][index - 1], expected, "selected headers must resize together");
        }
      }
    }
    assert.doesNotThrow(() => patchCore(original));
  }
  assert.throws(() => patchHeaderResizing("future engine"));
});

test("undo and redo restore the active drawing config as well as saved resize data", async () => {
  for (const file of ["index.js", "index.esm.js"]) {
    const source = patchHistoryLayout(await readFile(`node_modules/@fortune-sheet/react/dist/${file}`, "utf8"));
    const entries = [...source.matchAll(/var newContext = [^\n]+history\.(inversePatches|patches)\)[\s\S]*?\n        \}\);/g)];
    assert.equal(entries.length, 2);
    const getSheetIndex = (ctx, id) => ctx.luckysheetfile.findIndex(sheet => sheet.id === id);
    for (const dimension of ["columnlen", "rowlen"]) {
      const config = { [dimension]: { 2: 73 } };
      const before = { currentSheetId: "active", config, luckysheetfile: [{ id: "active", config }] };
      const [after, patches, inversePatches] = produceWithPatches(before, draft => {
        draft.config[dimension][2] = 150;
        draft.luckysheetfile[0].config = draft.config;
      });
      const history = { patches: patches.filter(patch => patch.path[0] === "luckysheetfile"),
        inversePatches: inversePatches.filter(patch => patch.path[0] === "luckysheetfile") };
      const apply = (entry, ctx_) => vm.runInNewContext(entry[0] + "\nnewContext;", {
        ctx_, history, applyPatches, produce: Object.assign((...args) => produce(...args), { applyPatches }),
        produce__default: { default: produce }, getSheetIndex, core: { getSheetIndex },
      });
      const undone = apply(entries[0], after);
      assert.equal(undone.config[dimension][2], 73);
      assert.equal(undone.luckysheetfile[0].config[dimension][2], 73);
      const redone = apply(entries[1], undone);
      assert.equal(redone.config[dimension][2], 150);
      assert.equal(redone.luckysheetfile[0].config[dimension][2], 150);
      assert.equal(before.config[dimension][2], 73, "history must not mutate earlier snapshots");
    }
    assert.match(source, /useLayoutEffect\(function \(\) \{\n    if \(!data\) return;/);
  }
  assert.throws(() => patchHistoryLayout("future engine"));
});

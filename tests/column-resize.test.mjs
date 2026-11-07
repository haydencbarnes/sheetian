import { test } from "node:test";
import assert from "node:assert/strict";
import { build } from "esbuild";
import vm from "node:vm";
import { produceWithPatches, enablePatches } from "immer";
import { readFile } from "node:fs/promises";
import { patchColumnResizePreview, patchHistoryLayout } from "../engine-compat.mjs";

enablePatches();
const built = await build({ entryPoints: ["column-resize.ts"], bundle: true, write: false, format: "cjs", platform: "node", packages: "external" });
const module = { exports: {} };
vm.runInNewContext(built.outputFiles[0].text, { module, exports: module.exports, require: () => ({}) });
const { previewColumnResize, followColumnResize, prepareColumnResizeCommit, resizeHistoryScrollTarget, sheetGeometryMatches } = module.exports;

test("live previews preserve saved data and use the original widths for every move", () => {
  for (const zoomRatio of [0.5, 0.9, 1, 1.3, 2]) for (const grouped of [false, true]) {
    const config = { columnlen: { 24: 80, 25: 120 } };
    const original = { currentSheetId: 'last', luckysheetfile: [{ id: 'last', config }], config, defaultcollen: 73, zoomRatio,
      luckysheet_cols_change_size: true, luckysheet_cols_change_size_start: [1000, 25, 500],
      luckysheet_select_save: grouped ? [{ column: [24, 25], column_select: true }] : [],
    };
    let current = original;
    for (const delta of [40, 100, -30, 0]) {
      const [next, patches] = produceWithPatches(current, ctx => { previewColumnResize(ctx, 500 + delta); });
      assert(patches.every(patch => patch.path[0] !== 'luckysheetfile'), 'Preview must not enter workbook history');
      assert.equal(next.luckysheetfile[0].config.columnlen[25], 120);
      const expected = Math.abs(delta) < 3 ? 120 : Math.ceil(Math.max(10, 120 + delta / zoomRatio));
      assert.equal(next.config.columnlen[25], expected);
      if (grouped && delta) assert.equal(next.config.columnlen[24], expected);
      const [repeated, repeatedPatches] = produceWithPatches(next, ctx => {
        assert.equal(previewColumnResize(ctx, 500 + delta), false);
      });
      assert.equal(repeatedPatches.length, 0, 'Repeated pointer positions must not restart geometry rendering');
      assert.equal(repeated, next);
      current = next;
    }
    const [commit] = produceWithPatches(current, ctx => {
      assert.equal(prepareColumnResizeCommit(ctx, 30), 25);
      assert.equal(ctx.config.columnlen[25], 120);
      assert.equal(ctx.luckysheet_cols_change_size_start[2], 470);
    });
    assert.equal(commit.luckysheetfile[0].config.columnlen[24], 80);
  }
});

test("equivalent config maps do not restart grid measurement at 90% zoom", () => {
  const data = Array.from({ length: 3 }, () => Array(4).fill(null));
  const ctx = { defaultrowlen: 19, defaultcollen: 73, zoomRatio: 0.9,
    config: { rowlen: { 0: 12 }, rowhidden: { 1: 0 }, columnlen: { 0: 19, 3: 156 }, colhidden: { 2: 0 } },
    visibledatarow: [12, 12, 30], rh_height: 110, visibledatacolumn: [18, 85, 85, 226], ch_width: 226 };
  assert.equal(sheetGeometryMatches(ctx, data), true);
  for (let n = 0; n < 100; n++) {
    ctx.config = { ...ctx.config, columnlen: { ...ctx.config.columnlen } };
    assert.equal(sheetGeometryMatches(ctx, data), true, 'New object identities with the same widths must settle');
  }
  ctx.config.columnlen[3] = 170;
  assert.equal(sheetGeometryMatches(ctx, data), false);
  ctx.config.columnlen[3] = 156; ctx.config.rowlen[0] = 20;
  assert.equal(sheetGeometryMatches(ctx, data), false);
});

test("following a column keeps its edge visible and clamps scrolling to the real grid", () => {
  const ctx = { currentSheetId: 'last', luckysheetfile: [{ id: 'last' }], visibledatacolumn: [100, 300, 1100], ch_width: 1100, scrollLeft: 200 };
  followColumnResize(ctx, 2, 500); assert.equal(ctx.scrollLeft, 600);
  ctx.visibledatacolumn[2] = ctx.ch_width = 800;
  followColumnResize(ctx, 2, 500); assert.equal(ctx.scrollLeft, 300);
  followColumnResize(ctx, 0, 500); assert.equal(ctx.scrollLeft, 94);
  ctx.luckysheetfile[0].frozen = { type: 'column', range: { column_focus: 0 } };
  ctx.scrollLeft = 250;
  followColumnResize(ctx, 0, 500); assert.equal(ctx.scrollLeft, 250);
  followColumnResize(ctx, 2, 500); assert.equal(ctx.scrollLeft, 300);
});

test("history waits for restored column geometry before following the final edge", () => {
  const ctx = { currentSheetId: 'last', luckysheetfile: [{ id: 'last', data: [Array(26).fill(null)] }],
    config: { columnlen: { 25: 120 }, colhidden: { 24: 0 } }, defaultcollen: 73, zoomRatio: 1.3,
    ch_width: 3000, sheetianFollowResizeEnd: true };
  assert.equal(resizeHistoryScrollTarget(ctx, 500), undefined, 'An old geometry commit must not settle history scrolling');
  const width = 24 * Math.round(74 * 1.3) + Math.round(121 * 1.3);
  ctx.ch_width = width;
  assert.equal(resizeHistoryScrollTarget(ctx, 500), width - 500);
  delete ctx.sheetianFollowResizeEnd;
  assert.equal(resizeHistoryScrollTarget(ctx, 500), undefined);
});

test("the preview patch guards both pinned engine distributions", async () => {
  for (const file of ['index.js', 'index.esm.js']) {
    const source = await readFile(`node_modules/@fortune-sheet/react/dist/${file}`, 'utf8');
    assert.doesNotThrow(() => patchColumnResizePreview(patchHistoryLayout(source)));
  }
  assert.throws(() => patchColumnResizePreview('future engine'));
});

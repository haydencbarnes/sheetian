import { test } from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import { readFile } from "node:fs/promises";
import lodash from "lodash";
import { patchWheelScroll } from "../engine-compat.mjs";

async function wheel(file = "index.esm.js") {
  const source = patchWheelScroll(await readFile(`node_modules/@fortune-sheet/core/dist/${file}`, "utf8"));
  const start = source.indexOf("function handleGlobalWheel(");
  const end = source.indexOf("function fixPositionOnFrozenCells(", start);
  const context = vm.createContext({ _: lodash, ___default: {default:lodash}, mouseWheelUniqueTimeout: null, scrollLockTimeout: null,
    clearTimeout() {}, setTimeout: () => 0 });
  vm.runInContext(source.slice(start, end), context);
  return context.handleGlobalWheel;
}
function scrollbar(ratio, max = Infinity) {
  let top = 0;
  return { ownerDocument: { defaultView: { devicePixelRatio: ratio } },
    get scrollTop() { return top; },
    set scrollTop(value) { top = Math.min(max, Math.max(0, Math.round(value * ratio) / ratio)); },
  };
}
const event = deltaY => ({ deltaY, deltaX: 0, preventDefault() {} });

test("upward wheel scrolling cannot stick at a row boundary rounded to physical pixels", async () => {
  const handler = await wheel();
  for (const ratio of [0.9128709435462952, 1, 1.25, 1.5, 2]) {
    for (const zoom of [0.5, 1, 1.3, 1.75]) {
      const rows = Array.from({length: 100}, (_,i) => (i+1)*Math.round(20*zoom));
      const ctx = { visibledatarow: rows, visibledatacolumn: [73], zoomRatio: zoom };
      const bar = scrollbar(ratio), cache = {};
      for (let i=0;i<12;i++) handler(ctx, event(100), cache, {scrollLeft:0}, bar);
      assert(bar.scrollTop > 0);
      for (let i=0;i<100 && bar.scrollTop > 0;i++) {
        const before = bar.scrollTop;
        handler(ctx, event(-100), cache, {scrollLeft:0}, bar);
        assert(bar.scrollTop < before, `Stuck at ${before}, pixel ratio ${ratio}, sheet zoom ${zoom}`);
      }
      assert.equal(bar.scrollTop, 0);
      handler(ctx, event(-100), cache, {scrollLeft:0}, bar);
      assert.equal(bar.scrollTop, 0);
    }
  }
});
test("upward scrolling works from an arbitrary bottom clamp with hidden and differently sized rows", async () => {
  const handler = await wheel();
  const ctx = { visibledatarow: [26,26,52,99,99,125,151,177,203,229,255], visibledatacolumn: [73], zoomRatio: 1.3 };
  const bar = scrollbar(0.9128709435462952, 181.8), cache = {};
  bar.scrollTop = 999;
  let count = 0;
  while (bar.scrollTop > 0 && count++ < 20) {
    const before = bar.scrollTop;
    handler(ctx, event(-100), cache, {scrollLeft:0}, bar);
    assert(bar.scrollTop < before);
  }
  assert.equal(bar.scrollTop, 0);
});
test("the scroll patch matches both pinned engine distributions and rejects source changes", async () => {
  assert.equal(typeof await wheel("index.js"), "function");
  assert.throws(() => patchWheelScroll("future engine"));
});

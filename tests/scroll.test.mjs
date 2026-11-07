import { test } from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import { readFile } from "node:fs/promises";
import lodash from "lodash";
import { patchWheelScroll, patchScrollSynchronization, patchColumnScrollExtent } from "../engine-compat.mjs";

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

test("horizontal scroll extent ends at the last visible column at every zoom", async () => {
  for (const file of ["index.js", "index.esm.js"]) {
    const source = patchColumnScrollExtent(await readFile(`node_modules/@fortune-sheet/core/dist/${file}`, "utf8"));
    const start = source.indexOf("function calcRowColSize(");
    const end = source.indexOf("function ensureSheetIndex(", start);
    const calcRowColSize = vm.runInNewContext(source.slice(start, end) + "; calcRowColSize;", { getFlowdata: () => [] });
    for (const zoomRatio of [0.5, 1, 1.3, 2]) for (const columnCount of [1, 26]) {
      const ctx = { zoomRatio, defaultcollen: 73, defaultrowlen: 19,
        config: { columnlen: { 0: 19, 1: 546, 2: 60, 3: 292, 25: 100 }, colhidden: { 8: 0, 24: 0 } } };
      calcRowColSize(ctx, 100, columnCount);
      const expected = Array.from({ length: columnCount }, (_, c) =>
        ctx.config.colhidden[c] != null ? 0 : Math.round(((ctx.config.columnlen[c] ?? 73) + 1) * zoomRatio)
      ).reduce((sum, width) => sum + width, 0);
      assert.equal(ctx.ch_width, expected, `${file}, zoom ${zoomRatio}, ${columnCount} columns`);
      assert.equal(ctx.visibledatacolumn.at(-1), expected);
      // The same calculation must shrink its extent after deleting columns.
      calcRowColSize(ctx, 100, 0);
      assert.equal(ctx.ch_width, 0);
    }
  }
  assert.throws(() => patchColumnScrollExtent("future engine"));
});

test("scrollbar drags commit their captured position synchronously without adding undo history", async () => {
  for (const file of ["index.js", "index.esm.js"]) {
    const patched = patchScrollSynchronization(await readFile(`node_modules/@fortune-sheet/react/dist/${file}`, "utf8"));
    const component = patched.slice(patched.indexOf("var ScrollBar = function ScrollBar("), patched.indexOf("var Dialog = function Dialog("));
    for (const axis of ["x", "y"]) {
      const property = axis === "x" ? "scrollLeft" : "scrollTop";
      const context = { scrollLeft: 0, scrollTop: 0 };
      const bar = { scrollLeft: 0, scrollTop: 0 }, effects = [];
      let flushing = false, commits = 0;
      const react = {
        createElement: (type, props) => ({ type, props }),
        useContext: () => ({ context, refs: { scrollbarX: { current: bar }, scrollbarY: { current: bar } },
          setContext(recipe, options) {
            assert(flushing, "a continuous scroll event must commit before it returns");
            assert.equal(options.noHistory, true);
            recipe(context); commits++;
          } }),
        useLayoutEffect: callback => effects.push(callback),
        useEffect: () => assert.fail("scroll offsets must not be written after paint"),
      };
      const render = vm.runInNewContext(component + "; ScrollBar;", {
        ...react, React: react, React__default: { default: react }, WorkbookContext: {},
        sheetianFlushScroll(callback) {
          flushing = true;
          // A pending React commit can change the element while flushSync
          // drains work. The event's captured position must still win.
          bar[property] = 17;
          try { callback(); } finally { flushing = false; }
        },
      });
      const tree = render({ axis });
      assert.equal(effects.length, 1);
      effects[0]();
      for (const position of [160, 25, 480, 0]) {
        bar[property] = position;
        tree.props.onScroll({ currentTarget: bar });
        assert.equal(context[property], position, `${file}, ${axis}, ${position}`);
      }
      assert.equal(commits, 4);
      bar[property] = context[property];
      tree.props.onScroll({ currentTarget: bar });
      assert.equal(commits, 4, "programmatic synchronization must not cause a feedback update");
      assert.equal(context[axis === "x" ? "scrollTop" : "scrollLeft"], 0);
    }
    for (const target of ["containerRef.current.scrollLeft", "containerRef.current.scrollTop", "refs.cellArea.current.scrollLeft"]) {
      assert(patched.includes(`useLayoutEffect(function () {\n    ${target}`), `${target} must synchronize before paint`);
    }
  }
  assert.throws(() => patchScrollSynchronization("future engine"));
});

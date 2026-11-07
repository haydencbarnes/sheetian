import { test } from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import { readFile } from "node:fs/promises";
import { build } from "esbuild";
import { patchContextMenuLayout } from "../engine-compat.mjs";

const built = await build({ entryPoints: ["menu-layout.ts"], bundle: true, write: false, format: "cjs", platform: "node" });
const module = { exports: {} };
vm.runInNewContext(built.outputFiles[0].text, { module, exports: module.exports });
const { fitContextMenu } = module.exports;

test("menus fit split panes, short panes and the visible part of a window", () => {
  for (const [left, top, width, height, x, y] of [
    [1159, 108, 818, 808, 753, 70], [1159, 108, 818, 808, 790, 790],
    [500, 100, 220, 260, 190, 230], [-40, -20, 400, 300, -50, -50],
  ]) {
    const workbook = { clientWidth: width, clientHeight: height,
      ownerDocument: { defaultView: { innerWidth: 1977, innerHeight: 984 } },
      getBoundingClientRect: () => ({ left, top }),
    };
    const menu = { style: {}, getBoundingClientRect() {
      return { width: Math.min(315, parseFloat(this.style.maxWidth)), height: Math.min(627, parseFloat(this.style.maxHeight)) };
    } };
    const position = fitContextMenu(menu, workbook, { x, y });
    const size = menu.getBoundingClientRect();
    assert(position.x >= Math.max(0, -left) && position.y >= Math.max(0, -top));
    assert(position.x + size.width <= width && position.y + size.height <= height);
    assert(left + position.x >= 0 && top + position.y >= 0);
    assert(left + position.x + size.width <= 1977 && top + position.y + size.height <= 984);
  }
});

test("menu focus stays in its workbook without scrolling, and clamping settles without history", async () => {
  for (const file of ["index.js", "index.esm.js"]) {
    const source = patchContextMenuLayout(await readFile(`node_modules/@fortune-sheet/react/dist/${file}`, "utf8"));
    const focus = source.match(/(?:React\.)?useEffect\((function \(\) \{\n    var element = containerRef.current;[\s\S]*?\n  \}), \[\]\);/)[1];
    let focuses = 0;
    const menu = { querySelector: () => first };
    const first = { closest: () => menu, focus(options) { focuses++; assert.equal(options.preventScroll, true); } };
    for (const element of [first, { closest: () => menu, focus: () => assert.fail("Only the first menu item should take focus") }, null]) {
      vm.runInNewContext(`(${focus})()`, { containerRef: { current: element },
        document: { querySelector: () => assert.fail("Menu focus must not query another workbook") } });
    }
    assert.equal(focuses, 1);
    const layout = source.match(/(?:React\.)?useLayoutEffect\((function \(\) \{\n    var menu = containerRef.current, workbook = refs.workbookContainer.current;[\s\S]*?\n  \})\);/)[1];
    const contextMenu = { x: 750, y: 90 };
    let updates = 0;
    const clamp = vm.runInNewContext(`(${layout})`, { contextMenu, containerRef: { current: menu },
      refs: { workbookContainer: { current: {} } }, fitContextMenu: () => ({ x: 400, y: 90 }),
      setContext(recipe, options) { updates++; assert.equal(options.noHistory, true); recipe({ contextMenu }); },
    });
    clamp(); clamp();
    assert.equal(updates, 1);
    assert.equal(contextMenu.x, 400);
  }
  assert.throws(() => patchContextMenuLayout("future engine"));
});

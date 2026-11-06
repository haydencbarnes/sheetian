import { test } from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import { readFile } from "node:fs/promises";
import { patchEngine, patchDelayedFocus, patchPasteOwnership, patchToolbarZoom, patchFormulaBar, patchScrollbarLayout } from "../engine-compat.mjs";

test("delayed focus is safe after the workbook unmounts", () => {
  let callback, calls = 0;
  const inputRef = { current: { isConnected: true } };
  vm.runInNewContext(patchDelayedFocus("setTimeout(() => { moveToEnd(inputRef.current); });"), {
    inputRef, setTimeout: fn => { callback = fn; }, moveToEnd: () => { calls++; },
  });
  callback(); assert.equal(calls, 1);
  inputRef.current = null;
  assert.doesNotThrow(callback); assert.equal(calls, 1);
});
test("inactive workbooks cannot consume a paste event for the active workbook", () => {
  let calls = 0;
  const document = { activeElement: {} };
  let owns = false;
  const workbookContainer = { current: { contains: () => owns } };
  const code = patchPasteOwnership("var onPaste = React.useCallback(function (e) {\n    var _document$activeEleme; consumePaste(); }); onPaste({});");
  const context = { document, workbookContainer, React: { useCallback: fn => fn }, consumePaste: () => calls++ };
  vm.runInNewContext(code, context); assert.equal(calls, 0);
  owns = true;
  vm.runInNewContext(code, context); assert.equal(calls, 1);
});
test("formula bar resizing measures and draws the viewport after its height commits", async () => {
  for (const file of ["index.js", "index.esm.js"]) {
    const patched = patchFormulaBar(await readFile(`node_modules/@fortune-sheet/react/dist/${file}`, "utf8"));
    const callback = patched.match(/onResize: (function \(height\) \{[\s\S]+?\n    \})/)[1];
    const editor = { style: { height: "29px" } }, canvas = { width: 800, height: 571 };
    const placeholder = { clientWidth: 800, get clientHeight() { return 600 - parseInt(editor.style.height); } };
    const ctx = { calculatebarHeight: 29, luckysheetTableContentHW: [800, 571], rowHeaderWidth: 46, columnHeaderHeight: 20,
      devicePixelRatio: 1, cellmainWidth: 754, cellmainHeight: 551 };
    let updates = 0, measurements = 0;
    const updateContextWithCanvas = (draft, element, measured) => {
      measurements++;
      assert.equal(element, canvas);
      assert.equal(draft.calculatebarHeight, parseInt(editor.style.height));
      draft.luckysheetTableContentHW = [measured.clientWidth, measured.clientHeight];
      draft.cellmainWidth = measured.clientWidth - draft.rowHeaderWidth;
      draft.cellmainHeight = measured.clientHeight - draft.columnHeaderHeight;
      element.width = measured.clientWidth * draft.devicePixelRatio;
      element.height = measured.clientHeight * draft.devicePixelRatio;
    };
    const resize = vm.runInNewContext(`(${callback})`, {
      refs: { workbookContainer: { current: { querySelector: selector => selector === ".fortune-fx-editor" ? editor : placeholder } }, canvas: { current: canvas } },
      setContext: (recipe, options) => { updates++; assert.equal(options.noHistory, true); recipe(ctx); },
      core: { updateContextWithCanvas }, updateContextWithCanvas,
    });
    resize(160);
    assert.equal(updates, 1);
    assert.equal(measurements, 0, "requesting a height must not clear or measure the old canvas");
    assert.equal(editor.style.height, "29px");
    // React applies the height before the layout effect measures the viewport.
    editor.style.height = "160px";
    const measure = patched.match(/(?:React\.)?useLayoutEffect\((function \(\) \{\n    var placeholder = placeholderRef.current;[\s\S]+?\n  \})\);/)[1];
    const commitLayout = vm.runInNewContext(`(${measure})`, {
      refs: { canvas: { current: canvas } }, placeholderRef: { current: placeholder }, context: ctx,
      setContext: (recipe, options) => { updates++; assert.equal(options.noHistory, true); recipe(ctx); },
      core: { updateContextWithCanvas }, updateContextWithCanvas,
    });
    commitLayout();
    assert.equal(updates, 2);
    assert.equal(measurements, 1);
    assert.equal(ctx.calculatebarHeight, 160);
    assert.deepEqual(ctx.luckysheetTableContentHW, [800, 440]);
    commitLayout();
    assert.equal(updates, 2, "stable geometry must not clear the bitmap or trigger another render");
    assert.match(patched, /useLayoutEffect\(function \(\) \{\n    var _refs\$globalCache\$fre, _freeze\$horizontal, _freeze\$vertical;/);
    assert.match(patched, /useLayoutEffect\(function \(\) \{\n    (?:core\.)?initFreeze\(context, refs.globalCache, context.currentSheetId\);/);
  }
});
test("scrollbars are siblings of the overlay and size themselves from the sheet frame", async () => {
  const react = { createElement: (type, props, ...children) => ({ type, props, children }) };
  for (const file of ["index.js", "index.esm.js"]) {
    const patched = patchScrollbarLayout(await readFile(`node_modules/@fortune-sheet/react/dist/${file}`, "utf8"));
    const sheet = patched.slice(patched.indexOf("var Sheet = function Sheet("), patched.indexOf("var Button = function Button("));
    const render = sheet.slice(sheet.lastIndexOf("\n  return ") + "\n  return ".length, sheet.lastIndexOf("\n};"));
    const tree = vm.runInNewContext(render, {
      React: react, React__default: { default: react },
      containerRef: {}, placeholderRef: {}, refs: { canvas: {} },
      SheetOverlay: "overlay", ScrollBar: "scrollbar",
    });
    assert.equal(tree.props.className, "fortune-sheet-container");
    assert.equal(tree.children[2].type, "overlay");
    assert.deepEqual(Array.from(tree.children.slice(3), child => [child.type, child.props.axis]), [["scrollbar", "x"], ["scrollbar", "y"]]);
    assert.match(patched, /right: "var\(--sheetian-scrollbar-size\)"/);
    assert.match(patched, /bottom: "var\(--sheetian-scrollbar-size\)"/);
  }
  assert.throws(() => patchScrollbarLayout("unexpected future engine"));
});
test("the compatibility patch still matches the pinned engine's two distributions", async () => {
  for (const file of ["index.js", "index.esm.js"]) {
    const patched = patchEngine(await readFile(`node_modules/@fortune-sheet/react/dist/${file}`, "utf8"));
    assert.match(patched, /if \(inputRef.current\?\.isConnected\)/);
    assert.match(patched, /height: context.calculatebarHeight/);
    assert.match(patched, /runMenuAction: function \(action, payload\)/);
    assert.match(patched, /createElement\(FormulaBarControls/);
    assert.match(patched, /name === "sheetian-zoom"/);
    assert.match(patched, /createElement\(ToolbarZoom/);
    assert.match(patched, /settings\.customToolbarItems, sheetWidth/);
    assert.doesNotMatch(patched, /createElement\(ZoomControl, null\)/);
    assert.match(patched, /"data-sheetian-column": allowEditRef.current \? hoverLocation.col_index : undefined/);
    assert.match(patched, /"data-sheetian-row": hoverLocation.row_index >= 0 && (?:core\.)?isAllowEdit\(context\)/);
    assert.ok(patched.includes(`&& ${file === "index.js" ? "core." : ""}isAllowEdit(context)`));
  }
  assert.throws(() => patchEngine("unexpected future engine"));
  assert.throws(() => patchToolbarZoom("unexpected future engine"));
});

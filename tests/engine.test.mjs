import { test } from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import { readFile } from "node:fs/promises";
import { patchEngine, patchDelayedFocus, patchPasteOwnership, patchToolbarZoom } from "../engine-compat.mjs";

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
test("the compatibility patch still matches the pinned engine's two distributions", async () => {
  for (const file of ["index.js", "index.esm.js"]) {
    const patched = patchEngine(await readFile(`node_modules/@fortune-sheet/react/dist/${file}`, "utf8"));
    assert.match(patched, /if \(inputRef.current\?\.isConnected\)/);
    assert.match(patched, /height: context.calculatebarHeight/);
    assert.match(patched, /runMenuAction: function \(action, payload\)/);
    assert.match(patched, /context.devicePixelRatio, context.calculatebarHeight/);
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

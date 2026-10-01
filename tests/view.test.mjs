import { test } from "node:test";
import assert from "node:assert/strict";
import { build } from "esbuild";
import { createRequire } from "node:module";
import vm from "node:vm";

const realRequire = createRequire(import.meta.url);
let renders = [], roots = [], observers = [];
class Element {
  dataset = {};
  ownerDocument = { defaultView: { ResizeObserver: class {
    disconnected = false;
    constructor() { observers.push(this); }
    observe() {}
    disconnect() { this.disconnected = true; }
  }, Event: class {}, dispatchEvent() {} }, activeElement: null };
  empty() {} addClass() {} createEl() {} createDiv() { return new Element(); } addEventListener() {} contains() { return false; } querySelector() { return null; }
}
class TextFileView {
  app = { scope: {} };
  contentEl = new Element();
  file = { path: "test.sheet" };
  saveRequests = 0;
  saved = [];
  requestSave() { this.saveRequests++; }
  async save() { this.saved.push(this.getViewData()); }
  async onUnloadFile() { await this.save(); this.clear(); }
}
const result = await build({ entryPoints: ["view.tsx"], bundle: true, write: false, format: "cjs", platform: "node", packages: "external" });
const module = { exports: {} };
vm.runInNewContext(result.outputFiles[0].text, { module, exports: module.exports, HTMLElement: Element, queueMicrotask,
  require: name => {
    if (name === "obsidian") return { TextFileView, Notice: class {}, Scope: class { register() {} } };
    if (name === "react") return { createRef: () => ({ current: null }), createElement: (_type, props) => props };
    if (name === "react-dom") return { flushSync: fn => fn() };
    if (name === "react-dom/client") return { createRoot: () => {
      const root = { unmounted: false, render: props => renders.push(props), unmount() { this.unmounted = true; } };
      roots.push(root); return root;
    } };
    if (name === "@fortune-sheet/react") return { Workbook: "Workbook" };
    return realRequire(name);
  }, console });
const { SpreadsheetView } = module.exports;
const make = () => new SpreadsheetView({}, () => ({ theme: "auto", currency: "$", folder: "" }));
const source = '[{"name":"Existing","id":"old","celldata":[{"r":0,"c":0,"v":{"v":"preserve"}}]}]';

test("opening and closing untouched files preserves the original bytes", async () => {
  const view = make();
  view.setViewData(source, true);
  renders.at(-1).onChange([{ name: "Normalized", data: [] }]);
  assert.equal(view.getViewData(), source);
  await view.onClose();
  assert.equal(view.saved[0], source);
});
test("file switch clears old output and ignores callbacks from the old workbook", () => {
  const view = make();
  view.setViewData(source, true);
  const old = renders.at(-1);
  old.onChange([{ name: "Old", data: [[{ v: "changed" }]] }]);
  old.onOp([{ op: "replace", path: [0] }]);
  view.setViewData('[{"name":"New","id":"new"}]', true);
  old.onChange([{ name: "Old", data: [[{ v: "wrong file" }]] }]);
  old.onOp([{ op: "replace", path: [0] }]);
  assert.equal(view.getViewData(), '[{"name":"New","id":"new"}]');
});
test("closing immediately after editing flushes the changed snapshot", async () => {
  const view = make();
  view.setViewData(source, true);
  const props = renders.at(-1);
  props.onChange([{ name: "Existing", id: "old", data: [[{ v: "new value" }]] }]);
  props.onOp([{ op: "replace", path: [0] }]);
  assert.ok(view.saveRequests > 0);
  await view.onUnloadFile(view.file);
  assert.equal(JSON.parse(view.saved[0])[0].celldata[0].v.v, "new value");
  assert.ok(roots.at(-1).unmounted);
  assert.ok(observers.at(-1).disconnected);
});
test("invalid JSON survives close unchanged and does not mount an editable workbook", async () => {
  const view = make(), count = renders.length;
  view.setViewData("damaged but valuable bytes", true);
  assert.equal(renders.length, count);
  await view.onClose();
  assert.equal(view.saved[0], "damaged but valuable bytes");
});
test("repeated loads disconnect every observer and unmount every obsolete root", () => {
  const view = make();
  view.setViewData(source, true);
  const firstRoot = roots.at(-1), firstObserver = observers.at(-1);
  view.setViewData(source, false);
  assert.ok(firstRoot.unmounted);
  assert.ok(firstObserver.disconnected);
  view.clear();
  assert.ok(roots.at(-1).unmounted);
  assert.ok(observers.at(-1).disconnected);
});

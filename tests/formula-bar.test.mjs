import { test } from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import { build } from "esbuild";

const built = await build({ entryPoints: ["formula-bar.tsx"], bundle: true, write: false, format: "cjs", platform: "node", packages: "external" });

function mount(height = 29) {
  const frames = new Map(), effects = [], cleanups = [], calls = [];
  let nextFrame = 0;
  const ownerWindow = {
    requestAnimationFrame(callback) { frames.set(++nextFrame, callback); return nextFrame; },
    cancelAnimationFrame(id) { frames.delete(id); },
    ResizeObserver: class { observe() {} disconnect() {} },
  };
  const book = { ownerDocument: { defaultView: ownerWindow } };
  const react = {
    useState: value => [value, () => {}],
    useRef: value => ({ current: value }),
    useLayoutEffect: callback => effects.push(callback),
    createElement: (type, props, ...children) => ({ type, props, children }),
  };
  const module = { exports: {} };
  vm.runInNewContext(built.outputFiles[0].text, { module, exports: module.exports,
    require: name => name === "react-dom" ? { flushSync: callback => callback() } : react });
  const tree = module.exports.FormulaBarControls({ height, workbookRef: { current: book }, onResize: value => calls.push(value) });
  for (const effect of effects) { const cleanup = effect(); if (cleanup) cleanups.push(cleanup); }
  const handle = tree.children[1].props;
  const target = { ownerDocument: book.ownerDocument, setPointerCapture() {}, hasPointerCapture: () => false };
  const event = (clientY, pointerId = 7) => ({ clientY, pointerId, button: 0, currentTarget: target, preventDefault() {}, stopPropagation() {} });
  return {
    handle, calls, frames, event,
    frame() { const callbacks = [...frames.values()]; frames.clear(); callbacks.forEach(callback => callback()); },
    unmount() { cleanups.forEach(cleanup => cleanup()); },
  };
}

test("dragging applies only the latest pointer position once per animation frame", () => {
  const control = mount();
  control.handle.onPointerDown(control.event(100));
  for (let y = 101; y <= 220; y++) control.handle.onPointerMove(control.event(y));
  assert.equal(control.frames.size, 1);
  assert.deepEqual(control.calls, []);
  control.frame();
  assert.deepEqual(control.calls, [149]);
  control.handle.onPointerMove(control.event(120));
  control.frame();
  assert.deepEqual(control.calls, [149, 49]);
});

test("ending or losing a drag flushes its last pending position and cancels the frame", () => {
  for (const finish of ["onPointerUp", "onPointerCancel", "onLostPointerCapture"]) {
    const control = mount();
    control.handle.onPointerDown(control.event(100));
    control.handle.onPointerMove(control.event(220));
    control.handle[finish](control.event(220));
    assert.deepEqual(control.calls, [149], finish);
    assert.equal(control.frames.size, 0, finish);
    control.handle.onPointerMove(control.event(250));
    control.frame();
    assert.deepEqual(control.calls, [149], finish);
  }
});

test("other pointers cannot move or finish the active drag", () => {
  const control = mount();
  control.handle.onPointerDown(control.event(100));
  control.handle.onPointerMove(control.event(220, 8));
  control.handle.onPointerUp(control.event(220, 8));
  assert.equal(control.frames.size, 0);
  control.handle.onPointerMove(control.event(120));
  control.frame();
  assert.deepEqual(control.calls, [49]);
});

test("dragging against either height limit does not publish unchanged heights", () => {
  for (const [height, y] of [[29, 0], [420, 500]]) {
    const control = mount(height);
    control.handle.onPointerDown(control.event(100));
    control.handle.onPointerMove(control.event(y));
    control.frame();
    assert.deepEqual(control.calls, []);
  }
});

test("closing the workbook cancels a pending resize", () => {
  const control = mount();
  control.handle.onPointerDown(control.event(100));
  control.handle.onPointerMove(control.event(220));
  control.unmount();
  assert.equal(control.frames.size, 0);
  control.frame();
  assert.deepEqual(control.calls, []);
});

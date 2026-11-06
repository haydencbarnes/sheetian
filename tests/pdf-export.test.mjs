import { test } from "node:test";
import assert from "node:assert/strict";
import { build } from "esbuild";
import { createRequire } from "node:module";
import vm from "node:vm";

const realRequire = createRequire(import.meta.url);
const result = await build({ entryPoints: ["pdf-export.ts"], bundle: true, write: false, platform: "node", format: "cjs", packages: "external" });
function load(remote, writeFile = async () => {}, temporaryFs = {}) {
  const module = { exports: {} };
  const temporaryFiles = new Map(), cleanedUp = [];
  const fs = {
    async mkdtemp(prefix) { return prefix + "test"; },
    async writeFile(path, contents, options) {
      if (options?.encoding === "utf8") {
        assert.equal(options.mode, 0o600);
        temporaryFiles.set(path, contents); return;
      }
      return writeFile(path, contents);
    },
    async rm(directory) {
      cleanedUp.push(directory);
      for (const path of temporaryFiles.keys()) if (path.startsWith(directory + "/")) temporaryFiles.delete(path);
    },
    ...temporaryFs,
  };
  vm.runInNewContext(result.outputFiles[0].text, { module, exports: module.exports,
    require: name => name === "electron" ? { remote } : name === "fs" ? { promises: fs } : realRequire(name),
  });
  return { ...module.exports, temporaryFiles, cleanedUp };
}
const { sheetPrintDocument } = load();

test("PDF includes sparse offscreen cells, formatted results, zero, false and Unicode without modifying the sheet", () => {
  const sheet = { name: "Résumé & budget", row: 200, column: 26, zoomRatio: 0.5, celldata: [
    { r: 0, c: 0, v: { v: 0 } }, { r: 0, c: 1, v: { v: false } },
    { r: 1, c: 0, v: { f: "=SUM(A1:B1)", v: 42, m: "$42.00", bl: 1, bg: "#123456", fc: "#ffffff" } },
    { r: 60, c: 3, v: { v: "Last visible value: 日本語" } },
  ] };
  const before = JSON.stringify(sheet), { html } = sheetPrintDocument(sheet);
  for (const text of [">0<", ">false<", "$42.00", "日本語", "Résumé &amp; budget", "#123456", "bold"]) assert.ok(html.includes(text), text);
  assert.ok(!html.includes("=SUM"));
  assert.equal((html.match(/<tr /g) ?? []).length, 61);
  assert.equal((html.match(/<col style/g) ?? []).length, 4);
  assert.equal(JSON.stringify(sheet), before);
});

test("PDF omits hidden rows/columns and contracts merged spans around them", () => {
  const sheet = { name: "Merge", data: [
    [{ v: "Merged title", mc: { r: 0, c: 0, rs: 2, cs: 3 } }, null, null, { v: "Hidden column" }],
    [null, { v: "Hidden row" }],
    [null, null, { v: "Visible" }],
  ], config: { rowhidden: { 1: 0 }, colhidden: { 1: 0, 3: 0 }, merge: { "0_0": { r: 0, c: 0, rs: 2, cs: 3 } } } };
  const { html } = sheetPrintDocument(sheet);
  assert.match(html, /rowspan="1" colspan="2"/);
  assert.ok(html.includes("Merged title") && html.includes("Visible"));
  assert.ok(!html.includes("Hidden column") && !html.includes("Hidden row"));
  assert.equal((html.match(/<td/g) ?? []).length, 3);
});

test("PDF retains rich text, line breaks, row/column sizes, borders and gridline preference", () => {
  const { html } = sheetPrintDocument({ name: "Styles", showGridLines: 0, data: [[
    { ct: { t: "inlineStr", s: [{ v: "One\n", bl: 1 }, { v: "Two", it: 1, fc: "red" }] } },
  ]], config: { rowlen: { 0: 40 }, columnlen: { 0: 150 }, borderInfo: [
    { rangeType: "range", range: [{ row: [0, 0], column: [0, 0] }], borderType: "border-all", style: 7, color: "#112233" },
  ] } });
  for (const text of ["One\n", "Two", "bold", "italic", "color:red", "height:41px", "width:151px", "3px double #112233", "solid transparent"]) assert.ok(html.includes(text), text);
});

test("PDF escapes workbook content and blocks scripts, external images and CSS resource injection", () => {
  const { html } = sheetPrintDocument({ name: '<script>alert("title")</script>', data: [[
    { v: '<img src="https://example.com">', bg: "red;background:url(https://example.com)", ct: { t: "inlineStr", s: [{ v: "<script>bad()</script>" }] } },
  ]], images: [{ src: "https://example.com/private.png", left: 0, top: 0, width: 10, height: 10 }] });
  assert.ok(!html.includes("<script>") && !html.includes("<img"));
  assert.ok(!html.includes("url(https:"));
  assert.ok(html.includes("&lt;script&gt;") && html.includes("default-src 'none'"));
});

test("PDF fits wide sheets independently of zoom and handles empty/oversized sheets", () => {
  const wide = { name: "Wide", celldata: [{ r: 0, c: 14, v: { v: "Last column" } }] };
  const first = sheetPrintDocument({ ...wide, zoomRatio: 0.1 });
  assert.equal(first.landscape, true);
  assert.equal(first.html, sheetPrintDocument({ ...wide, zoomRatio: 4 }).html);
  const empty = sheetPrintDocument({ name: "Empty" }).html;
  assert.match(empty, /no visible cells/);
  assert.match(empty, /width: 718px/);
  assert.match(sheetPrintDocument({ name: "Hidden", data: [[{ v: "hidden" }]], config: { colhidden: { 0: 0 } } }).html, /width: 718px/);
  assert.throws(() => sheetPrintDocument({ name: "Huge", celldata: [{ r: 1_000_000, c: 1, v: { v: "far" } }] }), /too large/);
});

test("PDF includes embedded raster images beyond the cell range", () => {
  const src = "data:image/png;base64,test";
  const { html, landscape } = sheetPrintDocument({ name: "Image", images: [{ src, left: 800, top: 1200, width: 100, height: 40 }] });
  assert.ok(html.includes(src));
  assert.ok(html.includes("min-height: 1240px") && html.includes("width: 900px"));
  assert.equal(landscape, true);
});

test("canceling PDF export creates neither a print window nor a file", async () => {
  const { downloadSheetPdf } = load({ getCurrentWindow() {}, dialog: { async showSaveDialog() { return { canceled: true }; } },
    BrowserWindow: class { constructor() { assert.fail("Opened renderer after cancellation"); } },
  }, async () => assert.fail("Wrote a canceled export"));
  assert.equal(await downloadSheetPdf({ name: "Cancel" }, "cancel.pdf"), null);
});

test("PDF waits for loading and fonts/images, writes binary output, and closes its isolated renderer", async () => {
  const calls = [], bytes = Buffer.from("%PDF-1.7"), sheet = { name: "Original" };
  const remote = { getCurrentWindow() { return "parent"; }, dialog: { async showSaveDialog(parent, options) {
    assert.equal(parent, "parent"); assert.equal(options.defaultPath, "book.pdf");
    sheet.name = "Changed while choosing a destination";
    return { canceled: false, filePath: "/tmp/book.pdf" };
  } }, BrowserWindow: class {
    constructor(options) {
      assert.equal(options.show, false); assert.equal(options.webPreferences.nodeIntegration, false); assert.equal(options.webPreferences.sandbox, true);
    }
    async loadFile(path) { calls.push("load"); assert.ok(temporaryFiles.get(path).includes("<h1>Original</h1>")); }
    webContents = {
      async executeJavaScript(code) { calls.push("ready"); assert.ok(code.includes("document.fonts.ready") && code.includes("image.decode()")); },
      async printToPDF(options) { calls.push("print"); assert.equal(options.printBackground, true); return bytes; },
    };
    isDestroyed() { return false; }
    destroy() { calls.push("destroy"); }
  } };
  const { downloadSheetPdf, temporaryFiles, cleanedUp } = load(remote, async (path, output) => { calls.push("write"); assert.equal(path, "/tmp/book.pdf"); assert.equal(output, bytes); });
  assert.equal(await downloadSheetPdf(sheet, "book.pdf"), "/tmp/book.pdf");
  assert.deepEqual(calls, ["load", "ready", "print", "write", "destroy"]);
  assert.equal(temporaryFiles.size, 0);
  assert.equal(cleanedUp.length, 1);
});

test("large PDF exports load a temporary file instead of exceeding Chromium's data URL limit", async () => {
  const sheet = { name: "Large", data: [[{ v: "x".repeat(2_200_000) + "final value" }]] };
  let loaded = false;
  const { downloadSheetPdf, temporaryFiles, cleanedUp } = load({ getCurrentWindow() {}, dialog: {
    async showSaveDialog() { return { filePath: "/tmp/large.pdf" }; },
  }, BrowserWindow: class {
    async loadURL() { assert.fail("Large workbooks must not load via a data URL"); }
    async loadFile(path) {
      const html = temporaryFiles.get(path);
      assert.ok(html.length > 2_097_152 && html.includes("final value")); loaded = true;
    }
    webContents = { async executeJavaScript() {}, async printToPDF() { return Buffer.from("%PDF-1.7"); } };
    isDestroyed() { return false; } destroy() {}
  } });
  await downloadSheetPdf(sheet, "large.pdf");
  assert.ok(loaded);
  assert.equal(temporaryFiles.size, 0);
  assert.equal(cleanedUp.length, 1);
});

test("PDF renderer and temporary content are cleaned up when preparation, loading, printing or writing fails", async () => {
  for (const failure of ["directory", "html", "load", "print", "write"]) {
    let destroyed = false;
    const fail = stage => { if (stage === failure) throw new Error(stage + " failed"); };
    const { downloadSheetPdf, temporaryFiles, cleanedUp } = load({ getCurrentWindow() {}, dialog: { async showSaveDialog() { return { filePath: "/tmp/failure.pdf" }; } },
      BrowserWindow: class {
        async loadFile() { fail("load"); }
        webContents = { async executeJavaScript() {}, async printToPDF() { fail("print"); return Buffer.from("pdf"); } };
        isDestroyed() { return false; } destroy() { destroyed = true; }
      },
    }, async () => fail("write"), failure === "directory" ? { async mkdtemp() { fail("directory"); } } : failure === "html" ? { async writeFile() { fail("html"); } } : {});
    await assert.rejects(downloadSheetPdf({ name: "Failure" }, "failure.pdf"), new RegExp(failure + " failed"));
    assert.equal(destroyed, true);
    assert.equal(temporaryFiles.size, 0);
    assert.equal(cleanedUp.length, failure === "directory" ? 0 : 1);
  }
});

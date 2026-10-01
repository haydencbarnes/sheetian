import { test } from "node:test";
import assert from "node:assert/strict";
import { build } from "esbuild";
import { createRequire } from "node:module";
import vm from "node:vm";
import postcss from "postcss";
import { readFile } from "node:fs/promises";
import { buildCss } from "../build-css.mjs";

const require = createRequire(import.meta.url);
async function load(file) {
  const result = await build({ entryPoints: [file], bundle: true, write: false, platform: "node", format: "cjs", packages: "external" });
  const module = { exports: {} };
  vm.runInNewContext(result.outputFiles[0].text, { module, exports: module.exports, require, Date, console });
  return module.exports;
}
const { parseSheets, serializeSheets, sheetToCsv } = await load("sheet-data.ts");
const { normalizeInput } = await load("input.ts");
const plain = value => JSON.parse(JSON.stringify(value));

test("saving preserves every sheet's metadata and formula dependency chain", () => {
  const metadata = { name: "Second", id: "b", order: 2, status: 1, hide: 0, row: 500, column: 50,
    images: [{ id: "image", src: "data:image/png;base64,test", width: 10, height: 10 }],
    filter: { 0: { rowhidden: { 2: 0 } } }, filter_select: { row: [0, 3], column: [0, 1] },
    frozen: { type: "both", range: { row_focus: 1, column_focus: 1 } },
    hyperlink: { "0_0": { linkType: "external", linkAddress: "https://example.com" } },
    dataVerification: { "0_1": { type: "number" } },
    luckysheet_conditionformat_save: [{ type: "colorScale" }],
    calcChain: [{ r: 100, c: 2, id: "b", custom: "keep" }], customFutureField: { value: 1 } };
  const output = JSON.parse(serializeSheets([{ name: "First", id: "a", celldata: [{ r: 0, c: 0, v: { v: 8 } }] },
    { ...metadata, data: [[{ v: 0 }, null, { v: false }, { v: "" }, { v: 16, f: "=First!A1*2" }]], celldata: [] }]));
  assert.deepEqual(output[1], { ...metadata, celldata: [
    { r: 0, c: 0, v: { v: 0 } }, { r: 0, c: 2, v: { v: false } },
    { r: 0, c: 3, v: { v: "" } }, { r: 0, c: 4, v: { v: 16, f: "=First!A1*2" } },
  ] });
  assert.equal(output[0].celldata[0].v.v, 8);
});
test("dense legacy files and sparse files reopen without losing cells", () => {
  const dense = parseSheets('[{"name":"Legacy","data":[[{"v":3},null]]}]');
  assert.equal(dense[0].celldata[0].v.v, 3);
  assert.equal(dense[0].data, undefined);
  const sparse = [{ name: "Sparse", id: "original", celldata: [{ r: 2, c: 1, v: { v: "ok" } }] }];
  assert.deepEqual(plain(parseSheets(serializeSheets(sparse)))[0].celldata, sparse[0].celldata);
  assert.equal(parseSheets(serializeSheets(sparse))[0].id, "original");
});
test("missing sheet IDs are stable and do not collide with existing IDs", () => {
  const sheets = parseSheets('[{"name":"New"},{"name":"Old","id":"sheet-1"}]');
  assert.notEqual(sheets[0].id, sheets[1].id);
  assert.equal(sheets[1].id, "sheet-1");
});
test("invalid file shapes are rejected instead of silently becoming empty sheets", () => {
  for (const input of ["oops", "{}", "[]", "[null]", '[{"name":"X","data":[null]}]',
    '[{"name":"X","celldata":[{"r":-1,"c":0,"v":null}]}]']) {
    assert.throws(() => parseSheets(input));
  }
});
test("currency entry produces numbers for formulas while explicit text is preserved", () => {
  assert.equal(normalizeInput("$300", null, "$").v, 300);
  assert.equal(normalizeInput("€1,234.50", null, "€").v, 1234.5);
  assert.equal(normalizeInput("£-12.25", null, "£").v, -12.25);
  assert.equal(normalizeInput("$300", { ct: { fa: "@" } }, "$"), undefined);
  for (const input of ["$1,2", "$300abc", "=A1*2", "300", "hello"]) {
    assert.equal(normalizeInput(input, null, "$"), undefined);
  }
});
test("date-formatted cells accept calendar dates with the Excel leap-day convention", () => {
  const date = { ct: { fa: "yyyy-MM-dd", t: "d" } };
  assert.equal(normalizeInput("2024-03-28", date, "$").v, 45379);
  assert.equal(normalizeInput("3/28/2024", date, "$").m, "2024-03-28");
  assert.equal(normalizeInput("1900-01-01", date, "$").v, 1);
  assert.equal(normalizeInput("1900-03-01", date, "$").v, 61);
  assert.equal(normalizeInput("2024-02-30", date, "$"), undefined);
});
test("CSV quotes commas, newlines and quotes and preserves sparse positions", () => {
  assert.equal(sheetToCsv({ name: "X", celldata: [
    { r: 0, c: 0, v: { v: 'a,"b"' } }, { r: 0, c: 1, v: { v: "a\nb" } },
    { r: 1, c: 2, v: { v: 0 } }, { r: 2, c: 0, v: { ct: { s: [{ v: "rich" }, { v: " text" }] } } },
  ] }), '"a,""b""","a\nb",\r\n,,0\r\nrich text,,\r\n');
});
test("all built CSS selectors are scoped to the spreadsheet view", async () => {
  await buildCss();
  const css = postcss.parse(await readFile("styles.css", "utf8"));
  css.walkRules(rule => {
    if (rule.parent.type === "atrule" && /keyframes$/i.test(rule.parent.name)) return;
    for (const selector of rule.selectors) {
      assert.match(selector, /\.obsidian-spreadsheet(?:-view)?\b/);
      assert.doesNotMatch(selector, /^html\b|^body\b|^:root\b/);
    }
  });
});

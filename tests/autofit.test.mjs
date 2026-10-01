import { test } from "node:test";
import assert from "node:assert/strict";
import { build } from "esbuild";

const result = await build({ entryPoints: ["autofit.ts"], bundle: true, write: false, format: "esm" });
const { fitColumnWidth, autoFitColumns, cellFont } = await import(`data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString("base64")}`);
const measure = (text, style) => text.length * (style.fs ?? 10) * (style.bl ? 1.2 : 1);

test("autofit measures displayed formula/number values, fonts, and the longest explicit line", () => {
  const sheet = { name: "Test", data: [
    [{ v: 1234, m: "$1,234.00" }],
    [{ f: "=SUM(B1:B20)", v: 25, m: "25" }],
    [{ v: "short\nlonger line", fs: 20, bl: 1 }],
  ] };
  assert.equal(fitColumnWidth(sheet, 0, measure), 276);
  assert.match(cellFont({ ff: 1, fs: 20, bl: 1, it: 1 }), /italic bold 20pt "Arial"/);
});
test("rich text adds runs on one line, then resets at line breaks", () => {
  const sheet = { name: "Rich", data: [[{ ct: { t: "inlineStr", s: [
    { v: "AB", fs: 20 }, { v: "C\nD", fs: 10 }, { v: "EF", fs: 30 },
  ] } }]] };
  assert.equal(fitColumnWidth(sheet, 0, measure), 82);
});
test("empty columns use the default width, merged spans do not inflate a single column", () => {
  const sheet = { name: "Merged", defaultColWidth: 90, celldata: [
    { r: 0, c: 0, v: { v: "A giant merged title should not resize column A", mc: { r: 0, c: 0, rs: 1, cs: 3 } } },
    { r: 1, c: 0, v: { v: "Body" } },
    { r: 2, c: 0, v: { mc: { r: 0, c: 0 } } },
  ] };
  assert.equal(fitColumnWidth(sheet, 0, measure), 52);
  assert.equal(fitColumnWidth(sheet, 1, measure), 90);
});
test("a column-header selection autofits all selected visible columns; cell selections do not", () => {
  const sheet = { name: "Selected", config: { colhidden: { 1: 0 } }, luckysheet_select_save: [
    { row: [0, 99], column: [0, 2], column_select: true },
    { row: [0, 99], column: [4, 4], column_select: true },
  ] };
  assert.deepEqual(autoFitColumns(sheet, 2), [0, 2, 4]);
  assert.deepEqual(autoFitColumns(sheet, 3), [3]);
  assert.deepEqual(autoFitColumns({ ...sheet, row: 50 }, 2, [{ row: [0,49], column: [0,2] }]), [0,2]);
  assert.deepEqual(autoFitColumns({ name: "Cells", luckysheet_select_save: [{ row: [0, 5], column: [0, 2] }] }, 2), [2]);
});

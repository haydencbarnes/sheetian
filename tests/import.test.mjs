import { test } from "node:test";
import assert from "node:assert/strict";
import { build } from "esbuild";
import { createRequire } from "node:module";
import vm from "node:vm";
import ExcelJS from "exceljs";
import { importFixture } from "./import-fixture.mjs";

const require = createRequire(import.meta.url);
const result = await build({entryPoints:["xlsx-import.ts"],bundle:true,write:false,platform:"node",format:"cjs",packages:"external"});
const module={exports:{}};
vm.runInNewContext(result.outputFiles[0].text,{module,exports:module.exports,require,Date,Uint8Array,console});
const {importXlsx}=module.exports;
const bytes=await importFixture();
const imported=await importXlsx(bytes);
const cell=(sheet,r,c)=>sheet.celldata.find(entry=>entry.r===r&&entry.c===c)?.v;

test("XLSX import retains formulas, translated shared formulas, cross-sheet references, and dependency chains",()=>{
  assert.deepEqual(Array.from(imported.sheets,s=>s.name),["Inputs","Summary","Hidden","Year's Data"]);
  const summary=imported.sheets[1];
  assert.equal(cell(summary,0,0).f,"=SUM(Inputs!A1:A2)");
  assert.equal(cell(summary,1,0).f,"='Inputs'!$A$1*2");
  assert.equal(cell(summary,2,1).f,"=Inputs!A2*2");
  assert.equal(cell(summary,6,0).f,"='Year''s Data'!A1+Inputs!A1");
  assert.equal(summary.calcChain.length,10);
  assert(summary.calcChain.every(entry=>entry.id===summary.id));
  assert.equal(imported.formulaCount,10);
  assert.equal(imported.missingResults.length,1);
  assert.equal(imported.sheets[2].hide,1);
});
test("cached zero/false values survive import; numeric dates and currency formatting remain editable",()=>{
  const inputs=imported.sheets[0],summary=imported.sheets[1];
  assert.equal(cell(summary,3,0).v,false);
  assert.equal(cell(summary,4,0).v,0);
  assert.equal(cell(inputs,0,1).v,45379);
  assert.equal(cell(inputs,0,1).ct.t,"d");
  assert.equal(cell(inputs,2,0).v,1234.5);
  assert.equal(cell(inputs,2,0).ct.fa,'$#,##0.00');
});
test("basic styles, rich text, merges, hyperlinks, sizes, hidden rows/columns, and freeze panes survive",()=>{
  const sheet=imported.sheets[0],title=cell(sheet,0,3);
  assert.equal(title.bl,1);assert.equal(title.fs,16);assert.equal(title.ff,"Arial");
  assert.equal(title.fc,"#FF0000");assert.equal(title.bg,"#DDFFDD");assert.equal(title.tb,"2");
  assert.equal(title.mc.cs,2);assert.equal(cell(sheet,0,4).mc.c,3);
  assert.equal(sheet.config.merge['0_3'].cs,2);
  assert.equal(sheet.hyperlink['0_2'].linkAddress,'https://example.com');
  assert.equal(sheet.config.rowhidden[4],0);assert.equal(sheet.config.colhidden[5],0);
  assert.equal(sheet.config.rowlen[0],39);assert.equal(sheet.config.columnlen[0],144);
  assert.equal(sheet.frozen.type,"both");
  assert.equal(cell(imported.sheets[1],1,2).ct.s[0].v,"Bold");
});
test("1904 dates are normalized to the engine's 1900 date system",async()=>{
  const book=new ExcelJS.Workbook();book.properties.date1904=true;
  const sheet=book.addWorksheet('Dates');sheet.getCell('A1').value=new Date('2024-03-28T00:00:00Z');sheet.getCell('A1').numFmt='yyyy-mm-dd';
  const parsed=await importXlsx(await book.xlsx.writeBuffer());assert.equal(cell(parsed.sheets[0],0,0).v,45379);
});
test("unsupported references/functions are reported while formula text and cached results are retained",async()=>{
  const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Unsupported');
  sheet.getCell('A1').value=5;book.definedNames.add('Unsupported!$A$1','Amount');
  sheet.getCell('B1').value={formula:'QUERY(A1:A2,"select *")',result:'cached'};
  sheet.getCell('B2').value={formula:'_xlfn.FILTER(A1:A2,A1:A2>0)',result:5,shareType:'array',ref:'B2:B3'};
  const parsed=await importXlsx(await book.xlsx.writeBuffer());
  assert.equal(cell(parsed.sheets[0],0,1).f,'=QUERY(A1:A2,"select *")');assert.equal(cell(parsed.sheets[0],0,1).v,'cached');
  assert(parsed.warnings.some(v=>v.includes('Google-only')));assert(parsed.warnings.some(v=>v.includes('Named ranges')));
  assert(parsed.warnings.some(v=>v.includes('Array/spill')));
});
test("invalid files and workbooks too large for the grid are rejected",async()=>{
  await assert.rejects(importXlsx(new Uint8Array([1,2,3])),/Could not read/);
  const book=new ExcelJS.Workbook(),sheet=book.addWorksheet('Oversize');sheet.getCell('ZZ10000').value=1;
  await assert.rejects(importXlsx(await book.xlsx.writeBuffer()),/1,000,000/);
});

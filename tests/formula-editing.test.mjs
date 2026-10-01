import { test } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { patchCore, patchFormulaEditing } from '../engine-compat.mjs';
const require=createRequire(import.meta.url);
const built=await build({entryPoints:['node_modules/@fortune-sheet/core/dist/index.js'],bundle:true,write:false,platform:'node',format:'cjs',packages:'external',plugins:[{name:'formula-editing',setup(builder){builder.onLoad({filter:/core[\\/]dist[\\/]index\.js$/},async args=>({contents:patchCore(await readFile(args.path,'utf8')),loader:'js'}));}}]});
const module={exports:{}};new Function('module','exports','require',built.outputFiles[0].text)(module,module.exports,require);
const {defaultContext,updateCell,groupValuesRefresh}=module.exports;
const formula='=SUM(\nA2:A3\n)';
function context(f=formula){const ctx=defaultContext({});ctx.currentSheetId='formulas';ctx.luckysheetfile=[{id:'formulas',name:'Formulas',status:1,data:[[{f,v:6,m:'$6.00',ct:{fa:'$0.00',t:'n'}}],[{v:2}],[{v:4}]],calcChain:[{r:0,c:0,id:'formulas'}],config:{}}];ctx.luckysheet_select_save=[{row:[0,0],column:[0,0],row_focus:0,column_focus:0}];ctx.luckysheetCellUpdate=[0,0];return ctx;}
const input=text=>({innerText:text,innerHTML:text,querySelectorAll:()=>[]});

test('committing unchanged multiline formulas preserves source, cached result, format and chain',()=>{
  for(const source of [formula,formula.replace(/\n/g,'\r\n'),'=UNSUPPORTED(\r\nA2:A3\r\n)']){
    const ctx=context(source),before=structuredClone(ctx.luckysheetfile[0]);updateCell(ctx,0,0,input(source.replace(/\r\n/g,'\n')));
    assert.deepEqual(ctx.luckysheetfile[0],before);assert.deepEqual(ctx.luckysheetCellUpdate,[]);
  }
});
test('editing a multiline formula recalculates it without converting it to rich text',()=>{
  const ctx=context();const edited='=SUM(\nA2:A3\n)+1';updateCell(ctx,0,0,input(edited));
  const cell=ctx.luckysheetfile[0].data[0][0];assert.equal(cell.f,edited);assert.equal(cell.v,7);assert.notEqual(cell.ct.t,'inlineStr');assert.equal(cell.ct.fa,'$0.00');
});
test('ordinary multiline text can intentionally replace a formula and retains its line breaks',()=>{
  const ctx=context();updateCell(ctx,0,0,input('First line\nSecond line'));
  const cell=ctx.luckysheetfile[0].data[0][0];assert.equal(cell.f,undefined);assert.equal(cell.ct.t,'inlineStr');assert.equal(cell.ct.s.map(s=>s.v).join(''),'First line\r\nSecond line');
});
test('input changes recalculate multiline formulas with whitespace around references',()=>{
  const source='=SUM(\n A2:A3 \n)',ctx=context(source);updateCell(ctx,1,0,input('8'));groupValuesRefresh(ctx);
  const cell=ctx.luckysheetfile[0].data[0][0];assert.equal(cell.f,source);assert.equal(cell.v,12);
});
test('multiline formula compatibility patch guards both pinned core distributions',async()=>{
  for(const name of ['index.js','index.esm.js'])assert.match(patchFormulaEditing(await readFile(`node_modules/@fortune-sheet/core/dist/${name}`,'utf8')), /inputText.length > 0 && !isFormula\(inputText\)/);
  assert.throws(()=>patchFormulaEditing('future engine'));
});

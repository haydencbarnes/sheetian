import { test } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { patchCore, patchFormulaEditing, patchTextFormula } from '../engine-compat.mjs';
const require=createRequire(import.meta.url);
const built=await build({entryPoints:['node_modules/@fortune-sheet/core/dist/index.js'],bundle:true,write:false,platform:'node',format:'cjs',packages:'external',plugins:[{name:'formula-editing',setup(builder){builder.onLoad({filter:/core[\\/]dist[\\/]index\.js$/},async args=>({contents:patchCore(await readFile(args.path,'utf8')),loader:'js'}));}}]});
const module={exports:{}};new Function('module','exports','require',built.outputFiles[0].text)(module,module.exports,require);
const {defaultContext,updateCell,groupValuesRefresh,execfunction}=module.exports;
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

const goalFormula='=IF(D10-C10>0,\n  "✅ Exceeded goal by $" & TEXT(D10-C10,"#,##0.00"),\n  IF(D10-C10=0,\n    "✅ Met goal exactly!",\n    "❌ Under goal by $" & TEXT(ABS(D10-C10),"#,##0.00")\n  )\n)\n';
function goalContext(actual=6234.56){
  const ctx=context(goalFormula),sheet=ctx.luckysheetfile[0];
  sheet.data=Array.from({length:11},()=>Array(5).fill(null));
  sheet.data[9][2]={v:5000};sheet.data[9][3]={v:actual};
  sheet.data[9][4]={f:goalFormula,v:'#ERROR!',ct:{fa:'General',t:'s'}};
  sheet.calcChain=[{r:9,c:4,id:sheet.id}];
  return ctx;
}
test('the imported multiline IF/ABS/TEXT formula calculates all three goal outcomes',()=>{
  for(const [actual,expected] of [[6234.56,'✅ Exceeded goal by $1,234.56'],[5000,'✅ Met goal exactly!'],[3765.44,'❌ Under goal by $1,234.56']]){
    const result=execfunction(goalContext(actual),goalFormula,9,4,'formulas');
    assert.equal(result[1],expected);assert.equal(result[2],goalFormula);
  }
});
test('TEXT formats numbers, percentages, dates, literals and numeric strings as text',()=>{
  const parser=context().formulaCache.parser;
  for(const [formula,expected] of [['TEXT(1234.567,"#,##0.00")','1,234.57'],['TEXT(-12.5,"0.00;[Red](0.00)")','(12.50)'],['TEXT(0.125,"0.0%")','12.5%'],['TEXT(46023,"yyyy-mm-dd")','2026-01-01'],['TEXT("1234.5","#,##0.00")','1,234.50'],['TEXT("hello","@")','hello'],['TEXT(1.5,"0.00") & " USD"','1.50 USD'],['TEXT(42,"")','']]){
    assert.deepEqual(parser.parse(formula),{error:null,result:expected},formula);
  }
});
test('TEXT reports invalid arguments and formats instead of swallowing errors',()=>{
  const parser=context().formulaCache.parser;
  for(const formula of ['TEXT(12,0)','TEXT(12,"[broken")'])assert.equal(parser.parse(formula).error,'#VALUE!',formula);
  assert.equal(parser.parse('TEXT(12)').error,'#N/A');
  assert.equal(parser.parse('TEXT(SQRT(-1),"0.00")').error,'#NUM!');
});
test('editing goal inputs recalculates the imported TEXT formula and preserves its source',()=>{
  const ctx=goalContext();updateCell(ctx,9,3,input('3765.44'));groupValuesRefresh(ctx);
  assert.equal(ctx.luckysheetfile[0].data[9][4].v,'❌ Under goal by $1,234.56');
  assert.equal(ctx.luckysheetfile[0].data[9][4].f,goalFormula);
});
test('formula text results retain numeric strings, currency, dates, booleans and empty strings',()=>{
  for(const [f,expected] of [['=TEXT(12.5,"0.00")','12.50'],['=TEXT(12.5,"$0.00")','$12.50'],['=TEXT(46023,"yyyy-mm-dd")','2026-01-01'],['="TRUE"','TRUE'],['=""',''],['="003"','003']]){
    const ctx=context();updateCell(ctx,0,0,input(f));
    const cell=ctx.luckysheetfile[0].data[0][0];assert.equal(cell.v,expected);assert.equal(cell.m,expected);assert.equal(cell.ct.t,'s');assert.equal(cell.f,f);
  }
});
test('recalculating over-closed formulas does not append more parentheses',()=>{
  const ctx=context(),source='=SUM(A2:A3))';
  for(let i=0;i<3;i++){const result=execfunction(ctx,source,0,0,'formulas');assert.equal(result[1],'#ERROR!');assert.equal(result[2],source);}
  assert.equal(execfunction(ctx,'=SUM(A2:A3',0,0,'formulas')[1],6);
});
test('TEXT and parenthesis compatibility patches guard both pinned core distributions',async()=>{
  for(const name of ['index.js','index.esm.js'])assert.match(patchTextFormula(await readFile(`node_modules/@fortune-sheet/core/dist/${name}`,'utf8')), /setFunction\("TEXT"/);
  assert.throws(()=>patchTextFormula('future engine'));
});

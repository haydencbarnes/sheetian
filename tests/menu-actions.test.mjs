import { test } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { patchCore, patchCopyWriter } from '../engine-compat.mjs';
const require=createRequire(import.meta.url);
const built=await build({entryPoints:['menu-actions.ts'],bundle:true,write:false,platform:'node',format:'cjs',packages:'external',plugins:[{name:'menu-core',setup(builder){builder.onResolve({filter:/^@fortune-sheet\/core$/},()=>({path:require.resolve('@fortune-sheet/core')}));builder.onLoad({filter:/core[\\/]dist[\\/]index\.js$/},async args=>({contents:patchCore(await readFile(args.path,'utf8')),loader:'js'}));}}]});
// Use the engine Context constructor for the test fixtures.
const coreBuild=await build({entryPoints:['node_modules/@fortune-sheet/core/dist/index.js'],bundle:true,write:false,platform:'node',format:'cjs',packages:'external'});
const evaluate=text=>{const module={exports:{}};new Function('module','exports','require',text)(module,module.exports,require);return module.exports;};
const {runMenuAction}=evaluate(built.outputFiles[0].text),{defaultContext}=evaluate(coreBuild.outputFiles[0].text);
const formula='=SUM(\nA2:A3\n)';
function context(){const ctx=defaultContext({});ctx.currentSheetId='menus';ctx.luckysheetfile=[{id:'menus',name:'Menus',status:1,data:[[{f:formula,v:6,m:'6',ct:{fa:'General',t:'n'}}],[{v:2,m:'2'}],[{v:4,m:'4'}]],calcChain:[{r:0,c:0,id:'menus'}],config:{}}];ctx.luckysheet_select_save=[{row:[0,0],column:[0,0],row_focus:0,column_focus:0}];ctx.luckysheetCellUpdate=[0,0];return ctx;}
const editor=text=>({innerText:text,innerHTML:text,querySelectorAll:()=>[]});
test('opening a menu commits a changed formula without moving its selection or replacing it with text',()=>{
 const ctx=context(),other=context();runMenuAction(ctx,'commit-formula',null,{editor:editor(formula+'+1')});
 assert.equal(ctx.luckysheetfile[0].data[0][0].f,formula+'+1');assert.equal(ctx.luckysheetfile[0].data[0][0].v,7);assert.deepEqual(ctx.luckysheetCellUpdate,[]);assert.equal(ctx.luckysheet_select_save[0].row[0],0);assert.equal(other.luckysheetfile[0].data[0][0].f,formula);
});
test('formula menu commit targets the editing cell in a disjoint selection',()=>{
 const ctx=context();ctx.luckysheetfile[0].data[1][0]={f:'=2+2',v:4,m:'4'};ctx.luckysheet_select_save.push({row:[1,1],column:[0,0],row_focus:1,column_focus:0});ctx.luckysheetCellUpdate=[1,0];runMenuAction(ctx,'commit-formula',null,{editor:editor('=2+3')});assert.equal(ctx.luckysheetfile[0].data[1][0].v,5);assert.equal(ctx.luckysheetfile[0].data[0][0].f,formula);
});
test('formula Enter commits even after the editing flag clears and navigates down',()=>{
 const ctx=context();ctx.luckysheetCellUpdate=[];ctx.visibledatarow=[20,40,60];ctx.visibledatacolumn=[73];runMenuAction(ctx,'commit-formula',null,{editor:editor(formula+'+1'),moveDown:true});assert.equal(ctx.luckysheetfile[0].data[0][0].v,7);assert.equal(ctx.luckysheet_select_save[0].row[0],1);
});
test('menu opening retains an unchanged formula and cached result exactly',()=>{
 const ctx=context(),before=structuredClone(ctx.luckysheetfile);runMenuAction(ctx,'commit-formula',null,{editor:editor(formula)});assert.deepEqual(ctx.luckysheetfile,before);
});
test('menu Copy and Cut use a supplied clipboard writer and keep formula-aware engine copy state',()=>{
 const ctx=context();let html;
 runMenuAction(ctx,'copy',null,{writeClipboard:value=>html=value});assert.match(html,/fortune-copy-action-table/);assert.match(html,/>6</);assert.equal(ctx.luckysheet_copy_save.dataSheetId,'menus');assert.deepEqual(ctx.luckysheet_copy_save.copyRange,[{row:[0,0],column:[0,0]}]);assert.equal(ctx.luckysheet_paste_iscut,false);
 runMenuAction(ctx,'cut',null,{writeClipboard:value=>html=value});assert.equal(ctx.luckysheet_paste_iscut,true);assert.equal(ctx.luckysheetfile[0].data[0][0].f,formula);
});
test('menu clipboard patch guards both pinned core distributions',async()=>{
 for(const name of ['index.js','index.esm.js'])assert.match(patchCopyWriter(await readFile(`node_modules/@fortune-sheet/core/dist/${name}`,'utf8')),/\(writeClipboard \|\| clipboard.writeHtml\)\(cpdata\)/);
 assert.throws(()=>patchCopyWriter('future engine'));
});

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { createRequire } from 'node:module';
const require=createRequire(import.meta.url);
const built=await build({entryPoints:['row-autofit.ts'],bundle:true,write:false,format:'cjs',platform:'node',packages:'external'});
const module={exports:{}};
new Function('module','exports','require',built.outputFiles[0].text)(module,module.exports,require);
const {fitRowHeight,autoFitRows}=module.exports;
const canvas=()=>({font:'10pt Arial',measureText(text){const size=Number(/(\d+)pt/.exec(this.font)?.[1]??10);return {width:String(text).length*size*.6,actualBoundingBoxAscent:size,actualBoundingBoxDescent:size*.25};}});
const sheet=cell=>({name:'Row',data:[[cell]],config:{columnlen:{0:60}}});

test('row AutoFit respects wrapping at current widths and explicit line breaks',()=>{
  const wrapped=sheet({v:'Words that must wrap into several lines',tb:'2'});
  const narrow=fitRowHeight(wrapped,0,canvas());
  assert(narrow>19);
  assert(fitRowHeight({...wrapped,config:{columnlen:{0:300}}},0,canvas())<narrow);
  assert(fitRowHeight(sheet({v:'one\ntwo\nthree'}),0,canvas())>fitRowHeight(sheet({v:'one'}),0,canvas()));
  assert.equal(fitRowHeight(sheet({v:'A very long single clipped line',tb:'0'}),0,canvas()),fitRowHeight(sheet({v:'short',tb:'0'}),0,canvas()));
});
test('fonts and rich text affect height; displayed formula values are measured',()=>{
  assert(fitRowHeight(sheet({v:'Text',fs:30}),0,canvas())>fitRowHeight(sheet({v:'Text',fs:10}),0,canvas()));
  const rich=sheet({ct:{t:'inlineStr',s:[{v:'Small ',fs:10},{v:'Large',fs:30}]}});
  assert(fitRowHeight(rich,0,canvas())>=46, 'Rich text must accommodate the tallest font and padding');
  assert.equal(fitRowHeight(sheet({v:123456789,m:'1',f:'=SUM(A2:A20)',tb:'2'}),0,canvas()),fitRowHeight(sheet({v:'1',tb:'2'}),0,canvas()));
});
test('horizontal merges use combined visible width; vertical spans and hidden columns are excluded',()=>{
  const cell={v:'A merged title wraps at the combined column width',tb:'2'};
  const base=sheet(cell),narrow=fitRowHeight(base,0,canvas());
  const merged={...base,data:[[{...cell,mc:{r:0,c:0,rs:1,cs:3}}, {mc:{r:0,c:0}}, {mc:{r:0,c:0}}]]};
  assert(fitRowHeight(merged,0,canvas())<narrow);
  assert.equal(fitRowHeight({...base,data:[[{...cell,mc:{r:0,c:0,rs:2,cs:1}}]],defaultRowHeight:24},0,canvas()),24);
  assert.equal(fitRowHeight({...base,config:{colhidden:{0:0}}},0,canvas()),19);
  assert.equal(fitRowHeight({name:'Empty',defaultRowHeight:25,celldata:[]},0,canvas()),25);
});
test('row-header selections fit selected visible rows; ordinary cell selections fit the clicked row',()=>{
  const s={name:'Selection',row:50,column:26,config:{rowhidden:{1:0}}};
  assert.deepEqual(autoFitRows(s,2,[{row:[0,2],column:[0,25]},{row:[4,4],column:[0,25]}]),[0,2,4]);
  assert.deepEqual(autoFitRows(s,3,[{row:[0,2],column:[0,25]}]),[3]);
  assert.deepEqual(autoFitRows(s,2,[{row:[0,2],column:[0,3]}]),[2]);
});

// Run through Obsidian CLI eval. Uses isolated files and restores the prior tab.
(async()=>{
  const delay=ms=>new Promise(r=>setTimeout(r,ms)),assert=(ok,message)=>{if(!ok)throw Error(message);};
  const prior=app.workspace.activeLeaf,root=`Sheetian formula editing checks ${Date.now()}`,results=[],files=[];
  await app.vault.createFolder(root);
  const normalize=f=>f?.replace(/\r|\n/g,'');
  const formula='=SUM(\nInputs!A1:A2\n)',crlf='=SUM(\r\nInputs!A1:A2\r\n)';
  const data=[{name:'Inputs',id:'inputs',order:0,status:1,celldata:[{r:0,c:0,v:{v:2}},{r:1,c:0,v:{v:4}}]},
    {name:'Formulas',id:'formulas',order:1,status:0,calcChain:[{r:0,c:0,id:'formulas'},{r:1,c:0,id:'formulas'}],celldata:[
      {r:0,c:0,v:{f:formula,v:6,m:'$6.00',ct:{fa:'$0.00',t:'n'},fs:12,ff:'Arial',tb:'2'}},
      {r:1,c:0,v:{f:crlf,v:6,m:'6',ct:{fa:'General',t:'n'}}}]}];
  const file=await app.vault.create(`${root}/formulas.sheet`,JSON.stringify(data));files.push(file);
  const open=async()=>{const l=app.workspace.getLeaf(true);await l.openFile(file);app.workspace.setActiveLeaf(l,{focus:true});await delay(400);l.view.workbook.current.activateSheet({id:'formulas'});await delay(120);return l;};
  let leaf=await open(),view=leaf.view;
  const begin=async(row,expectFormula=true)=>{
    view.workbook.current.setSelection([{row:[row,row],column:[0,0]}]);await delay(70);
    const area=view.contentEl.querySelector('.fortune-cell-area'),rect=area.getBoundingClientRect();
    area.dispatchEvent(new MouseEvent('dblclick',{clientX:rect.left+10,clientY:rect.top+10,button:0,bubbles:true}));await delay(80);
    const input=view.contentEl.querySelector('.luckysheet-cell-input');if(expectFormula)assert(input.innerText.startsWith('='),'Editor lost formula source');return input;
  };
  const press=async(input,key='Enter')=>{input.dispatchEvent(new KeyboardEvent('keydown',{key,code:key,keyCode:key==='Escape'?27:13,which:key==='Escape'?27:13,bubbles:true}));await delay(100);};
  let input=await begin(0);assert(input.innerText===formula,'Multiline formula editor text changed');await press(input);
  let cell=view.workbook.current.getSheet().data[0][0];assert(cell.f===formula&&cell.v===6&&cell.ct.t!=='inlineStr','Unchanged formula converted to text');assert(cell.ct.fa==='$0.00','Currency format lost');
  input=await begin(1);await press(input);cell=view.workbook.current.getSheet().data[1][0];assert(cell.f===crlf&&cell.v===6,'CRLF formula converted to text');
  results.push('Double-click and commit retain LF/CRLF multiline formula source, cached values and formatting');
  input=await begin(0);input.innerText=formula+'+1';await press(input);cell=view.workbook.current.getSheet().data[0][0];assert(normalize(cell.f)===normalize(formula+'+1')&&cell.v===7,'Edited multiline formula failed to recalculate');
  input=await begin(0);input.innerText='=999';await press(input,'Escape');assert(normalize(view.workbook.current.getSheet().data[0][0].f)===normalize(formula+'+1'),'Escape changed the formula');
  results.push('Editing recalculates multiline formulas; Escape preserves the previous formula');
  input=await begin(0);await leaf.detach();await delay(100);
  const saved=JSON.parse(await app.vault.read(file)),savedCell=saved[1].celldata.find(c=>c.r===0&&c.c===0).v;
  assert(normalize(savedCell.f)===normalize(formula+'+1')&&savedCell.v===7,'Closing during formula editing converted it to text');
  assert(saved[1].calcChain.length===2,'Formula chain lost');
  leaf=await open();view=leaf.view;assert(normalize(view.workbook.current.getSheet().data[0][0].f)===normalize(formula+'+1'),'Formula lost on reopen');
  view.workbook.current.activateSheet({id:'inputs'});await delay(100);input=await begin(0,false);input.innerText='8';await press(input);
  const target=view.workbook.current.getSheet({id:'formulas'});assert(target.data[0][0].v===13&&target.data[1][0].v===12,'Dependency recalculation failed after reopen');
  results.push('Closing during editing, saving, reopening and input edits retain formula dependency tracking');
  view.workbook.current.activateSheet({id:'formulas'});await delay(100);
  input=await begin(1);input.innerText='First line\nSecond line';await press(input);
  cell=view.workbook.current.getSheet().data[1][0];assert(!cell.f&&cell.ct.t==='inlineStr','Ordinary multiline text no longer works');
  results.push('Intentional replacement with ordinary multiline text remains supported');
  await leaf.detach();await delay(100);if(prior)app.workspace.setActiveLeaf(prior,{focus:true});
  for(const f of files)await app.vault.trash(f,true);await app.vault.trash(app.vault.getAbstractFileByPath(root),true);
  return JSON.stringify({passed:results.length,results});
})()

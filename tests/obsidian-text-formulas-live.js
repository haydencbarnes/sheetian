// Run through Obsidian CLI eval. Uses isolated files and restores the prior tab.
(async()=>{
  const delay=ms=>new Promise(r=>setTimeout(r,ms)),assert=(ok,msg)=>{if(!ok)throw Error(msg);};
  const prior=app.workspace.activeLeaf,root=`Sheetian TEXT checks ${Date.now()}`,results=[];
  await app.vault.createFolder(root);
  const formula='=IF(D10-C10>0,\n  "✅ Exceeded goal by $" & TEXT(D10-C10,"#,##0.00"),\n  IF(D10-C10=0,\n    "✅ Met goal exactly!",\n    "❌ Under goal by $" & TEXT(ABS(D10-C10),"#,##0.00")\n  )\n)\n';
  const data=[{id:'goals',name:'Goals',status:1,calcChain:[{r:9,c:4,id:'goals'}],celldata:[
    {r:9,c:2,v:{v:5000}},{r:9,c:3,v:{v:6234.56}},
    {r:9,c:4,v:{f:formula,v:'#ERROR!',ct:{fa:'General',t:'g'}}}]}];
  const file=await app.vault.create(`${root}/goals.sheet`,JSON.stringify(data));
  let leaf,view;
  const open=async()=>{leaf=app.workspace.getLeaf(true);await leaf.openFile(file);app.workspace.setActiveLeaf(leaf,{focus:true});await delay(450);view=leaf.view;};
  const edit=async(r,c,text)=>{
    view.workbook.current.setSelection([{row:[r,r],column:[c,c]}]);await delay(80);
    const area=view.contentEl.querySelector('.fortune-cell-area'),rect=area.getBoundingClientRect();
    area.dispatchEvent(new MouseEvent('dblclick',{clientX:rect.left+10,clientY:rect.top+10,button:0,bubbles:true}));await delay(90);
    const input=view.contentEl.querySelector('.luckysheet-cell-input');input.innerText=text;
    input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));await delay(120);
  };
  try{
    await open();
    // Force calculation with equivalent whitespace, then restore exact source.
    // An unchanged editor commit intentionally retains the cached value.
    view.workbook.current.setCellValue(9,4,formula+' ');await delay(150);
    view.workbook.current.setCellValue(9,4,formula);await delay(150);
    const cell=()=>view.workbook.current.getSheet().data[9][4];
    assert(cell().f===formula&&cell().v==='✅ Exceeded goal by $1,234.56','TEXT failed to replace a cached error');
    results.push('Recalculate an imported multiline IF/ABS/TEXT formula with a cached error');
    for(const [value,expected] of [['5000','✅ Met goal exactly!'],['3765.44','❌ Under goal by $1,234.56'],['6234.56','✅ Exceeded goal by $1,234.56']]){
      await edit(9,3,value);assert(cell().v===expected&&cell().f===formula&&cell().ct.t==='s','Goal outcome or formula result type changed');
    }
    results.push('Input edits calculate over, on and under goal, keeping the full message as text');
    await edit(9,4,formula);assert(cell().v==='✅ Exceeded goal by $1,234.56'&&cell().f===formula,'Double-click commit lost formula or result');
    await leaf.detach();await delay(150);await open();
    assert(cell().v==='✅ Exceeded goal by $1,234.56'&&cell().f===formula,'Saved formula/message changed after reopening');
    await edit(9,3,'5000');assert(cell().v==='✅ Met goal exactly!','Dependencies failed after reopening');
    results.push('Double-click, save, reopen and subsequent input changes preserve the formula and message');
    for(const [f,expected] of [['=TEXT(12.5,"0.00")','12.50'],['=TEXT(0.125,"0.0%")','12.5%'],['=TEXT(46023,"yyyy-mm-dd")','2026-01-01'],['=""','']]){
      await edit(0,0,f);const c=view.workbook.current.getSheet().data[0][0];assert(c.f===f&&c.v===expected&&c.ct.t==='s','TEXT return value was coerced');
    }
    results.push('Number, percentage, date and empty-string results retain their text type in the live editor');
    return JSON.stringify({passed:results.length,results});
  }finally{
    if(leaf)await leaf.detach();await delay(100);if(prior)app.workspace.setActiveLeaf(prior,{focus:true});
    await app.vault.trash(file,true);const folder=app.vault.getAbstractFileByPath(root);if(folder)await app.vault.trash(folder,true);
  }
})()

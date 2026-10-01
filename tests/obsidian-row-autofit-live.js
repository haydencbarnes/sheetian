// Run through Obsidian CLI eval. Creates/trashes only isolated test files.
(async()=>{
  const delay=ms=>new Promise(r=>setTimeout(r,ms));
  const assert=(ok,message)=>{if(!ok)throw Error(message);};
  const prior=app.workspace.activeLeaf,root=`Sheetian row autofit checks ${Date.now()}`,files=[],results=[];
  await app.vault.createFolder(root);
  const open=async file=>{const leaf=app.workspace.getLeaf(true);await leaf.openFile(file);app.workspace.setActiveLeaf(leaf,{focus:true});await delay(500);assert(leaf.view.getViewType()==='spreadsheet-view','Sheet view not ready');return leaf;};
  const doubleClick=async(view,row,frozen=false)=>{
    const api=view.workbook.current,sheet=api.getSheet(),header=view.contentEl.querySelector('.fortune-row-header'),rect=header.getBoundingClientRect();
    const heights=api.getRowHeight(Array.from({length:row+1},(_,i)=>i));
    const edge=Object.values(heights).reduce((a,b)=>a+b+1,0)*(sheet.zoomRatio||1);
    const x=rect.left+rect.width/2,y=rect.top+edge-(frozen?0:header.scrollTop)-2;
    header.dispatchEvent(new MouseEvent('mousemove',{clientX:x,clientY:y,bubbles:true}));await delay(70);
    const handle=header.querySelector('.fortune-rows-change-size');
    assert(Number(handle.dataset.sheetianRow)===row,`Wrong row target ${handle.dataset.sheetianRow}, expected ${row}; y=${y}, top=${rect.top}, currentTop=${header.getBoundingClientRect().top}, scroll=${header.scrollTop}, edge=${edge}`);
    for(let detail=1;detail<=2;detail++){
      handle.dispatchEvent(new MouseEvent('mousedown',{clientX:x,clientY:y,button:0,detail,bubbles:true}));await delay(30);
      handle.dispatchEvent(new MouseEvent('mouseup',{clientX:x,clientY:y,button:0,detail,bubbles:true}));await delay(30);
    }
    handle.dispatchEvent(new MouseEvent('dblclick',{clientX:x,clientY:y,button:0,detail:2,bubbles:true,cancelable:true}));await delay(150);
    return view.workbook.current.getRowHeight([row])[row];
  };
  const data=[{name:'Rows',id:'rows',row:100,column:26,config:{columnlen:{0:80,1:80},rowlen:{0:19,1:19,2:19,3:19,20:19},customHeight:{0:1,1:1,2:1,3:1,20:1}},celldata:[
    {r:0,c:0,v:{v:'Wrapped text needs several lines at the current column width',tb:'2',fs:14}},
    {r:1,c:0,v:{v:'One\nTwo\nThree',ct:{t:'inlineStr',s:[{v:'One\nTwo\nThree',fs:12}]}}},
    {r:2,c:0,v:{ct:{t:'inlineStr',s:[{v:'Small ',fs:10},{v:'Large',fs:30}]}}},
    {r:3,c:0,v:{v:'short'}},
    {r:20,c:0,v:{v:'Scrolled row\nsecond line\nthird line',ct:{t:'inlineStr',s:[{v:'Scrolled row\nsecond line\nthird line',fs:12}]}}},
  ]}];
  const file=await app.vault.create(`${root}/rows.sheet`,JSON.stringify(data));files.push(file);
  let leaf=await open(file),view=leaf.view;
  const height=await doubleClick(view,0);
  assert(height>50,'Wrapped row did not expand');
  assert(view.workbook.current.getRowHeight([1])[1]===19,'Unselected row resized');
  const overlay=view.contentEl.querySelector('.fortune-sheet-overlay');overlay.focus();
  overlay.dispatchEvent(new KeyboardEvent('keydown',{key:'z',code:'KeyZ',keyCode:90,ctrlKey:true,bubbles:true}));await delay(100);
  assert(view.workbook.current.getRowHeight([0])[0]<height,'Row AutoFit undo failed');
  overlay.dispatchEvent(new KeyboardEvent('keydown',{key:'y',code:'KeyY',keyCode:89,ctrlKey:true,bubbles:true}));await delay(100);
  assert(view.workbook.current.getRowHeight([0])[0]===height,'Row AutoFit redo failed');
  results.push('Double-click fits wrapped text at current column width; undo and redo work');
  await leaf.detach();await delay(100);
  assert(JSON.parse(await app.vault.read(file))[0].config.rowlen[0]===height,'Height not saved');
  leaf=await open(file);view=leaf.view;assert(view.workbook.current.getRowHeight([0])[0]===height,'Height lost on reopen');
  results.push('Row height persists across save and reopen');
  // Actual row-header selection events, including Shift for multiple rows.
  const header=view.contentEl.querySelector('.fortune-row-header'),rect=header.getBoundingClientRect();
  for(const [y,shiftKey] of [[rect.top+10,false],[rect.top+height+12,true]]){
    header.dispatchEvent(new MouseEvent('mousedown',{clientX:rect.left+10,clientY:y,button:0,shiftKey,bubbles:true}));
    header.dispatchEvent(new MouseEvent('mouseup',{clientX:rect.left+10,clientY:y,button:0,shiftKey,bubbles:true}));await delay(60);
  }
  assert(view.workbook.current.getSelection()[0].column[1]===25,'Row headers did not select whole rows');
  const multiline=await doubleClick(view,1);
  assert(multiline>35,'Explicit multiline row not fit');
  assert(view.workbook.current.getRowHeight([0])[0]===height,'Selected wrapped row lost fit');
  results.push('Multiple selected row headers fit independently, including explicit line breaks');
  view.workbook.current.setSelection([{row:[2,2],column:[0,0]}]);
  const rich=await doubleClick(view,2);assert(rich>30,'Rich text font size not accommodated');
  view.workbook.current.setColumnWidth({0:300},{},true);await delay(100);
  view.workbook.current.setSelection([{row:[0,0],column:[0,0]}]);
  const wider=await doubleClick(view,0);assert(wider<height,'Row did not shrink when wrapping needed fewer lines');
  results.push('Styled rich text fits; widening a column lets wrapped rows shrink on AutoFit');
  view.workbook.current.setSelection([{row:[20,20],column:[0,0]}]);
  view.workbook.current.scroll({scrollTop:450,scrollLeft:0});await delay(120);
  assert(await doubleClick(view,20)>35,'Scrolled row did not resize');
  view.workbook.current.freeze('row',{row:0,column:0});await delay(100);
  view.workbook.current.scroll({scrollTop:450,scrollLeft:0});await delay(100);
  assert(await doubleClick(view,0,true)===wider,'Frozen row targeted incorrectly');
  results.push('Scrolled and frozen row boundaries target the correct row');
  await leaf.detach();await delay(100);
  const zoomFile=await app.vault.create(`${root}/zoomed.sheet`,JSON.stringify([{...data[0],zoomRatio:1.3}]));files.push(zoomFile);
  leaf=await open(zoomFile);assert(await doubleClick(leaf.view,0)===height,'Zoom changed the saved row height');
  results.push('Zoomed headers fit using unscaled saved heights');
  await leaf.detach();await delay(100);
  if(prior)app.workspace.setActiveLeaf(prior,{focus:true});
  for(const f of files)await app.vault.trash(f,true);
  await app.vault.trash(app.vault.getAbstractFileByPath(root),true);
  return JSON.stringify({passed:results.length,results});
})()

// Run through Obsidian CLI eval. Tests isolated files and restores the prior tab.
(async()=>{
  const wait=ms=>new Promise(r=>setTimeout(r,ms)),assert=(v,m)=>{if(!v)throw Error(m);};
  const until=async fn=>{for(let n=0;n<50&&!fn();n++)await wait(20);};
  const prior=app.workspace.activeLeaf,root=`Sheetian formula bar checks ${Date.now()}`,results=[];
  const formula='=SUM(\nA2:A3\n)',source=JSON.stringify([{id:'formulas',name:'Formulas',status:1,calcChain:[{r:0,c:0,id:'formulas'}],celldata:[{r:0,c:0,v:{f:formula,v:6,m:'6'}},{r:1,c:0,v:{v:2}},{r:2,c:0,v:{v:4}}]}]);
  await app.vault.createFolder(root);const file=await app.vault.create(root+'/formulas.sheet',source),leaf=app.workspace.getLeaf(true);let second;
  try{
    await leaf.openFile(file);app.workspace.setActiveLeaf(leaf,{focus:true});await wait(450);
    const view=leaf.view,bar=()=>view.contentEl.querySelector('.fortune-fx-editor'),handle=()=>view.contentEl.querySelector('.sheetian-formula-resize'),toggle=()=>view.contentEl.querySelector('.sheetian-formula-toggle'),input=()=>view.contentEl.querySelector('.fortune-fx-input');
    const height=()=>bar().offsetHeight,key=async k=>{handle().dispatchEvent(new KeyboardEvent('keydown',{key:k,code:k,bubbles:true}));await wait(100);};
    view.workbook.current.setSelection([{row:[0,0],column:[0,0],row_focus:0,column_focus:0}]);await wait(100);
    assert(height()===29,'Initial bar is not compact');assert(input().innerText===formula,'Formula source changed');
    const before=view.getViewData(),placeholder=view.contentEl.querySelector('.fortune-sheet-canvas-placeholder'),initialGrid=placeholder.clientHeight;
    toggle().click();await wait(150);assert(height()===160&&toggle().getAttribute('aria-expanded')==='true','Expand button failed');
    assert(Math.abs(initialGrid-placeholder.clientHeight-131)<2,'Grid did not shrink with the bar');
    const canvas=view.contentEl.querySelector('.fortune-sheet-canvas');await until(()=>Math.abs(parseFloat(canvas.style.height)-placeholder.clientHeight)<2);assert(Math.abs(parseFloat(canvas.style.height)-placeholder.clientHeight)<2,'Canvas geometry did not update: '+JSON.stringify({canvas:canvas.style.height,placeholder:placeholder.clientHeight,bar:height()}));
    toggle().click();await wait(100);assert(height()===29,'Collapse button failed');
    results.push('Expand/collapse moves the grid and resizes the canvas correctly');
    await key('ArrowDown');assert(height()===39,'Keyboard growth failed');await key('End');assert(height()===Number(handle().getAttribute('aria-valuemax')),'Maximum clamp failed');await key('Home');assert(height()===29,'Minimum clamp failed');
    results.push('Keyboard resizing and pane-aware minimum/maximum bounds work');
    // Synthetic pointer events do not create a native capture stream. Stub only
    // capture for this isolated test; native pointer dragging is checked via CDP.
    const h=handle(),capture=h.setPointerCapture,has=h.hasPointerCapture,release=h.releasePointerCapture;
    h.setPointerCapture=()=>{};h.hasPointerCapture=()=>false;h.releasePointerCapture=()=>{};
    h.dispatchEvent(new PointerEvent('pointerdown',{pointerId:7,button:0,clientY:100,bubbles:true}));h.dispatchEvent(new PointerEvent('pointermove',{pointerId:7,clientY:220,bubbles:true}));await wait(100);assert(height()===149,'Drag expansion failed');
    h.dispatchEvent(new PointerEvent('pointermove',{pointerId:7,clientY:120,bubbles:true}));await wait(100);assert(height()===49,'Drag shrink failed');h.dispatchEvent(new PointerEvent('pointerup',{pointerId:7,bubbles:true}));
    h.setPointerCapture=capture;h.hasPointerCapture=has;h.releasePointerCapture=release;
    results.push('Pointer dragging grows and shrinks the formula bar');
    assert(view.getViewData()===before&&input().innerText===formula,'Resizing changed file contents or formula source');
    second=app.workspace.getLeaf(true);await second.openFile(file);await wait(400);assert(second.view.contentEl.querySelector('.fortune-fx-editor').offsetHeight===29,'Height leaked into another pane');await second.detach();second=null;app.workspace.setActiveLeaf(leaf,{focus:true});await wait(150);
    results.push('Resizing is independent per pane and leaves workbook data untouched');
    // Opening a second pane can restore DOM focus after clearing edit state.
    // Start a fresh focus session, as a user clicking back into the bar would.
    input().blur();await wait(50);input().focus();await wait(100);input().innerText=formula+'+1';
    toggle().dispatchEvent(new MouseEvent('mousedown',{button:0,bubbles:true,cancelable:true}));toggle().click();await wait(100);assert(input().innerText===formula+'+1','Resize discarded the pending formula');
    input().dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));await wait(150);
    await until(()=>view.workbook.current.getSheet().data[0][0].v===7);
    let cell=view.workbook.current.getSheet().data[0][0];assert(cell.f===formula+'+1'&&cell.v===7,'Formula-bar edit failed to calculate: '+JSON.stringify(cell));await view.save();
    const saved=JSON.parse(await app.vault.read(file))[0].celldata.find(c=>c.r===0&&c.c===0).v;assert(saved.f===formula+'+1'&&saved.v===7,'Formula-bar edit failed to save');
    results.push('Resizing during formula editing retains the draft, calculation and saved source');
    return JSON.stringify({passed:results.length,results});
  }finally{if(second)await second.detach();await leaf.detach();await wait(100);if(prior)app.workspace.setActiveLeaf(prior,{focus:true});await app.vault.trash(file,true);await app.vault.trash(app.vault.getAbstractFileByPath(root),true);}
})()

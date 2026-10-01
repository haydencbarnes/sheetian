// Run through Obsidian CLI eval. Only isolated test files are created/trashed.
(async () => {
  const delay = ms => new Promise(resolve => setTimeout(resolve,ms));
  const assert = (ok,message) => { if(!ok) throw new Error(message); };
  const prior = app.workspace.activeLeaf, root = `Sheetian scroll checks ${Date.now()}`;
  await app.vault.createFolder(root);
  const files=[], results=[];
  for(const zoom of [1,1.3]) {
    const raw=JSON.stringify([{name:'Scroll',id:`scroll-${zoom}`,row:200,column:26,zoomRatio:zoom,
      celldata:[{r:0,c:0,v:{v:'Top'}},{r:150,c:0,v:{v:'Bottom'}}]}]);
    const file=await app.vault.create(`${root}/${zoom}.sheet`,raw); files.push(file);
    const leaf=app.workspace.getLeaf(true);await leaf.openFile(file);app.workspace.setActiveLeaf(leaf,{focus:true});await delay(180);
    const view=leaf.view,area=view.contentEl.querySelector('.fortune-cell-area');
    const bar=view.contentEl.querySelector('.luckysheet-scrollbar-y');
    const wheel=async(deltaY,deltaX=0)=>{
      area.dispatchEvent(new WheelEvent('wheel',{deltaY,deltaX,bubbles:true,cancelable:true}));await delay(25);
    };
    for(let i=0;i<12;i++)await wheel(100);
    assert(bar.scrollTop>0,'Downward scrolling failed');
    const down=bar.scrollTop;
    let steps=0;
    while(bar.scrollTop>0 && steps++<300) {
      const before=bar.scrollTop;await wheel(-100);
      assert(bar.scrollTop<before,`Upward scrolling stuck at ${before}, zoom ${zoom}`);
      assert(Math.abs(area.scrollTop-bar.scrollTop)<1/window.devicePixelRatio,'Grid and vertical scrollbar diverged');
    }
    assert(bar.scrollTop===0,'Could not return to the top');
    await leaf.detach();await delay(100);
    assert(await app.vault.read(file)===raw,'Scrolling rewrote untouched workbook contents');
    results.push(`Down then back to top at ${zoom*100}% sheet zoom (from ${down})`);
  }
  const file=await app.vault.create(`${root}/frozen-hidden.sheet`,JSON.stringify([{name:'Variable rows',id:'var',row:100,column:26,zoomRatio:1.3,
    config:{rowlen:{1:30,2:40,3:50},rowhidden:{4:0,5:0,6:0}},celldata:[{r:0,c:0,v:{v:'Frozen heading'}}]}]));files.push(file);
  const leaf=app.workspace.getLeaf(true);await leaf.openFile(file);app.workspace.setActiveLeaf(leaf,{focus:true});await delay(150);
  const view=leaf.view;view.workbook.current.freeze('row',{row:1,column:0});await delay(80);
  view.workbook.current.scroll({scrollTop:99999,scrollLeft:0});await delay(100);
  const area=view.contentEl.querySelector('.fortune-cell-area'),bar=view.contentEl.querySelector('.luckysheet-scrollbar-y');
  assert(bar.scrollTop>0,'Bottom scroll setup failed');
  let steps=0;
  while(bar.scrollTop>0 && steps++<200){
    const before=bar.scrollTop;
    area.dispatchEvent(new WheelEvent('wheel',{deltaY:-100,bubbles:true,cancelable:true}));await delay(25);
    assert(bar.scrollTop<before,'Upward scrolling stuck with frozen/hidden/custom-height rows');
  }
  assert(bar.scrollTop===0,'Frozen rows prevented returning to top');
  results.push('Up from bottom with frozen, hidden, and custom-height rows');
  await delay(80);
  const horizontal=view.contentEl.querySelector('.luckysheet-scrollbar-x');
  area.dispatchEvent(new WheelEvent('wheel',{deltaX:100,deltaY:0,bubbles:true,cancelable:true}));await delay(80);
  assert(horizontal.scrollLeft>0,'Horizontal scrolling regressed');
  area.dispatchEvent(new WheelEvent('wheel',{deltaX:-100,deltaY:0,bubbles:true,cancelable:true}));await delay(80);
  assert(horizontal.scrollLeft===0,'Could not scroll left again');
  results.push('Horizontal wheel scrolling still works in both directions');
  await leaf.detach();await delay(100);
  if(prior)app.workspace.setActiveLeaf(prior,{focus:true});
  for(const f of files)await app.vault.trash(f,true);
  await app.vault.trash(app.vault.getAbstractFileByPath(root),true);
  return JSON.stringify({passed:results.length,results});
})()

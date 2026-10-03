// Run through Obsidian CLI eval. Uses isolated files and restores the prior tab.
(async()=>{
  const wait=ms=>new Promise(resolve=>setTimeout(resolve,ms)),assert=(value,message)=>{if(!value)throw Error(message);},results=[];
  const prior=app.workspace.activeLeaf,root=`Sheetian zoom checks ${Date.now()}`;
  const formula='=SUM(A2:A3)',source=JSON.stringify([
    {id:'first',name:'First',status:1,zoomRatio:1.3,calcChain:[{r:0,c:0,id:'first'}],celldata:[{r:0,c:0,v:{f:formula,v:5,m:'5',bl:1}},{r:1,c:0,v:{v:2}},{r:2,c:0,v:{v:3}}]},
    {id:'second',name:'Second',status:0,zoomRatio:0.75,celldata:[{r:0,c:0,v:{v:'Second worksheet'}}]},
  ]);
  await app.vault.createFolder(root);
  const file=await app.vault.create(root+'/zoom.sheet',source),other=await app.vault.create(root+'/other.sheet',source);
  let leaf=app.workspace.getLeaf(true),second;
  try {
    await leaf.openFile(file);app.workspace.setActiveLeaf(leaf,{focus:true});await wait(400);
    const view=()=>leaf.view,api=()=>view().workbook.current,dropdown=()=>view().contentEl.querySelector('.sheetian-toolbar-zoom');
    const choose=async value=>{const select=dropdown();select.value=String(value);select.dispatchEvent(new Event('change',{bubbles:true}));await wait(120);};
    const first=dropdown(),painter=view().contentEl.querySelector('.fortune-toolbar [data-tips="Format-Painter"]');
    assert(first&&first.previousElementSibling===painter,'Zoom is not beside Format Painter');
    assert(!view().contentEl.querySelector('.fortune-zoom-container'),'Footer zoom remains');
    assert(first.value==='130','Saved intermediate zoom is not shown');
    assert([...first.options].map(option=>option.value).join(',')==='50,75,90,100,125,130,150,175,200','Wrong presets');
    assert(view().getViewData()===source,'Mounting zoom rewrote an untouched file');
    results.push('Toolbar position, footer removal, presets and saved 130% display');
    const original=JSON.stringify(api().getSheet().data[0][0]);
    await choose(125);assert(api().getSheet().zoomRatio===1.25&&dropdown().value==='125','125% was rounded or not applied');
    dropdown().dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown',bubbles:true}));await wait(100);
    assert(api().getSheet().zoomRatio===1.5,'Arrow-key zoom failed');await choose(125);
    assert(JSON.stringify(api().getSheet().data[0][0])===original,'Zoom changed cell data');
    api().setCellValue(1,0,4);await wait(100);await choose(150);api().handleUndo();await wait(120);
    assert(api().getSheet().data[1][0].v===2,'Zoom consumed the undo step');
    results.push('Exact percentage zoom preserves formulas and cell-edit undo');
    api().activateSheet({id:'second'});await wait(100);assert(dropdown().value==='75','Second sheet zoom not reflected');
    await choose(175);assert(api().getSheet().zoomRatio===1.75,'Second sheet zoom failed');
    api().activateSheet({id:'first'});await wait(100);assert(dropdown().value==='150','Zoom leaked between sheets');
    const menu=view().sheetMenus(view().generation).find(menu=>menu.title==='View');
    menu.items.find(item=>item?.title==='Zoom 125%').action();await wait(100);assert(dropdown().value==='125','View menu zoom did not update toolbar');
    results.push('Each worksheet retains zoom and the View menu updates the dropdown');
    api().setSelection([{row:[0,0],column:[0,0],row_focus:0,column_focus:0}]);await wait(100);
    const fx=view().contentEl.querySelector('.fortune-fx-input');fx.blur();await wait(30);fx.focus();await wait(200);
    assert(document.activeElement===fx,'Formula input not focused: '+document.activeElement?.className);
    fx.innerText=formula+'+1';
    dropdown().dispatchEvent(new MouseEvent('mousedown',{button:0,bubbles:true,cancelable:true}));dropdown().focus();await choose(100);
    const cell=api().getSheet().data[0][0];
    if(cell.f!==formula+'+1'){
      assert(fx.innerText===formula+'+1','Zoom discarded an uncommitted formula: '+JSON.stringify({cell,input:fx.innerText}));fx.focus();
      fx.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));await wait(150);
    }
    assert(api().getSheet().data[0][0].f===formula+'+1'&&api().getSheet().data[0][0].v===6,'Formula edit was lost during zoom');
    results.push('Zoom while editing retains the formula draft and calculated value');
    fx.focus();await wait(100);fx.innerText=formula+'+2';
    // Background CLI eval updates activeElement without emitting native blur.
    // Supply that missing event to exercise the keyboard-focus path as well.
    const background=!document.hasFocus();dropdown().focus();
    if(background)fx.dispatchEvent(new FocusEvent('blur',{relatedTarget:dropdown()}));
    await choose(125);assert(api().getSheet().data[0][0].f===formula+'+2'&&api().getSheet().data[0][0].v===7,'Keyboard focus lost the formula draft');
    await choose(100);results.push('Keyboard focus into zoom commits the formula draft');
    await view().save();const saved=JSON.parse(await app.vault.read(file));
    assert(saved.find(sheet=>sheet.id==='first').zoomRatio===1&&saved.find(sheet=>sheet.id==='second').zoomRatio===1.75,'Zoom was not saved');
    await leaf.detach();leaf=app.workspace.getLeaf(true);await leaf.openFile(file);app.workspace.setActiveLeaf(leaf,{focus:true});await wait(400);
    assert(dropdown().value==='100'&&api().getSheet().data[0][0].v===7,'Zoom/formula did not survive reopen');
    api().activateSheet({id:'second'});await wait(100);assert(dropdown().value==='175','Second sheet zoom lost after reopen');
    results.push('Per-sheet zoom and edited formulas survive save/reopen');
    second=app.workspace.getLeaf('split','vertical');await second.openFile(other);app.workspace.setActiveLeaf(second,{focus:true});await wait(300);
    const secondDropdown=second.view.contentEl.querySelector('.sheetian-toolbar-zoom');secondDropdown.value='200';secondDropdown.dispatchEvent(new Event('change',{bubbles:true}));await wait(120);
    assert(second.view.workbook.current.getSheet().zoomRatio===2&&api().getSheet().zoomRatio===1.75,'Zoom leaked into another pane');
    assert(secondDropdown.getBoundingClientRect().right<=second.view.contentEl.getBoundingClientRect().right,'Dropdown clipped in split pane');
    results.push('Split-pane zoom is independent and remains visible');
    return JSON.stringify({passed:results.length,results});
  } finally {
    if(second)await second.detach();await leaf.detach();await wait(100);if(prior)app.workspace.setActiveLeaf(prior,{focus:true});
    await app.vault.trash(file,true);await app.vault.trash(other,true);await app.vault.trash(app.vault.getAbstractFileByPath(root),true);
  }
})()

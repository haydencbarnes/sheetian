// Run through Obsidian CLI eval. Uses isolated files and restores the prior tab.
(async () => {
  const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (ok, message) => { if (!ok) throw new Error(message); };
  const prior = app.workspace.activeLeaf;
  const root = `Sheetian autofit checks ${Date.now()}`;
  await app.vault.createFolder(root);
  const files = [], results = [];
  const open = async file => {
    const leaf = app.workspace.getLeaf(true);
    await leaf.openFile(file); app.workspace.setActiveLeaf(leaf, { focus: true }); await delay(180);
    return leaf;
  };
  const doubleClick = async (view, column, frozen = false) => {
    const api = view.workbook.current, sheet = api.getSheet();
    const header = view.contentEl.querySelector('.fortune-col-header');
    const rect = header.getBoundingClientRect();
    const widths = api.getColumnWidth(Array.from({length:column+1}, (_,i)=>i));
    const edge = Object.values(widths).reduce((a,b)=>a+b,0) * (sheet.zoomRatio || 1);
    const x = rect.left + edge - (frozen ? 0 : header.scrollLeft) - 2;
    const y = rect.top + rect.height / 2;
    header.dispatchEvent(new MouseEvent('mousemove', {clientX:x, clientY:y, bubbles:true})); await delay(70);
    const handle = header.querySelector('.fortune-cols-change-size');
    assert(Number(handle.dataset.sheetianColumn) === column, `Wrong resize target: ${handle.dataset.sheetianColumn}, expected ${column}`);
    for (let detail = 1; detail <= 2; detail++) {
      handle.dispatchEvent(new MouseEvent('mousedown', {clientX:x, clientY:y, button:0, detail, bubbles:true}));
      await delay(30);
      handle.dispatchEvent(new MouseEvent('mouseup', {clientX:x, clientY:y, button:0, detail, bubbles:true}));
      await delay(30);
    }
    handle.dispatchEvent(new MouseEvent('dblclick', {clientX:x, clientY:y, button:0, detail:2, bubbles:true, cancelable:true}));
    await delay(150);
    return view.workbook.current.getColumnWidth([column])[column];
  };
  const data = [{name:'Autofit',id:'autofit',row:50,column:26,config:{columnlen:{0:73,1:73}},celldata:[
    {r:0,c:0,v:{v:'A long heading that needs more column space',bl:1,fs:16}},
    {r:1,c:0,v:{v:'Short\nLonger second line'}},
    {r:0,c:1,v:{v:1234,m:'$1,234.00',ct:{fa:'$#,##0.00',t:'n'}}},
    {r:0,c:6,v:{v:'Scrolled column text must fit too',fs:14}},
  ]}];
  const file = await app.vault.create(`${root}/autofit.sheet`,JSON.stringify(data)); files.push(file);
  let leaf = await open(file), view = leaf.view;
  const canvas = document.createElement('canvas').getContext('2d');
  canvas.font = 'normal bold 16pt "Times New Roman", "Helvetica Neue", Helvetica, Arial, sans-serif';
  const required = Math.ceil(canvas.measureText(data[0].celldata[0].v.v).width + 12);
  const width = await doubleClick(view,0);
  assert(width === required, `Autofit ${width} does not fit displayed text ${required}`);
  assert(view.workbook.current.getColumnWidth([1])[1] === 73, 'Unselected neighbor resized');
  // Undo/redo goes through the workbook's real keyboard handler.
  const overlay = view.contentEl.querySelector('.fortune-sheet-overlay'); overlay.focus();
  overlay.dispatchEvent(new KeyboardEvent('keydown',{key:'z',code:'KeyZ',keyCode:90,ctrlKey:true,bubbles:true})); await delay(100);
  assert(view.workbook.current.getColumnWidth([0])[0] < width, 'Autofit undo failed');
  overlay.dispatchEvent(new KeyboardEvent('keydown',{key:'y',code:'KeyY',keyCode:89,ctrlKey:true,bubbles:true})); await delay(100);
  assert(view.workbook.current.getColumnWidth([0])[0] === width, 'Autofit redo failed');
  results.push('Double-click fits styled/multiline contents; undo and redo work');
  await leaf.detach(); await delay(100);
  assert(JSON.parse(await app.vault.read(file))[0].config.columnlen[0] === width, 'Width not saved');
  leaf = await open(file); view = leaf.view;
  assert(view.workbook.current.getColumnWidth([0])[0] === width, 'Width lost on reopen');
  results.push('Column widths persist across save and reopen');
  // Select whole columns using actual header events, including Shift selection.
  const header = view.contentEl.querySelector('.fortune-col-header'), rect = header.getBoundingClientRect();
  for (const [x,shiftKey] of [[rect.left+10,false],[rect.left+width+10,true]]) {
    header.dispatchEvent(new MouseEvent('mousedown',{clientX:x,clientY:rect.top+10,button:0,shiftKey,bubbles:true}));
    header.dispatchEvent(new MouseEvent('mouseup',{clientX:x,clientY:rect.top+10,button:0,shiftKey,bubbles:true})); await delay(60);
  }
  assert(view.workbook.current.getSelection()[0].row[1] === 49,'Headers did not select columns');
  await doubleClick(view,1);
  assert(view.workbook.current.getColumnWidth([0])[0] === width, 'Selected column A lost its fit');
  canvas.font = 'normal normal 10pt "Times New Roman", "Helvetica Neue", Helvetica, Arial, sans-serif';
  assert(view.workbook.current.getColumnWidth([1])[1] === Math.ceil(canvas.measureText('$1,234.00').width+12), 'Selected column B not fit');
  results.push('Multiple selected column headers fit independently');
  view.workbook.current.setSelection([{row:[0,0],column:[6,6]}]);
  view.workbook.current.scroll({scrollLeft:400,scrollTop:0}); await delay(120);
  const scrolled = await doubleClick(view,6);
  assert(scrolled > 73,'Scrolled column did not resize');
  results.push('Horizontal scrolling targets the correct column');
  view.workbook.current.freeze('column',{row:0,column:0}); await delay(100);
  view.workbook.current.scroll({scrollLeft:450,scrollTop:0}); await delay(100);
  assert(await doubleClick(view,0,true) === width,'Frozen column fit wrong target');
  results.push('Frozen column boundary remains accurate after scrolling');
  await leaf.detach(); await delay(100);
  const zoomedFile = await app.vault.create(`${root}/zoomed.sheet`,JSON.stringify([{...data[0],zoomRatio:1.25}])); files.push(zoomedFile);
  leaf = await open(zoomedFile);
  assert(await doubleClick(leaf.view,0) === width,'Zoom changed the saved column width');
  results.push('Zoomed headers fit using unscaled saved widths');
  await leaf.detach(); await delay(100);
  if(prior) app.workspace.setActiveLeaf(prior,{focus:true});
  for(const f of files) await app.vault.trash(f,true);
  await app.vault.trash(app.vault.getAbstractFileByPath(root),true);
  return JSON.stringify({passed:results.length,results});
})()

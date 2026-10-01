// Run this script through Obsidian CLI eval in a desktop vault.
// Creates an isolated test folder, restores the prior tab, and trashes only its own files on success.
(async () => {
  const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
  const results = []; window._spreadsheetTestResults = results;
  const assert = (condition, message) => { if (!condition) throw new Error(message); };
  const priorLeaf = app.workspace.activeLeaf;
  const root = `Spreadsheet plugin checks ${Date.now()}`;
  assert(!app.vault.getAbstractFileByPath(root), 'Test folder already exists');
  await app.vault.createFolder(root);
  const files = [];
  const makeFile = async (name, data) => {
    const file = await app.vault.create(`${root}/${name}`, typeof data === 'string' ? data : JSON.stringify(data));
    files.push(file); return file;
  };
  const open = async file => {
    const leaf = app.workspace.getLeaf(true);
    await leaf.openFile(file); app.workspace.setActiveLeaf(leaf,{focus:true}); await delay(160);
    assert(leaf.view.getViewType() === 'spreadsheet-view', 'Wrong view');
    assert(leaf.view.workbook.current, 'Workbook did not initialize');
    return leaf;
  };
  const close = async leaf => { await leaf.detach(); await delay(150); };
  const edit = async (view, row, column, text, commit = true) => {
    view.workbook.current.setSelection([{row:[row,row],column:[column,column]}]);
    await delay(60);
    const area = view.contentEl.querySelector('.fortune-cell-area');
    const rect = area.getBoundingClientRect();
    area.dispatchEvent(new MouseEvent('dblclick', {clientX:rect.left+10,clientY:rect.top+10,button:0,bubbles:true}));
    await delay(60);
    const input = view.contentEl.querySelector('.luckysheet-cell-input');
    assert(Number(getComputedStyle(input.closest('.luckysheet-input-box')).zIndex) >= 0, 'Editor failed to open');
    input.innerText = text;
    if (commit) {
      input.dispatchEvent(new KeyboardEvent('keydown', {key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));
      await delay(100);
    }
  };

  const raw = '[ {"name":"Original","id":"original","celldata":[{"r":0,"c":0,"v":{"v":"keep exactly"}}]} ]';
  const untouched = await makeFile('untouched.sheet', raw);
  let leaf = await open(untouched);
  assert(!leaf.view.edited, 'Opening marked file edited');
  assert(leaf.view.contentEl.querySelector('.fortune-location-box')?.textContent?.includes('NaN') !== true, 'Invalid initial selection');
  await close(leaf);
  assert(await app.vault.read(untouched) === raw, 'Untouched contents rewritten');
  results.push('Untouched close: bytes preserved');

  const inactiveFile = await makeFile('inactive.sheet', raw);
  const inactiveLeaf = await open(inactiveFile);
  const metadata = {name:'Budget',id:'budget',status:1,order:0,row:100,column:26,
    celldata:[{r:0,c:0,v:{v:'Item'}},{r:0,c:1,v:{v:'Amount'}},{r:1,c:0,v:{v:'Supplies'}},{r:1,c:1,v:{v:300}}],
    images:[{id:'test-image',src:'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Zl1kAAAAASUVORK5CYII=',width:10,height:10,left:400,top:10}],
    filter_select:{row:[0,3],column:[0,1]},filter:{},
    hyperlink:{'1_0':{linkType:'external',linkAddress:'https://example.com'}},
    frozen:{type:'row',range:{row_focus:0,column_focus:0}},customFutureField:{keep:'yes'}};
  const workbookFile = await makeFile('editing.sheet', [metadata,{name:'Second',id:'second',status:0,order:1,celldata:[{r:0,c:0,v:{v:'Second sheet kept'}}]}]);
  leaf = await open(workbookFile);
  let view = leaf.view;
  await edit(view,2,1,'$125.50');
  assert(view.workbook.current.getCellValue(2,1) === 125.5, 'Manual currency remains text');
  await edit(view,3,1,'=B2+B3');
  assert(view.workbook.current.getCellValue(3,1) === 425.5, 'Currency formula incorrect');
  view.workbook.current.setCellFormat(5,1,'ct',{fa:'yyyy-MM-dd',t:'d'});
  await delay(70); await edit(view,5,1,'2024-03-28');
  assert(view.workbook.current.getSheet().data[5][1].v === 45379, 'Date input incorrect');
  results.push('Manual currency + formula + preformatted date');

  // Paste through the same ClipboardEvent handler used for Excel HTML tables.
  view.workbook.current.setSelection([{row:[9,9],column:[0,0]}]); await delay(60);
  const overlay = view.contentEl.querySelector('.fortune-sheet-overlay');
  view.contentEl.querySelector('.fortune-container').dispatchEvent(new MouseEvent('mousedown',{button:0,bubbles:true})); await delay(60); overlay.focus();
  overlay.dispatchEvent(new KeyboardEvent('keydown',{key:'v',code:'KeyV',keyCode:86,ctrlKey:true,bubbles:true})); await delay(60); overlay.focus();
  const clipboard = new DataTransfer();
  clipboard.setData('text/html','<html><body><table><tr><td style="color:#000000;background:#ffffff">Excel visible</td><td>42</td></tr></table></body></html>');
  clipboard.setData('text/plain','Excel visible\t42');
  overlay.dispatchEvent(new ClipboardEvent('paste',{clipboardData:clipboard,bubbles:true,cancelable:true})); await delay(120);
  assert(view.workbook.current.getCellValue(9,0) === 'Excel visible', 'HTML paste text missing');
  assert(Number(view.workbook.current.getCellValue(9,1)) === 42, 'HTML paste number missing');
  assert(view.workbook.current.getCellValue(9,0,{type:'fc'}) !== '#ffffff', 'Paste text invisible');
  assert(!JSON.stringify(inactiveLeaf.view.workbook.current.getAllSheets()).includes('Excel visible'), 'Inactive workbook stole paste');
  await close(inactiveLeaf);
  assert(await app.vault.read(inactiveFile)===raw, 'Inactive workbook changed');
  results.push('Excel-style HTML paste reaches only the active workbook');

  const scrollbar = view.contentEl.querySelector('.luckysheet-scrollbar-x');
  view.workbook.current.scroll({scrollLeft:0,scrollTop:0}); await delay(80);
  view.workbook.current.setSelection([{row:[1,1],column:[0,0]}]); await delay(60);
  overlay.focus();
  for (const key of ['ArrowDown','ArrowRight','ArrowLeft']) {
    overlay.dispatchEvent(new KeyboardEvent('keydown',{key,code:key,keyCode:{ArrowDown:40,ArrowRight:39,ArrowLeft:37}[key],bubbles:true})); await delay(60);
    assert(scrollbar.scrollLeft === 0, `Viewport jumped after ${key}`);
  }
  results.push('Arrow navigation keeps left viewport stable with sidebar open');

  await view.exportCsv(); await delay(80);
  const csv = app.vault.getAbstractFileByPath(`${root}/editing.csv`); assert(csv, 'CSV not created'); files.push(csv);
  assert((await app.vault.read(csv)).includes('Excel visible'), 'CSV missing values');
  await view.exportCsv();
  const csv2 = app.vault.getAbstractFileByPath(`${root}/editing-1.csv`); assert(csv2, 'CSV overwrite avoidance failed'); files.push(csv2);
  results.push('CSV export and existing-file protection');

  await edit(view,7,2,'typed immediately before closing',false);
  await close(leaf);
  const saved = JSON.parse(await app.vault.read(workbookFile));
  assert(saved[0].celldata.some(c=>c.v?.v==='typed immediately before closing'), 'Last edit lost on close');
  assert(saved[0].images[0].id==='test-image', 'Images lost');
  assert(saved[0].filter_select.row[1]===3 && saved[0].hyperlink['1_0'], 'Filter or link metadata lost');
  assert(saved[0].frozen.type==='row' && saved[0].customFutureField.keep==='yes', 'Metadata lost');
  assert(saved[1].celldata[0].v.v==='Second sheet kept', 'Second sheet lost');
  results.push('Close during edit preserves cell, images, filters, links, freeze, and second sheet');
  leaf = await open(workbookFile); view = leaf.view;
  assert(view.workbook.current.getCellValue(3,1)===425.5, 'Formula lost on reopen');
  const second = await makeFile('other.sheet',[{name:'Other',id:'other',celldata:[{r:0,c:0,v:{v:'Other file intact'}}]}]);
  await leaf.openFile(second); await delay(120); await close(leaf);
  assert((await app.vault.read(second)).includes('Other file intact'), 'File switch overwrote target');
  assert((await app.vault.read(workbookFile)).includes('typed immediately before closing'), 'File switch overwrote original');
  results.push('Reopen and same-tab file switch');

  const invalidRaw = 'Invalid JSON worth preserving';
  const invalid = await makeFile('invalid.sheet',invalidRaw);
  leaf = app.workspace.getLeaf(true); await leaf.openFile(invalid); await delay(100);
  assert(leaf.view.contentEl.querySelector('.spreadsheet-error'), 'Invalid file not protected');
  await close(leaf); assert(await app.vault.read(invalid)===invalidRaw, 'Invalid file overwritten');
  results.push('Malformed files preserved with error view');

  // Restore the user\'s prior tab, and remove only files created by this test.
  if (priorLeaf) app.workspace.setActiveLeaf(priorLeaf,{focus:true});
  for (const file of files) await app.vault.trash(file,true);
  await app.vault.trash(app.vault.getAbstractFileByPath(root),true);
  return JSON.stringify({passed:results.length,results});
})()

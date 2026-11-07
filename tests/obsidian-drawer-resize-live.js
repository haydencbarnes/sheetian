// Run through Obsidian CLI eval. Exercise the same sidebar sizing used by its
// drag handle and inspect the canvas between React commits and browser frames.
(async () => {
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (value, message) => { if (!value) throw Error(message); };
  const prior = app.workspace.activeLeaf, sidebar = app.workspace.leftSplit;
  const originalWidth = parseFloat(sidebar.containerEl.style.width), collapsed = sidebar.collapsed;
  const root = `Sheetian drawer resize checks ${Date.now()}`;
  const cells = [];
  for (let r = 0; r < 300; r++) for (let c = 0; c < 80; c++) cells.push({ r, c, v: { v: r + c, bg: '#dceefe' } });
  const source = JSON.stringify([{ id: 'drawer', name: 'Drawer resize', row: 300, column: 80, celldata: cells,
    frozen: { type: 'rangeBoth', range: { row_focus: 0, column_focus: 0 } } }]);
  await app.vault.createFolder(root);
  const file = await app.vault.create(root + '/drawer.sheet', source), leaf = app.workspace.getLeaf(true);
  let observer, frame;
  try {
    if (collapsed) sidebar.expand();
    await leaf.openFile(file); app.workspace.setActiveLeaf(leaf, { focus: true }); await wait(400);
    const view = leaf.view, workbook = view.workbook.current, reactRoot = view.root;
    const canvas = view.contentEl.querySelector('.fortune-sheet-canvas');
    const placeholder = view.contentEl.querySelector('.fortune-sheet-canvas-placeholder');
    const overlay = view.contentEl.querySelector('.fortune-sheet-overlay');
    workbook.setSelection([{ row: [10, 10], column: [8, 8] }]);
    workbook.scroll({ scrollLeft: 250, scrollTop: 200 }); await wait(150);
    const selection = JSON.stringify(workbook.getSelection());
    const x = view.contentEl.querySelector('.luckysheet-scrollbar-x');
    const y = view.contentEl.querySelector('.luckysheet-scrollbar-y');
    const scroll = [x.scrollLeft, y.scrollTop], context = canvas.getContext('2d'), samples = [];
    const sample = source => {
      if (!canvas.isConnected) return;
      const alphas = [0.25, 0.5, 0.9].map(fraction => context.getImageData(
        Math.min(250, canvas.width - 1), Math.floor((canvas.height - 1) * fraction), 1, 1).data[3]);
      samples.push({ source, alphas, width: placeholder.clientWidth, canvas: canvas.clientWidth, overlay: overlay.clientWidth });
    };
    sample('initial'); assert(samples[0].alphas.every(alpha => alpha === 255), 'Fixture was not fully drawn');
    let bitmapResets = 0;
    observer = new MutationObserver(records => {
      bitmapResets += records.filter(record => record.attributeName === 'width').length;
      sample('mutation');
    });
    observer.observe(canvas, { attributes: true, attributeFilter: ['width', 'height', 'style'] });
    const tick = () => { sample('frame'); frame = requestAnimationFrame(tick); }; frame = requestAnimationFrame(tick);
    // A layout notification with unchanged dimensions must also keep its bitmap.
    for (let n = 0; n < 8; n++) {
      window.dispatchEvent(new Event('resize')); sample('resize-return'); await wait(20);
    }
    const unchangedSizeResets = bitmapResets;
    const widths = new Set(), handle = sidebar.resizeHandleEl, capture = handle.setPointerCapture;
    const pointer = (target, type, clientX) => target.dispatchEvent(new PointerEvent(type, {
      pointerId: 71, clientX, clientY: 200, button: 0, buttons: type === 'pointerup' ? 0 : 1, bubbles: true,
    }));
    handle.setPointerCapture = () => {};
    try {
      pointer(handle, 'pointerdown', handle.getBoundingClientRect().x + 1);
      sample('grab');
      for (let n = 0; n < 48; n++) await new Promise(resolve => requestAnimationFrame(() => {
        pointer(window, 'pointermove', sidebar.containerEl.getBoundingClientRect().x + (n % 2 === 0 ? 240 : 510));
        widths.add(placeholder.clientWidth);
        resolve();
      }));
    } finally {
      pointer(window, 'pointerup', handle.getBoundingClientRect().x + 1);
      handle.setPointerCapture = capture;
    }
    await wait(150); observer.disconnect(); cancelAnimationFrame(frame);
    const blank = samples.filter(sample => sample.alphas.some(alpha => alpha !== 255));
    const mismatched = samples.filter(sample => sample.source === 'mutation' &&
      (sample.canvas !== sample.width || sample.overlay !== sample.width));
    const outcome = { samples: samples.length, widths: [...widths], bitmapResets, unchangedSizeResets, blankUpdates: blank.length,
      blankFrames: blank.filter(sample => sample.source === 'frame').length,
      mismatchedUpdates: mismatched.length, examples: [...blank, ...mismatched].slice(0, 5) };
    window.__sheetianDrawerResizeOutcome = outcome;
    assert(widths.size > 1, 'Sidebar resizing did not change the spreadsheet width');
    assert(unchangedSizeResets === 0, 'Unchanged window resize notifications cleared the canvas: ' + JSON.stringify(outcome));
    assert(blank.length === 0 && mismatched.length === 0, 'Drawer resizing exposed an incomplete grid: ' + JSON.stringify(outcome));
    assert(view.root === reactRoot && canvas.isConnected, 'Drawer resizing remounted the workbook');
    assert(JSON.stringify(workbook.getSelection()) === selection && x.scrollLeft === scroll[0] && y.scrollTop === scroll[1], 'Drawer resizing changed selection or scrolling');
    assert(view.getViewData() === source, 'Drawer resizing changed workbook contents');
    return JSON.stringify({ passed: 6, ...outcome });
  } finally {
    observer?.disconnect(); if (frame !== undefined) cancelAnimationFrame(frame);
    sidebar.setSize(originalWidth); if (collapsed) sidebar.collapse();
    await leaf.detach(); await wait(100); if (prior) app.workspace.setActiveLeaf(prior, { focus: true });
    await app.vault.trash(file, true); await app.vault.trash(app.vault.getAbstractFileByPath(root), true);
  }
})()

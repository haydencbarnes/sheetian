// Run through Obsidian CLI eval. Use trusted Chromium input at the actual
// window boundary, so native scrollbars and mouse release participate too.
(async () => {
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (value, message) => { if (!value) throw Error(message); };
  const prior = app.workspace.activeLeaf, root = `Sheetian native final column checks ${Date.now()}`;
  const source = window.__sheetianNativeReproSource ?? JSON.stringify([{ id: 'native', name: 'Native last column', row: 1168, column: 26,
    zoomRatio: 0.9, config: { columnlen: { 0: 19, 1: 546, 2: 60, 3: 60, 4: 292, 5: 267, 25: 156 } },
    celldata: [{ r: 0, c: 25, v: { v: 'Final column' } }] }]);
  delete window.__sheetianNativeReproSource;
  await app.vault.createFolder(root);
  const file = await app.vault.create(root + '/native.sheet', source), leaf = app.workspace.getLeaf(true);
  const errors = [], result = { stage: 'opening', errors, drags: [] };
  window.__sheetianNativeDragOutcome = result;
  const error = event => errors.push(`${result.stage}: ${event.message}`);
  window.addEventListener('error', error);
  let frame;
  try {
    await leaf.openFile(file); app.workspace.setActiveLeaf(leaf, { focus: true }); await wait(500);
    const view = leaf.view, wc = require('electron').remote.getCurrentWindow().webContents;
    const input = (type, x, y, down = false) => wc.debugger.sendCommand('Input.dispatchMouseEvent', {
      type, x: Math.floor(x), y: Math.floor(y), button: type === 'mouseMoved' ? 'none' : 'left', buttons: down ? 1 : 0,
      ...(type === 'mouseMoved' ? {} : { clickCount: 1 }),
    });
    const monitor = () => {
      const canvas = view.contentEl.querySelector('.fortune-sheet-canvas');
      if (!canvas) {
        if (!errors.some(message => message.includes('Workbook disappeared'))) errors.push(`${result.stage}: Workbook disappeared`);
      } else if (canvas.width > 100 && canvas.height > 100 && canvas.getContext('2d').getImageData(100, 100, 1, 1).data[3] !== 255) {
        if (!errors.some(message => message.includes('Blank canvas'))) errors.push(`${result.stage}: Blank canvas`);
      }
      frame = requestAnimationFrame(monitor);
    }; frame = requestAnimationFrame(monitor);
    for (let n = 0; n < 3; n++) {
      result.stage = `drag ${n + 1}: hover`;
      const bar = view.contentEl.querySelector('.luckysheet-scrollbar-x');
      bar.scrollLeft = bar.scrollWidth; bar.dispatchEvent(new Event('scroll')); await wait(80);
      const canvas = view.contentEl.querySelector('.fortune-sheet-canvas');
      assert(canvas, 'Workbook disappeared before dragging');
      const header = view.contentEl.querySelector('.fortune-col-header'), rect = header.getBoundingClientRect();
      const x = canvas.getBoundingClientRect().right - 3, y = rect.top + 10;
      await input('mouseMoved', x, y); await wait(80);
      const handle = header.querySelector('.fortune-cols-change-size');
      assert(Number(handle.dataset.sheetianColumn) === 25, 'Native pointer did not reach the last-column handle');
      const before = view.workbook.current.getColumnWidth([25])[25];
      await input('mousePressed', x, y, true); await wait(30);
      result.stage = `drag ${n + 1}: moving`;
      const end = Math.min(innerWidth - 1, x + 12);
      for (let step = 1; step <= 12; step++) {
        await input('mouseMoved', x + (end - x) * step / 12, y, true); await wait(20);
      }
      result.stage = `drag ${n + 1}: holding`;
      await wait(1800);
      result.stage = `drag ${n + 1}: release`;
      await input('mouseReleased', end, y); await wait(250);
      assert(errors.length === 0 && view.contentEl.querySelector('.fortune-sheet-canvas'),
        'Native drag/release crashed the workbook: ' + JSON.stringify(result));
      const after = view.workbook.current.getColumnWidth([25])[25];
      assert(after > before + 100, 'Column stalled at the boundary');
      assert(Math.abs(bar.scrollWidth - bar.clientWidth - bar.scrollLeft) < 1, 'Last column edge is no longer visible');
      result.stage = `drag ${n + 1}: undo`;
      view.workbook.current.handleUndo(); await wait(100);
      assert(view.workbook.current.getColumnWidth([25])[25] === before, 'One undo did not restore the width');
      result.stage = `drag ${n + 1}: redo`;
      view.workbook.current.handleRedo(); await wait(100);
      assert(view.workbook.current.getColumnWidth([25])[25] === after, 'One redo did not restore the width');
      assert(Math.abs(bar.scrollWidth - bar.clientWidth - bar.scrollLeft) < 1 && errors.length === 0, 'Redo lost the final edge or blanked the canvas');
      result.drags.push({ before, after });
    }
    result.stage = 'complete';
    return JSON.stringify({ passed: result.drags.length * 5, ...result });
  } finally {
    if (frame !== undefined) cancelAnimationFrame(frame);
    window.removeEventListener('error', error);
    await leaf.detach(); await wait(80); if (prior) app.workspace.setActiveLeaf(prior, { focus: true });
    await app.vault.trash(file, true); await app.vault.trash(app.vault.getAbstractFileByPath(root), true);
  }
})()

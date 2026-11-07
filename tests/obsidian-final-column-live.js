// Run through Obsidian CLI eval. Check the visible boundary throughout a
// drag, including holding at the edge, reversing, committing and undo/redo.
(async () => {
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (value, message) => { if (!value) throw Error(message); };
  const prior = app.workspace.activeLeaf, root = `Sheetian final column checks ${Date.now()}`;
  const files = [], results = [];
  await app.vault.createFolder(root);
  try {
    for (const [zoom, frozen] of [[1, false], [1.3, false], [1.3, true]]) {
      const source = JSON.stringify([{ id: 'last', name: 'Visible final column', row: 100, column: 26, zoomRatio: zoom,
        config: { columnlen: { 25: 120 } }, celldata: [{ r: 0, c: 25, v: { v: 'Final column stays visible' } }],
        ...(frozen ? { frozen: { type: 'column', range: { column_focus: 0, row_focus: 0 } } } : {}) }]);
      const file = await app.vault.create(`${root}/${zoom}-${frozen}.sheet`, source); files.push(file);
      const leaf = app.workspace.getLeaf(true);
      try {
        await leaf.openFile(file); app.workspace.setActiveLeaf(leaf, { focus: true }); await wait(300);
        const view = leaf.view;
        view.contentEl.style.width = '600px'; view.contentEl.style.flex = 'none'; await wait(100);
        const bar = view.contentEl.querySelector('.luckysheet-scrollbar-x');
        bar.scrollLeft = bar.scrollWidth; bar.dispatchEvent(new Event('scroll')); await wait(60);
        const header = view.contentEl.querySelector('.fortune-col-header'), rect = header.getBoundingClientRect();
        const x = view.contentEl.querySelector('.fortune-sheet-canvas').getBoundingClientRect().right - 3, y = rect.top + 10;
        const mouse = (target, type, delta = 0) => target.dispatchEvent(new MouseEvent(type, {
          clientX: x + delta, clientY: y, button: 0, buttons: type === 'mouseup' ? 0 : 1, bubbles: true, cancelable: true,
        }));
        mouse(header, 'mousemove'); await wait(50);
        const handle = header.querySelector('.fortune-cols-change-size');
        assert(Number(handle.dataset.sheetianColumn) === 25, 'Could not grab the final column boundary');
        const initialScroll = bar.scrollLeft, baseWidth = view.workbook.current.getColumnWidth([25])[25];
        const sample = label => {
          const grip = handle.getBoundingClientRect(), canvas = view.contentEl.querySelector('.fortune-sheet-canvas').getBoundingClientRect();
          const line = view.contentEl.querySelector('.fortune-change-size-line').getBoundingClientRect();
          assert(grip.right <= canvas.right + 2 && grip.left >= canvas.left,
            'Resize handle left the visible grid: ' + JSON.stringify({ label, zoom, frozen, grip: grip.toJSON(), canvas: canvas.toJSON() }));
          assert(line.right <= canvas.right + 2 && line.left >= canvas.left, 'Resize guide left the visible grid');
          assert(view.workbook.current.getColumnWidth([25])[25] === baseWidth && view.getViewData() === source,
            'A preview changed saved data or created undo history');
          assert(view.contentEl.scrollLeft === 0, 'Preview scrolled the whole plugin pane');
          return bar.scrollLeft;
        };
        mouse(handle, 'mousedown'); await wait(30);
        mouse(document, 'mousemove', 40); await wait(40);
        const moved = sample('moving');
        assert(moved > initialScroll, 'Viewport did not follow the widening final column');
        await wait(220);
        const held = sample('holding at edge');
        assert(held > moved + 15, 'Holding at the edge did not continue widening the column');
        mouse(document, 'mousemove', -60); await wait(70);
        const reversed = sample('reversed'); await wait(100);
        assert(Math.abs(bar.scrollLeft - reversed) < 1, 'Resizing kept growing after the pointer moved away from the edge');
        mouse(document, 'mousemove', 100); await wait(70); sample('final preview');
        mouse(document, 'mouseup', 100); await wait(100);
        const finalWidth = view.workbook.current.getColumnWidth([25])[25];
        assert(finalWidth > baseWidth, 'Drag did not commit the expanded width');
        const atEnd = () => Math.abs(bar.scrollWidth - bar.clientWidth - bar.scrollLeft) < 1;
        assert(atEnd(), 'Committed final column is offscreen');
        view.workbook.current.handleUndo(); await wait(100);
        assert(view.workbook.current.getColumnWidth([25])[25] === baseWidth, 'A single undo did not restore the original width');
        view.workbook.current.handleRedo(); await wait(100);
        assert(view.workbook.current.getColumnWidth([25])[25] === finalWidth, 'A single redo did not restore the committed width');
        assert(atEnd(), 'Redo moved the final column edge offscreen');
        results.push({ zoom, frozen, initialScroll, moved, held, finalWidth, finalEdgeVisible: atEnd() });
      } finally { await leaf.detach(); await wait(80); }
    }
    return JSON.stringify({ passed: results.length * 9, results });
  } finally {
    if (prior) app.workspace.setActiveLeaf(prior, { focus: true });
    for (const file of files) if (app.vault.getAbstractFileByPath(file.path)) await app.vault.trash(file, true);
    const folder = app.vault.getAbstractFileByPath(root); if (folder) await app.vault.trash(folder, true);
  }
})()

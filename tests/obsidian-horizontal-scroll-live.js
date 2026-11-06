// Run through Obsidian CLI eval. Check each scroll commit, then its settled
// bitmap, using isolated workbooks and restoring the user's original pane.
(async () => {
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (value, message) => { if (!value) throw Error(message); };
  const prior = app.workspace.activeLeaf, root = `Sheetian horizontal scroll checks ${Date.now()}`;
  const files = [], results = [];
  await app.vault.createFolder(root);
  try {
    for (const [zoom, frozen] of [[1, false], [1.3, false], [1.3, true]]) {
      const celldata = [], colors = ['#ffe0e0', '#e0ffe0', '#e0e0ff', '#ffffe0'];
      for (let r = 0; r < 60; r++) for (let c = 0; c < 40; c++) {
        celldata.push({ r, c, v: { v: `Column ${c + 1}, row ${r + 1}`, bg: colors[c % colors.length] } });
      }
      const source = JSON.stringify([{ id: 'scroll', name: 'Horizontal scroll', row: 100, column: 40, zoomRatio: zoom,
        config: { columnlen: { 0: 19, 1: 546, 2: 60, 3: 60, 4: 292, 5: 267 }, colhidden: { 8: 0 } }, celldata }]);
      const file = await app.vault.create(`${root}/${zoom}-${frozen}.sheet`, source); files.push(file);
      const leaf = app.workspace.getLeaf(true);
      await leaf.openFile(file); app.workspace.setActiveLeaf(leaf, { focus: true }); await wait(250);
      const view = leaf.view;
      if (frozen) { view.workbook.current.freeze('both', { row: 1, column: 1 }); await wait(100); }
      const content = view.getViewData(), bar = view.contentEl.querySelector('.luckysheet-scrollbar-x');
      const area = view.contentEl.querySelector('.fortune-cell-area'), header = view.contentEl.querySelector('.fortune-col-header');
      const canvas = view.contentEl.querySelector('.fortune-sheet-canvas'), ctx = canvas.getContext('2d');
      const bitmap = () => require('crypto').createHash('sha256').update(ctx.getImageData(0, 0, canvas.width, canvas.height).data).digest('hex');
      let maximumDrift = 0, changedBitmaps = 0, previous = bitmap();
      for (const requested of [145.5, 725, 40, 1200.5, 90, bar.scrollWidth, 330.5, 0]) {
        bar.scrollLeft = requested;
        // Dispatch the same event as a native drag without waiting for React
        // or a second frame to hide a stale canvas/selection/header commit.
        bar.dispatchEvent(new Event('scroll'));
        // The engine's header is 1.5px wider than the cell viewport. Its
        // native scroll range therefore clamps slightly earlier at the end.
        const headerPosition = Math.min(bar.scrollLeft, Math.max(0, header.scrollWidth - header.getBoundingClientRect().width));
        const drift = Math.max(Math.abs(bar.scrollLeft - area.scrollLeft), Math.abs(headerPosition - header.scrollLeft));
        maximumDrift = Math.max(maximumDrift, drift);
        assert(drift < 1 / window.devicePixelRatio, `Scroll layers diverged by ${drift}px, zoom ${zoom}, frozen ${frozen}, requested ${requested}, bar ${bar.scrollLeft}, area ${area.scrollLeft}, header ${header.scrollLeft}`);
        const committed = bitmap();
        if (committed !== previous) changedBitmaps++;
        await wait(35);
        assert(bitmap() === committed, `Canvas repainted after the scroll commit, zoom ${zoom}, frozen ${frozen}`);
        previous = committed;
      }
      assert(changedBitmaps === 8, 'Scrollbar movement did not repaint the grid');
      assert(view.getViewData() === content, 'Scrolling changed workbook contents or undo history');
      results.push({ zoom, frozen, positions: 8, maximumDrift, settledBitmapMatchesCommit: true, contentsUnchanged: true });
      await leaf.detach(); await wait(80);
      if (!frozen) assert(await app.vault.read(file) === source, 'Scrolling rewrote an untouched workbook');
    }
    return JSON.stringify({ passed: results.length, results });
  } finally {
    for (const leaf of app.workspace.getLeavesOfType('spreadsheet-view')) if (leaf.view.file?.parent?.path === root) await leaf.detach();
    await wait(100);
    if (prior) app.workspace.setActiveLeaf(prior, { focus: true });
    for (const file of files) if (app.vault.getAbstractFileByPath(file.path)) await app.vault.trash(file, true);
    const folder = app.vault.getAbstractFileByPath(root); if (folder) await app.vault.trash(folder, true);
  }
})()

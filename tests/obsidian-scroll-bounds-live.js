// Run through Obsidian CLI eval. Verify the end of a finite sheet and the
// surrounding Obsidian view using isolated files, including narrow panes.
(async () => {
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (value, message) => { if (!value) throw Error(message); };
  const prior = app.workspace.activeLeaf, root = `Sheetian scroll bounds checks ${Date.now()}`;
  const results = [], files = [];
  await app.vault.createFolder(root);
  try {
    for (const [zoom, frozen] of [[1, false], [0.75, false], [1.3, true]]) {
      const celldata = [];
      for (let r = 0; r < 100; r++) for (let c = 0; c < 26; c++) {
        celldata.push({ r, c, v: { v: `R${r + 1} C${c + 1}`, bg: c === 25 ? '#dceefe' : '#ffffff' } });
      }
      const source = JSON.stringify([{ id: 'bounds', name: 'Scroll bounds', row: 100, column: 26, zoomRatio: zoom,
        config: { columnlen: { 0: 19, 1: 546, 2: 60, 3: 60, 4: 292, 5: 267 }, colhidden: { 8: 0 } }, celldata }]);
      const file = await app.vault.create(`${root}/${zoom}-${frozen}.sheet`, source); files.push(file);
      const leaf = app.workspace.getLeaf(true);
      try {
        await leaf.openFile(file); app.workspace.setActiveLeaf(leaf, { focus: true }); await wait(250);
        const view = leaf.view;
        if (frozen) { view.workbook.current.freeze('both', { row: 1, column: 1 }); await wait(80); }
        const before = view.getViewData(), bar = view.contentEl.querySelector('.luckysheet-scrollbar-x');
        // Hidden toolbar tooltips can extend beyond a narrow pane. They must
        // never turn Obsidian's outer view into another scroll container.
        const tooltip = document.createElement('div');
        tooltip.style.cssText = 'position:absolute;left:100%;width:150px;height:20px;visibility:hidden';
        view.contentEl.querySelector('.fortune-toolbar').append(tooltip);
        const widths = view.workbook.current.getColumnWidth(Array.from({ length: 26 }, (_, c) => c));
        const sheet = view.workbook.current.getSheet();
        const totalWidth = Object.entries(widths).reduce((sum, [column, width]) =>
          sum + (sheet.config.colhidden?.[column] != null ? 0 : Math.round((width + 1) * zoom)), 0);
        bar.scrollLeft = bar.scrollWidth; bar.dispatchEvent(new Event('scroll')); await wait(50);
        const rightGap = bar.scrollLeft + bar.clientWidth - totalWidth;
        const overflow = getComputedStyle(view.contentEl).overflowX;
        const outcome = { zoom, frozen, rightGap, overflow, outerScrollLeft: view.contentEl.scrollLeft };
        results.push(outcome);
        assert(Math.abs(rightGap) < 1, 'Horizontal scrolling extends beyond the final column: ' + JSON.stringify(outcome));
        assert(overflow === 'clip' && view.contentEl.scrollLeft === 0,
          'Obsidian adds a second horizontal scrollbar: ' + JSON.stringify(outcome));
        assert(view.getViewData() === before, 'Scroll bounds checks changed workbook contents');
        tooltip.remove();
      } finally { await leaf.detach(); await wait(80); }
    }
    return JSON.stringify({ passed: results.length * 3, results });
  } finally {
    if (prior) app.workspace.setActiveLeaf(prior, { focus: true });
    for (const file of files) if (app.vault.getAbstractFileByPath(file.path)) await app.vault.trash(file, true);
    const folder = app.vault.getAbstractFileByPath(root); if (folder) await app.vault.trash(folder, true);
  }
})()

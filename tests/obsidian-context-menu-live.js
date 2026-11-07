// Run through Obsidian CLI eval. Menus must fit their pane, keep the workbook
// stationary while taking focus, and retain numeric resize undo/redo.
(async () => {
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (value, message) => { if (!value) throw Error(message); };
  const prior = app.workspace.activeLeaf, root = `Sheetian context menu checks ${Date.now()}`;
  const results = [];
  await app.vault.createFolder(root);
  const source = JSON.stringify([{ id: 'menus', name: 'Menu layout', row: 100, column: 26,
    celldata: [{ r: 0, c: 25, v: { v: 'Last column' } }] }]);
  const file = await app.vault.create(root + '/menus.sheet', source), leaf = app.workspace.getLeaf(true);
  try {
    await leaf.openFile(file); app.workspace.setActiveLeaf(leaf, { focus: true }); await wait(350);
    const view = leaf.view, workbook = view.contentEl.querySelector('.fortune-container');
    const menu = () => view.contentEl.querySelector('.fortune-context-menu');
    const close = async () => {
      view.contentEl.querySelector('.fortune-popover-backdrop')?.dispatchEvent(new MouseEvent('mousedown', { button: 0, bubbles: true }));
      await wait(50);
    };
    const check = label => {
      const pane = workbook.getBoundingClientRect(), rect = menu().getBoundingClientRect();
      const outcome = { label, paneWidth: pane.width, menuWidth: rect.width, menuHeight: rect.height,
        outerScroll: [view.contentEl.scrollLeft, view.contentEl.scrollTop] };
      assert(rect.left >= pane.left - 1 && rect.right <= pane.right + 1 && rect.top >= pane.top - 1 && rect.bottom <= pane.bottom + 1,
        'Menu escaped the spreadsheet pane: ' + JSON.stringify(outcome));
      assert(outcome.outerScroll.every(value => value === 0), 'Menu shifted the entire workbook: ' + JSON.stringify(outcome));
      assert(menu().contains(document.activeElement), 'Menu focused a different workbook');
      const before = pane.toJSON();
      menu().querySelector('input')?.focus();
      view.contentEl.scrollLeft = 200; view.contentEl.scrollTop = 200;
      assert(view.contentEl.scrollLeft === 0 && view.contentEl.scrollTop === 0 && workbook.getBoundingClientRect().x === before.x,
        'Outer pane still permits focus-driven scrolling');
      results.push(outcome);
    };
    for (const narrow of [false, true]) {
      view.contentEl.style.width = `${narrow ? 280 : Math.min(818, view.containerEl.clientWidth)}px`;
      view.contentEl.style.height = `${narrow ? 380 : view.contentEl.clientHeight}px`;
      view.contentEl.style.flex = 'none'; await wait(150);
      const xbar = view.contentEl.querySelector('.luckysheet-scrollbar-x');
      xbar.scrollLeft = xbar.scrollWidth; xbar.dispatchEvent(new Event('scroll')); await wait(80);
      view.workbook.current.setSelection([{ row: [0, 99], column: [25, 25], column_select: true }]); await wait(80);
      const header = view.contentEl.querySelector('.fortune-col-header'), rect = header.getBoundingClientRect();
      header.dispatchEvent(new MouseEvent('mousemove', { clientX: rect.right - 25, clientY: rect.top + 10, bubbles: true })); await wait(60);
      const arrow = header.querySelector('.header-arrow'), arrowRect = arrow.getBoundingClientRect();
      arrow.dispatchEvent(new MouseEvent('click', { clientX: arrowRect.left + 6, clientY: arrowRect.top + 6, bubbles: true })); await wait(100);
      check(narrow ? 'narrow header arrow' : 'header arrow');
      if (!narrow) {
        if (window.__sheetianMenuScreenshotPath) {
          const crop = view.contentEl.getBoundingClientRect();
          const screenshot = await require('electron').remote.getCurrentWindow().webContents.capturePage({
            x: Math.floor(crop.x), y: Math.floor(crop.y), width: Math.ceil(crop.width), height: Math.ceil(crop.height),
          });
          require('fs').writeFileSync(window.__sheetianMenuScreenshotPath, screenshot.toPNG());
        }
        const before = view.workbook.current.getColumnWidth([25])[25];
        const widthItem = [...menu().querySelectorAll('.luckysheet-cols-menuitem')].find(el => /Column\s*Width/i.test(el.textContent));
        const input = widthItem?.querySelector('input'); assert(input, 'Column width input is missing');
        input.value = '160'; input.closest('.luckysheet-cols-menuitem').click(); await wait(100);
        assert(view.workbook.current.getColumnWidth([25])[25] === 160, 'Numeric column resize failed');
        view.workbook.current.handleUndo(); await wait(80);
        assert(view.workbook.current.getColumnWidth([25])[25] === before, 'Numeric column resize undo failed');
        view.workbook.current.handleRedo(); await wait(80);
        assert(view.workbook.current.getColumnWidth([25])[25] === 160, 'Numeric column resize redo failed');
        assert(view.contentEl.scrollLeft === 0, 'Numeric resizing shifted the workbook');
      }
      await close();
      const area = view.contentEl.querySelector('.fortune-cell-area'), cellRect = area.getBoundingClientRect();
      area.dispatchEvent(new MouseEvent('contextmenu', { clientX: cellRect.right - 10, clientY: cellRect.bottom - 10,
        button: 2, bubbles: true, cancelable: true })); await wait(100);
      check(narrow ? 'narrow bottom-right cell' : 'bottom-right cell');
      await close();
    }
    assert(view.workbook.current.getSheet().data[0][25].v === 'Last column', 'Menu actions changed cell data');
    return JSON.stringify({ passed: results.length + 4, results });
  } finally {
    await leaf.detach(); await wait(80); if (prior) app.workspace.setActiveLeaf(prior, { focus: true });
    await app.vault.trash(file, true); await app.vault.trash(app.vault.getAbstractFileByPath(root), true);
  }
})()

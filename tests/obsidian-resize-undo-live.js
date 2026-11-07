// Run through Obsidian CLI eval. Exercise native header drags and compare
// saved dimensions with the visible grid before/after undo and redo.
(async () => {
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (value, message) => { if (!value) throw Error(message); };
  const prior = app.workspace.activeLeaf, root = `Sheetian resize undo checks ${Date.now()}`;
  const results = [], files = [];
  await app.vault.createFolder(root);
  try {
    for (const zoom of [1, 1.3]) {
      const source = JSON.stringify([{ id: 'resize', name: 'Resize undo', row: 100, column: 26, zoomRatio: zoom,
        config: { columnlen: { 0: 73, 2: 73, 25: 100 }, rowlen: { 0: 19, 2: 19 } },
        celldata: [
          { r: 0, c: 0, v: { v: 'Text long enough to exercise AutoFit and its undo history' } },
          { r: 4, c: 4, v: { v: 'Long column contents for fitting from the Tools menu' } },
          { r: 4, c: 5, v: { v: 'A second column to fit in the same undo step' } },
          ...[4, 5, 6].map(r => ({ r, c: 0, v: { v: 'First line\nSecond line\nThird line' } })),
        ] }]);
      const file = await app.vault.create(`${root}/${zoom}.sheet`, source); files.push(file);
      const leaf = app.workspace.getLeaf(true);
      try {
        await leaf.openFile(file); app.workspace.setActiveLeaf(leaf, { focus: true }); await wait(300);
        const view = leaf.view;
        const snapshot = (axis, indices) => ({
          sizes: axis === 'column' ? view.workbook.current.getColumnWidth(indices) : view.workbook.current.getRowHeight(indices),
          extent: parseFloat(view.contentEl.querySelector(`.luckysheet-scrollbar-${axis === 'column' ? 'x' : 'y'} > div`).style[axis === 'column' ? 'width' : 'height']),
        });
        const key = async (code, shiftKey = false, ctrlKey = false) => {
          const overlay = view.contentEl.querySelector('.fortune-sheet-overlay'); overlay.focus();
          overlay.dispatchEvent(new KeyboardEvent('keydown', { key: code === 'KeyZ' ? 'z' : 'y', code,
            keyCode: code === 'KeyZ' ? 90 : 89, metaKey: !ctrlKey, ctrlKey, shiftKey, bubbles: true, cancelable: true }));
          await wait(100);
        };
        const choose = async (menu, title) => {
          const button = [...view.contentEl.querySelectorAll('.sheetian-menu-button')].find(el => el.textContent === menu);
          button.click(); await wait(50);
          const item = [...document.querySelectorAll('.menu-item')].find(el => el.querySelector('.menu-item-title')?.textContent === title);
          assert(item, `Missing ${menu} → ${title}`); item.click(); await wait(100);
        };
        const boundary = async (axis, index) => {
          const column = axis === 'column';
          const bar = view.contentEl.querySelector(`.luckysheet-scrollbar-${column ? 'x' : 'y'}`);
          bar[column ? 'scrollLeft' : 'scrollTop'] = column && index === 25 ? bar.scrollWidth : 0;
          bar.dispatchEvent(new Event('scroll')); await wait(80);
          const header = view.contentEl.querySelector(column ? '.fortune-col-header' : '.fortune-row-header');
          const rect = header.getBoundingClientRect();
          const sizes = column ? view.workbook.current.getColumnWidth(Array.from({ length: index + 1 }, (_, i) => i))
            : view.workbook.current.getRowHeight(Array.from({ length: index + 1 }, (_, i) => i));
          const edge = Object.values(sizes).reduce((sum, size) => sum + Math.round((size + 1) * zoom), 0);
          const position = edge - header[column ? 'scrollLeft' : 'scrollTop'] - 3;
          const x = column ? rect.left + position : rect.left + rect.width / 2;
          const y = column ? rect.top + rect.height / 2 : rect.top + position;
          const mouse = (target, type, movement = 0) => target.dispatchEvent(new MouseEvent(type, {
            clientX: x + (column ? movement : 0), clientY: y + (column ? 0 : movement),
            button: 0, buttons: type === 'mouseup' ? 0 : 1, bubbles: true, cancelable: true,
          }));
          mouse(header, 'mousemove'); await wait(60);
          const handle = header.querySelector(column ? '.fortune-cols-change-size' : '.fortune-rows-change-size');
          assert(Number(handle.dataset[column ? 'sheetianColumn' : 'sheetianRow']) === index, 'Wrong resize target');
          const handleRect = handle.getBoundingClientRect();
          assert(column ? x >= handleRect.left && x <= handleRect.right : y >= handleRect.top && y <= handleRect.bottom,
            'Resize handle did not follow the updated header boundary');
          return { handle, mouse };
        };
        const drag = async (axis, index, delta) => {
          const { handle, mouse } = await boundary(axis, index);
          mouse(handle, 'mousedown'); await wait(30);
          mouse(document, 'mousemove', delta); await wait(30);
          mouse(document, 'mouseup', delta); await wait(80);
        };
        const roundTrip = async (action, axis, indices, resize, menu = false, ctrl = false) => {
          const before = snapshot(axis, indices);
          await resize();
          const resized = snapshot(axis, indices);
          // A click on the resized boundary must not add a second undo step.
          if (action.startsWith('drag')) await drag(axis, indices.at(-1), 0);
          if (menu) await choose('Edit', 'Undo'); else await key('KeyZ', false, ctrl);
          const undone = snapshot(axis, indices);
          if (menu) await choose('Edit', 'Redo'); else await key(ctrl ? 'KeyY' : 'KeyZ', !ctrl, ctrl);
          const redone = snapshot(axis, indices);
          const result = { zoom, action, axis, indices, before, resized, undone, redone };
          results.push(result);
          assert(JSON.stringify(resized) !== JSON.stringify(before), 'Resize had no effect: ' + JSON.stringify(result));
          assert(JSON.stringify(undone) === JSON.stringify(before), 'One undo did not restore saved and visible dimensions: ' + JSON.stringify(result));
          assert(JSON.stringify(redone) === JSON.stringify(resized), 'One redo did not restore saved and visible dimensions: ' + JSON.stringify(result));
          if (action.startsWith('drag')) for (const index of indices) {
            assert(resized.sizes[index] > before.sizes[index], 'Outward drag did not expand the header: ' + JSON.stringify(result));
          }
        };
        for (const [axis, index, delta] of [['column', 25, 70], ['column', 2, 55], ['row', 2, 35]]) {
          await roundTrip('drag', axis, [index], () => drag(axis, index, delta), false, axis === 'row');
        }
        for (const axis of ['column', 'row']) {
          view.workbook.current.setSelection([{ row: axis === 'row' ? [1, 2] : [0, 99],
            column: axis === 'column' ? [1, 2] : [0, 25], [axis === 'column' ? 'column_select' : 'row_select']: true }]);
          await wait(50);
          await roundTrip('drag selected headers', axis, [1, 2], () => drag(axis, 2, 50), true);
        }
        for (const [axis, index] of [['column', 0], ['row', 4]]) {
          view.workbook.current.setSelection([{ row: [index, index], column: [index, index] }]); await wait(50);
          await roundTrip('double-click AutoFit', axis, [index], async () => {
            const { handle, mouse } = await boundary(axis, index);
            for (let n = 0; n < 2; n++) {
              mouse(handle, 'mousedown'); await wait(30); mouse(document, 'mouseup'); await wait(30);
            }
            mouse(handle, 'dblclick'); await wait(100);
          });
        }
        for (const axis of ['column', 'row']) {
          const indices = axis === 'column' ? [4, 5] : [5, 6];
          view.workbook.current.setSelection([{ row: axis === 'row' ? indices : [0, 99],
            column: axis === 'column' ? indices : [0, 25] }]); await wait(50);
          await roundTrip('Tools AutoFit', axis, indices,
            () => choose('Tools', axis === 'column' ? 'Fit selected columns' : 'Fit selected rows'), true);
        }
      } finally { await leaf.detach(); await wait(80); }
    }
    window.__sheetianResizeUndoOutcome = results;
    return JSON.stringify({ passed: results.length * 3, results: results.map(({ zoom, action, axis, indices }) => ({ zoom, action, axis, indices })) });
  } finally {
    if (prior) app.workspace.setActiveLeaf(prior, { focus: true });
    for (const file of files) if (app.vault.getAbstractFileByPath(file.path)) await app.vault.trash(file, true);
    const folder = app.vault.getAbstractFileByPath(root); if (folder) await app.vault.trash(folder, true);
  }
})()

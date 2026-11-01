// Run via Obsidian CLI eval in a desktop vault. Uses temporary workbooks and
// preserves the user's clipboard. Call window.__sheetianClipboardLive.cleanup()
// after capturing the review screenshot; no test files remain afterward.
(async () => {
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const assert = (value, message) => { if (!value) throw Error(message); };
  const clipboard = require('electron').clipboard;
  const previousClipboard = { text: clipboard.readText(), html: clipboard.readHTML(), rtf: clipboard.readRTF(), image: clipboard.readImage() };
  const previousLeaf = app.workspace.activeLeaf;
  const root = `Sheetian clipboard checks ${Date.now()}`;
  const leaves = [], results = [];
  let view;
  const api = () => view.workbook.current;
  const sheet = () => api().getSheet();
  const cell = (r, c) => sheet().data[r][c];
  const compact = formula => formula?.replace(/\s/g, '');
  const cleanup = async () => {
    for (const menu of document.querySelectorAll('.menu')) if (menu.win?.hide) menu.win.hide();
    for (const leaf of leaves.slice().reverse()) if (leaf.parent) await leaf.detach();
    await wait(100);
    clipboard.write(previousClipboard);
    if (previousLeaf?.parent) app.workspace.setActiveLeaf(previousLeaf, { focus: true });
    for (const file of app.vault.getAbstractFileByPath(root)?.children?.slice() ?? []) await app.vault.trash(file, true);
    const folder = app.vault.getAbstractFileByPath(root);
    if (folder) await app.vault.trash(folder, true);
    return { removedTemporaryFiles: !app.vault.getAbstractFileByPath(root), restoredClipboard: clipboard.readText() === previousClipboard.text };
  };
  const state = window.__sheetianClipboardLive = { root, results, cleanup, api, select: null, choose: null };
  const open = async file => {
    const leaf = app.workspace.getLeaf(true); leaves.push(leaf);
    await leaf.openFile(file); app.workspace.setActiveLeaf(leaf, { focus: true });
    for (let n = 0; n < 100 && !leaf.view.workbook?.current; n++) await wait(30);
    assert(leaf.view.workbook?.current, 'Workbook did not initialize');
    view = leaf.view; await wait(180); return leaf;
  };
  const select = state.select = async (r, c, r2 = r, c2 = c) => {
    api().setSelection([{ row: [r, r2], column: [c, c2], row_focus: r, column_focus: c }]);
    await wait(70);
  };
  const choose = state.choose = async title => {
    const edit = [...view.contentEl.querySelectorAll('.sheetian-menu-button')].find(el => el.textContent === 'Edit');
    edit.dispatchEvent(new MouseEvent('mousedown', { button: 0, bubbles: true, cancelable: true })); edit.click(); await wait(50);
    const item = [...document.querySelectorAll('.menu-item')].find(el => el.querySelector('.menu-item-title')?.textContent === title);
    assert(item && !item.classList.contains('is-disabled'), `Missing/enabled Edit item: ${title}`);
    item.click(); await wait(130);
  };
  const run = async (name, test) => {
    try { await test(); results.push({ name, passed: true }); }
    catch (error) { results.push({ name, passed: false, error: error.message }); }
  };
  const sourceCell = { f: '=SUM(\nA2:A3\n)+$D$1', v: 11, m: '$11.00', ct: { fa: '$0.00', t: 'n' }, bl: 1, fs: 14, fc: '#125faa', bg: '#dceefe', ht: 2, tb: '2' };
  const entry = (r, c, v) => ({ r, c, v });
  const contents = [{ id: 'clipboard', name: 'Clipboard source', status: 1, row: 40, column: 16,
    calcChain: [{ r: 0, c: 0, id: 'clipboard' }, { r: 5, c: 0, id: 'clipboard' }, { r: 17, c: 4, id: 'clipboard' }],
    config: { borderInfo: [{ rangeType: 'cell', value: { row_index: 0, col_index: 0, b: { style: '1', color: '#123456' } } }], merge: { '0_7': { r: 0, c: 7, rs: 1, cs: 2 } } },
    celldata: [entry(0, 0, sourceCell), entry(1, 0, { v: 2 }), entry(2, 0, { v: 4 }), entry(1, 1, { v: 7 }), entry(2, 1, { v: 9 }),
      entry(0, 2, { v: 'constant' }), entry(0, 3, { v: 5 }), entry(1, 5, { v: 8 }), entry(2, 5, { v: 9 }),
      entry(5, 0, { f: '=99', v: 99, it: 1, bg: '#ffeebb', ct: { fa: '0.00', t: 'n' } }),
      entry(8, 0, { v: 'old A9', bg: '#ffeebb' }), entry(8, 1, { v: 'old B9', bg: '#ffeebb' }), entry(8, 2, { v: 'keep constant' }),
      entry(9, 0, { v: 3 }), entry(10, 0, { v: 4 }), entry(9, 1, { v: 6 }), entry(10, 1, { v: 8 }),
      entry(12, 1, { v: 4 }), entry(13, 1, { v: 5 }), entry(17, 4, { f: '=$D$1*2', v: 10, bl: 1, bg: '#ddeeff' }),
      entry(0, 7, { v: 'Merged cell', mc: { r: 0, c: 7, rs: 1, cs: 2 } }), entry(0, 8, { mc: { r: 0, c: 7 } }),
      ...[0, false, '', '00123', '=literal'].map((v, c) => entry(20, c, { v, ct: { fa: '@', t: 's' } })),
    ] }, { id: 'second', name: 'Second worksheet', status: 0, row: 30, column: 16,
    celldata: [entry(1, 0, { v: 8 }), entry(2, 0, { v: 12 }), entry(0, 3, { v: 1 })] }];
  try {
    await app.vault.createFolder(root);
    const sourceFile = await app.vault.create(`${root}/clipboard-source.sheet`, JSON.stringify(contents));
    const sourceLeaf = await open(sourceFile); state.sourceLeaf = sourceLeaf;

    await run('Normal copy/paste: formulas, relative/absolute references, fonts, colors, wrap, number format and borders', async () => {
      await select(0, 0); await choose('Copy');
      assert(clipboard.readText() === '$11.00' && clipboard.readHTML().includes('data-sheetian-clipboard'), 'Desktop clipboard lacks text or the cell snapshot');
      await select(0, 1); await choose('Paste');
      const value = cell(0, 1);
      assert(compact(value.f) === '=SUM(B2:B3)+$D$1' && value.v === 21, 'Pasted formula/reference/result mismatch: ' + JSON.stringify(value));
      for (const key of ['bl', 'fs', 'fc', 'bg', 'ht', 'tb']) assert(value[key] === sourceCell[key], 'Lost style ' + key);
      assert(value.ct.fa === '$0.00' && value.m === '$21.00', 'Number format was lost');
      assert(sheet().config.borderInfo.some(b => b.value?.row_index === 0 && b.value?.col_index === 1 && b.value?.b?.color === '#123456'), 'Bottom border was lost');
    });
    await run('Paste values only: raw result, destination styles and removal of old formulas', async () => {
      await select(0, 0); await choose('Copy'); await select(5, 0); await choose('Paste values only');
      assert(cell(5, 0).v === 11 && !cell(5, 0).f && cell(5, 0).bg === '#ffeebb' && cell(5, 0).it === 1 && cell(5, 0).ct.fa === '0.00', 'Values-only changed styles or kept a formula');
      assert(!sheet().calcChain.some(e => e.r === 5 && e.c === 0), 'Old formula dependency remains');
    });
    await run('Paste formulas only: destination styles and constants preserved; new formulas recalculate', async () => {
      await select(0, 0, 0, 2); await choose('Copy'); await select(8, 0); await choose('Paste formulas only');
      assert(compact(cell(8, 0).f) === '=SUM(A10:A11)+$D$1' && cell(8, 0).v === 12, 'Formula-only result/reference mismatch');
      assert(cell(8, 0).bg === '#ffeebb' && cell(8, 1).v === 19 && cell(8, 2).v === 'keep constant', 'Formula-only changed styles/constants');
      // The engine's low-level value setter does not recalculate dependents.
      // Commit through the visible editor, exactly as an ordinary edit does.
      await select(9, 0);
      const editor = view.contentEl.querySelector('.fortune-fx-input'); editor.focus(); await wait(100);
      editor.innerText = '10'; editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true, cancelable: true }));
      await wait(120); assert(cell(8, 0).v === 19, 'Pasted formula did not recalculate');
    });
    await run('Paste formatting only and Undo/Redo leave values unchanged', async () => {
      await select(0, 0); await choose('Copy'); await select(5, 0); const before = JSON.stringify(cell(5, 0));
      await choose('Paste formatting only'); const after = JSON.stringify(cell(5, 0));
      assert(cell(5, 0).v === 11 && !cell(5, 0).f && cell(5, 0).bg === sourceCell.bg && cell(5, 0).bl === 1 && !cell(5, 0).it, 'Formatting-only changed content or missed styles');
      await choose('Undo'); assert(JSON.stringify(cell(5, 0)) === before, 'Undo did not restore exact cell');
      await choose('Redo'); assert(JSON.stringify(cell(5, 0)) === after, 'Redo did not restore formatted cell');
    });
    await run('Copy values only and Copy formulas only control normal paste', async () => {
      await select(0, 0); await choose('Copy values only'); assert(clipboard.readText() === '$11.00', 'Copy values text mismatch');
      await select(11, 0); await choose('Paste'); assert(cell(11, 0).v === 11 && !cell(11, 0).f && !cell(11, 0).bg, 'Copy values then Paste copied a formula or styles');
      await select(0, 0); await choose('Copy formulas only'); assert(compact(clipboard.readText()) === '"=SUM(A2:A3)+$D$1"', 'Copy formulas did not supply quoted multiline formula text');
      await select(11, 1); await choose('Paste'); assert(compact(cell(11, 1).f) === '=SUM(B13:B14)+$D$1' && cell(11, 1).v === 14, 'Copy formulas then Paste failed');
    });
    await run('Values-only preserves zero, false, empty text, numeric strings and literal equals signs', async () => {
      await select(20, 0, 20, 4); await choose('Copy'); await select(23, 0); await choose('Paste values only');
      [0, false, '', '00123', '=literal'].forEach((value, c) => assert(cell(23, c).v === value && !cell(23, c).f, 'Typed value changed at column ' + c));
    });
    await run('Copied cells are snapshots even after the source changes', async () => {
      await select(0, 2); await choose('Copy'); api().setCellValue(0, 2, 'edited after copy'); await wait(90);
      await select(4, 3); await choose('Paste'); assert(cell(4, 3).v === 'constant', 'Paste reread the live source');
      api().setCellValue(0, 2, 'constant'); await wait(70);
    });
    await run('Merged cells retain their span and relocated merge metadata', async () => {
      await select(0, 7, 0, 8); await choose('Copy'); await select(4, 7); await choose('Paste');
      assert(cell(4, 7).v === 'Merged cell' && sheet().config.merge['4_7']?.cs === 2 && cell(4, 8).mc.r === 4 && cell(4, 8).mc.c === 7, 'Merged cell paste failed');
    });
    await run('Excel HTML: headers, class styles, borders and merged cells', async () => {
      clipboard.write({ html: '<style>.xl{background-color:#cceeaa;color:#224466;font-weight:bold;text-align:right;border-bottom:2px solid #123456}</style><table><tr><th class="xl">Excel header</th><td>42</td></tr><tr><td rowspan="2" colspan="2">Merged HTML</td></tr><tr></tr></table>', text: 'Excel header\t42\nMerged HTML' });
      await select(8, 7); await choose('Paste');
      assert(cell(8, 7).v === 'Excel header' && cell(8, 7).bl === 1 && cell(8, 8).v === 42 && sheet().config.merge['9_7']?.rs === 2, 'HTML values/styles/merge failed');
      assert(sheet().config.borderInfo.some(b => b.value?.row_index === 8 && b.value?.col_index === 7 && b.value?.b?.color), 'HTML border missing');
    });
    await run('TSV paste: formulas, quoted multiline cells, tabs and trailing newline', async () => {
      clipboard.writeText('2\t=H13*3\n"line one\nline two"\t"a\tb"\n'); await select(12, 7); await choose('Paste');
      assert(cell(12, 7).v === 2 && cell(12, 8).f === '=H13*3' && cell(12, 8).v === 6 && cell(13, 7).v === 'line one\nline two' && cell(13, 8).v === 'a\tb' && !cell(14, 7)?.v, 'TSV formula or quoted cell boundaries failed');
    });
    await run('Copy between worksheet tabs preserves styles and recalculates formulas', async () => {
      await select(0, 0); await choose('Copy'); api().activateSheet({ id: 'second' }); await wait(180);
      await select(0, 0); await choose('Paste'); assert(cell(0, 0).f === sourceCell.f && cell(0, 0).v === 21 && cell(0, 0).bg === sourceCell.bg, 'Cross-worksheet paste failed');
      api().activateSheet({ id: 'clipboard' }); await wait(180);
    });
    await run('Copy between workbooks uses the actual desktop clipboard; inactive panes cannot consume the paste', async () => {
      await select(0, 0); await choose('Copy'); const sourceView = view, inactiveBefore = JSON.stringify(api().getAllSheets());
      const targetFile = await app.vault.create(`${root}/clipboard-target.sheet`, JSON.stringify([{ id: 'clipboard', name: 'Other workbook', status: 1, row: 30, column: 16, celldata: [entry(1, 0, { v: 10 }), entry(2, 0, { v: 20 }), entry(0, 3, { v: 1 })] }]));
      const targetLeaf = await open(targetFile); await select(0, 0);
      const overlay = view.contentEl.querySelector('.fortune-sheet-overlay'); overlay.focus();
      const data = new DataTransfer(); data.setData('text/html', clipboard.readHTML()); data.setData('text/plain', clipboard.readText());
      overlay.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true })); await wait(150);
      assert(cell(0, 0).f === sourceCell.f && cell(0, 0).v === 31 && cell(0, 0).bg === sourceCell.bg, 'Cross-workbook clipboard paste failed');
      assert(JSON.stringify(sourceView.workbook.current.getAllSheets()) === inactiveBefore, 'Inactive workbook consumed the paste');
      await targetLeaf.detach(); app.workspace.setActiveLeaf(sourceLeaf, { focus: true }); view = sourceView; await wait(100);
    });
    await run('Keyboard Copy, normal Paste and Paste values only use the enhanced clipboard', async () => {
      await select(0, 0); let overlay = view.contentEl.querySelector('.fortune-sheet-overlay'); overlay.focus();
      overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'c', code: 'KeyC', keyCode: 67, metaKey: true, bubbles: true, cancelable: true })); await wait(90);
      await select(0, 5); overlay.focus(); const transfer = new DataTransfer(); transfer.setData('text/html', clipboard.readHTML()); transfer.setData('text/plain', clipboard.readText());
      overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'v', code: 'KeyV', keyCode: 86, metaKey: true, bubbles: true, cancelable: true }));
      overlay.dispatchEvent(new ClipboardEvent('paste', { clipboardData: transfer, bubbles: true, cancelable: true })); await wait(120);
      assert(compact(cell(0, 5).f) === '=SUM(F2:F3)+$D$1' && cell(0, 5).v === 22 && cell(0, 5).bg === sourceCell.bg, 'Keyboard normal paste failed');
      await select(16, 5); overlay.focus(); overlay.dispatchEvent(new KeyboardEvent('keydown', { key: 'v', code: 'KeyV', keyCode: 86, metaKey: true, shiftKey: true, bubbles: true, cancelable: true })); await wait(120);
      assert(cell(16, 5).v === 11 && !cell(16, 5).f && !cell(16, 5).bg, 'Keyboard paste values failed');
    });
    await run('Cut/Paste moves formulas and styling once and supports Undo/Redo', async () => {
      await select(17, 4); await choose('Cut'); assert(cell(17, 4)?.f === '=$D$1*2', 'Cut removed the source before paste');
      await select(17, 5); await choose('Paste'); assert(!cell(17, 4)?.v && cell(17, 5).f === '=$D$1*2' && cell(17, 5).v === 10 && cell(17, 5).bl === 1, 'Cut/Paste lost content or styling');
      await choose('Undo'); assert(cell(17, 4)?.f === '=$D$1*2' && !cell(17, 5)?.v, 'Undo cut failed');
      await choose('Redo'); assert(!cell(17, 4)?.v && cell(17, 5).f === '=$D$1*2', 'Redo cut failed');
    });
    await run('Right-click menu exposes every clipboard option and pastes values correctly', async () => {
      const context = async (r, c, title) => {
        await select(r, c);
        const area = view.contentEl.querySelector('.fortune-cell-area');
        const rect = view.contentEl.querySelector('.luckysheet-cell-selected-focus').getBoundingClientRect();
        area.dispatchEvent(new MouseEvent('contextmenu', { clientX: rect.left + 3, clientY: rect.top + 3, button: 2, bubbles: true, cancelable: true }));
        await wait(80);
        const items = [...view.contentEl.querySelectorAll('.fortune-context-menu .luckysheet-cols-menuitem')];
        for (const expected of ['Cut', 'Copy', 'Copy values only', 'Copy formulas only', 'Paste', 'Paste values only', 'Paste formulas only', 'Paste formatting only']) {
          assert(items.some(item => item.textContent === expected), 'Missing right-click option: ' + expected);
        }
        items.find(item => item.textContent === title).click(); await wait(130);
      };
      await context(0, 0, 'Copy'); await context(3, 0, 'Paste values only');
      assert(cell(3, 0).v === 11 && !cell(3, 0).f && !cell(3, 0).bg, 'Right-click values-only paste failed');
    });
    await run('Saved workbooks retain pasted formulas, formatting and merge metadata', async () => {
      await view.save(); const saved = JSON.parse(await app.vault.read(sourceFile)); const first = saved.find(s => s.id === 'clipboard');
      const pasted = first.celldata.find(e => e.r === 0 && e.c === 1)?.v;
      assert(compact(pasted?.f) === '=SUM(B2:B3)+$D$1' && pasted.bg === sourceCell.bg && first.config.merge['4_7'].cs === 2, 'Saved clipboard edits lost data');
    });
    await select(0, 0);
    state.view = view; state.sourceFile = sourceFile;
    return JSON.stringify({ passed: results.filter(r => r.passed).length, failed: results.filter(r => !r.passed), results });
  } catch (error) {
    await cleanup(); throw error;
  }
})()

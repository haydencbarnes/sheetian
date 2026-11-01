import { test } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { enablePatches, produceWithPatches, applyPatches } from 'immer';
import { patchCore, patchClipboardMenu, patchCutPaste, patchPasteFormulaCalculation } from '../engine-compat.mjs';

const require = createRequire(import.meta.url);
const built = await build({
  stdin: { contents: 'export * from "./clipboard"; export {defaultContext, updateCell, groupValuesRefresh} from "@fortune-sheet/core";', resolveDir: process.cwd(), loader: 'ts' },
  bundle: true, write: false, platform: 'node', format: 'cjs', packages: 'external',
  plugins: [{ name: 'clipboard-core', setup(builder) {
    builder.onResolve({ filter: /^@fortune-sheet\/core$/ }, () => ({ path: require.resolve('@fortune-sheet/core') }));
    builder.onLoad({ filter: /core[\\/]dist[\\/]index\.js$/ }, async args => ({ contents: patchCore(await readFile(args.path, 'utf8')), loader: 'js' }));
  } }],
});
const module = { exports: {} };
new Function('module', 'exports', 'require', built.outputFiles[0].text)(module, module.exports, require);
const { defaultContext, updateCell, groupValuesRefresh, copyCells, pasteCells, readSheetClipboard, toTsv, parseTsv } = module.exports;
const range = (r, c, r2 = r, c2 = c) => ({ row: [r, r2], column: [c, c2], row_focus: r, column_focus: c });
function context(id = 'source') {
  const ctx = defaultContext({});
  ctx.currentSheetId = id;
  ctx.config = {};
  ctx.luckysheetfile = [{ id, name: id, status: 1, data: Array.from({ length: 12 }, () => Array(10).fill(null)), config: ctx.config, calcChain: [] }];
  ctx.luckysheet_select_save = [range(0, 0)];
  return ctx;
}
const data = ctx => ctx.luckysheetfile[0].data;
const copy = (ctx, mode = 'all', cut = false) => { let contents; copyCells(ctx, mode, cut, value => contents = value); return contents; };
const select = (ctx, ...args) => { ctx.luckysheet_select_save = [range(...args)]; };
const editor = text => ({ innerText: text, innerHTML: text, querySelectorAll: () => [] });

test('normal copy is an immutable snapshot of formulas, number format, fonts, colors, wrapping, and borders across workbooks', () => {
  const source = context(), target = context('target');
  const cell = { f: '=A2+$A$3+A$4+$A5', v: 20, m: '$20.00', ct: { fa: '$0.00', t: 'n' }, bl: 1, it: 1, fs: 16, ff: 'Arial', fc: '#112233', bg: '#abcdef', ht: 2, tb: '2' };
  data(source)[0][0] = structuredClone(cell);
  source.config.borderInfo = [{ rangeType: 'cell', value: { row_index: 0, col_index: 0, b: { style: '1', color: '#123456' } } }];
  const contents = copy(source);
  assert.equal(contents.text, '$20.00');
  assert.match(contents.html, /data-sheetian-clipboard/);
  data(source)[0][0] = { v: 'changed after copy' };
  for (let r = 0; r < 6; r++) data(target)[r][0] = { v: r + 1 };
  data(target)[2][1] = { v: 10 };
  select(target, 1, 1);
  pasteCells(target, contents);
  const pasted = data(target)[1][1];
  assert.equal(pasted.f, '=B3+$A$3+B$4+$A6');
  for (const key of ['bl', 'it', 'fs', 'ff', 'fc', 'bg', 'ht', 'tb']) assert.equal(pasted[key], cell[key]);
  assert.equal(pasted.ct.fa, '$0.00');
  assert.equal(target.config.borderInfo.at(-1).value.b.color, '#123456');
  assert.ok(target.luckysheetfile[0].calcChain.some(entry => entry.r === 1 && entry.c === 1));
});

test('pasted formulas recalculate and keep working when their inputs change', () => {
  const ctx = context(); data(ctx)[0][0] = { f: '=A2*2', v: 8 }; data(ctx)[1][1] = { v: 7 };
  const contents = copy(ctx); select(ctx, 0, 1); pasteCells(ctx, contents);
  assert.equal(data(ctx)[0][1].v, 14);
  updateCell(ctx, 1, 1, editor('9')); groupValuesRefresh(ctx);
  assert.equal(data(ctx)[0][1].v, 18);
});

test('values-only uses raw values and retains destination formatting while removing old formula dependencies', () => {
  const source = context(), target = context('target');
  const values = [0, false, '', '00123', '=literal', 0.25];
  values.forEach((v, c) => { data(source)[0][c] = { v, f: c === 5 ? '=1/4' : undefined, ct: { fa: 'General', t: typeof v === 'number' ? 'n' : 's' } }; data(target)[1][c] = { f: '=99', v: 99, bg: '#ffeeaa', bl: 1, ct: { fa: '0.00%', t: 'n' } }; target.luckysheetfile[0].calcChain.push({ r: 1, c, id: 'target' }); });
  select(source, 0, 0, 0, 5); const contents = copy(source); select(target, 1, 0); pasteCells(target, contents, 'values');
  values.forEach((v, c) => { const cell = data(target)[1][c]; assert.equal(cell.v, v); assert.equal(cell.f, undefined); assert.equal(cell.bg, '#ffeeaa'); assert.equal(cell.bl, 1); assert.equal(cell.ct.fa, '0.00%'); });
  assert.equal(target.luckysheetfile[0].calcChain.length, 0);
});

test('formulas-only shifts relative references, preserves destination formatting, and leaves constants and blanks alone', () => {
  const ctx = context(); data(ctx)[0][0] = { f: '=A2*2', v: 4, bg: '#ff0000' }; data(ctx)[0][1] = { v: 'constant' };
  select(ctx, 0, 0, 0, 2); const contents = copy(ctx);
  data(ctx)[3][0] = { v: 'old', bg: '#00ff00', ct: { fa: '0.00', t: 's' } }; data(ctx)[3][1] = { v: 'keep constant' }; data(ctx)[3][2] = { v: 'keep blank' }; data(ctx)[4][0] = { v: 5 };
  select(ctx, 3, 0); pasteCells(ctx, contents, 'formulas');
  assert.equal(data(ctx)[3][0].f, '=A5*2'); assert.equal(data(ctx)[3][0].v, 10); assert.equal(data(ctx)[3][0].bg, '#00ff00'); assert.equal(data(ctx)[3][0].ct.fa, '0.00');
  assert.equal(data(ctx)[3][1].v, 'keep constant'); assert.equal(data(ctx)[3][2].v, 'keep blank');
});

test('formatting-only copies and clears styling without changing destination formulas or values', () => {
  const ctx = context(); data(ctx)[0][0] = { v: 50, bg: '#aabbcc', bl: 1, ct: { fa: '$0.00', t: 'n' } };
  const contents = copy(ctx); data(ctx)[2][0] = { f: '=1+2', v: 3, m: '3', it: 1, fc: '#ff0000' }; select(ctx, 2, 0); pasteCells(ctx, contents, 'formatting');
  const pasted = data(ctx)[2][0]; assert.equal(pasted.f, '=1+2'); assert.equal(pasted.v, 3); assert.equal(pasted.bg, '#aabbcc'); assert.equal(pasted.bl, 1); assert.equal(pasted.it, undefined); assert.equal(pasted.fc, undefined); assert.equal(pasted.m, '$3.00');
});

test('copy values/formulas explicitly controls subsequent normal paste and supplies useful text for other apps', () => {
  for (const mode of ['values', 'formulas']) {
    const ctx = context(); data(ctx)[0][0] = { f: '=A2+1', v: 3, m: '3' }; data(ctx)[1][1] = { v: 5 };
    const contents = copy(ctx, mode); assert.equal(contents.text, mode === 'formulas' ? '=A2+1' : '3');
    select(ctx, 0, 1); pasteCells(ctx, contents); assert.equal(data(ctx)[0][1].f, mode === 'formulas' ? '=B2+1' : undefined); assert.equal(data(ctx)[0][1].v, mode === 'formulas' ? 6 : 3);
  }
});

test('plain text paste supports formulas, rectangular TSV, quoted multiline cells and trailing line endings', () => {
  const ctx = context(); select(ctx, 1, 0);
  pasteCells(ctx, { html: '', text: '2\t=A2*3\n"line one\nline two"\t"a\tb"\n' });
  assert.equal(data(ctx)[1][0].v, 2); assert.equal(data(ctx)[1][1].f, '=A2*3'); assert.equal(data(ctx)[1][1].v, 6);
  assert.equal(data(ctx)[2][0].v, 'line one\nline two'); assert.equal(data(ctx)[2][1].v, 'a\tb'); assert.equal(data(ctx)[3][0], null);
  assert.equal(ctx.config.borderInfo, undefined, 'Plain text paste should not create empty border records');
  select(ctx, 5, 0); pasteCells(ctx, { html: '', text: '=1+2' }, 'values'); assert.equal(data(ctx)[5][0].v, '=1+2'); assert.equal(data(ctx)[5][0].f, undefined);
});

test('TSV quoting round-trips tabs, quotes, newlines, empty cells and uneven rows', () => {
  const rows = [['a\tb', 'x"y', 'line\nline', ''], ['', '', '', 'last']]; assert.deepEqual(parseTsv(toTsv(rows)), rows);
  assert.deepEqual(parseTsv('a\tb\r\nc\r\n'), [['a', 'b'], ['c', '']]);
});

test('copies merged cells and repeats source blocks over a larger selection', () => {
  const ctx = context(); data(ctx)[0][0] = { v: 'merged', mc: { r: 0, c: 0, rs: 1, cs: 2 } }; data(ctx)[0][1] = { mc: { r: 0, c: 0 } }; ctx.config.merge = { '0_0': { r: 0, c: 0, rs: 1, cs: 2 } };
  select(ctx, 0, 0, 0, 1); const contents = copy(ctx); select(ctx, 3, 0, 4, 3); pasteCells(ctx, contents);
  for (const [r, c] of [[3, 0], [3, 2], [4, 0], [4, 2]]) { assert.equal(data(ctx)[r][c].v, 'merged'); assert.deepEqual(ctx.config.merge[`${r}_${c}`], { r, c, rs: 1, cs: 2 }); assert.deepEqual(data(ctx)[r][c+1].mc, { r, c }); }
});

test('expanded destination protection, partial merges and disjoint targets refuse the entire paste', () => {
  const source = context(); select(source, 0, 0, 1, 0); data(source)[0][0] = { v: 'one' }; data(source)[1][0] = { v: 'two' }; const contents = copy(source);
  for (const kind of ['protected', 'merged', 'disjoint']) {
    const ctx = context(); select(ctx, 2, 0);
    if (kind === 'protected') ctx.config.rowReadOnly = { 3: 1 };
    if (kind === 'merged') ctx.config.merge = { '2_0': { r: 2, c: 0, rs: 1, cs: 2 } };
    if (kind === 'disjoint') ctx.luckysheet_select_save.push(range(5, 0));
    const before = structuredClone(data(ctx)); pasteCells(ctx, contents); assert.deepEqual(data(ctx), before);
  }
});

test('disjoint copies preserve each formula and its source position', () => {
  const ctx = context(); data(ctx)[0][0] = { f: '=A2', v: 1 }; data(ctx)[0][2] = { f: '=C2', v: 2 };
  ctx.luckysheet_select_save = [range(0, 0), range(0, 2)]; const contents = copy(ctx); select(ctx, 3, 0); pasteCells(ctx, contents);
  assert.equal(data(ctx)[3][0].f, '=A5'); assert.equal(data(ctx)[3][1].f, '=B5');
});

test('paste is a single reversible mutation including formulas, merges, and borders', () => {
  enablePatches(); const ctx = context(); data(ctx)[0][0] = { v: 42, bg: '#112233' }; const contents = copy(ctx); select(ctx, 3, 1);
  const before = structuredClone(ctx.luckysheetfile);
  const [after, patches, inverse] = produceWithPatches(ctx, draft => pasteCells(draft, contents));
  assert.ok(patches.length); assert.equal(data(after)[3][1].v, 42); assert.deepEqual(applyPatches(after, inverse).luckysheetfile, before);
  assert.equal(data(applyPatches(applyPatches(after, inverse), patches))[3][1].v, 42);
});

test('cut moves cells once without a hidden DOM editor and retains formula source', () => {
  const ctx = context(); data(ctx)[0][0] = { f: '=B1*2', v: 8, bl: 1, bg: '#123456' }; data(ctx)[0][1] = { v: 4 };
  ctx.luckysheetfile[0].calcChain = [{ r: 0, c: 0, id: 'source' }];
  const contents = copy(ctx, 'all', true); assert.equal(ctx.luckysheet_paste_iscut, true);
  select(ctx, 3, 0); pasteCells(ctx, contents); groupValuesRefresh(ctx);
  assert.equal(data(ctx)[0][0]?.v, undefined); assert.equal(data(ctx)[3][0].v, 8); assert.equal(data(ctx)[3][0].bl, 1);
  assert.equal(data(ctx)[3][0].f, '=B1*2'); assert.equal(ctx.luckysheet_paste_iscut, false);
  updateCell(ctx, 0, 1, editor('9')); groupValuesRefresh(ctx); assert.equal(data(ctx)[3][0].v, 18);
  data(ctx)[6][1] = { v: 3 }; select(ctx, 6, 0); pasteCells(ctx, contents); assert.equal(data(ctx)[3][0].v, 18); assert.equal(data(ctx)[6][0].v, 6); assert.equal(data(ctx)[6][0].f, '=B7*2');
});

test('recalculation uses each worksheet\'s formula chain and data', () => {
  const ctx = context(); data(ctx)[0][0] = { v: 5 }; const contents = copy(ctx);
  ctx.luckysheetfile.push({ id: 'other', name: 'Other', data: [[{ f: '=B1+1', v: 3 }, { v: 2 }]], config: {}, calcChain: [{ r: 0, c: 0, id: 'other' }] });
  select(ctx, 2, 0); pasteCells(ctx, contents);
  assert.equal(ctx.luckysheetfile[1].data[0][0].f, '=B1+1'); assert.equal(ctx.luckysheetfile[1].data[0][0].v, 3);
});

test('paste recalculates affected dependents and new formulas without touching unrelated cached results', () => {
  const ctx = context(); data(ctx)[0][0] = { v: 3 }; const contents = copy(ctx);
  data(ctx)[2][0] = { v: 1 }; data(ctx)[2][1] = { f: '=A3*2', v: 2 }; data(ctx)[5][0] = { f: '=UNSUPPORTED(Z99)', v: 'cached unsupported result' };
  ctx.luckysheetfile[0].calcChain = [{ r: 2, c: 1, id: 'source' }, { r: 5, c: 0, id: 'source' }];
  select(ctx, 2, 0); pasteCells(ctx, contents); assert.equal(data(ctx)[2][1].v, 6); assert.equal(data(ctx)[5][0].v, 'cached unsupported result');
  select(ctx, 8, 0); pasteCells(ctx, { html: '', text: '=B9+1\t=2+3' }); assert.equal(data(ctx)[8][0].v, 6); assert.equal(data(ctx)[8][1].v, 5);
});

test('malformed clipboard metadata is ignored', () => {
  assert.equal(readSheetClipboard('<table data-sheetian-clipboard="bad-json">'), null);
  assert.equal(readSheetClipboard('<table data-sheetian-clipboard="'+encodeURIComponent(JSON.stringify({version:1,id:'x',sheetId:'x',mode:'all',cells:[[{row:0,column:0,cell:{f:3}}]]}))+'">'), null);
});

test('context menu patch guards both pinned distributions and exposes all clipboard modes', async () => {
  for (const file of ['index.js', 'index.esm.js']) {
    const patched = patchClipboardMenu(await readFile(`node_modules/@fortune-sheet/react/dist/${file}`, 'utf8'));
    assert.match(patched, /"sheetian-copy-formulas": "Copy formulas only"/); assert.match(patched, /"sheetian-paste-formatting": "Paste formatting only"/); assert.match(patched, /refs.cellInput.current/);
    assert.match(patchCutPaste(await readFile(`node_modules/@fortune-sheet/core/dist/${file}`, 'utf8')), /triggerType === "sheetian-cut"/);
    assert.match(patchPasteFormulaCalculation(await readFile(`node_modules/@fortune-sheet/core/dist/${file}`, 'utf8')), /sheetianPasteFormulaKeys/);
  }
  assert.throws(() => patchClipboardMenu('future engine'));
  assert.throws(() => patchCutPaste('future engine'));
  assert.throws(() => patchPasteFormulaCalculation('future engine'));
});

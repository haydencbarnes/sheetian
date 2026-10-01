import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { build } from 'esbuild';
import { createRequire } from 'node:module';
import { patchBooleanLiterals } from '../engine-compat.mjs';

const require = createRequire(import.meta.url);
test('boolean literals calculate without changing strings or sheet/cell references', async () => {
  const result = await build({ stdin: { contents: 'export { Parser } from "@fortune-sheet/formula-parser";', resolveDir: process.cwd() },
    bundle: true, write: false, platform: 'node', format: 'cjs',
    plugins: [{ name: 'boolean-literals', setup(builder) {
      builder.onLoad({ filter: /formula-parser[\\/](?:es|lib)[\\/]parser\.js$/ }, async args => ({
        contents: patchBooleanLiterals(await readFile(args.path, 'utf8')), loader: 'js' }));
    } }] });
  const module = { exports: {} };
  new Function('module', 'exports', 'require', result.outputFiles[0].text)(module, module.exports, require);
  const parser = new module.exports.Parser();
  for (const [expression, expected] of [['TRUE', true], ['false', false], ['IF(1>2,TRUE,FALSE)', false],
    ['IF(2>1,TRUE,FALSE)', true], ['"TRUE"', 'TRUE'], ['TRUE()', true], ['FALSE()', false]]) {
    assert.deepEqual(parser.parse(expression), { error: null, result: expected }, expression);
  }
  const references = [];
  parser.on('callCellValue', (cell, _options, done) => { references.push(cell); done(7); });
  assert.equal(parser.parse('TRUE1').result, 7);
  assert.equal(parser.parse("'TRUE'!A1").result, 7);
  assert.equal(references.length, 2);
  assert.equal(references[1].sheetName, 'TRUE');
  assert.equal(parser.parse('$TRUE').error, '#NAME?');
});
test('boolean patch guards both pinned parser distributions', async () => {
  for (const type of ['es', 'lib']) assert.match(patchBooleanLiterals(await readFile(`node_modules/@fortune-sheet/formula-parser/${type}/parser.js`, 'utf8')), /label\.toUpperCase\(\)/);
  assert.throws(() => patchBooleanLiterals('future parser'));
});

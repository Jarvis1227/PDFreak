import test from 'node:test';
import assert from 'node:assert/strict';
import { parsePages, outputName, textRows, convertPdf } from '../src/converter.js';

test('page selection is bounded, deduplicated and in document order', () => {
  assert.deepEqual(parsePages('5, 1-3, 2', 8), [1, 2, 3, 5]);
  assert.deepEqual(parsePages('', 3), [1, 2, 3]);
  for (const range of ['0', '3-1', '1,', '1.5', '1-9', '1e2']) assert.throws(() => parsePages(range, 8));
  assert.throws(() => parsePages('', 61)); assert.throws(() => parsePages('1-99999999999999999', 100000000000000000));
  assert.throws(() => parsePages('1-60,61', 80));
});
test('output filenames cannot introduce path separators', () => {
  assert.equal(outputName('report.PDF'), 'report.docx');
  assert.equal(outputName('../my:document.pdf'), '.._my_document.docx');
});
const content = { styles: { f1: { fontFamily: 'sans-serif' } }, items: [
  { str: 'world', transform: [12, 0, 0, 12, 90, 700], width: 30, fontName: 'f1' },
  { str: 'Hello', transform: [12, 0, 0, 12, 50, 700], width: 30, fontName: 'f1' },
  { str: 'Second line', transform: [12, 0, 0, 12, 50, 676], width: 70, fontName: 'f1' }
] };
const viewport = { width: 612, height: 792, transform: [1, 0, 0, -1, 0, 792] };
test('fragments become lines in visual reading order', () => {
  const rows = textRows(content, viewport);
  assert.equal(rows.length, 2); assert.deepEqual(rows[0].items.map(i => i.text), ['Hello', 'world']); assert.equal(rows[0].y, 92);
});
const fakePdf = text => ({ getPage: async () => ({ getViewport: () => viewport, getTextContent: async () => text }) });
test('editable conversion produces a real ZIP based Word file', async () => {
  const progress = []; const blob = await convertPdf(fakePdf(content), [1], 'editable', value => progress.push(value));
  const bytes = new Uint8Array(await blob.arrayBuffer());
  assert.deepEqual([...bytes.slice(0, 2)], [80, 75]); assert.equal(progress.at(-1), 100); assert.ok(blob.size > 1000);
});
test('scans are reported rather than silently exported as blank pages', async () => {
  await assert.rejects(convertPdf(fakePdf({ items: [] }), [2], 'editable'), /Page 2 has no selectable text/);
});
test('cancellation prevents conversion', async () => {
  const controller = new AbortController(); controller.abort();
  await assert.rejects(convertPdf(fakePdf(content), [1], 'editable', undefined, controller.signal), { name: 'AbortError' });
});

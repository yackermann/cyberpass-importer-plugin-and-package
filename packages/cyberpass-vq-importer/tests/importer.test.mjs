import assert from 'node:assert/strict';
import test from 'node:test';
import * as api from '../dist/index.js';
import * as XLSX from '../vendor/sheetjs/xlsx.mjs';
const { importVq, parseVendorQuestionnaire, readVendorQuestionnaire, responseChoice, bindVqDropzone, bindVqFileInput, VqImportError } = api;

function workbook(sheets = { VQ: [['SR No.', 'Vendor Response'], ['1.1', 'Text']] }) {
  const wb = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
  return wb;
}
const importOptions = { reader: { read: () => workbook() } };
const file = new File(['ignored by test decoder'], 'sample.xls');
const tick = () => new Promise(resolve => setImmediate(resolve));

test('standalone decoder accepts binary XLS and exposes rows, prediction and normalized lookup', async () => {
  const wb = workbook({ VQ: [['SR No.', 'Vendor Response'], ['SR001', 'Café\n中文'], ['1.2', ' n/A: not supported '], ['1.3', '']] });
  const bytes = XLSX.write(wb, { type: 'array', bookType: 'biff8' });
  for (const input of [new File([bytes], 'VQ.xls'), new Blob([bytes]), bytes, new Uint8Array(bytes)]) {
    const vq = await importVq(input);
    assert.deepEqual(vq.rows.map(({ requirementId, value, predictedResponse }) => ({ requirementId, value, predictedResponse })), [
      { requirementId: '1', value: 'Café\n中文', predictedResponse: 'YES' },
      { requirementId: '1.2', value: 'n/A: not supported', predictedResponse: 'N/A' },
      { requirementId: '1.3', value: '', predictedResponse: 'NO' }
    ]);
    assert.equal(vq.get('Requirement #001'), vq.rows[0]);
    assert.equal(vq.get('missing'), undefined);
    assert.deepEqual([...vq], vq.rows);
    assert.equal(vq.rows[0].row, 2);
    assert.equal(vq.rows[0].responseColumn, 'B');
    assert.ok(Object.isFrozen(vq.rows[0]));
  }
  assert.equal(api.CyberPassVqImporter, undefined);
  assert.equal(api.installImportExcelButton, undefined);
});

test('uses paired headers and prefers Vendor Response over other answers', () => {
  const wb = workbook({
    Overview: [['Version', 'Notes'], [1, 'Do not import']],
    VQ: [['Title mentions Vendor Response'], ['SR No.', 'Answer', 'Vendor Response'], ['1.1', 'NO', 'Actual text'], ['', '', '1.9'], ['See requirement 1.2', '', 'Not a row']]
  });
  const rows = parseVendorQuestionnaire(wb);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].value, 'Actual text');
  assert.equal(rows[0].predictedResponse, 'YES');
});

test('sparse worksheets with delayed and repeated headers retain correct coordinates', () => {
  const wb = { SheetNames: ['VQ'], Sheets: { VQ: {
    B30: { v: 'SR No.' }, AA30: { v: 'Vendor Response' },
    B31: { v: '1.1' }, AA31: { v: 'Text' },
    B40: { v: 'SR No.' }, AA40: { v: 'Vendor Response' }, B41: { v: '1.2' },
    XFD1048576: { v: 'Unrelated formatting tail' }
  } } };
  const rows = parseVendorQuestionnaire(wb);
  assert.deepEqual(rows.map(row => [row.requirementId, row.row, row.responseColumn, row.predictedResponse]), [['1.1', 31, 'AA', 'YES'], ['1.2', 41, 'AA', 'NO']]);
});

test('rejects ambiguous duplicates with locations and supports explicit policies and sheets', () => {
  const wb = workbook({ A: [['SR No.', 'Vendor Response'], ['SR001', 'First']], B: [['Requirement', 'Vendor Response'], ['1', 'Last']] });
  assert.throws(() => parseVendorQuestionnaire(wb), error => error.code === 'DUPLICATE_REQUIREMENT' && /A!A2.*B!A2/.test(error.message));
  assert.equal(parseVendorQuestionnaire(wb, { duplicates: 'first' })[0].value, 'First');
  assert.equal(parseVendorQuestionnaire(wb, { duplicates: 'last' })[0].value, 'Last');
  assert.equal(parseVendorQuestionnaire(wb, { sheetNames: ['B'] })[0].value, 'Last');
  assert.throws(() => parseVendorQuestionnaire(wb, { sheetNames: ['Missing'] }), { code: 'SHEET_NOT_FOUND' });
});

test('surfaces missing columns and read failures as typed errors', async () => {
  assert.throws(() => parseVendorQuestionnaire(workbook({ Notes: [['SR No.', 'Description'], ['1.1', 'Requirement prose']] })), { code: 'NO_REQUIREMENTS' });
  const cause = new Error('Encrypted or damaged workbook');
  await assert.rejects(readVendorQuestionnaire(file, { reader: { read() { throw cause; } } }), error => error instanceof VqImportError && error.code === 'READ_FAILED' && error.cause === cause && error.message.includes('sample.xls'));
  assert.deepEqual(['', '   ', 'n/a', ' N/a: reason', 'N/Aanything', 'NO', 'Text'].map(responseChoice), ['NO', 'NO', 'N/A', 'N/A', 'N/A', 'YES', 'YES']);
});

// EventTarget harness: these helpers only bind events on caller-owned elements.
function dropEvent(files, type = 'drop') {
  const event = new Event(type, { cancelable: true });
  Object.defineProperty(event, 'dataTransfer', { value: { files, types: files.length ? ['Files'] : ['text/plain'], dropEffect: 'none' } });
  return event;
}

test('drop helper delivers data, reports multiple files, ignores text, and unbinds', async () => {
  const element = new EventTarget();
  const received = [], errors = [];
  const unbind = bindVqDropzone(element, { options: importOptions, onImport: vq => received.push(vq), onError: error => errors.push(error) });
  const textDrop = dropEvent([]);
  element.dispatchEvent(textDrop);
  assert.equal(textDrop.defaultPrevented, false);
  const drag = dropEvent([file], 'dragover'); element.dispatchEvent(drag);
  assert.equal(drag.dataTransfer.dropEffect, 'copy');
  const drop = dropEvent([file]); element.dispatchEvent(drop);
  assert.equal(drop.defaultPrevented, true);
  await tick();
  assert.equal(received[0].get('1.1').value, 'Text');
  element.dispatchEvent(dropEvent([file, file])); await tick();
  assert.equal(errors[0].code, 'FILE_COUNT');
  unbind(); element.dispatchEvent(dropEvent([file])); await tick();
  assert.equal(received.length, 1);
});

test('file-input helper resets the input, suppresses stale results, and cleans up pending work', async () => {
  const input = Object.assign(new EventTarget(), { type: 'file', accept: '', files: [file], value: 'sample.xls' });
  const received = [], errors = [], pending = [];
  const unbind = bindVqFileInput(input, {
    options: { reader: { read: () => new Promise(resolve => pending.push(resolve)) } },
    onImport: vq => received.push(vq), onError: error => errors.push(error)
  });
  input.dispatchEvent(new Event('change')); await tick();
  assert.equal(input.value, '');
  assert.match(input.accept, /\.xls/);
  input.dispatchEvent(new Event('change')); await tick();
  pending[1](workbook()); await tick();
  pending[0](workbook()); await tick();
  assert.equal(received.length, 1);
  input.dispatchEvent(new Event('change')); await tick();
  unbind(); pending[2](workbook()); await tick();
  assert.equal(received.length, 1);
  assert.equal(input.accept, '');
  assert.deepEqual(errors, []);
});

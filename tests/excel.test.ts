import assert from 'node:assert/strict';
import test from 'node:test';
import * as XLSX from '../packages/cyberpass-vq-importer/vendor/sheetjs/xlsx.mjs';
import reader from '../vendor/xlsx.mjs';
import { readWorkbookFile } from '../src/excel/parser';
import { readVendorQuestionnaire, responseChoice } from '../packages/cyberpass-vq-importer/src/workbook';

for (const [bookType, extension] of [
  ['biff8', 'xls'], ['biff5', 'xls'], ['xlsx', 'xlsx'], ['xlsm', 'xlsm'], ['xlsb', 'xlsb']
]) {
  test(`imports ${bookType} bytes through the extension and reusable library`, async () => {
    // The BIFF5 writer supports ASCII only. BIFF8 exercises Unicode XLS text.
    const response = bookType === 'biff5' ? 'First line\nSecond line' : 'Café € — 中文\nSecond line';
    const rows = [
      ['Vendor Questionnaire'],
      [],
      ['', 'SR No.', '', '', '', '', '', '', 'Vendor Response'],
      ['', '1.1', '', '', '', '', '', '', response],
      ['', '1.2', '', '', '', '', '', '', 'n/a - Unsupported'],
      ['', '1.3', '', '', '', '', '', '', '']
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'FIDO Security Requirements');
    const bytes = XLSX.write(workbook, { type: 'array', bookType });
    if (extension === 'xls') {
      // Verify a real OLE compound binary file, rather than renamed OOXML.
      assert.deepEqual([...new Uint8Array(bytes).slice(0, 8)], [208, 207, 17, 224, 161, 177, 26, 225]);
    }
    const file = new File([bytes], `VQ.${extension.toUpperCase()}`);
    const requirements = await readWorkbookFile(file);
    assert.deepEqual(requirements.map(row => [row.requirementId, row.value, row.column, row.responseColumn]), [
      ['1.1', response, 'B', 'I'], ['1.2', 'n/a - Unsupported', 'B', 'I'], ['1.3', '', 'B', 'I']
    ]);
    assert.deepEqual(requirements.map(row => responseChoice(row.value)), ['YES', 'N/A', 'NO']);
    const importedRows = await readVendorQuestionnaire(file);
    assert.deepEqual(importedRows, requirements.map(row => ({ ...row, predictedResponse: responseChoice(row.value) })));
    assert.deepEqual(await readVendorQuestionnaire(file, { reader }), importedRows);
  });
}

test('reader failures identify the workbook and retain the cause', async () => {
  const file = { name: 'broken.xls', arrayBuffer: async () => { throw new Error('File read failed'); } } as File;
  await assert.rejects(readWorkbookFile(file), /Could not read workbook "broken.xls": File read failed/);
});

import { CyberPassVqImporter } from './cyberpass.js';
import { readVendorQuestionnaire } from './workbook.js';
import type { ImportExcelButtonOptions, XlsxReader } from './types.js';

/** Mount a file picker button that reads a VQ workbook and fills the current page. */
export function installImportExcelButton(xlsx: XlsxReader, options: ImportExcelButtonOptions = {}): HTMLButtonElement {
  const document = options.document ?? globalThis.document;
  const button = options.button ?? document.createElement('button');
  button.type = 'button';
  if (!options.button) document.body.appendChild(button);
  button.textContent = options.buttonText ?? 'Import Excel VQ';
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = options.accept ?? '.xls,.xlsx,.xlsm,.xlsb';
  input.hidden = true;
  document.body.appendChild(input);
  const importer = new CyberPassVqImporter(document);
  button.addEventListener('click', () => input.click());
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    if (!file) return;
    try {
      const requirements = await readVendorQuestionnaire(file, xlsx);
      if (!requirements.length) throw new Error('No requirement rows detected. Check the SR No. and Vendor Response columns.');
      options.onLoaded?.(requirements);
      const result = await importer.fill(requirements, { replaceExisting: options.replaceExisting });
      options.onFilled?.(result);
      if (result.failed) throw new Error(`Import completed with ${result.failed} failed field(s): ${result.errors.slice(0, 3).join(' | ')}`);
    } catch (error) {
      options.onError?.(error instanceof Error ? error : new Error(String(error)));
    } finally {
      input.value = '';
    }
  });
  return button;
}

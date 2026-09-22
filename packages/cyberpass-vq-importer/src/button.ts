import { CyberPassVqImporter } from './cyberpass.js';
import { readVendorQuestionnaire } from './workbook.js';
import type { ImportExcelButtonOptions, XlsxReader } from './types.js';

/** Mount a file picker button that reads a VQ workbook and fills the current page. */
export function installImportExcelButton(xlsx: XlsxReader, options: ImportExcelButtonOptions = {}): HTMLButtonElement {
  const document = options.document ?? globalThis.document;
  const buttonTarget = typeof options.button === 'string' ? document.querySelector<HTMLElement>(options.button) : options.button;
  if (typeof options.button === 'string' && !buttonTarget) throw new Error(`Import Excel VQ button was not found: ${options.button}`);
  if (buttonTarget && buttonTarget.tagName !== 'BUTTON') throw new Error('Import Excel VQ button selector must target a <button>.');
  const button = (buttonTarget as HTMLButtonElement | undefined) ?? document.createElement('button');
  const status = typeof options.status === 'string' ? document.querySelector<HTMLElement>(options.status) : options.status;
  if (typeof options.status === 'string' && !status) throw new Error(`Import Excel VQ status element was not found: ${options.status}`);
  button.type = 'button';
  if (!options.button) document.body.appendChild(button);
  button.textContent = options.buttonText ?? 'Import Excel VQ';
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = options.accept ?? '.xls,.xlsx,.xlsm,.xlsb';
  input.hidden = true;
  document.body.appendChild(input);
  const importer = new CyberPassVqImporter(document);
  const setStatus = (message: string) => { if (status) status.textContent = message; };
  button.addEventListener('click', () => input.click());
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    if (!file) return;
    try {
      const requirements = await readVendorQuestionnaire(file, xlsx);
      if (!requirements.length) throw new Error('No requirement rows detected. Check the SR No. and Vendor Response columns.');
      setStatus(`Loaded ${requirements.length} workbook responses.`);
      options.onLoaded?.(requirements);
      const result = await importer.fill(requirements, { replaceExisting: options.replaceExisting });
      setStatus(`Filled ${result.filled} · skipped ${result.skipped} · ${result.mismatches} mismatches · ${result.failed} failed.`);
      options.onFilled?.(result);
      if (result.failed) throw new Error(`Import completed with ${result.failed} failed field(s): ${result.errors.slice(0, 3).join(' | ')}`);
    } catch (error) {
      const normalized = error instanceof Error ? error : new Error(String(error));
      setStatus(`Excel import failed: ${normalized.message}`);
      options.onError?.(normalized);
    } finally {
      input.value = '';
    }
  });
  return button;
}

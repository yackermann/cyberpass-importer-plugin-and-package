import type { ImportExcelButtonOptions, XlsxReader } from './types.js';
/** Mount a file picker button that reads a VQ workbook and fills the current page. */
export declare function installImportExcelButton(xlsx: XlsxReader, options?: ImportExcelButtonOptions): HTMLButtonElement;

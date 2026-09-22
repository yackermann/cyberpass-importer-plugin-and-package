import type { ExcelRequirement, WorkbookLike, XlsxReader } from './types.js';
/** Parse a SheetJS-compatible workbook into normalized CyberPass requirements. */
export declare function parseVendorQuestionnaire(workbook: WorkbookLike): ExcelRequirement[];
/** Read and parse a local .xls/.xlsx/.xlsm/.xlsb file using the caller's SheetJS instance. */
export declare function readVendorQuestionnaire(file: Blob, xlsx: XlsxReader): Promise<ExcelRequirement[]>;
export declare function responseChoice(value: string): 'YES' | 'NO' | 'N/A';

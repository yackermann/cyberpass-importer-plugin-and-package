import type { ImportOptions, ParseOptions, ResponseChoice, VqInput, VqRow, WorkbookLike } from './types.js';
/** Deterministic prediction from Vendor Response text; no compliance assessment. */
export declare function responseChoice(value: string): ResponseChoice;
/** Extract rows only from sheets containing both a requirement and response header. */
export declare function parseVendorQuestionnaire(workbook: WorkbookLike, options?: ParseOptions): VqRow[];
/** Decode a local File/Blob, ArrayBuffer, or Uint8Array into plain VQ rows. */
export declare function readVendorQuestionnaire(input: VqInput, options?: ImportOptions): Promise<VqRow[]>;

export type ResponseChoice = 'YES' | 'NO' | 'N/A';
/** One imported requirement, including its location in the original workbook. */
export interface VqRow {
    requirementId: string;
    rawRequirementId: string;
    value: string;
    predictedResponse: ResponseChoice;
    sheetName: string;
    /** One-based Excel row. */
    row: number;
    /** Excel column letters, e.g. B and I. */
    column: string;
    responseColumn: string;
}
export type ExcelRequirement = VqRow;
export interface SheetLike {
    [cell: string]: unknown;
}
export interface WorkbookLike {
    SheetNames: string[];
    Sheets: Record<string, SheetLike>;
}
export interface XlsxReader {
    read(data: ArrayBuffer | Uint8Array, options?: Record<string, unknown>): WorkbookLike | Promise<WorkbookLike>;
}
export type VqInput = Blob | ArrayBuffer | Uint8Array;
export interface ParseOptions {
    /** Process only these worksheets. Default: all worksheets with VQ headers. */
    sheetNames?: readonly string[];
    /** Repeated normalized IDs fail by default. Explicitly select first/last if needed. */
    duplicates?: 'error' | 'first' | 'last';
}
export interface ImportOptions extends ParseOptions {
    /** Optional alternative to the included SheetJS decoder. */
    reader?: XlsxReader;
}

export type ResponseChoice = 'YES' | 'NO' | 'N/A';
export interface ExcelRequirement {
    requirementId: string;
    rawRequirementId: string;
    value: string;
    sheetName: string;
    row: number;
    column?: string;
    responseColumn?: string;
}
export interface WorkbookLike {
    SheetNames: string[];
    Sheets: Record<string, SheetLike>;
}
export interface SheetLike {
    '!ref'?: string;
    [cell: string]: unknown;
}
export interface XlsxReader {
    read(data: ArrayBuffer | Uint8Array, options?: Record<string, unknown>): WorkbookLike | Promise<WorkbookLike>;
}
export interface CyberPassControl {
    tag: string;
    type?: string;
    name?: string;
    id?: string;
    ariaLabel?: string;
    value?: string;
}
export interface CyberPassField {
    requirementId: string;
    label?: string;
    containerText: string;
    controls: CyberPassControl[];
}
export interface ScanResult {
    fields: CyberPassField[];
}
export interface FillOptions {
    replaceExisting?: boolean;
    /** Scroll the page while discovering lazy-rendered requirement cards. */
    discoverLazyRequirements?: boolean;
}
export interface FillResult {
    filled: number;
    skipped: number;
    failed: number;
    mismatches: number;
    errors: string[];
}
export interface ImportExcelButtonOptions {
    document?: Document;
    button?: HTMLButtonElement;
    buttonText?: string;
    accept?: string;
    replaceExisting?: boolean;
    onLoaded?: (requirements: ExcelRequirement[]) => void;
    onFilled?: (result: FillResult) => void;
    onError?: (error: Error) => void;
}

export type VqErrorCode = 'READ_FAILED' | 'NO_REQUIREMENTS' | 'DUPLICATE_REQUIREMENT' | 'SHEET_NOT_FOUND' | 'FILE_COUNT';
export declare class VqImportError extends Error {
    readonly code: VqErrorCode;
    constructor(code: VqErrorCode, message: string, options?: ErrorOptions);
}

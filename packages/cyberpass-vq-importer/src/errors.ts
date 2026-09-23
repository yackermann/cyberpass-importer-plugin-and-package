export type VqErrorCode = 'READ_FAILED' | 'NO_REQUIREMENTS' | 'DUPLICATE_REQUIREMENT' | 'SHEET_NOT_FOUND' | 'FILE_COUNT';

export class VqImportError extends Error {
  constructor(public readonly code: VqErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'VqImportError';
  }
}

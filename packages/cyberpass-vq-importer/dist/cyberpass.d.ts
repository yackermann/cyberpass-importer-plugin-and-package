import type { ExcelRequirement, FillOptions, FillResult, ScanResult } from './types.js';
export declare class CyberPassVqImporter {
    private readonly document;
    constructor(document?: Document);
    scan(options?: FillOptions): Promise<ScanResult>;
    fill(requirements: ExcelRequirement[], options?: FillOptions): Promise<FillResult>;
}

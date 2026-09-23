import type { ImportOptions, VqInput, VqRow } from './types.js';
export declare class VqDocument implements Iterable<Readonly<VqRow>> {
    readonly rows: readonly Readonly<VqRow>[];
    private readonly byId;
    constructor(rows: readonly VqRow[]);
    /** Accepts normalized or prefixed IDs. Missing requirements return undefined. */
    get(requirementId: string): Readonly<VqRow> | undefined;
    [Symbol.iterator](): Iterator<Readonly<VqRow>>;
}
/** Import data for the host application to review, map, and apply. */
export declare function importVq(input: VqInput, options?: ImportOptions): Promise<VqDocument>;

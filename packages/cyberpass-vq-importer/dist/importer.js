import { normalizeRequirementId } from './normalizer.js';
import { readVendorQuestionnaire } from './workbook.js';
export class VqDocument {
    rows;
    byId;
    constructor(rows) {
        this.rows = Object.freeze(rows.map(row => Object.freeze({ ...row })));
        this.byId = new Map(this.rows.map(row => [row.requirementId, row]));
    }
    /** Accepts normalized or prefixed IDs. Missing requirements return undefined. */
    get(requirementId) {
        return this.byId.get(normalizeRequirementId(requirementId) ?? '');
    }
    [Symbol.iterator]() { return this.rows[Symbol.iterator](); }
}
/** Import data for the host application to review, map, and apply. */
export async function importVq(input, options) {
    return new VqDocument(await readVendorQuestionnaire(input, options));
}

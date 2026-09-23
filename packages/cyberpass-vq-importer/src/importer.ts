import { normalizeRequirementId } from './normalizer.js';
import { readVendorQuestionnaire } from './workbook.js';
import type { ImportOptions, VqInput, VqRow } from './types.js';

export class VqDocument implements Iterable<Readonly<VqRow>> {
  readonly rows: readonly Readonly<VqRow>[];
  private readonly byId: Map<string, Readonly<VqRow>>;

  constructor(rows: readonly VqRow[]) {
    this.rows = Object.freeze(rows.map(row => Object.freeze({ ...row })));
    this.byId = new Map(this.rows.map(row => [row.requirementId, row]));
  }

  /** Accepts normalized or prefixed IDs. Missing requirements return undefined. */
  get(requirementId: string): Readonly<VqRow> | undefined {
    return this.byId.get(normalizeRequirementId(requirementId) ?? '');
  }

  [Symbol.iterator](): Iterator<Readonly<VqRow>> { return this.rows[Symbol.iterator](); }
}

/** Import data for the host application to review, map, and apply. */
export async function importVq(input: VqInput, options?: ImportOptions): Promise<VqDocument> {
  return new VqDocument(await readVendorQuestionnaire(input, options));
}

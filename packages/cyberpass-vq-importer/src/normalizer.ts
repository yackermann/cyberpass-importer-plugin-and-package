/** Normalize SR001, SAR 12, Requirement #12, and dotted requirement numbers. */
export function normalizeRequirementId(input: unknown): string | null {
  const value = String(input ?? '').trim();
  const match = value.match(/^(?:(?:S(?:A)?R\s*[-_:#]?|requirement\s*#?|req\s*#?)\s*)?(\d+(?:\.\d+)*)$/i);
  return match ? match[1].split('.').map(part => part.replace(/^0+(?=\d)/, '')).join('.') : null;
}

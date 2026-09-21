/** Normalize SR001, SR 12, Requirement #12, and dotted requirement numbers. */
export function normalizeRequirementId(input: unknown): string | null {
  const text = String(input ?? '').trim();
  if (!text) return null;
  const sr = text.match(/\bSR\s*[-_:]?\s*0*(\d+(?:\.\d+)*)\b/i);
  const req = text.match(/\b(?:requirement|req)\s*#?\s*0*(\d+(?:\.\d+)*)\b/i);
  const bare = text.match(/^\s*0*(\d+(?:\.\d+)*)\s*$/);
  const value = sr?.[1] ?? req?.[1] ?? bare?.[1];
  if (!value) return null;
  return value.split('.').map(part => { const n=Number(part); return Number.isFinite(n) ? String(Number(n.toPrecision(12))) : part; }).join('.');
}

export function findRequirementIds(input: unknown): string[] {
  const text = String(input ?? '');
  const out = new Set<string>();
  const re = /\b(?:SR\s*[-_:]?\s*0*\d+(?:\.\d+)*|Requirement\s*#?\s*0*\d+(?:\.\d+)*)\b/gi;
  for (const match of text.matchAll(re)) { const id = normalizeRequirementId(match[0]); if (id) out.add(id); }
  return [...out];
}

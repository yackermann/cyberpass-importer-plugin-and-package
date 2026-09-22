/** Normalize SR001, SR 12, Requirement #12, and dotted requirement numbers. */
export function normalizeRequirementId(input) {
    const value = String(input ?? '').trim();
    if (!value)
        return null;
    const sr = value.match(/\bSR\s*[-_:]?\s*0*(\d+(?:\.\d+)*)\b/i);
    const requirement = value.match(/\b(?:requirement|req)\s*#?\s*0*(\d+(?:\.\d+)*)\b/i);
    const bare = value.match(/^\s*0*(\d+(?:\.\d+)*)\s*$/);
    const id = sr?.[1] ?? requirement?.[1] ?? bare?.[1];
    if (!id)
        return null;
    return id.split('.').map(part => String(Number(part))).join('.');
}

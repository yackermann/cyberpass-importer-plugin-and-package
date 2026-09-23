import { normalizeRequirementId } from './normalizer.js';
import { VqImportError } from './errors.js';
function text(value) {
    return value == null ? '' : String(value).replace(/\u00a0/g, ' ').trim();
}
function header(value) {
    return value.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
const requirementHeaders = new Set(['sr', 'sr no', 'sr number', 'sr id', 'sar', 'sar no', 'requirement', 'requirement id', 'requirement no', 'requirement number']);
const responseHeaders = ['vendor response', 'response', 'answer', 'rationale'];
// Sparse iteration avoids allocating the worksheet's entire used rectangle.
function sheetRows(sheet) {
    const rows = new Map();
    for (const [address, raw] of Object.entries(sheet)) {
        const match = /^([A-Z]+)([1-9]\d*)$/i.exec(address);
        if (!match || !raw || typeof raw !== 'object')
            continue;
        const cell = raw;
        const rowNumber = Number(match[2]);
        let row = rows.get(rowNumber);
        if (!row)
            rows.set(rowNumber, row = new Map());
        row.set(match[1].toUpperCase(), text(cell.w ?? cell.v));
    }
    return new Map([...rows].sort(([a], [b]) => a - b));
}
/** Deterministic prediction from Vendor Response text; no compliance assessment. */
export function responseChoice(value) {
    return value.trimStart().toLowerCase().startsWith('n/a') ? 'N/A' : value.trim() ? 'YES' : 'NO';
}
/** Extract rows only from sheets containing both a requirement and response header. */
export function parseVendorQuestionnaire(workbook, options = {}) {
    const result = new Map();
    for (const sheetName of options.sheetNames ?? workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet)
            throw new VqImportError('SHEET_NOT_FOUND', `Worksheet "${sheetName}" was not found.`);
        let columns;
        for (const [rowNumber, cells] of sheetRows(sheet)) {
            const entries = [...cells];
            const requirementColumn = entries.find(([, value]) => requirementHeaders.has(header(value)))?.[0];
            // Prefer Vendor Response even if Answer or Response appears earlier.
            const responseColumn = responseHeaders.map(label => entries.find(([, value]) => header(value) === label)?.[0]).find(Boolean);
            if (requirementColumn && responseColumn) {
                columns = { requirement: requirementColumn, response: responseColumn };
                continue;
            }
            if (!columns)
                continue;
            const rawRequirementId = cells.get(columns.requirement) ?? '';
            const requirementId = normalizeRequirementId(rawRequirementId);
            if (!requirementId)
                continue;
            const value = cells.get(columns.response) ?? '';
            const item = { requirementId, rawRequirementId, value, predictedResponse: responseChoice(value), sheetName, row: rowNumber, column: columns.requirement, responseColumn: columns.response };
            const previous = result.get(requirementId);
            if (previous) {
                if (options.duplicates === 'first')
                    continue;
                if (options.duplicates !== 'last') {
                    throw new VqImportError('DUPLICATE_REQUIREMENT', `Requirement ${requirementId} appears at ${previous.sheetName}!${previous.column}${previous.row} and ${sheetName}!${columns.requirement}${rowNumber}. Select a worksheet or an explicit duplicate policy.`);
                }
            }
            result.set(requirementId, item);
        }
    }
    if (!result.size)
        throw new VqImportError('NO_REQUIREMENTS', 'No requirement rows detected. Expected SR No. (or Requirement ID) and Vendor Response headers on the same row.');
    return [...result.values()];
}
/** Decode a local File/Blob, ArrayBuffer, or Uint8Array into plain VQ rows. */
export async function readVendorQuestionnaire(input, options = {}) {
    let workbook;
    try {
        const bytes = 'arrayBuffer' in input ? await input.arrayBuffer() : input;
        const reader = options.reader ?? (await import('../vendor/reader.mjs')).default;
        workbook = await reader.read(bytes, { type: 'array', cellText: true, cellDates: true });
    }
    catch (cause) {
        const detail = cause instanceof Error ? cause.message : String(cause);
        const name = 'name' in input ? ` "${String(input.name)}"` : '';
        throw new VqImportError('READ_FAILED', `Could not read workbook${name}: ${detail}`, { cause });
    }
    return parseVendorQuestionnaire(workbook, options);
}

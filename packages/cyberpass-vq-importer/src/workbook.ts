import { normalizeRequirementId } from './normalizer.js';
import type { ExcelRequirement, SheetLike, WorkbookLike, XlsxReader } from './types.js';

interface Cell { v?: unknown; w?: string; }

function text(value: unknown): string {
  return value == null ? '' : String(value).replace(/\u00a0/g, ' ').trim();
}

function columnName(ref: string): string {
  return (ref.match(/^[A-Z]+/i)?.[0] ?? '').toUpperCase();
}

function columnNumber(ref: string): number {
  let number = 0;
  for (const character of columnName(ref)) number = number * 26 + character.charCodeAt(0) - 64;
  return number;
}

function rowNumber(ref: string): number {
  return Number(ref.match(/\d+/)?.[0] ?? 0);
}

function sheetToRows(sheet: SheetLike): { rows: string[][]; refs: string[][] } {
  const cells = Object.keys(sheet).filter(key => !key.startsWith('!'));
  const maxRow = Math.max(0, ...cells.map(rowNumber));
  const maxColumn = Math.max(0, ...cells.map(columnNumber));
  const rows = Array.from({ length: maxRow }, () => Array(maxColumn).fill('')) as string[][];
  const refs = Array.from({ length: maxRow }, () => Array(maxColumn).fill('')) as string[][];
  for (const ref of cells) {
    const cell = sheet[ref] as Cell | undefined;
    const row = rowNumber(ref) - 1;
    const column = columnNumber(ref) - 1;
    if (row < 0 || column < 0) continue;
    rows[row][column] = text(cell?.w ?? cell?.v);
    refs[row][column] = ref;
  }
  return { rows, refs };
}

const requirementHeader = /sr\s*(?:no|number)|requirement\s*(?:id|no|number)/i;
const responseHeader = /vendor\s*response|response|answer|rationale/i;

/** Parse a SheetJS-compatible workbook into normalized CyberPass requirements. */
export function parseVendorQuestionnaire(workbook: WorkbookLike): ExcelRequirement[] {
  const output: ExcelRequirement[] = [];
  for (const sheetName of workbook.SheetNames ?? []) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const { rows, refs } = sheetToRows(sheet);
    if (!rows.length) continue;
    const header = rows.slice(0, Math.min(12, rows.length)).find(row => row.some(value => requirementHeader.test(value) || responseHeader.test(value))) ?? rows[0];
    const requirementColumn = header.findIndex(value => requirementHeader.test(value));
    const responseColumn = header.findIndex(value => responseHeader.test(value));
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
      const row = rows[rowIndex];
      let raw = requirementColumn >= 0 ? text(row[requirementColumn]) : '';
      let requirementColumnIndex = requirementColumn;
      if (!raw) {
        for (let column = 0; column < row.length; column++) {
          if (normalizeRequirementId(row[column])) {
            raw = text(row[column]);
            requirementColumnIndex = column;
            break;
          }
        }
      }
      const requirementId = normalizeRequirementId(raw);
      if (!requirementId || rowIndex < 1) continue;
      let value = responseColumn >= 0 ? text(row[responseColumn]) : '';
      if (responseColumn < 0) {
        const candidates = row.map((cell, column) => ({ value: text(cell), column }))
          .filter(item => item.column !== requirementColumnIndex && item.value);
        value = candidates.at(-1)?.value ?? '';
      }
      output.push({
        requirementId,
        rawRequirementId: raw,
        value,
        sheetName,
        row: rowIndex + 1,
        column: refs[rowIndex]?.[requirementColumnIndex] ? columnName(refs[rowIndex][requirementColumnIndex]) : undefined,
        responseColumn: responseColumn >= 0 ? columnName(refs[rowIndex]?.[responseColumn] ?? '') : undefined
      });
    }
  }
  const deduplicated = new Map<string, ExcelRequirement>();
  for (const item of output) {
    const previous = deduplicated.get(item.requirementId);
    if (!previous || (!previous.value && item.value)) deduplicated.set(item.requirementId, item);
  }
  return [...deduplicated.values()];
}

/** Read and parse a local .xlsx/.xlsm file using the caller's SheetJS instance. */
export async function readVendorQuestionnaire(file: Blob, xlsx: XlsxReader): Promise<ExcelRequirement[]> {
  const workbook = await xlsx.read(await file.arrayBuffer(), { type: 'array', cellText: true, cellDates: true });
  return parseVendorQuestionnaire(workbook);
}

export function responseChoice(value: string): 'YES' | 'NO' | 'N/A' {
  return /^\s*n\/a\b/i.test(value) ? 'N/A' : value.trim() ? 'YES' : 'NO';
}

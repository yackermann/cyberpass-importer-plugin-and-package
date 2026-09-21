import type { ExcelRequirement } from '../shared/types';
import { normalizeRequirementId } from './normalizer';

interface Cell { v?: unknown; w?: string; }
interface SheetLike { '!ref'?: string; [key: string]: unknown; }

function text(v: unknown): string { return v == null ? '' : String(v).replace(/\u00a0/g, ' ').trim(); }
function colName(ref: string): string { return (ref.match(/^[A-Z]+/i)?.[0] ?? '').toUpperCase(); }
function rowNumber(ref: string): number { return Number(ref.match(/\d+/)?.[0] ?? 0); }
function sheetToRows(sheet: SheetLike): { rows: string[][]; refs: string[][] } {
  const cells = Object.keys(sheet).filter(k => !k.startsWith('!'));
  const maxRow = Math.max(0, ...cells.map(rowNumber));
  const maxCol = Math.max(0, ...cells.map(k => { let n=0; for (const c of colName(k)) n=n*26+c.charCodeAt(0)-64; return n; }));
  const rows = Array.from({length:maxRow},()=>Array(maxCol).fill(''));
  const refs = Array.from({length:maxRow},()=>Array(maxCol).fill(''));
  for (const ref of cells) { const c=sheet[ref] as Cell; let n=0; for (const ch of colName(ref)) n=n*26+ch.charCodeAt(0)-64; rows[rowNumber(ref)-1][n-1]=text(c?.w ?? c?.v); refs[rowNumber(ref)-1][n-1]=ref; }
  return {rows, refs};
}

/** Extract from a SheetJS workbook. Column detection is header-driven and intentionally isolated here. */
export function parseWorkbook(workbook: { SheetNames: string[]; Sheets: Record<string, SheetLike> }): ExcelRequirement[] {
  const output: ExcelRequirement[] = [];
  for (const sheetName of workbook.SheetNames ?? []) {
    const sheet = workbook.Sheets[sheetName]; if (!sheet) continue;
    const {rows, refs} = sheetToRows(sheet);
    if (!rows.length) continue;
    const header = rows.slice(0, Math.min(12, rows.length)).find(r => r.some(v => /sr\s*(no|number)|requirement\s*(id|no|number)|vendor response|response|answer/i.test(v)) ) ?? rows[0];
    const reqCol = header.findIndex(v => /sr\s*(no|number)|requirement\s*(id|no|number)/i.test(v));
    const responseCol = header.findIndex(v => /vendor\s*response|response|answer|rationale/i.test(v));
    for (let r=0; r<rows.length; r++) {
      const row=rows[r];
      let raw=''; let c=-1;
      if (reqCol >= 0) { raw = text(row[reqCol]); c=reqCol; }
      if (!raw) for (let i=0;i<row.length;i++) { if (normalizeRequirementId(row[i])) { raw=text(row[i]); c=i; break; } }
      const id = normalizeRequirementId(raw);
      if (!id || r < 1) continue;
      let value = responseCol >= 0 ? text(row[responseCol]) : '';
      if (responseCol < 0 && !value) {
        const candidates = row.map((v,i)=>({v:text(v),i})).filter(x=>x.i!==c && x.v);
        value = candidates.length ? candidates[candidates.length-1].v : '';
      }
      output.push({requirementId:id, rawRequirementId:raw, value, sheetName, row:r+1, column:refs[r]?.[c] ? colName(refs[r][c]) : undefined, responseColumn: responseCol >= 0 ? colName(refs[r]?.[responseCol] ?? '') : undefined});
    }
  }
  const dedup = new Map<string, ExcelRequirement>();
  for (const item of output) { const prior=dedup.get(item.requirementId); if (!prior || (!prior.value && item.value)) dedup.set(item.requirementId,item); }
  return [...dedup.values()];
}

export async function readWorkbookFile(file: File): Promise<ExcelRequirement[]> {
  const name=file.name.toLowerCase();
  if (name.endsWith('.csv') || name.endsWith('.tsv')) {
    const raw=await file.text(); const delim=name.endsWith('.tsv')?'\t':','; const rows=raw.split(/\r?\n/).map(line=>line.split(delim));
    return parseWorkbook({SheetNames:['CSV'],Sheets:{CSV: rows.reduce((s,row,r)=>{row.forEach((v,c)=>{s[`${String.fromCharCode(65+c)}${r+1}`]={v,w:v};});return s;},{'!ref':`A1:${String.fromCharCode(64+(rows[0]?.length||1))}${rows.length}`} as SheetLike)}});
  }
  const xlsx = (globalThis as any).XLSX;
  if (!xlsx?.read) throw new Error('SheetJS runtime is missing. Run the build to bundle xlsx into vendor/xlsx.mjs.');
  const workbook=xlsx.read(await file.arrayBuffer(), {type:'array', cellText:true, cellDates:true});
  return parseWorkbook(workbook);
}

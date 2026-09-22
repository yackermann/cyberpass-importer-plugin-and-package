// Full SheetJS reader with legacy Excel character encodings, bundled locally.
import { read, set_cptable } from './sheetjs/xlsx.mjs';
import * as cptable from './sheetjs/cpexcel.full.mjs';
set_cptable(cptable);
export const XLSX = { read };
export default XLSX;

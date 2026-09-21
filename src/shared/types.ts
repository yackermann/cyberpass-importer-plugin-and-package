export interface ExcelRequirement {
  requirementId: string;
  rawRequirementId: string;
  value: string;
  sheetName: string;
  row: number;
  column?: string;
  responseColumn?: string;
}

export interface CyberPassControl {
  tag: string;
  type?: string;
  name?: string;
  id?: string;
  ariaLabel?: string;
  value?: string;
  selectorHint?: string;
}

export interface CyberPassField {
  requirementId: string;
  element?: HTMLElement;
  label?: string;
  containerText: string;
  controls: CyberPassControl[];
}

export interface MappingRow {
  requirementId: string;
  excel?: ExcelRequirement;
  cyberPass?: CyberPassField;
  status: 'matched' | 'unmatched-excel' | 'unmatched-page' | 'empty' | 'mismatch' | 'filled' | 'skipped-existing' | 'failed';
  currentValue?: string;
  reason?: string;
}

export interface ScanResult {
  fields: CyberPassField[];
  debug: unknown[];
}

export interface FillOptions { replaceExisting: boolean; }
export interface FillResult { filled: number; skipped: number; failed: number; mismatches: number; errors: string[]; }

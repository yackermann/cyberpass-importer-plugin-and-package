export interface ParsedSheet { name: string; rows: unknown[][]; }
export interface WorkbookLike { SheetNames: string[]; Sheets: Record<string, unknown>; }

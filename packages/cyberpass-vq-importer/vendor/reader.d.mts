export declare const XLSX: {
  read(data: ArrayBuffer | Uint8Array, options?: Record<string, unknown>): {
    SheetNames: string[];
    Sheets: Record<string, { [cell: string]: unknown }>;
  };
};
export default XLSX;

import type { VqDocument } from './importer.js';
import type { ImportOptions } from './types.js';
export declare const VQ_FILE_ACCEPT = ".xls,.xlsx,.xlsm,.xlsb";
export interface VqFileEvents {
    options?: ImportOptions;
    onImport: (vq: VqDocument, file: File) => void | Promise<void>;
    onError: (error: Error) => void;
}
/** Attach to an existing file input; returns an unbind function for component cleanup. */
export declare function bindVqFileInput(input: HTMLInputElement, events: VqFileEvents): () => void;
/** Handle file drops only on the given element; leaves presentation to the host. */
export declare function bindVqDropzone(element: HTMLElement, events: VqFileEvents): () => void;

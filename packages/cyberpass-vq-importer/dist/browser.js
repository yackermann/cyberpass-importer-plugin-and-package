import { importVq } from './importer.js';
import { VqImportError } from './errors.js';
export const VQ_FILE_ACCEPT = '.xls,.xlsx,.xlsm,.xlsb';
function receiver(events) {
    let generation = 0;
    return {
        dispose() { generation++; },
        async receive(files) {
            const current = ++generation;
            if (!files?.length)
                return;
            try {
                if (files.length !== 1)
                    throw new VqImportError('FILE_COUNT', 'Choose one VQ workbook at a time.');
                const file = files[0];
                const vq = await importVq(file, events.options);
                if (current === generation)
                    await events.onImport(vq, file);
            }
            catch (error) {
                if (current === generation)
                    events.onError(error instanceof Error ? error : new Error(String(error)));
            }
        }
    };
}
/** Attach to an existing file input; returns an unbind function for component cleanup. */
export function bindVqFileInput(input, events) {
    if (input.type !== 'file')
        throw new TypeError('bindVqFileInput requires an input with type="file".');
    const accept = input.accept;
    if (!accept)
        input.accept = VQ_FILE_ACCEPT;
    const handler = receiver(events);
    const change = () => {
        const pending = handler.receive(input.files);
        input.value = ''; // Allow choosing the same workbook again.
        void pending;
    };
    input.addEventListener('change', change);
    return () => {
        handler.dispose();
        input.removeEventListener('change', change);
        if (!accept)
            input.accept = accept;
    };
}
/** Handle file drops only on the given element; leaves presentation to the host. */
export function bindVqDropzone(element, events) {
    const handler = receiver(events);
    const dragover = (event) => {
        if (!event.dataTransfer?.types.includes('Files'))
            return;
        event.preventDefault();
        event.dataTransfer.dropEffect = 'copy';
    };
    const drop = (event) => {
        if (!event.dataTransfer?.files.length)
            return;
        event.preventDefault();
        void handler.receive(event.dataTransfer.files);
    };
    element.addEventListener('dragover', dragover);
    element.addEventListener('drop', drop);
    return () => {
        handler.dispose();
        element.removeEventListener('dragover', dragover);
        element.removeEventListener('drop', drop);
    };
}

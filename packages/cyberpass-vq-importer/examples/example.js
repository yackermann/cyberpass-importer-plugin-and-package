// Served from the package directory, these relative imports need no bundler/CDN.
import { importVq, bindVqFileInput, bindVqDropzone } from '../dist/index.js';

const element = id => document.getElementById(id);
const input = element('file-input');
const dropzone = element('dropzone');
const status = element('status');
const sample = element('sample');
let documentVq = null;
let busy = false;

function setStatus(message, kind = '') {
  status.textContent = message;
  status.dataset.kind = kind;
}

function setBusy(value) {
  busy = value;
  input.disabled = sample.disabled = value;
  dropzone.setAttribute('aria-busy', String(value));
}

function reset() {
  documentVq = null;
  element('results').hidden = true;
  element('rows').replaceChildren();
  element('lookup').value = '';
  element('lookup-result').textContent = 'Enter an ID to see vq.get(id).';
}

function start() {
  reset();
  setBusy(true);
  setStatus('Reading workbook…');
}

function lookup() {
  const id = element('lookup').value.trim();
  const row = documentVq?.get(id);
  element('lookup-result').textContent = !id
    ? 'Enter an ID to see vq.get(id).'
    : row ? JSON.stringify(row, null, 2) : `No requirement found for "${id}".`;
}

// These callbacks belong to the example application. The library only imports data.
const events = {
  onImport(vq, file) {
    documentVq = vq;
    const fragment = document.createDocumentFragment();
    for (const row of vq.rows) {
      const tr = document.createElement('tr');
      const id = document.createElement('td');
      id.textContent = row.requirementId;
      const response = document.createElement('td');
      response.className = 'response';
      response.textContent = row.value || '(empty)';
      const prediction = document.createElement('td');
      const badge = document.createElement('span');
      badge.className = `badge ${{ YES: 'yes', NO: 'no', 'N/A': 'na' }[row.predictedResponse]}`;
      badge.textContent = row.predictedResponse;
      prediction.append(badge);
      const source = document.createElement('td');
      source.className = 'source';
      source.textContent = `${row.sheetName} · ${row.responseColumn}${row.row}`;
      tr.append(id, response, prediction, source);
      fragment.append(tr);
    }
    // Workbook values are rendered as text, never interpreted as HTML.
    element('rows').replaceChildren(fragment);
    element('file-name').textContent = file.name;
    const count = choice => vq.rows.filter(row => row.predictedResponse === choice).length;
    element('summary').textContent = `${vq.rows.length} requirements · ${count('YES')} YES · ${count('NO')} NO · ${count('N/A')} N/A`;
    element('results').hidden = false;
    setStatus(`Imported ${vq.rows.length} requirements from ${file.name}.`);
    setBusy(false);
  },
  onError(error) {
    reset();
    setBusy(false);
    setStatus(error.message, 'error');
  }
};

// Start the example's loading state before the library's event listeners run.
input.addEventListener('change', () => { if (input.files?.length) start(); });
dropzone.addEventListener('drop', event => {
  dropzone.classList.remove('dragging');
  if (!event.dataTransfer?.files.length) return;
  if (busy) {
    event.preventDefault();
    event.stopImmediatePropagation();
    return;
  }
  start();
});
dropzone.addEventListener('dragover', event => {
  if (!busy && event.dataTransfer?.types.includes('Files')) dropzone.classList.add('dragging');
});
dropzone.addEventListener('dragleave', event => {
  if (!dropzone.contains(event.relatedTarget)) dropzone.classList.remove('dragging');
});
const unbindInput = bindVqFileInput(input, events);
const unbindDrop = bindVqDropzone(dropzone, events);

sample.addEventListener('click', async () => {
  start();
  try {
    const response = await fetch('./sample-vq.xlsx');
    if (!response.ok) throw new Error(`Could not load sample workbook (HTTP ${response.status}).`);
    const file = new File([await response.arrayBuffer()], 'sample-vq.xlsx');
    events.onImport(await importVq(file), file);
  } catch (error) {
    events.onError(error);
  }
});
element('lookup').addEventListener('input', lookup);
element('clear').addEventListener('click', () => {
  reset();
  setStatus('Cleared. Choose another workbook or try the sample.');
});
element('export').addEventListener('click', () => {
  if (!documentVq) return;
  const url = URL.createObjectURL(new Blob([JSON.stringify(documentVq.rows, null, 2)], { type: 'application/json' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = 'vq-rows.json';
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
});
window.addEventListener('pagehide', event => {
  if (!event.persisted) { unbindInput(); unbindDrop(); }
});
if (location.protocol === 'file:') setStatus('Serve this package over HTTP. See examples/README.md for the local server command.', 'error');

# Integrating the generic VQ importer

CyberPass and other questionnaire platforms can use this package to decode a local Excel VQ into application data. It includes XLS/XLSX/XLSM/XLSB reading, requirement normalization, source locations, response prediction, and optional file input/drop-zone bindings.

## Minimal integration

```ts
import { importVq } from '@yuriy-ackermann/cyberpass-vq-importer';

const vq = await importVq(file);
const rows = vq.rows;
const requirement = vq.get('1.1');
```

A row has `requirementId`, `value`, `predictedResponse` (`YES` / `NO` / `N/A`), and its source worksheet/cell coordinates. The application decides how to map those into field names and answer enums.

For example, a frontend might map the data into its own model:

```ts
const proposedChanges = vq.rows.map(row => ({
  id: row.requirementId,
  comment: row.value,
  response: row.predictedResponse
}));
```

`proposedChanges` is an ordinary array, not a patch format imposed by the library. The host owns matching against known requirements, highlighting conflicts, overwrite policy, form updates, validation, and persistence. Predictions describe the response text and should be reviewed according to the host's workflow.

## Native file input

In a framework event handler, pass the selected File to `importVq` and consume the promise. For plain DOM integration:

```ts
import { bindVqFileInput } from '@yuriy-ackermann/cyberpass-vq-importer';

const unbind = bindVqFileInput(fileInput, {
  onImport: vq => setImportedRows(vq.rows),
  onError: error => showError(error.message)
});
```

`fileInput` is your existing `<input type="file">`. `setImportedRows` and `showError` are your application's functions. Call `unbind()` when the view unmounts.

## Drag and drop

```ts
import { bindVqDropzone } from '@yuriy-ackermann/cyberpass-vq-importer';

const unbind = bindVqDropzone(dropzone, {
  onImport: vq => setImportedRows(vq.rows),
  onError: error => showError(error.message)
});
```

The element, labels, keyboard alternative, loading UI, and presentation belong to the host. This helper reads one dropped file and returns data. It does not inspect questionnaire controls, scroll, click dropdowns, write comments, or submit anything.

## Receiving the package

Until it is published, install the archive provided by the author:

```sh
npm install ./yuriy-ackermann-cyberpass-vq-importer-0.2.0.tgz
```

Once published under the configured scope:

```sh
npm install @yuriy-ackermann/cyberpass-vq-importer
```

The package includes its Excel decoder, so importing a workbook needs only the package import shown above. Data stays in memory unless the host chooses to persist or transmit it. See the [README](./README.md) for workbook rules, duplicate handling, errors, all exports, and the version 0.1 migration.

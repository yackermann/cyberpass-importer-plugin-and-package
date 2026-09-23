# Vendor Questionnaire importer

`@yuriy-ackermann/cyberpass-vq-importer` imports standardized VQ Excel workbooks into typed data. The name is retained for continuity; the package works independently of CyberPass and any frontend framework.

Import a file, iterate its requirement rows, or look up a requirement by ID. Each row includes the Vendor Response text and a deterministic Yes/No/N/A prediction. The host application owns review, field mapping, form updates, and saving.

## Runnable HTML example

See [examples/index.html](./examples/index.html) and the [run instructions](./examples/README.md). Serve this package directory with `python3 -m http.server 8080 --bind 127.0.0.1`, then open `http://127.0.0.1:8080/examples/`. The example includes a fictional sample workbook, file selection, drag and drop, a row table, `.get()` lookup, and JSON export. It ships in the package archive.

## Install

After publishing (or install a locally packed `.tgz`):

```sh
npm install @yuriy-ackermann/cyberpass-vq-importer
```

The package includes SheetJS 0.20.3 with legacy character encodings. No separate `xlsx` install, global variable, or reader configuration is required. The decoder loads on the first file import. Supported Excel formats: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`. Files are processed locally; the package has no upload or persistence behavior.

## Import and use the data

```ts
import { importVq } from '@yuriy-ackermann/cyberpass-vq-importer';

const vq = await importVq(file);

const rows = vq.rows;
const requirement = vq.get('1.1'); // VqRow | undefined

for (const { requirementId, value, predictedResponse } of vq) {
  // Apply these values through your application's own state/form API.
}
```

`file` can be a browser `File`/`Blob`, an `ArrayBuffer`, or a `Uint8Array` (including a Node.js Buffer). The core importer works without a DOM.

Example row:

```ts
{
  requirementId: '1.1',
  rawRequirementId: 'SR 1.1',
  value: 'N/A — this feature is not supported.',
  predictedResponse: 'N/A',
  sheetName: 'FIDO Security Requirements',
  row: 4,
  column: 'B',
  responseColumn: 'I'
}
```

`row` is a one-based Excel row number. Columns are Excel letters. `.get('SR 001.01')` and `.get('1.1')` look up the same normalized ID. The document's rows are immutable. To get a mutable plain array instead:

```ts
import { readVendorQuestionnaire } from '@yuriy-ackermann/cyberpass-vq-importer';

const rows = await readVendorQuestionnaire(file);
```

## Response prediction

| Vendor Response text | `predictedResponse` |
| --- | --- |
| Starts with `N/A`, ignoring case and leading whitespace | `N/A` |
| Any other non-empty content | `YES` |
| Empty or whitespace only | `NO` |

This is a text-based suggestion, not a compliance assessment. Even the literal text `No` is non-empty and predicts `YES`, following the import rule. Use `responseChoice(value)` to apply this rule independently.

## File input and drag and drop

Use the importer directly in your framework's event handler, or bind existing elements:

```ts
import {
  bindVqFileInput,
  bindVqDropzone,
  type VqFileEvents
} from '@yuriy-ackermann/cyberpass-vq-importer';

// fileInput and dropzone are elements owned by your application.
// setImportedRows and showError are your application's callbacks.
const events: VqFileEvents = {
  onImport: vq => setImportedRows(vq.rows),
  onError: error => showError(error.message)
};

const unbindInput = bindVqFileInput(fileInput, events);
const unbindDrop = bindVqDropzone(dropzone, events);

// When your view unmounts:
unbindInput();
unbindDrop();
```

The helpers accept one file at a time, pass `VqDocument` and the original `File` to `onImport`, and report failures through the required `onError`. They do not render UI or set application fields. Provide a visible file input or accessible button as a keyboard alternative to dropping a file.

The input helper sets the accepted extensions if none were specified and resets the input so the same file can be selected again. The drop helper handles only file drops on its element. Each binding suppresses results from an older pending import when a new file arrives; unbinding removes listeners and suppresses pending callbacks. Separate bindings are independent. Callbacks already running are not canceled.

## Workbook mapping and options

Worksheets must contain a header row with both a requirement column and response column. The parser recognizes `SR No.`, `SR Number`, `SR ID`, `SR`, `SAR`, `SAR No.`, `Requirement`, `Requirement ID`, `Requirement No.`, or `Requirement Number`. Header matching ignores case and punctuation.

`Vendor Response` takes precedence over `Response`, `Answer`, and `Rationale`. Rows before the header, blank requirement IDs, non-ID prose, and sheets without paired headers are ignored. Repeated header rows are supported. The parser never substitutes requirement description text for a missing response column. An empty response cell for a valid requirement remains an imported row with a `NO` prediction.

Cell display text is used when available, with raw values as fallback. Surrounding whitespace is trimmed and nonbreaking spaces become spaces; internal line breaks and Unicode are retained. Formula results use cached workbook values; formulas and macros are not executed. Formatting, images, attachments, and requirement prose are not included in the returned rows.

```ts
const vq = await importVq(file, {
  sheetNames: ['FIDO Security Requirements'],
  duplicates: 'error' // default; alternatives: 'first' or 'last'
});
```

Repeated normalized IDs raise an error with both source locations by default. Select worksheets or an explicit duplicate policy when a workbook intentionally repeats IDs. These options also work with `parseVendorQuestionnaire` and `readVendorQuestionnaire`; browser helpers accept them under `events.options`.

An optional `reader` in `ImportOptions` accepts a SheetJS-compatible sync or async decoder. It replaces the built-in decoder for that call. Use `parseVendorQuestionnaire(workbook, options)` if the application already has a decoded workbook.

## Errors

```ts
import { importVq, VqImportError } from '@yuriy-ackermann/cyberpass-vq-importer';

try {
  const vq = await importVq(file);
  setImportedRows(vq.rows);
} catch (error) {
  if (error instanceof VqImportError) {
    showError(error.message); // error.code is available for custom handling
  } else {
    throw error;
  }
}
```

| Code | Meaning |
| --- | --- |
| `READ_FAILED` | Workbook could not be decoded/read; original error in `cause` |
| `NO_REQUIREMENTS` | No rows under recognized paired headers |
| `DUPLICATE_REQUIREMENT` | Repeated normalized requirement ID |
| `SHEET_NOT_FOUND` | Requested worksheet missing |
| `FILE_COUNT` | A browser helper received multiple files |

Password-encrypted files require an unencrypted copy. Host applications decide how errors and predictions are presented to users.

## API

| Export | Result |
| --- | --- |
| `importVq(input, options?)` | `Promise<VqDocument>` with `.rows`, `.get(id)`, iteration |
| `readVendorQuestionnaire(input, options?)` | `Promise<VqRow[]>` |
| `parseVendorQuestionnaire(workbook, options?)` | `VqRow[]` |
| `normalizeRequirementId(value)` | Normalized string or `null` |
| `responseChoice(value)` | `YES`, `NO`, or `N/A` |
| `bindVqFileInput(element, events)` | Cleanup function |
| `bindVqDropzone(element, events)` | Cleanup function |
| `VQ_FILE_ACCEPT` | File input accept string |
| `VqImportError` | Error class with stable code |

## Migration from 0.1

Version 0.2 removes `CyberPassVqImporter`, `installImportExcelButton`, and the DOM/form-specific types and behavior. Replace them with `importVq` or file event helpers, then map `vq.rows` into your application's state. A custom reader is now passed as `{ reader: XLSX }`, not as a positional argument. Rows include `predictedResponse`. Missing headers and duplicate IDs now produce explicit errors.

## Build and publish

From this package directory:

```sh
npm install
npm run typecheck
npm test
npm pack --dry-run
npm pack
```

`npm test` builds the package and tests its public API. `npm pack` creates an installable archive to send to an integration team. To publish, confirm you control the npm scope in `package.json` (or change it), then run `npm publish --access public`. The package is ESM with TypeScript declarations. See [INTEGRATION.md](./INTEGRATION.md) for the application handoff.

## Attribution and trademark notice

Made by **Yuriy Ackermann**. Package code is MIT licensed. The included SheetJS decoder has its own Apache-2.0 license in `vendor/sheetjs/LICENSE`.

FIDO, FIDO Alliance, CyberPass, and related names, marks, and logos belong to their respective owners. This is an independent, unofficial tool, not sponsored, endorsed, or affiliated with FIDO Alliance or CyberPass. No ownership of those trademarks is claimed by the author.

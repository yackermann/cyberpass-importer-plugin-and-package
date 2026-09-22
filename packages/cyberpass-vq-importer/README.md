# `@yuriy-ackermann/cyberpass-vq-importer`

Browser-first TypeScript library for adding an **Import Excel VQ** workflow to a FIDO CyberPass vendor questionnaire.

It provides:

- Header-driven parsing of `.xls`, `.xlsx`, `.xlsm`, and `.xlsb` workbooks through a caller-supplied SheetJS reader.
- Normalized matching of `SR001`, `SR 1`, `Requirement 1.1`, and dotted requirement IDs.
- The FIDO/CyberPass response rule: an empty Vendor Response selects **No**, content beginning with `N/A` selects **N/A**, and other content selects **Yes**.
- A CyberPass DOM adapter that discovers lazy-loaded requirement cards, clicks the Ant Design response menu, fills the comment textarea, and reports mismatches and failures.
- `installImportExcelButton`, a small integration helper that mounts a file picker button on an existing page.

## Install

```bash
npm install @yuriy-ackermann/cyberpass-vq-importer xlsx
```

`xlsx` is a peer dependency because the host application chooses how to bundle or load its workbook reader.

## Quick integration

```ts
import * as XLSX from 'xlsx';
import { installImportExcelButton } from '@yuriy-ackermann/cyberpass-vq-importer';

installImportExcelButton(XLSX, {
  button: '#import-excel-vq',
  status: '#import-excel-vq-status'
});
```

The helper updates the status element as it loads and fills. Use `onLoaded`, `onFilled`, or `onError` only when the application needs additional telemetry or custom notifications. `button` and `status` may also receive the already queried DOM elements.

The helper does not submit, save, or advance the procedure. It only updates the currently rendered questionnaire controls. It progressively scrolls the page to discover cards that CyberPass lazy-loads.

## Separate parsing and filling

For applications that want their own button and confirmation flow:

```ts
import * as XLSX from 'xlsx';
import {
  CyberPassVqImporter,
  readVendorQuestionnaire,
  responseChoice
} from '@yuriy-ackermann/cyberpass-vq-importer';

const requirements = await readVendorQuestionnaire(file, XLSX);
const importer = new CyberPassVqImporter(document);
const fields = await importer.scan();
const result = await importer.fill(requirements, { replaceExisting: false });
```

`result.errors` contains field-specific messages such as `1.6: could not set Response to N/A`. Existing non-empty values are skipped and counted in `mismatches` unless `replaceExisting: true` is explicitly used.

## Workbook contract

The parser searches the first 12 rows of every worksheet for headers matching:

- Requirement: `SR No.`, `SR Number`, `Requirement`, `Requirement ID`, or equivalent.
- Response: `Vendor Response`, `Response`, `Answer`, or `Rationale`.

The parser returns one normalized row per requirement. Duplicate IDs are deduplicated, preferring a row with a non-empty response. A workbook with no recognized rows returns an empty array so the host can show a useful error.

## CyberPass page contract

The adapter is intended for the FIDO CyberPass questionnaire step. It recognizes headings such as `.input-node-view-builder-header` containing `Requirement 1.1`, then uses the nearest `.input-node-view-builder-container`. It expects:

- A response control whose ID ends in `.answer`.
- An editable comment control, normally the textarea whose ID ends in `.description`.
- Ant Design response options labelled `✅ Yes`, `❌ No`, and `🚫 N/A` (native `<select>` controls are also supported).

If CyberPass changes these DOM contracts, the host should pin the package version and update the adapter after testing the new rendered HTML.

## Building and publishing

From this package directory:

```bash
npm install
npm run typecheck
npm run build
npm publish --access public
```

Before publishing, update the version in `package.json`, review the generated `dist/` files, and run the integration against a test CyberPass procedure. The package is browser code; bundle it with the CyberPass application rather than loading it from an untrusted CDN.

## Attribution and trademark notice

Made by **Yuriy Ackermann**.

FIDO, FIDO Alliance, CyberPass, and related names, marks, and logos belong to their respective owners. This package is an independent, unofficial integration and is not sponsored, endorsed, administered by, or affiliated with FIDO Alliance, CyberPass, or their respective owners. No ownership of those trademarks is claimed by the author.

Legacy `.xls` imports require a full SheetJS reader. If using the SheetJS ESM build, register its `cpexcel.full.mjs` tables with `set_cptable` for older workbook character encodings. The bundled reader in the browser extension already includes these tables.

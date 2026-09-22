# CyberPass integration handoff

This note is for the CyberPass application team integrating the library into the FIDO vendor-questionnaire step.

## 1. Add the package

```bash
npm install @yuriy-ackermann/cyberpass-vq-importer xlsx
```

The package is browser code and should be bundled with the CyberPass frontend. Do not load a package script from an arbitrary CDN.

## 2. Add a button

```html
<button id="import-excel-vq" type="button">Import Excel VQ</button>
<div id="import-excel-vq-status" role="status"></div>
```

```ts
import * as XLSX from 'xlsx';
import { installImportExcelButton } from '@yuriy-ackermann/cyberpass-vq-importer';

installImportExcelButton(XLSX, {
  button: '#import-excel-vq',
  status: '#import-excel-vq-status'
});
```

The library updates the status element automatically. Add `onLoaded`, `onFilled`, or `onError` only for custom notifications or telemetry.

The button may be rendered beside the questionnaire header. If it is created after the questionnaire route mounts, call `installImportExcelButton` after the button exists.

## 3. Confirm behavior

Use a test procedure and verify:

1. An empty Vendor Response selects `❌ No`.
2. A response beginning with `N/A` selects `🚫 N/A` case-insensitively.
3. Any other non-empty response selects `✅ Yes`.
4. The Vendor Response text is inserted into the Comment field.
5. Existing values remain unchanged by default and are counted as mismatches when different.
6. `replaceExisting: true` requires an explicit product decision because it overwrites existing questionnaire values.
7. The page may scroll while lazy-rendered requirements are discovered.
8. No Save, Submit, or Next action is performed by the library.

## 4. Browser and data behavior

Workbook bytes are read in the browser by the host's SheetJS instance. The library does not upload the file, call a server, or persist workbook contents. The host controls any telemetry and should avoid logging workbook response text.

## 5. Versioning

Pin a package version in production. The adapter intentionally targets the current CyberPass questionnaire DOM contract: requirement headings, `.input-node-view-builder-container`, `.answer` response controls, and `.description` textareas. Update the package after any intentional DOM change and run the checklist above.

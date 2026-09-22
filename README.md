# CyberPass Excel Importer

A local-only Chrome/Edge Manifest V3 extension that restores the CyberPass vendor-questionnaire workflow on `https://app.fido.cyber-pass.org/procedures/{id}?step=fido_user_authenticator_vendor_questionnaire`: drop an Excel workbook, preview the requirement mapping, then fill the active assessment without submitting it.

## Install

1. Run `npm install` in this directory to install the build tools. SheetJS 0.20.3 and its legacy character encodings are vendored locally; no runtime download is needed.
2. Run `npm run build`.
3. Open `chrome://extensions` (or `edge://extensions`), enable **Developer mode**, choose **Load unpacked**, and select this directory.
4. Open a CyberPass assessment at the vendor-questionnaire step, click the extension, and drop a `.xls`, `.xlsx`, `.xlsm`, or `.xlsb` workbook. The extension is inactive on all other sites and procedure steps.

The bundled full SheetJS reader handles legacy binary `.xls` workbooks (including Excel 97–2003), `.xlsx`, `.xlsm`, and `.xlsb` locally. Vendor Response text, line breaks, and international characters are preserved. Reader errors include the filename and are displayed in the popup. `npm test` exercises binary XLS imports and regression checks for the modern formats. See [vendor/sheetjs/README.md](./vendor/sheetjs/README.md) for provenance and license.

## Workflow

- **Fill all** progressively scrolls the questionnaire so lazy-loaded requirements are discovered, then fills matching fields. Existing non-empty values that differ are shown as warnings and outlined in orange.
- The page adapter watches DOM mutations and scrolling, so per-field helpers are added as new requirement batches appear.
- **Fill all** writes only empty fields by default. Enable **Replace existing values** to overwrite after reviewing the warnings. The CyberPass answer control is derived from the workbook response: content beginning with `N/A` selects `N/A`, other content selects `YES`, and an empty response selects `NO`.
- **Override requirement colour** applies a local, page-scoped style that makes `.input-node-view-builder-description` text black. The setting is saved locally and reapplied on page reload.
- Parsed workbook requirements are saved in browser-session storage keyed to the procedure ID, so closing the popup or refreshing the page does not discard the mapping. The drop zone shows the saved filename and provides **Clear** to remove it for the current procedure. The background service worker stores the parsed workbook in browser-session storage keyed by procedure ID. The popup and content script request it from the worker, and the content script restores the saved mapping automatically after a page refresh. The mapping is cleared when the browser session ends and is never reused for a different procedure.
- Each mapped field also gets a small **Fill from Excel** helper button so individual fields can be filled one at a time. Clicking an empty comment textarea shows an autofill suggestion with the workbook’s `Vendor Response` value; select **Fill from Excel** to insert it. A confirmation is shown before replacing existing content.
- The extension never clicks Submit, Save, Next, or any other workflow button.

## Workbook mapping

`src/excel/parser.ts` keeps extraction separate from the rest of the extension. It looks for a requirement column headed `SR No.`/`Requirement`, recognizes `SR001`, `SR 12`, `Requirement #12`, and dotted IDs such as `6.5`, and prefers a `Vendor Response`/`Response`/`Answer` column. The supplied FIDO workbook maps `FIDO Security Requirements!B:B` (SR No.) to `I:I` (Vendor Response).

## Architecture

- `src/excel/normalizer.ts` — shared requirement-ID normalization.
- `src/excel/parser.ts` — SheetJS-compatible workbook extraction.
- `src/content/cyberpass.ts` — generic CyberPass DOM adapter: scan, safe native-value filling, mismatch marking, and per-field helpers.
- `src/content/content.ts` — message bridge.
- `src/popup/*` — file drop, preview, controls, sanitized debug export.
- `src/shared/*` — shared types and messages.

## Reusable CyberPass integration library

The publishable package in [`packages/cyberpass-vq-importer`](./packages/cyberpass-vq-importer) exposes the workbook parser and the browser DOM adapter for CyberPass itself. It is designed so CyberPass can add a native **Import Excel VQ** button without adopting the extension UI or Chrome APIs.

```bash
npm install @yuriy-ackermann/cyberpass-vq-importer xlsx
```

See the package [README](./packages/cyberpass-vq-importer/README.md) for the integration snippet, API, workbook contract, error handling, build, and publishing instructions. Build it from this repository with `npm run build:library`.

## Debugging DOM detection

The content adapter recognizes `.input-node-view-builder-header` text such as `Requirement 6.5` and searches its nearest `.input-node-view-builder-container` for `textarea`, text inputs, selects, and contenteditable controls. Its debug representation contains requirement IDs and control metadata (`tag`, `id`, `name`, `aria-label`) without field values, cookies, tokens, passwords, or unrelated page data.

To make matching deterministic, provide a saved DOM fragment for one or two representative requirements (including the requirement heading and its editable response control), plus whether the response should go into the answer select, comment textarea, or another control. A screenshot alone is useful for layout but does not reveal the control attributes needed for a deterministic adapter.

## Attribution and trademark notice

Made by **Yuriy Ackermann**.

FIDO, FIDO Alliance, CyberPass, and any related names, marks, or logos are trademarks of their respective owners. This project is an independent, unofficial tool. It is not sponsored, endorsed, administered by, or affiliated with FIDO Alliance, CyberPass, or their respective owners. No ownership of those trademarks is claimed by this project or its author.

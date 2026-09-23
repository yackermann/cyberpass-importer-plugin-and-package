# HTML example

This plain HTML and JavaScript page uses the package's public API. It has a file picker, drag and drop, sample workbook, imported-row table, requirement lookup, errors, and JSON download. All sample responses are fictional.

From the **package directory** (`packages/cyberpass-vq-importer` in this repository):

```sh
python3 -m http.server 8080 --bind 127.0.0.1
```

Open **http://127.0.0.1:8080/examples/** and click **Try sample workbook**, or choose/drop your VQ file. Stop the server with Ctrl+C. Serve the package directory, not just `examples/`, because the page imports `../dist/index.js` and the bundled decoder. JavaScript modules require an HTTP server; double-clicking the HTML file is not supported.

The published archive includes the built JavaScript. When modifying TypeScript in a source checkout, run `npm install` and `npm run build` before serving.

- `index.html`: page structure and styles.
- `example.js`: host application behavior, using `importVq`, `bindVqFileInput`, and `bindVqDropzone`.
- `sample-vq.xlsx`: a small workbook covering YES, NO, N/A, multiline text, and Unicode.

The page renders data returned by the package. It does not connect to a questionnaire platform or apply values to forms. Use the callbacks in `example.js` as the integration point for your own application state.

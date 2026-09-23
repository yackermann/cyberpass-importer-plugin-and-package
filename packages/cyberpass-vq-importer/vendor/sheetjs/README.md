# SheetJS Community Edition 0.20.3

Unmodified files from the [official release](https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz), following the [SheetJS installation instructions](https://docs.sheetjs.com/docs/getting-started/installation/nodejs/).

- `xlsx.mjs` comes from `package/xlsx.mjs`.
- `cpexcel.full.mjs` comes from `package/dist/cpexcel.full.mjs` and supplies legacy character encodings.
- `LICENSE` comes from `package/LICENSE` (Apache-2.0).

`../reader.mjs` registers the codepage tables and exposes the reader. The extension bundles this code locally into `dist/popup.js`; files are never fetched at runtime. The generic import library includes this decoder by default and optionally accepts the host application's own reader.

SHA-256:

```text
1a0fb062ee9781b13f6687371b202aaefc53b6ce55b530c027e01f9c087b77db  xlsx.mjs
7a7bba23b6b6f23b5c69fbb631f78d1d455f74b57e4aa54e9a2f81a8ab844964  cpexcel.full.mjs
```

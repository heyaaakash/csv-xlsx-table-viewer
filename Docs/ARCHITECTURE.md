# Architecture

Spreadsheet Viewer is a Node-based VS Code extension with a custom-editor webview.

## How a file is opened

1. `package.json` registers CSV, TSV, XLSX, and XLS files with the custom editor.
2. `src/extension.js` reads the opened document, parses delimited text or loads workbook sheets with `xlsx`, and sends the data and editor settings to the webview.
3. The webview renders the table and sends edits and view actions back to the extension host.
4. VS Code's custom-document lifecycle handles save, revert, and backup.

## Build layout

`src/extension.js` is the extension entry point. `scripts/build.js` bundles it and `xlsx` into `dist/<version>/extension.js`; the VS Code API stays external. The manifest points to the current version's bundle, and packaging runs the build first.

This extension uses Node APIs and does not support VS Code's browser host.

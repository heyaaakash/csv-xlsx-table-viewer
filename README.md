# Spreadsheet Viewer

**Inspect and edit CSV, TSV, XLSX, and XLS files without leaving VS Code.** Open a file as a table, find the row or value you need, make a change, and save it back in its original format.

<p align="center">
  <a href="https://marketplace.visualstudio.com/items?itemName=aakashrohilla.csv-xlsx-table-viewer">
    <img src="https://shieldcn.dev/vscode/installs/aakashrohilla/csv-xlsx-table-viewer.svg" alt="Marketplace Installs" />
  </a>
  <a href="https://marketplace.visualstudio.com/items?itemName=aakashrohilla.csv-xlsx-table-viewer">
    <img src="https://shieldcn.dev/vscode/rating/aakashrohilla/csv-xlsx-table-viewer.svg" alt="Marketplace Rating" />
  </a>
  <a href="https://marketplace.visualstudio.com/items?itemName=aakashrohilla.csv-xlsx-table-viewer">
    <img src="https://shieldcn.dev/vscode/v/aakashrohilla/csv-xlsx-table-viewer.svg" alt="Marketplace Version" />
  </a>
</p>

[Install from the Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=aakashrohilla.csv-xlsx-table-viewer) · [Source code](https://github.com/heyaaakash/csv-xlsx-table-viewer) · [Read the product story](Docs/PRODUCT_STORY.md)

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

![Spreadsheet Viewer showing a table in VS Code](https://1vrxnmjjog.ucarecd.net/6614e201-8c92-4f36-ace6-e8efef029d1d/image.png)

## Start in under a minute

1. [Install CSV XLSX Table Viewer](https://marketplace.visualstudio.com/items?itemName=aakashrohilla.csv-xlsx-table-viewer) from the Marketplace.
2. Open a `.csv`, `.tsv`, `.xlsx`, or `.xls` file in VS Code.
3. Search, sort, switch workbook sheets, or edit a cell and save.

You can also install from the VS Code Extensions view (`⌘⇧X` on macOS, `Ctrl+Shift+X` on Windows/Linux) by searching for **CSV XLSX Table Viewer**, or from a terminal where the `code` command is available:

```sh
code --install-extension aakashrohilla.csv-xlsx-table-viewer
```

If a file opens as plain text, choose **Open as Table** from the editor title or Explorer context menu.

## A small workflow, kept in your editor

Spreadsheet Viewer turns a file that usually sends you to another app into an editor surface beside your code. CSV and TSV files become searchable tables. Excel workbooks keep their sheets together. When you make a change, save writes the data back to the original file and format.

![Spreadsheet Viewer preview mode](https://1vrxnmjjog.ucarecd.net/11d2a5fc-f7b4-4a20-9618-6081dd9c6875/image.png)

### What you can do

- Open CSV, TSV, XLSX, and XLS in a table editor.
- Search and sort data, with raw and summary views for inspection.
- Edit cells and save supported files in place.
- Move between workbook sheets with clickable tabs or `Alt+←` / `Alt+→`.
- Undo and redo edits and sorts with VS Code's standard shortcuts.
- Match the editor to your workspace with automatic, colorful, or minimal styling and comfortable or compact density.

### Where it draws the line

- Large inputs use a preview-oriented path based on configurable row and byte limits. Editing a previewed file loads it for editing and may use substantial memory.
- Spreadsheet formulas are not calculated; cells are treated as values.
- Adding, deleting, or renaming rows, columns, and sheets is not supported.
- Some advanced Excel features may not survive editing and saving unchanged; validate important workbooks before relying on them.

The extension requires VS Code `^1.85.0` and a Node extension host. VS Code for the Web is not supported.

## Settings

Open **Preferences: Open Settings** and search for **Spreadsheet Viewer**.

| Setting | Default | What it controls |
| --- | --- | --- |
| `csvViewer.themeMode` | `auto` | `auto`, `colorful`, or `minimal` appearance |
| `csvViewer.density` | `comfortable` | `comfortable` or `compact` row spacing |
| `csvViewer.maxRowsForClientSort` | `200000` | Maximum rows allowed for client-side sorting |
| `csvViewer.previewRowLimit` | `200000` | Row threshold for switching to preview-oriented mode |
| `csvViewer.previewByteLimit` | `104857600` | Byte threshold (100 MiB) for preview-oriented mode |

## Build and run locally

Requirements: Node.js, npm, and VS Code.

```sh
git clone https://github.com/heyaaakash/csv-xlsx-table-viewer.git
cd csv-xlsx-table-viewer
npm ci
npm run build
```

Run these commands from a clone of the source repository. Open the project in VS Code and press `F5` to launch an Extension Development Host. The build bundles the extension into `dist/1.0.0/extension.js`. To create a VSIX, install the [VS Code Extension Manager](https://code.visualstudio.com/api/working-with-extensions/publishing-extension) with `npm install -g @vscode/vsce`, prepare its compatible README, then package and install it with **Extensions: Install from VSIX...**:

```sh
npm run prepare:marketplace-readme
vsce package --githubBranch main --readme-path .vscode/README.marketplace.md --out dist/1.0.0/packages/csv-xlsx-table-viewer-1.0.0.vsix
```

See [how it is built](Docs/ARCHITECTURE.md) and the [release process](Docs/RELEASING.md).

## Privacy

The extension reads and writes files you open. It does not send product analytics; VS Code may collect usage data according to your VS Code settings.

For product questions, use the **Q&A** tab on the [Marketplace listing](https://marketplace.visualstudio.com/items?itemName=aakashrohilla.csv-xlsx-table-viewer). See [security reporting](SECURITY.md) for vulnerability reports. Do not attach real spreadsheets, API logs, credentials, or private workspace data to reports.

## Built with a product point of view

The goal is deliberately narrow: make a quick spreadsheet check part of the coding workflow while being clear about where an in-editor table stops behaving like a full spreadsheet application. Read the [product story](Docs/PRODUCT_STORY.md) for the user problem and product trade-offs. See [contribution guidance](CONTRIBUTING.md) before opening a pull request.

## License

MIT. See [LICENSE](LICENSE).

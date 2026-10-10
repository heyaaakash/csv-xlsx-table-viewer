# Product story

## The user moment

You are already working in VS Code when an export, report, or workbook arrives. You need to find a value or make a small correction. Opening another app breaks the flow; reading raw delimiters makes the data hard to scan.

The goal is not to replace Excel. It is to make quick inspection and focused edits feel at home in the editor.

## The product bet

Treat spreadsheet files as editor content: open the file as a table, inspect or edit it, then save it in its original format. CSV, TSV, XLSX, and XLS share that path, with workbook sheets available in the same editor.

**Open a data file → inspect it → make a focused edit → save in place.**

## Product choices

- **Stay in the coding workflow.** File-open behavior, editor actions, and the Explorer menu put the table beside the code.
- **Keep the promise bounded.** Row and byte thresholds help control large-file rendering, but editing can still use substantial memory.
- **Be clear about spreadsheet semantics.** The extension edits cell values; it does not calculate formulas or manage workbook structure.
- **Preserve the source file.** Saving writes to the opened file in its format; advanced Excel feature fidelity is not guaranteed.

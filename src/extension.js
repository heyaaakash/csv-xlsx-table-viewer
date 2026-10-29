'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const vscode = require('vscode');
const xlsx = require('xlsx');

const VIEW_TYPE = 'csvViewer.csvEditor';
const MAX_RAW_PREVIEW_LINES = 2000;

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeForScript(value) {
  return JSON.stringify(value).replace(/</g, '\\u003c');
}

function getExtension(uri) {
  return path.extname(uri.fsPath).toLowerCase();
}

function getConfig() {
  const config = vscode.workspace.getConfiguration('csvViewer');
  return {
    themeMode: config.get('themeMode', 'auto'),
    density: config.get('density', 'comfortable'),
    hasHeader: config.get('hasHeader', 'auto'),
    maxRowsForClientSort: config.get('maxRowsForClientSort', 200000),
    previewRowLimit: config.get('previewRowLimit', 5000),
    previewByteLimit: config.get('previewByteLimit', 10485760),
  };
}

function detectHasHeader(rows) {
  if (!rows || rows.length < 2) {
    return false;
  }
  const row0 = rows[0];
  const row1 = rows[1];
  if (!row0 || row0.length === 0) {
    return false;
  }

  const nonBlank0 = row0.filter((c) => String(c ?? '').trim() !== '');
  if (nonBlank0.length === 0) {
    return false;
  }

  const isNumeric = (val) => {
    const s = String(val ?? '').trim();
    return s !== '' && !isNaN(Number(s));
  };

  const row0AllNumeric = nonBlank0.every(isNumeric);
  if (row0AllNumeric) {
    return false;
  }

  const row0AnyNumeric = nonBlank0.some(isNumeric);
  const row1AnyNumeric = row1.some(isNumeric);
  if (!row0AnyNumeric && row1AnyNumeric) {
    return true;
  }

  return true;
}

function detectDelimiter(sampleText) {
  const candidates = [',', '\t', ';', '|'];
  const lines = sampleText.split(/\r?\n/).filter(Boolean).slice(0, 25);
  let best = ',';
  let bestScore = -1;

  for (const candidate of candidates) {
    let score = 0;
    for (const line of lines) {
      const count = line.split(candidate).length - 1;
      if (count > 0) {
        score += count;
      }
    }

    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  return best;
}

function detectLineEnding(text) {
  const crlf = (text.match(/\r\n/g) || []).length;
  const lf = (text.match(/(?<!\r)\n/g) || []).length;
  return crlf >= lf ? '\r\n' : '\n';
}

function parseDelimitedText(text, delimiter, maxRows = Infinity) {
  const rows = [];
  let row = [];
  let cell = '';
  let inQuotes = false;
  let truncated = false;

  if (text.charCodeAt(0) === 0xfeff) {
    text = text.slice(1);
  }

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];

    if (ch === '"') {
      if (inQuotes && text[i + 1] === '"') {
        cell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === delimiter && !inQuotes) {
      row.push(cell);
      cell = '';
      continue;
    }

    if ((ch === '\n' || ch === '\r') && !inQuotes) {
      if (ch === '\r' && text[i + 1] === '\n') {
        i++;
      }
      row.push(cell);
      if (row.some((value) => String(value).trim().length > 0)) {
        rows.push(row);
        if (rows.length >= maxRows) {
          truncated = i < text.length - 1;
          break;
        }
      }
      row = [];
      cell = '';
      continue;
    }

    cell += ch;
  }

  if (!truncated) {
    row.push(cell);
    if (row.some((value) => String(value).trim().length > 0)) {
      rows.push(row);
    }
  }

  return { rows, truncated };
}

function rawPreviewFromText(text) {
  const lines = text.split(/\r?\n/).slice(0, MAX_RAW_PREVIEW_LINES);
  return lines
    .map((line, index) => (index === 0 ? `<span class="hdr">${escapeHtml(line)}</span>` : escapeHtml(line)))
    .join('\n');
}

function columnLabel(index) {
  let n = index + 1;
  let label = '';

  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }

  return label;
}

function serializeDelimitedRows(rows, delimiter, lineEnding) {
  const allRows = [];
  const maxColumns = Math.max(...rows.map((row) => row.length), 0);

  const encode = (value) => {
    const text = value == null ? '' : String(value);
    if (text.includes('"') || text.includes(delimiter) || /[\r\n]/.test(text)) {
      return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
  };

  for (const row of rows) {
    const encoded = [];
    for (let i = 0; i < maxColumns; i++) {
      encoded.push(encode(row[i] ?? ''));
    }
    allRows.push(encoded.join(delimiter));
  }

  return allRows.join(lineEnding);
}

async function getUriSize(uri) {
  if (uri.scheme === 'file') {
    const stats = await fs.promises.stat(uri.fsPath);
    return stats.size;
  }

  const stats = await vscode.workspace.fs.stat(uri);
  return stats.size;
}

async function readUriBytes(uri, maxBytes) {
  if (uri.scheme === 'file') {
    const handle = await fs.promises.open(uri.fsPath, 'r');
    try {
      const stat = await handle.stat();
      const bytesToRead = maxBytes ? Math.min(stat.size, maxBytes) : stat.size;
      const buffer = Buffer.alloc(Math.max(bytesToRead, 1));
      const { bytesRead } = await handle.read(buffer, 0, bytesToRead, 0);
      return buffer.subarray(0, bytesRead);
    } finally {
      await handle.close();
    }
  }

  const bytes = await vscode.workspace.fs.readFile(uri);
  return maxBytes ? bytes.subarray(0, maxBytes) : bytes;
}

function workbookToSnapshot(uri, bytes, previewRowLimit, editable) {
  const workbook = xlsx.read(Buffer.from(bytes), {
    type: 'buffer',
    cellDates: true,
    sheetStubs: true,
    dense: true,
    sheetRows: editable ? undefined : previewRowLimit + 1,
  });

  const sheets = workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const matrix = xlsx.utils.sheet_to_json(worksheet, {
      header: 1,
      blankrows: false,
      defval: '',
      raw: false,
    });
    const rows = (editable ? matrix : matrix.slice(0, previewRowLimit)).map((row) => row.map((cell) => String(cell ?? '')));
    const totalRows = editable ? matrix.length : rows.length + 1;
    const totalColumns = Math.max(...rows.map((row) => row.length), 0);

    return {
      name: sheetName,
      rows,
      totalRows,
      totalColumns,
      limited: !editable,
      editable,
    };
  });

  return {
    uri,
    sourceUri: uri.toString(),
    kind: 'workbook',
    fileName: path.basename(uri.fsPath),
    fileExt: getExtension(uri),
    editable,
    dirty: false,
    delimiter: null,
    lineEnding: '\n',
    rawPreview: workbook.SheetNames.join('\n'),
    sheets,
    activeSheetIndex: 0,
    limited: !editable,
  };
}

function delimitedToSnapshot(uri, text, previewRowLimit, byteLength, previewByteLimit) {
  const ext = getExtension(uri);
  const delimiter = ext === '.tsv' ? '\t' : detectDelimiter(text);
  const lineEnding = detectLineEnding(text);
  const parsed = parseDelimitedText(text, delimiter, previewRowLimit + 1);
  const rows = parsed.rows;
  const loadedAllRows = !parsed.truncated && byteLength <= previewByteLimit;
  const totalColumns = Math.max(...rows.map((row) => row.length), 0);

  return {
    uri,
    sourceUri: uri.toString(),
    kind: 'delimited',
    fileName: path.basename(uri.fsPath),
    fileExt: ext,
    editable: loadedAllRows,
    dirty: false,
    delimiter,
    lineEnding,
    rawPreview: rawPreviewFromText(text),
    sheets: [
      {
        name: 'Data',
        rows,
        totalRows: rows.length,
        totalColumns,
        limited: !loadedAllRows,
        editable: loadedAllRows,
      },
    ],
    activeSheetIndex: 0,
    limited: !loadedAllRows,
  };
}

async function loadSnapshot(uri) {
  const { previewRowLimit, previewByteLimit } = getConfig();
  const ext = getExtension(uri);

  if (ext === '.csv' || ext === '.tsv') {
    const fileSize = await getUriSize(uri);
    const bytes = await readUriBytes(uri, fileSize <= previewByteLimit ? fileSize : previewByteLimit);
    const text = Buffer.from(bytes).toString('utf8');
    return delimitedToSnapshot(uri, text, previewRowLimit, bytes.length, previewByteLimit);
  }

  if (ext === '.xlsx' || ext === '.xls') {
    const fileSize = await getUriSize(uri);
    const bytes = await readUriBytes(uri);
    return workbookToSnapshot(uri, bytes, previewRowLimit, fileSize <= previewByteLimit);
  }

  throw new Error(`Unsupported file type: ${ext || 'unknown'}`);
}

function cloneSnapshot(snapshot) {
  return JSON.parse(JSON.stringify(snapshot));
}

function buildLoadingHtml(fileName, nonce) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'nonce-${nonce}';">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    html,body{height:100%;margin:0;background:var(--vscode-editor-background,#1e1e1e);color:var(--vscode-editor-foreground,#d4d4d4);font-family:var(--vscode-font-family,sans-serif)}
    .wrap{height:100%;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:16px}
    .spinner{width:36px;height:36px;border:3px solid var(--vscode-editorWidget-border,#454545);border-top-color:var(--vscode-button-background,#0e639c);border-radius:50%;animation:spin .8s linear infinite}
    @keyframes spin{to{transform:rotate(360deg)}}
    .label{color:var(--vscode-descriptionForeground,#888);font-size:13px}
  </style>
  <title>${escapeHtml(fileName)}</title>
</head>
<body>
  <div class="wrap">
    <div class="spinner"></div>
    <div class="label">Loading <strong>${escapeHtml(fileName)}</strong>…</div>
  </div>
</body>
</html>`;
}

function buildErrorHtml(fileName, message, nonce) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'nonce-${nonce}';">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(fileName)}</title>
  <style>
    html,body{height:100%;margin:0;background:var(--vscode-editor-background,#1e1e1e);color:var(--vscode-editor-foreground,#d4d4d4);font-family:var(--vscode-font-family,sans-serif)}
    .wrap{height:100%;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:12px;padding:24px;text-align:center}
    .card{max-width:720px;border:1px solid var(--vscode-editorWidget-border,#454545);border-radius:10px;padding:20px;background:var(--vscode-sideBar-background,#252526)}
    .title{font-size:16px;font-weight:700;margin-bottom:8px}
    .msg{color:var(--vscode-descriptionForeground,#888);white-space:pre-wrap;line-height:1.6}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <div class="title">Couldn’t open ${escapeHtml(fileName)}</div>
      <div class="msg">${escapeHtml(message)}</div>
    </div>
  </div>
</body>
</html>`;
}

function buildEditorHtml(nonce, state) {
  const stateJson = escapeForScript(state);
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'nonce-${nonce}';">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(state.fileName)}</title>
  <style>
    :root{--bg:var(--vscode-editor-background,#1e1e1e);--fg:var(--vscode-editor-foreground,#d4d4d4);--muted:var(--vscode-descriptionForeground,#8a8a8a);--border:var(--vscode-editorWidget-border,#454545);--surface:var(--vscode-editor-background,#1e1e1e);--accent:var(--vscode-button-background,#0e639c);--hover:var(--vscode-list-hoverBackground,rgba(255,255,255,.07));}
    *{box-sizing:border-box}
    html,body{height:100%;margin:0;background:var(--bg);color:var(--fg);font-family:var(--vscode-font-family,sans-serif);font-size:13px;overflow:hidden;display:flex;flex-direction:column}
    .viewport{flex:1;width:100%;overflow:auto;outline:none;position:relative}

    /* Banner Alert */
    .banner {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 20px;
      background: color-mix(in srgb, var(--accent) 15%, var(--bg));
      color: var(--fg);
      border-bottom: 1px solid var(--border);
      font-size: 13px;
      font-weight: 500;
      gap: 16px;
      animation: slideDown 0.25s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .banner.hidden {
      display: none;
    }
    @keyframes slideDown {
      from { transform: translateY(-100%); opacity: 0; }
      to { transform: translateY(0); opacity: 1; }
    }
    .banner-text {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .banner-btn {
      background: var(--accent);
      color: var(--vscode-button-foreground, #ffffff);
      border: none;
      border-radius: 4px;
      padding: 6px 12px;
      cursor: pointer;
      font-size: 12px;
      font-weight: 600;
      transition: background 0.15s ease, transform 0.1s ease;
      white-space: nowrap;
    }
    .banner-btn:hover {
      background: var(--vscode-button-hoverBackground, #1177bb);
    }
    .banner-btn:active {
      transform: scale(0.97);
    }
    .banner-btn:disabled {
      background: var(--muted);
      cursor: not-allowed;
      transform: none;
      opacity: 0.7;
    }
    .viewport::-webkit-scrollbar{width:10px;height:10px}
    .viewport::-webkit-scrollbar-thumb{background:var(--border);border-radius:999px}
    table{border-collapse:separate;border-spacing:0;width:max-content;min-width:100%;table-layout:fixed}
    thead th{position:sticky;top:0;z-index:10;background:var(--surface);border-bottom:2px solid var(--border);border-right:1px solid var(--border);padding:8px 12px;text-align:left;white-space:nowrap;user-select:none}
    thead th.corner{left:0;top:0;z-index:15;position:sticky;width:64px;text-align:right;background:var(--surface)}
    tbody td{border-bottom:1px solid var(--border);border-right:1px solid var(--border);padding:7px 12px;white-space:nowrap;max-width:360px;overflow:hidden;text-overflow:ellipsis;background:var(--bg);position:relative}
    tbody tr:hover td{background:var(--hover)}
    tbody td.rowhead{position:sticky;left:0;z-index:5;background:var(--bg);color:var(--muted);text-align:right;width:64px;user-select:none;white-space:nowrap}
    td.editing{outline:2px solid var(--accent) !important;outline-offset:-2px;background:color-mix(in srgb, var(--accent) 14%, var(--bg)) !important;z-index:5;color:var(--fg) !important}
    td.cell-selected{outline:2px solid var(--accent);outline-offset:-2px;background:color-mix(in srgb, var(--accent) 8%, var(--bg)) !important;z-index:2}
    td.cell-selected::after {
      content: '';
      position: absolute;
      bottom: -3px;
      right: -3px;
      width: 6px;
      height: 6px;
      background: var(--accent);
      border: 1px solid var(--bg);
      z-index: 10;
      pointer-events: none;
    }
    td.empty{color:var(--muted);padding:20px 12px;text-align:center}

    /* Resizers */
    .col-resizer {
      position: absolute;
      top: 0;
      right: 0;
      width: 6px;
      height: 100%;
      cursor: col-resize;
      z-index: 10;
      user-select: none;
    }
    .col-resizer:hover, .col-resizer.resizing {
      background: var(--vscode-focusBorder, #007fd4);
    }
    .row-resizer {
      position: absolute;
      bottom: 0;
      left: 0;
      width: 100%;
      height: 6px;
      cursor: row-resize;
      z-index: 10;
      user-select: none;
    }
    .row-resizer:hover, .row-resizer.resizing {
      background: var(--vscode-focusBorder, #007fd4);
    }

    /* Search Highlighting */
    .search-match {
      background-color: var(--vscode-editor-findMatchHighlightBackground, rgba(234, 92, 0, 0.3)) !important;
      outline: 1px dashed var(--vscode-editor-findMatchHighlightBorder, #ea5c00);
      color: var(--fg) !important;
    }
    .search-match-active {
      background-color: var(--vscode-editor-findMatchBackground, #ea5c00) !important;
      outline: 2px solid var(--vscode-editor-findMatchBorder, #ea5c00) !important;
      z-index: 2;
      color: var(--fg) !important;
    }

    /* Find Widget */
    .find-widget {
      position: absolute;
      top: 12px;
      right: 24px;
      z-index: 100;
      display: flex;
      align-items: center;
      gap: 6px;
      background: color-mix(in srgb, var(--vscode-editorWidget-background, #252526) 90%, transparent);
      backdrop-filter: blur(8px);
      border: 1px solid var(--vscode-editorWidget-border, #454545);
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
      border-radius: 6px;
      padding: 6px 8px;
      font-family: inherit;
      transition: opacity 0.15s ease, transform 0.15s ease;
    }
    .find-widget.hidden {
      opacity: 0;
      transform: translateY(-8px);
      pointer-events: none;
    }
    .find-widget input {
      background: var(--vscode-input-background, #3c3c3c);
      color: var(--vscode-input-foreground, #cccccc);
      border: 1px solid var(--vscode-input-border, #3c3c3c);
      border-radius: 4px;
      padding: 4px 8px;
      font-size: 12px;
      outline: none;
      width: 180px;
    }
    .find-widget input:focus {
      border-color: var(--vscode-focusBorder, #007fd4);
    }
    .find-widget button {
      background: transparent;
      color: var(--vscode-foreground, #cccccc);
      border: none;
      border-radius: 4px;
      width: 24px;
      height: 24px;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      font-size: 12px;
    }
    .find-widget button:hover {
      background: var(--vscode-toolbar-hoverBackground, rgba(255, 255, 255, 0.08));
    }
    .find-widget span {
      color: var(--muted);
      font-size: 11px;
      padding: 0 4px;
      user-select: none;
      white-space: nowrap;
    }

    /* Cell Popover */
    .cell-popover {
      position: absolute;
      z-index: 1000;
      background: var(--vscode-editorWidget-background, #252526);
      border: 1px solid var(--vscode-editorWidget-border, #454545);
      border-radius: 6px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.4);
      padding: 8px;
      min-width: 400px;
      min-height: 120px;
      max-width: 85vw;
      max-height: 75vh;
      width: 480px;
      height: 180px;
      overflow: hidden;
      resize: both;
      display: flex;
      flex-direction: column;
      outline: none;
      transition: opacity 0.15s ease;
    }
    .cell-popover.hidden {
      display: none;
    }
    .cell-popover textarea {
      flex: 1;
      width: 100%;
      height: 100%;
      background: transparent;
      color: var(--fg);
      border: none;
      outline: none;
      resize: none;
      font-family: var(--vscode-editor-font-family, monospace);
      font-size: 12px;
      line-height: 1.4;
      padding: 4px;
      margin: 0;
      overflow: auto;
      white-space: pre-wrap;
      word-break: break-word;
    }
    .cell-popover textarea::-webkit-scrollbar {
      width: 8px;
      height: 8px;
    }
    .cell-popover textarea::-webkit-scrollbar-thumb {
      background: var(--border);
      border-radius: 4px;
    }

    /* Formula Bar */
    .formula-bar {
      display: flex;
      align-items: center;
      background: var(--vscode-editorWidget-background, #252526);
      border-bottom: 1px solid var(--border);
      padding: 4px 12px;
      gap: 10px;
      height: 36px;
      font-family: inherit;
      flex-shrink: 0;
    }
    .formula-cell-address {
      font-weight: 600;
      color: var(--accent);
      min-width: 56px;
      text-align: center;
      font-size: 12px;
      user-select: none;
      font-family: var(--vscode-editor-font-family, monospace);
    }
    .formula-bar-divider {
      width: 1px;
      height: 20px;
      background: var(--border);
      flex-shrink: 0;
    }
    .formula-input {
      flex: 1;
      min-width: 120px;
      background: transparent;
      border: none;
      outline: none;
      color: var(--fg);
      font-family: var(--vscode-editor-font-family, monospace);
      font-size: 12px;
      padding: 4px 8px;
    }
    .formula-input:focus {
      background: var(--vscode-input-background, #3c3c3c);
      border-radius: 4px;
      outline: 1px solid var(--vscode-focusBorder, #007fd4);
    }

    /* Header Row Toggle Pill */
    .header-toggle-pill {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: var(--vscode-editor-background, #1e1e1e);
      color: var(--fg);
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 2px 8px;
      cursor: pointer;
      font-size: 11px;
      font-weight: 500;
      user-select: none;
      transition: background 0.15s ease, border-color 0.15s ease, box-shadow 0.15s ease;
      white-space: nowrap;
      height: 26px;
      flex-shrink: 0;
    }
    .header-toggle-pill:hover {
      background: var(--hover);
      border-color: var(--vscode-focusBorder, #007fd4);
    }
    .header-toggle-pill .toggle-icon {
      font-size: 12px;
      line-height: 1;
    }
    .header-toggle-pill .toggle-text {
      font-size: 11px;
      font-weight: 500;
      color: var(--fg);
    }
    .header-toggle-pill .toggle-badge {
      font-size: 10px;
      font-weight: 700;
      letter-spacing: 0.5px;
      text-transform: uppercase;
      padding: 2px 6px;
      border-radius: 4px;
      transition: background 0.15s ease, color 0.15s ease;
      line-height: 1;
    }
    .header-toggle-pill.active .toggle-badge {
      background: var(--accent);
      color: var(--vscode-button-foreground, #ffffff);
    }
    .header-toggle-pill:not(.active) .toggle-badge {
      background: rgba(127, 127, 127, 0.25);
      color: var(--muted);
    }
    .header-toggle-pill.active {
      border-color: color-mix(in srgb, var(--accent) 60%, var(--border));
      background: color-mix(in srgb, var(--accent) 10%, var(--vscode-editor-background, #1e1e1e));
    }

    /* Column Headers */
    .th-content {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      max-width: calc(100% - 8px);
      overflow: hidden;
      vertical-align: middle;
    }
    .col-title {
      font-weight: 600;
      letter-spacing: 0.2px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .col-letter-sub {
      display: inline-block;
      font-size: 10px;
      font-weight: 500;
      opacity: 0.75;
      padding: 1px 4px;
      border-radius: 3px;
      background: rgba(127, 127, 127, 0.2);
      vertical-align: middle;
      line-height: 1.2;
      flex-shrink: 0;
    }

    /* Sorting */
    thead th[data-col] {
      cursor: pointer;
    }
    thead th[data-col]:hover {
      background: var(--hover);
    }
    .sort-indicator {
      margin-left: 6px;
      font-size: 11px;
      color: var(--muted);
      display: inline-block;
      transition: color 0.15s ease;
      user-select: none;
    }
    thead th[data-col]:hover .sort-indicator {
      color: var(--accent);
    }

    /* Sheet Tabs */
    .sheet-tabs {
      display: none;
      align-items: center;
      gap: 2px;
      padding: 4px 8px;
      background: color-mix(in srgb, var(--vscode-sideBar-background, #252526) 92%, var(--bg));
      border-top: 1px solid var(--border);
      overflow-x: auto;
      overflow-y: hidden;
      flex-shrink: 0;
      height: 32px;
    }
    .sheet-tabs.visible {
      display: flex;
    }
    .sheet-tabs::-webkit-scrollbar {
      height: 6px;
    }
    .sheet-tabs::-webkit-scrollbar-thumb {
      background: var(--border);
      border-radius: 999px;
    }
    .sheet-tab {
      padding: 4px 14px;
      border-radius: 4px 4px 0 0;
      cursor: pointer;
      font-size: 12px;
      white-space: nowrap;
      color: var(--muted);
      border: 1px solid transparent;
      border-bottom: 2px solid transparent;
      user-select: none;
      transition: background 0.12s ease, color 0.12s ease, border-color 0.12s ease;
    }
    .sheet-tab:hover {
      background: var(--hover);
      color: var(--fg);
    }
    .sheet-tab.active {
      background: color-mix(in srgb, var(--accent) 22%, var(--bg));
      color: var(--fg);
      border-color: var(--border);
      border-bottom-color: var(--accent);
      font-weight: 600;
    }
  </style>
</head>
<body>
  <div id="banner" class="banner hidden"></div>
  <div id="formula-bar" class="formula-bar">
    <div id="formula-cell-address" class="formula-cell-address"></div>
    <div class="formula-bar-divider"></div>
    <input type="text" id="formula-input" class="formula-input" placeholder="Enter value..." />
    <div class="formula-bar-divider"></div>
    <button id="header-toggle-btn" class="header-toggle-pill" type="button" aria-pressed="true" title="First row is currently used as the header. Click to treat as normal data row.">
      <span class="toggle-icon">🏷️</span>
      <span class="toggle-text">First row as header</span>
      <span id="header-toggle-status" class="toggle-badge">ON</span>
    </button>
  </div>
  <div id="viewport" class="viewport" tabindex="0"></div>

  <!-- Find Widget -->
  <div id="find-widget" class="find-widget hidden">
    <input type="text" id="find-input" placeholder="Find..." />
    <span id="find-count">0 of 0</span>
    <button id="find-prev" title="Previous Match (Shift+Enter)">▲</button>
    <button id="find-next" title="Next Match (Enter)">▼</button>
    <button id="find-close" title="Close (Escape)">✕</button>
  </div>

  <!-- Cell Popover -->
  <div id="cell-popover" class="cell-popover hidden">
    <textarea id="cell-popover-textarea"></textarea>
  </div>

  <!-- Sheet Tabs -->
  <div id="sheet-tabs" class="sheet-tabs"></div>

  <script nonce="${nonce}">
    const vscode = acquireVsCodeApi();
    const initialState = ${stateJson};
  </script>
  <script nonce="${nonce}">
    (() => {
      function escapeHtml(value) {
        return String(value)
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;');
      }

      function columnLabel(index) {
        let n = index + 1;
        let label = '';
        while (n > 0) {
          const rem = (n - 1) % 26;
          label = String.fromCharCode(65 + rem) + label;
          n = Math.floor((n - 1) / 26);
        }
        return label;
      }

      const state = {
        ...initialState,
        hasHeader: Boolean(initialState.hasHeader),
        activeSheetIndex: Math.max(0, initialState.activeSheetIndex || 0),
        selected: null,
        colWidths: [],
        rowHeights: [],
        visibleRowCount: 1000,
        sortColumn: -1,
        sortDirection: null,
      };

      const viewport = document.getElementById('viewport');
      const formulaInput = document.getElementById('formula-input');
      const cellAddress = document.getElementById('formula-cell-address');
      let editingCell = null;

      viewport.addEventListener('scroll', () => {
        if (viewport.scrollHeight - viewport.scrollTop - viewport.clientHeight < 400) {
          const sheet = currentSheet();
          const total = sheet.rows ? sheet.rows.length : 0;
          if (state.visibleRowCount < total) {
            state.visibleRowCount = Math.min(total, state.visibleRowCount + 1000);
            renderGrid();
          }
        }
      });

      function currentSheet() {
        return state.sheets[state.activeSheetIndex] || state.sheets[0] || { name: 'Sheet', rows: [], totalColumns: 0 };
      }

      function columnCountForSheet(sheet, rows) {
        return Math.max(
          sheet.totalColumns || 0,
          ...rows.map((row) => row.length)
        );
      }

      function renderGrid() {
        const sheet = currentSheet();
        const rows = sheet.rows || [];
        const columnCount = columnCountForSheet(sheet, rows);
        const hasHeader = Boolean(state.hasHeader && rows.length > 0);
        const headers = [];
        for (let i = 0; i < columnCount; i++) {
          const colLetter = columnLabel(i);
          if (hasHeader && rows[0] && rows[0][i] !== undefined && String(rows[0][i]).trim() !== '') {
            headers.push({
              title: String(rows[0][i]),
              letter: colLetter,
            });
          } else {
            headers.push({
              title: colLetter,
              letter: colLetter,
            });
          }
        }

        let html = '<table>';

        // Add colgroup for resizing
        html += '<colgroup>';
        html += '<col style="width: 64px;">'; // corner column
        for (let i = 0; i < columnCount; i++) {
          const width = state.colWidths[i] || 120;
          html += '<col style="width: ' + width + 'px;">';
        }
        html += '</colgroup>';

        html += '<thead><tr><th class="corner">#</th>';
        for (let index = 0; index < columnCount; index++) {
          const hue = (index * 60) % 360;
          const style = 'color: color-mix(in srgb, hsl(' + hue + ', 85%, 45%) 85%, var(--fg));';
          let sortIndicator = ' ⇅';
          if (state.sortColumn === index) {
            if (state.sortDirection === 'asc') sortIndicator = ' ▲';
            else if (state.sortDirection === 'desc') sortIndicator = ' ▼';
          }
          const headerObj = headers[index];
          const displayTitle = escapeHtml(headerObj.title);
          const letterSub = hasHeader && headerObj.title !== headerObj.letter
            ? '<span class="col-letter-sub">' + escapeHtml(headerObj.letter) + '</span>'
            : '';

          html += '<th style="' + style + '" data-col="' + index + '" title="' + displayTitle + (hasHeader && headerObj.title !== headerObj.letter ? ' (' + escapeHtml(headerObj.letter) + ')' : '') + '">' +
            '<div class="th-content">' +
            '<span class="col-title">' + displayTitle + '</span>' +
            letterSub +
            '<span class="sort-indicator">' + sortIndicator + '</span>' +
            '</div>' +
            '<div class="col-resizer" data-col="' + index + '"></div></th>';
        }
        html += '</tr></thead><tbody>';

        const startRowIndex = hasHeader ? 1 : 0;
        const displaySlice = rows.slice(startRowIndex, startRowIndex + state.visibleRowCount);

        if (!displaySlice.length) {
          html += '<tr><td class="empty" colspan="' + (columnCount + 1) + '">No rows to display</td></tr>';
        }

        displaySlice.forEach((row, sliceIndex) => {
          const rowIndex = startRowIndex + sliceIndex;
          const displayRowNum = hasHeader ? rowIndex : (rowIndex + 1);
          const height = state.rowHeights[rowIndex] || 28;
          html += '<tr style="height: ' + height + 'px;">';
          html += '<td class="rowhead">' + displayRowNum + '<div class="row-resizer" data-row="' + rowIndex + '"></div></td>';
          for (let colIndex = 0; colIndex < columnCount; colIndex++) {
            const value = row[colIndex] ?? '';
            const selected = state.selected && state.selected.row === rowIndex && state.selected.col === colIndex ? ' cell-selected' : '';
            const hue = (colIndex * 60) % 360;
            const style = 'color: color-mix(in srgb, hsl(' + hue + ', 85%, 45%) 85%, var(--fg));';
            html += '<td class="cell' + selected + '" style="' + style + '" data-row="' + rowIndex + '" data-col="' + colIndex + '" title="' + escapeHtml(value) + '">' + escapeHtml(value) + '</td>';
          }
          html += '</tr>';
        });

        html += '</tbody></table>';
        viewport.innerHTML = html;
        document.title = state.fileName + (sheet.name ? ' — ' + sheet.name : '');

        applySearchMatches();
      }

      function updateSelection(row, col) {
        state.selected = { row, col };
        const sheet = currentSheet();
        const value = sheet.rows[row]?.[col] ?? '';
        if (formulaInput) {
          formulaInput.value = value;
        }
        if (cellAddress) {
          const displayRowNum = state.hasHeader ? row : (row + 1);
          cellAddress.textContent = columnLabel(col) + displayRowNum;
        }
        return value;
      }

      function startEditing(cell) {
        if (!cell) {
          return;
        }
        if (editingCell && editingCell !== cell) {
          commitEditing(editingCell, false);
        }
        editingCell = cell;
        cell.dataset.original = cell.innerText;
        cell.contentEditable = 'true';
        cell.classList.add('editing');
        cell.focus({ preventScroll: true });

        if (formulaInput) {
          formulaInput.value = cell.innerText;
        }

        const selection = window.getSelection();
        if (selection) {
          const range = document.createRange();
          range.selectNodeContents(cell);
          selection.removeAllRanges();
          selection.addRange(range);
        }
      }

      function commitEditing(cell, revert) {
        if (!cell) {
          return;
        }
        const row = Number(cell.dataset.row);
        const col = Number(cell.dataset.col);
        const original = cell.dataset.original ?? '';
        const value = revert ? original : cell.innerText;
        cell.contentEditable = 'false';
        cell.classList.remove('editing');
        delete cell.dataset.original;

        if (!revert && value !== original) {
          const sheet = currentSheet();
          if (sheet.rows[row]) {
            sheet.rows[row][col] = value;
          }

          vscode.postMessage({
            type: 'cellEdit',
            sheetIndex: state.activeSheetIndex,
            row,
            col,
            value,
          });

          cell.title = value;
        }

        if (formulaInput && state.selected && state.selected.row === row && state.selected.col === col) {
          formulaInput.value = revert ? original : value;
        }

        if (editingCell === cell) {
          editingCell = null;
        }
      }

      const banner = document.getElementById('banner');
      function updateBanner() {
        if (!banner) return;
        if (state.limited) {
          const sheet = currentSheet();
          const loadedRows = sheet.rows ? sheet.rows.length : 0;
          banner.innerHTML =
            '<div class="banner-text">' +
              '<span>⚠️ Viewing a preview of the first <strong>' + loadedRows + '</strong> rows. The file has been truncated to avoid performance lag.</span>' +
            '</div>' +
            '<button id="banner-btn" class="banner-btn">Load Entire File</button>';
          banner.classList.remove('hidden');

          document.getElementById('banner-btn').addEventListener('click', () => {
            const btn = document.getElementById('banner-btn');
            btn.disabled = true;
            btn.textContent = 'Loading...';
            vscode.postMessage({ type: 'loadAll' });
          });
        } else {
          banner.classList.add('hidden');
          banner.innerHTML = '';
        }
      }

      function renderSheetTabs() {
        const tabs = document.getElementById('sheet-tabs');
        if (!tabs) return;
        const sheets = state.sheets || [];
        if (sheets.length <= 1) {
          tabs.classList.remove('visible');
          tabs.innerHTML = '';
          return;
        }
        tabs.classList.add('visible');
        let html = '';
        for (let i = 0; i < sheets.length; i++) {
          const active = i === state.activeSheetIndex ? ' active' : '';
          html += '<div class="sheet-tab' + active + '" data-sheet="' + i + '" title="' + escapeHtml(sheets[i].name || ('Sheet' + (i + 1))) + '">' + escapeHtml(sheets[i].name || ('Sheet' + (i + 1))) + '</div>';
        }
        tabs.innerHTML = html;
      }

      function setActiveSheet(index) {
        const sheets = state.sheets || [];
        if (index < 0 || index >= sheets.length || index === state.activeSheetIndex) return;
        if (editingCell) {
          commitEditing(editingCell, false);
        }
        if (popoverCurrentCell) {
          closeCellPopover(false);
        }
        state.activeSheetIndex = index;
        state.selected = null;
        state.visibleRowCount = 1000;
        resetSortState();
        vscode.postMessage({ type: 'setActiveSheet', index: index });
        renderAll();
      }

      document.getElementById('sheet-tabs').addEventListener('click', (event) => {
        const tab = event.target.closest('.sheet-tab');
        if (tab) {
          const index = Number(tab.dataset.sheet);
          setActiveSheet(index);
        }
      });

      function updateHeaderToggleUI() {
        const btn = document.getElementById('header-toggle-btn');
        const badge = document.getElementById('header-toggle-status');
        if (!btn) return;
        if (state.hasHeader) {
          btn.classList.add('active');
          btn.title = 'First row is currently used as the header. Click to treat as normal data row.';
          btn.setAttribute('aria-pressed', 'true');
          if (badge) {
            badge.textContent = 'ON';
          }
        } else {
          btn.classList.remove('active');
          btn.title = 'First row is currently treated as data. Click to use as column header.';
          btn.setAttribute('aria-pressed', 'false');
          if (badge) {
            badge.textContent = 'OFF';
          }
        }
      }

      function renderAll() {
        updateBanner();
        updateHeaderToggleUI();
        renderGrid();
        renderSheetTabs();
      }

      const headerToggleBtn = document.getElementById('header-toggle-btn');
      if (headerToggleBtn) {
        headerToggleBtn.addEventListener('click', () => {
          state.hasHeader = !state.hasHeader;
          updateHeaderToggleUI();
          if (state.hasHeader && state.selected && state.selected.row === 0) {
            state.selected = { row: 1, col: state.selected.col };
          }
          renderGrid();
          vscode.postMessage({
            type: 'toggleHeader',
            hasHeader: state.hasHeader
          });
        });
      }

      // Resize handling
      let currentResizing = null;

      document.addEventListener('mousedown', (event) => {
        if (event.target.classList.contains('col-resizer')) {
          event.preventDefault();
          const colIndex = Number(event.target.dataset.col);
          const cols = viewport.querySelectorAll('colgroup col');
          const colEl = cols[colIndex + 1];
          const startWidth = colEl ? parseFloat(colEl.style.width || 120) : 120;
          currentResizing = {
            type: 'col',
            index: colIndex,
            startPos: event.clientX,
            startSize: startWidth,
            element: colEl,
            handle: event.target
          };
          event.target.classList.add('resizing');
          document.body.style.cursor = 'col-resize';
        } else if (event.target.classList.contains('row-resizer')) {
          event.preventDefault();
          const rowIndex = Number(event.target.dataset.row);
          const trs = viewport.querySelectorAll('tbody tr');
          const trEl = trs[rowIndex];
          const startHeight = trEl ? trEl.getBoundingClientRect().height : 28;
          currentResizing = {
            type: 'row',
            index: rowIndex,
            startPos: event.clientY,
            startSize: startHeight,
            element: trEl,
            handle: event.target
          };
          event.target.classList.add('resizing');
          document.body.style.cursor = 'row-resize';
        }
      });

      document.addEventListener('mousemove', (event) => {
        if (!currentResizing) {
          return;
        }

        if (currentResizing.type === 'col') {
          const deltaX = event.clientX - currentResizing.startPos;
          const newWidth = Math.max(40, currentResizing.startSize + deltaX);
          currentResizing.element.style.width = newWidth + 'px';
          state.colWidths[currentResizing.index] = newWidth;
        } else if (currentResizing.type === 'row') {
          const deltaY = event.clientY - currentResizing.startPos;
          const newHeight = Math.max(18, currentResizing.startSize + deltaY);
          currentResizing.element.style.height = newHeight + 'px';
          state.rowHeights[currentResizing.index] = newHeight;
        }
      });

      document.addEventListener('mouseup', () => {
        if (currentResizing) {
          currentResizing.handle.classList.remove('resizing');
          document.body.style.cursor = '';
          currentResizing = null;
        }
      });

      // Search / Find widget logic
      const findWidget = document.getElementById('find-widget');
      const findInput = document.getElementById('find-input');
      const findCount = document.getElementById('find-count');
      const findPrev = document.getElementById('find-prev');
      const findNext = document.getElementById('find-next');
      const findClose = document.getElementById('find-close');

      let searchMatches = [];
      let activeMatchIndex = -1;

      function showFindWidget() {
        findWidget.classList.remove('hidden');
        findInput.focus();
        findInput.select();
        performSearch();
      }

      function hideFindWidget() {
        findWidget.classList.add('hidden');
        clearSearchMatches();
        viewport.focus();
      }

      function clearSearchMatches() {
        viewport.querySelectorAll('.search-match, .search-match-active').forEach(el => {
          el.classList.remove('search-match', 'search-match-active');
        });
        searchMatches = [];
        activeMatchIndex = -1;
        findCount.textContent = '0 of 0';
      }

      function performSearch() {
        clearSearchMatches();
        const query = findInput.value.toLowerCase().trim();
        if (!query) {
          return;
        }

        const sheet = currentSheet();
        const rows = sheet.rows || [];

        searchMatches = [];
        const startRow = (state.hasHeader && rows.length > 1) ? 1 : 0;
        for (let rowIndex = startRow; rowIndex < rows.length; rowIndex++) {
          const row = rows[rowIndex];
          if (!row) continue;
          row.forEach((value, colIndex) => {
            const text = String(value ?? '').toLowerCase();
            if (text.includes(query)) {
              searchMatches.push({
                row: rowIndex,
                col: colIndex
              });
            }
          });
        }

        if (searchMatches.length > 0) {
          activeMatchIndex = 0;
          updateActiveMatch();
        } else {
          findCount.textContent = '0 of 0';
        }
      }

      function updateActiveMatch() {
        if (activeMatchIndex < 0 || activeMatchIndex >= searchMatches.length) {
          return;
        }

        const activeMatch = searchMatches[activeMatchIndex];

        // Ensure the active match row is visible
        if (activeMatch.row >= state.visibleRowCount + (state.hasHeader ? 1 : 0)) {
          state.visibleRowCount = Math.max(state.visibleRowCount, activeMatch.row + 100);
          renderGrid();
        }

        // Remove active class from all cells
        viewport.querySelectorAll('.search-match-active').forEach(el => {
          el.classList.remove('search-match-active');
        });

        // Add active class to the current active match cell
        const activeCell = viewport.querySelector('td[data-row="' + activeMatch.row + '"][data-col="' + activeMatch.col + '"]');
        if (activeCell) {
          activeCell.classList.add('search-match-active');
          activeCell.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        }

        findCount.textContent = (activeMatchIndex + 1) + ' of ' + searchMatches.length;
      }

      function applySearchMatches() {
        const query = findInput.value.toLowerCase().trim();
        if (!query || findWidget.classList.contains('hidden')) {
          return;
        }

        // Highlight matching cells currently visible in the DOM
        const cells = viewport.querySelectorAll('tbody td.cell');
        cells.forEach(cell => {
          const text = cell.innerText.toLowerCase();
          if (text.includes(query)) {
            cell.classList.add('search-match');
          } else {
            cell.classList.remove('search-match');
          }
        });

        if (searchMatches.length > 0) {
          if (activeMatchIndex >= searchMatches.length || activeMatchIndex < 0) {
            activeMatchIndex = 0;
          }
          // Highlight active match cell specifically
          const activeMatch = searchMatches[activeMatchIndex];
          const activeCell = viewport.querySelector('td[data-row="' + activeMatch.row + '"][data-col="' + activeMatch.col + '"]');
          if (activeCell) {
            activeCell.classList.add('search-match-active');
          }
          findCount.textContent = (activeMatchIndex + 1) + ' of ' + searchMatches.length;
        } else {
          activeMatchIndex = -1;
          findCount.textContent = '0 of 0';
        }
      }

      function navigateSearch(direction) {
        if (searchMatches.length === 0) return;
        activeMatchIndex = (activeMatchIndex + direction + searchMatches.length) % searchMatches.length;
        updateActiveMatch();
      }

      findInput.addEventListener('input', performSearch);

      findInput.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
          event.preventDefault();
          if (event.shiftKey) {
            navigateSearch(-1);
          } else {
            navigateSearch(1);
          }
        } else if (event.key === 'Escape') {
          event.preventDefault();
          hideFindWidget();
        }
      });

      findPrev.addEventListener('click', () => navigateSearch(-1));
      findNext.addEventListener('click', () => navigateSearch(1));
      findClose.addEventListener('click', hideFindWidget);

      // Cell Popover logic
      const cellPopover = document.getElementById('cell-popover');
      const cellPopoverTextarea = document.getElementById('cell-popover-textarea');
      let popoverCurrentCell = null;

      function selectCell(row, col) {
        if (editingCell) {
          commitEditing(editingCell, false);
        }
        if (popoverCurrentCell) {
          closeCellPopover(false);
        }

        const sheet = currentSheet();
        const rows = sheet.rows || [];
        const columnCount = columnCountForSheet(sheet, rows);

        if (rows.length === 0) return;

        const minRow = (state.hasHeader && rows.length > 1) ? 1 : 0;
        const targetRow = Math.max(minRow, Math.min(row, rows.length - 1));
        const targetCol = Math.max(0, Math.min(col, columnCount - 1));

        state.selected = { row: targetRow, col: targetCol };

        let needsRender = false;
        if (targetRow >= state.visibleRowCount + (state.hasHeader ? 1 : 0)) {
          state.visibleRowCount = Math.min(rows.length, targetRow + 500);
          needsRender = true;
        }

        if (needsRender) {
          renderGrid();
        } else {
          viewport.querySelectorAll('td.cell-selected').forEach(c => c.classList.remove('cell-selected'));
          const cell = viewport.querySelector('td[data-row="' + targetRow + '"][data-col="' + targetCol + '"]');
          if (cell) {
            cell.classList.add('cell-selected');
          }
        }

        const cell = viewport.querySelector('td[data-row="' + targetRow + '"][data-col="' + targetCol + '"]');
        if (cell) {
          cell.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        }

        // Update Formula Bar
        if (cellAddress && formulaInput) {
          const colLabel = columnLabel(targetCol);
          const displayRowNum = state.hasHeader ? targetRow : (targetRow + 1);
          cellAddress.textContent = colLabel + displayRowNum;
          const value = sheet.rows[targetRow]?.[targetCol] ?? '';
          formulaInput.value = value;
        }
      }

      function openCellPopover(cell) {
        if (editingCell) {
          commitEditing(editingCell, false);
        }

        popoverCurrentCell = cell;
        const row = Number(cell.dataset.row);
        const col = Number(cell.dataset.col);
        const sheet = currentSheet();
        const value = sheet.rows[row]?.[col] ?? '';

        cellPopoverTextarea.value = value;
        cellPopover.classList.remove('hidden');

        // Reset any inline sizes before measuring
        cellPopover.style.width = '';
        cellPopover.style.height = '';

        const cellRect = cell.getBoundingClientRect();
        let popoverWidth = cellPopover.offsetWidth;
        let popoverHeight = cellPopover.offsetHeight;

        // Constraint height & width to viewport
        const maxW = window.innerWidth - 32;
        const maxH = window.innerHeight - 32;
        if (popoverWidth > maxW) {
          popoverWidth = maxW;
          cellPopover.style.width = maxW + 'px';
        }
        if (popoverHeight > maxH) {
          popoverHeight = maxH;
          cellPopover.style.height = maxH + 'px';
        }

        let top = cellRect.bottom + 4;
        let left = cellRect.left;

        // Position vertically
        if (top + popoverHeight > window.innerHeight - 8) {
          // Try placing it above the cell
          top = cellRect.top - popoverHeight - 4;
          if (top < 8) {
            // If it doesn't fit either above or below, place it at the top and set height
            top = 8;
            cellPopover.style.height = (window.innerHeight - 16) + 'px';
          }
        }

        // Position horizontally
        if (left + popoverWidth > window.innerWidth - 8) {
          left = window.innerWidth - popoverWidth - 16;
        }
        left = Math.max(8, left);

        cellPopover.style.top = top + 'px';
        cellPopover.style.left = left + 'px';

        // Focus the textarea and select text
        cellPopoverTextarea.focus();
        cellPopoverTextarea.select();
        cellPopoverTextarea.scrollTop = 0;
        cellPopoverTextarea.scrollLeft = 0;
      }

      function closeCellPopover(revert = false) {
        if (cellPopover.classList.contains('hidden')) {
          return;
        }
        if (popoverCurrentCell) {
          const cell = popoverCurrentCell;
          const row = Number(cell.dataset.row);
          const col = Number(cell.dataset.col);
          const sheet = currentSheet();
          const original = sheet.rows[row]?.[col] ?? '';
          const value = revert ? original : cellPopoverTextarea.value;

          if (!revert && value !== original) {
            if (sheet.rows[row]) {
              sheet.rows[row][col] = value;
            }
            vscode.postMessage({
              type: 'cellEdit',
              sheetIndex: state.activeSheetIndex,
              row,
              col,
              value,
            });
            cell.textContent = value;
            cell.title = value;
          }

          if (formulaInput && state.selected && state.selected.row === row && state.selected.col === col) {
            formulaInput.value = value;
          }
        }
        cellPopover.classList.add('hidden');
        popoverCurrentCell = null;
      }

      cellPopoverTextarea.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' && !event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey) {
          event.preventDefault();
          closeCellPopover(false);
          viewport.focus({ preventScroll: true });
        } else if (event.key === 'Escape') {
          event.preventDefault();
          closeCellPopover(true);
          viewport.focus({ preventScroll: true });
        }
      });

      viewport.addEventListener('dblclick', (event) => {
        const cell = event.target.closest('tbody td.cell');
        if (cell) {
          event.preventDefault();
          const overflows = (cell.scrollWidth - cell.clientWidth > 1) || (cell.scrollHeight - cell.clientHeight > 1);
          if (overflows) {
            openCellPopover(cell);
          } else {
            startEditing(cell);
          }
        }
      });

      document.addEventListener('click', (event) => {
        if (cellPopover.classList.contains('hidden')) {
          return;
        }
        if (!cellPopover.contains(event.target) && !event.target.closest('td.cell')) {
          closeCellPopover(false);
        }
      });

      document.addEventListener('copy', (event) => {
        const activeTag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
        if (activeTag === 'input' || activeTag === 'textarea') {
          return;
        }

        if (state.selected) {
          const sheet = currentSheet();
          const value = sheet.rows[state.selected.row]?.[state.selected.col] ?? '';
          event.clipboardData.setData('text/plain', value);
          event.preventDefault();
        }
      });

      viewport.addEventListener('click', (event) => {
        const cell = event.target.closest('td[data-row]');
        if (!cell) {
          return;
        }
        event.stopPropagation(); // Prevent document click handler

        viewport.focus({ preventScroll: true });

        const row = Number(cell.dataset.row);
        const col = Number(cell.dataset.col);
        selectCell(row, col);
      });

      document.addEventListener('keydown', (event) => {
        if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') {
          event.preventDefault();
          showFindWidget();
          return;
        }

        const ctrlOrCmd = event.ctrlKey || event.metaKey;
        if (!ctrlOrCmd) return;
        const key = event.key.toLowerCase();
        const isUndo = key === 'z' && !event.shiftKey;
        const isRedo = (key === 'z' && event.shiftKey) || key === 'y';

        if (!isUndo && !isRedo) return;

        // Don't hijack undo/redo while inline-editing a cell or while focus is in a text input/textarea/popover.
        const tag = document.activeElement && document.activeElement.tagName ? document.activeElement.tagName.toLowerCase() : '';
        if (editingCell || tag === 'input' || tag === 'textarea') {
          return;
        }

        event.preventDefault();
        vscode.postMessage({ type: isUndo ? 'undo' : 'redo' });
      });

      viewport.addEventListener('keydown', (event) => {
        if (event.altKey && state.sheetCount > 1) {
          if (event.key === 'ArrowRight') {
            event.preventDefault();
            setActiveSheet(Math.min(state.sheets.length - 1, state.activeSheetIndex + 1));
            return;
          }
          if (event.key === 'ArrowLeft') {
            event.preventDefault();
            setActiveSheet(Math.max(0, state.activeSheetIndex - 1));
            return;
          }
        }

        // Arrow navigation when NOT editing inline and NOT editing in popover
        if (!editingCell && (!popoverCurrentCell || cellPopover.classList.contains('hidden'))) {
          if (state.selected) {
            let handled = false;
            const { row, col } = state.selected;

            if (event.key === 'ArrowUp') {
              selectCell(row - 1, col);
              handled = true;
            } else if (event.key === 'ArrowDown') {
              selectCell(row + 1, col);
              handled = true;
            } else if (event.key === 'ArrowLeft') {
              selectCell(row, col - 1);
              handled = true;
            } else if (event.key === 'ArrowRight') {
              selectCell(row, col + 1);
              handled = true;
            } else if (event.key === 'Tab') {
              if (event.shiftKey) {
                selectCell(row, col - 1);
              } else {
                selectCell(row, col + 1);
              }
              handled = true;
            } else if (event.key === 'Enter') {
              const cell = viewport.querySelector('td[data-row="' + row + '"][data-col="' + col + '"]');
              if (cell) {
                const overflows = (cell.scrollWidth - cell.clientWidth > 1) || (cell.scrollHeight - cell.clientHeight > 1);
                if (overflows) {
                  openCellPopover(cell);
                } else {
                  startEditing(cell);
                }
              }
              handled = true;
            } else if (event.key === 'Backspace' || event.key === 'Delete') {
              const sheet = currentSheet();
              if (sheet.rows[row]) {
                sheet.rows[row][col] = '';
                const cell = viewport.querySelector('td[data-row="' + row + '"][data-col="' + col + '"]');
                if (cell) {
                  cell.textContent = '';
                  cell.title = '';
                }
                vscode.postMessage({
                  type: 'cellEdit',
                  sheetIndex: state.activeSheetIndex,
                  row,
                  col,
                  value: '',
                });
              }
              handled = true;
            } else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
              const cell = viewport.querySelector('td[data-row="' + row + '"][data-col="' + col + '"]');
              if (cell) {
                const overflows = (cell.scrollWidth - cell.clientWidth > 1) || (cell.scrollHeight - cell.clientHeight > 1);
                if (overflows) {
                  openCellPopover(cell);
                  cellPopoverTextarea.value = event.key;
                } else {
                  startEditing(cell);
                  cell.innerText = event.key;
                  const selection = window.getSelection();
                  if (selection) {
                    const range = document.createRange();
                    range.selectNodeContents(cell);
                    range.collapse(false);
                    selection.removeAllRanges();
                    selection.addRange(range);
                  }
                }
              }
              handled = true;
            }

            if (handled) {
              event.preventDefault();
              return;
            }
          }
        }

        const cell = event.target.closest('td[data-row]');
        if (!cell || cell !== editingCell) {
          return;
        }

        if (event.key === 'Enter') {
          if (event.shiftKey || event.altKey) {
            // Allow inserting line breaks using default behavior
          } else {
            event.preventDefault();
            const row = Number(cell.dataset.row);
            const col = Number(cell.dataset.col);
            commitEditing(cell, false);
            cell.blur();
            viewport.focus({ preventScroll: true });
            selectCell(row + 1, col);
          }
        } else if (event.key === 'Tab') {
          event.preventDefault();
          const row = Number(cell.dataset.row);
          const col = Number(cell.dataset.col);
          commitEditing(cell, false);
          cell.blur();
          viewport.focus({ preventScroll: true });
          if (event.shiftKey) {
            selectCell(row, col - 1);
          } else {
            selectCell(row, col + 1);
          }
        } else if (event.key === 'Escape') {
          event.preventDefault();
          commitEditing(cell, true);
          cell.blur();
          viewport.focus({ preventScroll: true });
        }
      });

      viewport.addEventListener('blur', (event) => {
        const cell = event.target.closest('td[data-row]');
        if (cell && cell === editingCell) {
          commitEditing(cell, false);
        }
      }, true);

      viewport.addEventListener('focusout', (event) => {
        const cell = event.target.closest('td[data-row]');
        if (cell && cell === editingCell) {
          commitEditing(cell, false);
        }
      });

      viewport.addEventListener('input', (event) => {
        const cell = event.target.closest('td.editing');
        if (cell && cell === editingCell) {
          if (formulaInput) {
            formulaInput.value = cell.innerText;
          }
        }
      });

      function compareValues(a, b) {
        if (a === b) return 0;
        if (a === '') return 1;
        if (b === '') return -1;
        const numA = Number(a);
        const numB = Number(b);
        if (!isNaN(numA) && !isNaN(numB)) {
          return numA - numB;
        }
        return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
      }

      function sortGrid(colIndex) {
        const sheet = currentSheet();
        const allRows = sheet.rows || [];
        if (allRows.length === 0) return;

        const hasHeader = Boolean(state.hasHeader && allRows.length > 0);
        const headerRow = hasHeader ? allRows[0] : null;
        const rows = hasHeader ? allRows.slice(1) : allRows;
        if (rows.length === 0) return;

        rows.forEach((row, idx) => {
          if (row._originalIndex === undefined) {
            row._originalIndex = idx;
          }
        });

        let nextDir = 'asc';
        if (state.sortColumn === colIndex) {
          if (state.sortDirection === 'asc') nextDir = 'desc';
          else if (state.sortDirection === 'desc') nextDir = null;
        }

        state.sortColumn = colIndex;
        state.sortDirection = nextDir;

        if (nextDir === 'asc') {
          rows.sort((a, b) => compareValues(String(a[colIndex] ?? '').trim(), String(b[colIndex] ?? '').trim()));
        } else if (nextDir === 'desc') {
          rows.sort((a, b) => compareValues(String(b[colIndex] ?? '').trim(), String(a[colIndex] ?? '').trim()));
        } else {
          rows.sort((a, b) => a._originalIndex - b._originalIndex);
          state.sortColumn = -1;
        }

        sheet.rows = headerRow ? [headerRow, ...rows] : rows;

        renderGrid();

        // Notify VS Code of the row sorting change
        vscode.postMessage({
          type: 'sortSheet',
          sheetIndex: state.activeSheetIndex,
          colIndex: colIndex,
          direction: nextDir,
          hasHeader: hasHeader
        });
      }

      // Bind Formula Input Events
      if (formulaInput) {
        formulaInput.addEventListener('input', () => {
          if (state.selected) {
            const { row, col } = state.selected;
            const value = formulaInput.value;
            const sheet = currentSheet();
            if (sheet.rows[row]) {
              sheet.rows[row][col] = value;
            }
            const cell = viewport.querySelector('td[data-row="' + row + '"][data-col="' + col + '"]');
            if (cell) {
              cell.textContent = value;
              cell.title = value;
            }
          }
        });

        formulaInput.addEventListener('keydown', (event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            if (state.selected) {
              const { row, col } = state.selected;
              const value = formulaInput.value;
              vscode.postMessage({
                type: 'cellEdit',
                sheetIndex: state.activeSheetIndex,
                row,
                col,
                value,
              });
              viewport.focus({ preventScroll: true });
            }
          } else if (event.key === 'Escape') {
            event.preventDefault();
            if (state.selected) {
              const { row, col } = state.selected;
              const sheet = currentSheet();
              const originalValue = sheet.rows[row]?.[col] ?? '';
              formulaInput.value = originalValue;
              const cell = viewport.querySelector('td[data-row="' + row + '"][data-col="' + col + '"]');
              if (cell) {
                cell.textContent = originalValue;
              }
              viewport.focus({ preventScroll: true });
            }
          }
        });

        formulaInput.addEventListener('blur', () => {
          if (state.selected) {
            const { row, col } = state.selected;
            const value = formulaInput.value;
vscode.postMessage({
                type: 'cellEdit',
                sheetIndex: state.activeSheetIndex,
                row,
                col,
                value,
              });
          }
        });
      }

      // Bind Column Header click for sorting
      viewport.addEventListener('click', (event) => {
        const th = event.target.closest('thead th[data-col]');
        if (th) {
          if (event.target.classList.contains('col-resizer')) {
            return;
          }
          const colIndex = Number(th.dataset.col);
          sortGrid(colIndex);
        }
      });

      function resetSortState() {
        state.sortColumn = -1;
        state.sortDirection = null;
      }

      window.addEventListener('message', (event) => {
        const message = event.data;
        if (message.type === 'refresh') {
          renderAll();
        } else if (message.type === 'updateState') {
          Object.assign(state, message.state);
          // Server-side changes (undo/redo/loadAll) reset client sort visuals.
          resetSortState();
          // Make sure the selected row is visible after a server-driven row change.
          if (state.selected && state.selected.row >= state.visibleRowCount + (state.hasHeader ? 1 : 0)) {
            state.visibleRowCount = Math.min(
              (currentSheet().rows || []).length,
              state.selected.row + 500
            );
          }
          renderAll();
          // Refresh the formula bar to reflect the (possibly reverted) selected cell.
          if (state.selected && cellAddress && formulaInput) {
            const sheet = currentSheet();
            const { row, col } = state.selected;
            const displayRowNum = state.hasHeader ? row : (row + 1);
            cellAddress.textContent = columnLabel(col) + displayRowNum;
            formulaInput.value = sheet.rows[row]?.[col] ?? '';
          }
        }
      });

      renderAll();
      vscode.postMessage({ type: 'ready' });
    })();
  </script>
</body>
</html>`;
}

class SpreadsheetDocument {
  constructor(snapshot) {
    const cloned = cloneSnapshot(snapshot);
    this.uri = snapshot.uri;
    Object.assign(this, cloned);
    this.uri = snapshot.uri;
    this.version = 1;
    this.undoStack = [];
    this.redoStack = [];
    this.maxUndoEntries = 500;
    this.hasHeader = snapshot.hasHeader;
  }

  dispose() {
    // No-op cleanup required by Custom Document interface
  }

  sortSheet(colIndex, direction, sheetIndex, hasHeader) {
    const targetIndex = sheetIndex == null ? this.activeSheetIndex : sheetIndex;
    const sheet = this.sheets[targetIndex] || this.activeSheet;
    const allRows = sheet.rows || [];
    const beforeRows = allRows.map((row) => row.slice());

    const isHeaderMode = Boolean((hasHeader !== undefined ? hasHeader : this.hasHeader) && allRows.length > 0);
    const headerRow = isHeaderMode ? allRows[0] : null;
    const rows = isHeaderMode ? allRows.slice(1) : allRows;

    rows.forEach((row, idx) => {
      if (row._originalIndex === undefined) {
        row._originalIndex = idx;
      }
    });

    const compareValues = (a, b) => {
      if (a === b) return 0;
      if (a === '') return 1;
      if (b === '') return -1;
      const numA = Number(a);
      const numB = Number(b);
      if (!isNaN(numA) && !isNaN(numB)) {
        return numA - numB;
      }
      return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
    };

    if (direction === 'asc') {
      rows.sort((a, b) => compareValues(String(a[colIndex] ?? '').trim(), String(b[colIndex] ?? '').trim()));
    } else if (direction === 'desc') {
      rows.sort((a, b) => compareValues(String(b[colIndex] ?? '').trim(), String(a[colIndex] ?? '').trim()));
    } else {
      rows.sort((a, b) => a._originalIndex - b._originalIndex);
    }

    sheet.rows = headerRow ? [headerRow, ...rows] : rows;

    const afterRows = sheet.rows.map((row) => row.slice());
    this._pushUndo({ type: 'sort', sheetIndex: targetIndex, beforeRows, afterRows, colIndex, direction });
    this.dirty = true;
    this.version += 1;
  }

  setActiveSheet(index) {
    if (typeof index !== 'number' || isNaN(index)) return;
    if (index < 0 || index >= this.sheets.length) return;
    this.activeSheetIndex = index;
  }

  _pushUndo(op) {
    this.undoStack.push(op);
    if (this.undoStack.length > this.maxUndoEntries) {
      this.undoStack.shift();
    }
    this.redoStack.length = 0;
  }

  undo() {
    const op = this.undoStack.pop();
    if (!op) return false;
    if (op.type === 'cell') {
      const sheet = this.sheets[op.sheetIndex];
      if (sheet && sheet.rows[op.row]) {
        sheet.rows[op.row][op.col] = op.oldValue;
      }
    } else if (op.type === 'sort') {
      const sheet = this.sheets[op.sheetIndex];
      if (sheet) {
        sheet.rows = op.beforeRows.map((row) => row.slice());
      }
    }
    this.redoStack.push(op);
    this.dirty = true;
    this.version += 1;
    return true;
  }

  redo() {
    const op = this.redoStack.pop();
    if (!op) return false;
    if (op.type === 'cell') {
      const sheet = this.sheets[op.sheetIndex];
      if (sheet) {
        while (sheet.rows.length <= op.row) sheet.rows.push([]);
        while (sheet.rows[op.row].length <= op.col) sheet.rows[op.row].push('');
        sheet.rows[op.row][op.col] = op.newValue;
      }
    } else if (op.type === 'sort') {
      const sheet = this.sheets[op.sheetIndex];
      if (sheet) {
        sheet.rows = op.afterRows.map((row) => row.slice());
      }
    }
    this.undoStack.push(op);
    this.dirty = true;
    this.version += 1;
    return true;
  }

  get isSpreadsheet() {
    return this.kind === 'workbook';
  }

  get activeSheet() {
    return this.sheets[this.activeSheetIndex] || this.sheets[0];
  }

  async loadAll() {
    const uri = this.uri instanceof vscode.Uri ? this.uri : vscode.Uri.parse(this.sourceUri || String(this.uri));
    const ext = this.fileExt || getExtension(uri);

    if (this.kind === 'delimited') {
      const fileSize = await getUriSize(uri);
      const bytes = await readUriBytes(uri);
      const text = Buffer.from(bytes).toString('utf8');
      const fresh = delimitedToSnapshot(uri, text, Infinity, bytes.length, Infinity);
      Object.assign(this, cloneSnapshot(fresh));
      this.uri = uri;
      this.editable = true;
      this.undoStack = [];
      this.redoStack = [];
      return;
    }

    if (this.kind === 'workbook') {
      const bytes = await readUriBytes(uri);
      const fresh = workbookToSnapshot(uri, bytes, Infinity, true);
      Object.assign(this, cloneSnapshot(fresh));
      this.uri = uri;
      this.editable = true;
      this.undoStack = [];
      this.redoStack = [];
      return;
    }

    throw new Error(`Unsupported file type: ${ext || 'unknown'}`);
  }

  async ensureEditable() {
    if (this.editable) {
      return;
    }
    await this.loadAll();
  }

  toWebviewState() {
    const config = getConfig();
    const sheet = this.activeSheet;
    const rows = sheet ? sheet.rows || [] : [];
    let hasHeader;
    if (this.hasHeader !== undefined) {
      hasHeader = this.hasHeader;
    } else if (config.hasHeader === 'always') {
      hasHeader = true;
    } else if (config.hasHeader === 'never') {
      hasHeader = false;
    } else {
      hasHeader = detectHasHeader(rows);
    }

    return {
      fileName: this.fileName,
      kind: this.kind,
      editable: Boolean(this.editable),
      dirty: Boolean(this.dirty),
      sheetCount: this.sheets.length,
      sheets: this.sheets,
      activeSheetIndex: this.activeSheetIndex,
      rawPreview: this.rawPreview,
      delimiter: this.delimiter,
      lineEnding: this.lineEnding,
      limited: Boolean(this.limited),
      hasHeader,
      settings: config,
    };
  }

  updateCell(row, col, value, sheetIndex) {
    const targetIndex = sheetIndex == null ? this.activeSheetIndex : sheetIndex;
    const sheet = this.sheets[targetIndex] || this.activeSheet;
    if (!sheet) {
      return;
    }
    while (sheet.rows.length <= row) {
      sheet.rows.push([]);
    }
    while (sheet.rows[row].length <= col) {
      sheet.rows[row].push('');
    }
    const oldValue = sheet.rows[row][col];
    const newValue = String(value ?? '');
    if (oldValue === newValue) {
      return;
    }
    sheet.rows[row][col] = newValue;
    this._pushUndo({
      type: 'cell',
      sheetIndex: targetIndex,
      row,
      col,
      oldValue,
      newValue,
    });
    this.dirty = true;
    this.version += 1;
  }

  serialize() {
    if (this.kind === 'delimited') {
      const sheet = this.activeSheet;
      return { type: 'text', content: serializeDelimitedRows(sheet.rows, this.delimiter, this.lineEnding) };
    }

    if (this.kind === 'workbook') {
      const workbook = xlsx.utils.book_new();
      for (let i = 0; i < this.sheets.length; i++) {
        const sheet = this.sheets[i];
        const worksheet = xlsx.utils.aoa_to_sheet(sheet.rows, { sheetStubs: true });
        xlsx.utils.book_append_sheet(workbook, worksheet, sheet.name || `Sheet${i + 1}`);
      }

      return {
        type: 'buffer',
        content: xlsx.write(workbook, {
          type: 'buffer',
          bookType: this.fileExt === '.xls' ? 'xls' : 'xlsx',
          cellDates: true,
        }),
      };
    }

    throw new Error('Unsupported document type.');
  }

  async saveTo(destination) {
    const content = this.serialize();
    const buffer = content.type === 'buffer' ? content.content : Buffer.from(content.content, 'utf8');
    await vscode.workspace.fs.writeFile(destination, buffer);
    this.dirty = false;
  }

  async revert() {
    const fresh = await loadSnapshot(this.uri);
    const cloned = cloneSnapshot(fresh);
    Object.assign(this, cloned);
    this.uri = fresh.uri;
    this.version += 1;
    this.undoStack = [];
    this.redoStack = [];
  }
}

class CsvEditorProvider {
  constructor(context) {
    this.context = context;
    this._changeEmitter = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this._changeEmitter.event;
  }

  static register(context) {
    return vscode.window.registerCustomEditorProvider(
      VIEW_TYPE,
      new CsvEditorProvider(context),
      {
        webviewOptions: { retainContextWhenHidden: true },
        supportsMultipleEditorsPerDocument: true,
      }
    );
  }

  async openCustomDocument(uri, openContext) {
    if (openContext && openContext.backupId) {
      const backupUri = vscode.Uri.parse(openContext.backupId);
      const backupBytes = await vscode.workspace.fs.readFile(backupUri);
      const payload = JSON.parse(Buffer.from(backupBytes).toString('utf8'));
      payload.snapshot.uri = vscode.Uri.parse(payload.uri);
      return new SpreadsheetDocument(payload.snapshot);
    }

    const snapshot = await loadSnapshot(uri);
    return new SpreadsheetDocument(snapshot);
  }

  async resolveCustomEditor(document, webviewPanel) {
    const fileName = document.fileName;
    const nonce = crypto.randomBytes(16).toString('base64');

    webviewPanel.webview.options = {
      enableScripts: true,
    };

    webviewPanel.webview.html = buildLoadingHtml(fileName, nonce);

    setTimeout(() => {
      try {
        webviewPanel.webview.html = buildEditorHtml(nonce, document.toWebviewState());
      } catch (error) {
        const messageText = error instanceof Error ? error.message : String(error);
        webviewPanel.webview.html = buildErrorHtml(fileName, messageText, nonce);
        return;
      }

      const messageSubscription = webviewPanel.webview.onDidReceiveMessage(async (message) => {
        try {
          if (message.type === 'ready') {
            return;
          }

          if (message.type === 'cellEdit') {
            const wasLimited = document.limited;
            await document.ensureEditable();
            const sheetIndex = Number.isFinite(message.sheetIndex)
              ? Number(message.sheetIndex)
              : document.activeSheetIndex;
            document.updateCell(Number(message.row), Number(message.col), message.value, sheetIndex);
            this._changeEmitter.fire({ document });

            if (wasLimited) {
              webviewPanel.webview.postMessage({
                type: 'updateState',
                state: document.toWebviewState()
              });
            }
            return;
          }

          if (message.type === 'toggleHeader') {
            document.hasHeader = Boolean(message.hasHeader);
            return;
          }

          if (message.type === 'sortSheet') {
            await document.ensureEditable();
            const sheetIndex = Number.isFinite(message.sheetIndex)
              ? Number(message.sheetIndex)
              : document.activeSheetIndex;
            document.sortSheet(Number(message.colIndex), message.direction, sheetIndex, message.hasHeader);
            this._changeEmitter.fire({ document });
            return;
          }

          if (message.type === 'setActiveSheet') {
            document.setActiveSheet(Number(message.index));
            return;
          }

          if (message.type === 'undo') {
            await document.ensureEditable();
            const applied = document.undo();
            if (applied) {
              this._changeEmitter.fire({ document });
              webviewPanel.webview.postMessage({
                type: 'updateState',
                state: document.toWebviewState()
              });
            }
            return;
          }

          if (message.type === 'redo') {
            const applied = document.redo();
            if (applied) {
              this._changeEmitter.fire({ document });
              webviewPanel.webview.postMessage({
                type: 'updateState',
                state: document.toWebviewState()
              });
            }
            return;
          }

          if (message.type === 'loadAll') {
            await document.loadAll();
            webviewPanel.webview.postMessage({
              type: 'updateState',
              state: document.toWebviewState()
            });
            return;
          }
        } catch (error) {
          const messageText = error instanceof Error ? error.message : String(error);
          vscode.window.showErrorMessage(messageText);
        }
      });

      webviewPanel.onDidDispose(() => {
        messageSubscription.dispose();
      });
    }, 0);
  }

  async saveCustomDocument(document) {
    if (!(document instanceof SpreadsheetDocument)) {
      return;
    }

    await document.ensureEditable();
    await document.saveTo(document.uri);
  }

  async saveCustomDocumentAs(document, destination) {
    if (!(document instanceof SpreadsheetDocument)) {
      return;
    }

    if (!document.editable) {
      await document.ensureEditable();
    }

    await document.saveTo(destination);
  }

  async revertCustomDocument(document) {
    if (!(document instanceof SpreadsheetDocument)) {
      return;
    }

    await document.revert();
    this._changeEmitter.fire({ document });
  }

  async backupCustomDocument(document, context) {
    if (!(document instanceof SpreadsheetDocument)) {
      throw new Error('Cannot back up an unknown document type.');
    }

    const backupUri = context.destination;
    const payload = {
      uri: document.uri.toString(),
      snapshot: cloneSnapshot(document),
    };

    await vscode.workspace.fs.writeFile(backupUri, Buffer.from(JSON.stringify(payload), 'utf8'));

    return {
      id: backupUri.toString(),
      delete: () => vscode.workspace.fs.delete(backupUri, { useTrash: false }).then(undefined, () => undefined),
    };
  }
}

async function openTableCommand(resourceUri) {
  let targetUri = resourceUri;

  if (!targetUri && vscode.window.activeTextEditor) {
    targetUri = vscode.window.activeTextEditor.document.uri;
  }

  if (!targetUri) {
    const picked = await vscode.window.showOpenDialog({
      canSelectMany: false,
      filters: {
        'Data files': ['csv', 'tsv', 'xlsx', 'xls'],
      },
    });

    if (!picked || picked.length === 0) {
      return;
    }

    targetUri = picked[0];
  }

  const ext = getExtension(targetUri);
  if (!['.csv', '.tsv', '.xlsx', '.xls'].includes(ext)) {
    vscode.window.showWarningMessage('Open as Table currently supports CSV, TSV, XLSX, and XLS files.');
    return;
  }

  await vscode.commands.executeCommand('vscode.openWith', targetUri, VIEW_TYPE);
}

function activate(ctx) {
  ctx.subscriptions.push(CsvEditorProvider.register(ctx));
  ctx.subscriptions.push(vscode.commands.registerCommand('csvViewer.openTable', openTableCommand));
}

function deactivate() {}

module.exports = { activate, deactivate };

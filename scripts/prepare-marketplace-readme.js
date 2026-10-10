'use strict';

const fs = require('fs');
const path = require('path');

const manifest = require('../package.json');
const sourcePath = path.join(__dirname, '..', 'README.md');
const outputPath = path.join(__dirname, '..', '.vscode', 'README.marketplace.md');
const readme = fs.readFileSync(sourcePath, 'utf8');
const badgeUrls = [
  'https://shieldcn.dev/vscode/installs/aakashrohilla/csv-xlsx-table-viewer.svg',
  'https://shieldcn.dev/vscode/rating/aakashrohilla/csv-xlsx-table-viewer.svg',
  'https://shieldcn.dev/vscode/v/aakashrohilla/csv-xlsx-table-viewer.svg',
];

const start = readme.indexOf('<p align="center">');
const endTag = '</p>';
const end = start === -1 ? -1 : readme.indexOf(endTag, start);

if (start === -1 || end === -1) {
  throw new Error('Could not find the centered Marketplace badge block in README.md');
}

const badgeBlock = readme.slice(start, end + endTag.length);
if (!badgeUrls.every((url) => badgeBlock.includes(url))) {
  throw new Error('The canonical Marketplace badge block does not contain all expected ShieldCN URLs');
}

const packageReadme = `${readme.slice(0, start)}${readme.slice(end + endTag.length)}`;
fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, packageReadme);

console.log(`Prepared Marketplace README at ${path.relative(process.cwd(), outputPath)} (ShieldCN SVG badges omitted for vsce compatibility).`);

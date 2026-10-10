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

const badgeBlock = readme
  .match(/<p align="center">[\s\S]*?<\/p>/g)
  ?.find((block) => badgeUrls.every((url) => block.includes(url)));

if (!badgeBlock) {
  throw new Error('The canonical Marketplace badge block does not contain all expected ShieldCN URLs');
}

const packageReadme = readme.replace(badgeBlock, '');
fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, packageReadme);

console.log(`Prepared Marketplace README at ${path.relative(process.cwd(), outputPath)} (ShieldCN SVG badges omitted for vsce compatibility).`);

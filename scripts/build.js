'use strict';

const fs = require('fs');
const path = require('path');
const esbuild = require('esbuild');

const manifest = require('../package.json');
const outputPath = path.join('PrivateAssets', 'builds', manifest.version, 'extension.js');
const expectedMain = `./PrivateAssets/builds/${manifest.version}/extension.js`;

if (manifest.main !== expectedMain) {
  throw new Error(`package.json main must be ${expectedMain} for version ${manifest.version}`);
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.mkdirSync(path.join('PrivateAssets', 'builds', manifest.version, 'packages'), { recursive: true });

esbuild.buildSync({
  entryPoints: ['src/extension.js'],
  bundle: true,
  platform: 'node',
  format: 'cjs',
  external: ['vscode'],
  outfile: outputPath,
});

console.log(`Built ${outputPath}`);

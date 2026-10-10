'use strict';

const fs = require('fs');
const path = require('path');
const esbuild = require('esbuild');

const manifest = require('../package.json');
const outputPath = path.join('dist', manifest.version, 'extension.js');

if (manifest.main !== `./${outputPath}`) {
  throw new Error(`package.json main must be ./${outputPath} for version ${manifest.version}`);
}

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.mkdirSync(path.join('dist', manifest.version, 'packages'), { recursive: true });

esbuild.buildSync({
  entryPoints: ['src/extension.js'],
  bundle: true,
  platform: 'node',
  format: 'cjs',
  external: ['vscode'],
  outfile: outputPath,
});

console.log(`Built ${outputPath}`);

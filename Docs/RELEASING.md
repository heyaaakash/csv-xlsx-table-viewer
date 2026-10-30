# Release process

The Marketplace extension ID is `aakashrohilla.csv-xlsx-table-viewer`. Keep that identity stable across releases.

## Prepare and inspect

1. Update the version in `package.json` and `package-lock.json`, and point `main` to `./PrivateAssets/builds/<version>/extension.js`. Put user-facing release notes in the GitHub Release description.
2. From a clean checkout, run `npm ci` and `npm run build`.
3. Generate the Marketplace README and package the extension. For version `1.1.0`:

   ```sh
   npm run prepare:marketplace-readme
   vsce package --githubBranch main --readme-path .vscode/README.marketplace.md --out PrivateAssets/builds/1.1.0/packages/csv-xlsx-table-viewer-1.1.0.vsix
   ```

4. Inspect the VSIX. Confirm its ID, version, README, generated bundle, icon, license, and included docs; check that local data and development-only files are absent.
5. Install that exact VSIX in a clean VS Code profile. Check opening, editing, saving, and reopening representative files, and record the host and limitations. A successful build is not a behavioral test.

The package-only README omits the ShieldCN SVG badges because `vsce` only accepts images from trusted badge providers. The canonical GitHub README retains the requested badges. The generated README is excluded from the VSIX file list and added once as `extension/README.md`.

Upload through the Marketplace publisher account. After approval, verify the public listing and publish concise release notes with the matching source tag. Keep downloadable VSIX files in GitHub Releases rather than in the source tree.

# Development

```
npm install
npm run compile   # type check + bundle to dist/
npm run lint
npm test          # runs the tests in a downloaded VS Code
```

Press F5 in VS Code to launch the extension with `src/testdata` opened.

Office documents are converted by `script/saveAsPdf.ps1`, run in a hidden
`powershell.exe` process. No native modules are used, so no rebuild is needed
when VS Code (Electron) is updated.

Packaging (creates `mpdf-<version>.vsix`):
```
npm run package
```

## Release

Releases are published by GitHub Actions (`.github/workflows/release.yml`)
when a `v*` tag is pushed. It runs the tests, publishes to the Marketplace and
creates a GitHub release with the `.vsix` and the CHANGELOG section as notes.

1. Add the new version's section to `CHANGELOG.md` (`## [x.y.z] - date`) and commit.
2. Bump the version and push the tag:
   ```
   npm version patch        # or minor / major; commits and tags vx.y.z
   git push --follow-tags
   ```

The workflow needs the repository secret `VSCE_PAT`: a Personal Access Token
of the Microsoft account that owns the `kkzk` publisher, created at
https://dev.azure.com with Organization "All accessible organizations" and
scope "Marketplace > Manage". Tokens expire (at most one year); create a new
one and update the secret before then.

Manual fallback: `npx vsce login kkzk` then `npm run publish`, or upload the
`.vsix` at https://marketplace.visualstudio.com/manage/publishers/kkzk.

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

packaging:
```
npm run package
```

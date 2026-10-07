# Change Log

## [Unreleased]

## [0.2.0] - 2026-10-07

- Convert via a hidden PowerShell process instead of winax (no native rebuild on VS Code updates).
- Conversions run one at a time and always close Excel/Word, even on failure.
- Show the merged PDF in a built-in preview panel (PDF.js); no other PDF extension is needed. It keeps the page position and zoom when the PDF is re-merged.
- Merge with pdf-lib; conversion and merge errors are shown to the user.
- Fix: documents could not be reordered by drag and drop after reloading VS Code.
- Fix: dropping on empty space or on a worksheet put the document in the wrong position, and the new order was not saved.
- Fix: adding the same file twice created a duplicate entry.
- Tooling: TypeScript 6, ESLint 9 flat config, esbuild bundle, @vscode/test-cli. Requires VS Code 1.90 or later.

## [0.0.2] - 2022-05-04

- Add icon and images.
## [0.0.1] -2022-05-03

- Initial release
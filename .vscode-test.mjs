import { defineConfig } from '@vscode/test-cli';

export default defineConfig({
	files: 'out/test/**/*.test.js',
	workspaceFolder: 'src/testdata',
	mocha: { ui: 'tdd', timeout: 20000 },
});

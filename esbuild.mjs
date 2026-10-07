import * as esbuild from 'esbuild';
import * as fs from 'fs/promises';

const production = process.argv.includes('--production');
const watch = process.argv.includes('--watch');

/**
 * Prints the lines the background problem matcher in .vscode/tasks.json waits for.
 * Shared by both builds so that "finished" is printed only when neither is running.
 */
let running = 0;
function problemMatcherPlugin() {
	return {
		name: 'problem-matcher',
		setup(build) {
			build.onStart(() => {
				if (running++ === 0) {
					console.log('[watch] build started');
				}
			});
			build.onEnd(result => {
				for (const { text, location } of result.errors) {
					console.error(`✘ [ERROR] ${text}`);
					if (location) {
						console.error(`    ${location.file}:${location.line}:${location.column}:`);
					}
				}
				if (--running === 0) {
					console.log('[watch] build finished');
				}
			});
		},
	};
}

/** Files pdf.js loads at runtime in the preview webview. */
async function copyPdfjsAssets() {
	const from = 'node_modules/pdfjs-dist';
	const to = 'dist/pdfjs';
	await fs.rm(to, { recursive: true, force: true });
	await fs.mkdir(`${to}/wasm`, { recursive: true });
	await fs.copyFile(`${from}/legacy/build/pdf.worker.min.mjs`, `${to}/pdf.worker.min.mjs`);
	await fs.copyFile(`${from}/LICENSE`, `${to}/LICENSE`);
	await fs.cp(`${from}/cmaps`, `${to}/cmaps`, { recursive: true });
	await fs.cp(`${from}/standard_fonts`, `${to}/standard_fonts`, { recursive: true });
	// JPEG 2000 / JBIG2 decoders and color management; the JavaScript sandbox (quickjs) is not needed.
	for (const name of ['openjpeg.wasm', 'jbig2.wasm', 'qcms_bg.wasm', 'LICENSE_OPENJPEG', 'LICENSE_PDFJS_OPENJPEG', 'LICENSE_JBIG2', 'LICENSE_PDFJS_JBIG2', 'LICENSE_QCMS', 'LICENSE_PDFJS_QCMS']) {
		await fs.copyFile(`${from}/wasm/${name}`, `${to}/wasm/${name}`);
	}
}

const common = {
	bundle: true,
	minify: production,
	sourcemap: !production,
	sourcesContent: false,
	logLevel: 'silent',
	plugins: [problemMatcherPlugin()],
};

const contexts = await Promise.all([
	esbuild.context({
		...common,
		entryPoints: ['src/extension.ts'],
		format: 'cjs',
		platform: 'node',
		target: 'node20',
		outfile: 'dist/extension.js',
		external: ['vscode'],
	}),
	esbuild.context({
		...common,
		entryPoints: ['src/webview/preview.ts'],
		format: 'iife',
		platform: 'browser',
		target: 'chrome120',
		outfile: 'dist/webview/preview.js',
	}),
]);

await copyPdfjsAssets();

if (watch) {
	await Promise.all(contexts.map(ctx => ctx.watch()));
} else {
	await Promise.all(contexts.map(ctx => ctx.rebuild()));
	await Promise.all(contexts.map(ctx => ctx.dispose()));
}

import * as vscode from 'vscode';
import * as crypto from 'crypto';

/** A webview panel that shows the merged PDF with the bundled pdf.js (no other extension needed). */
export class PdfPreview implements vscode.Disposable {
	private panel?: vscode.WebviewPanel;
	private ready = false;
	private pending?: Uint8Array;
	private last?: { title: string; data: Uint8Array };

	constructor(private readonly extensionUri: vscode.Uri) { }

	/** Shows `data` in the preview, opening the panel beside the editor if needed. */
	show(title: string, data: Uint8Array): void {
		this.last = { title, data };
		if (!this.panel) {
			this.panel = this.createPanel();
		}
		this.panel.title = title;
		if (!this.panel.visible) {
			this.panel.reveal(undefined, true);
		}
		this.post(data);
	}

	/** Re-opens the panel with the last shown PDF. Returns false when nothing has been merged yet. */
	reveal(): boolean {
		if (!this.last) {
			return false;
		}
		this.show(this.last.title, this.last.data);
		return true;
	}

	dispose(): void {
		this.panel?.dispose();
	}

	private post(data: Uint8Array): void {
		if (this.ready) {
			this.panel?.webview.postMessage({ type: 'load', data });
		} else {
			this.pending = data;
		}
	}

	private createPanel(): vscode.WebviewPanel {
		const dist = vscode.Uri.joinPath(this.extensionUri, 'dist');
		const panel = vscode.window.createWebviewPanel(
			'mpdf.preview',
			'mPDF Preview',
			{ viewColumn: vscode.ViewColumn.Beside, preserveFocus: true },
			{ enableScripts: true, retainContextWhenHidden: true, localResourceRoots: [dist] },
		);
		panel.iconPath = vscode.Uri.joinPath(this.extensionUri, 'media', 'mPDF.svg');
		panel.webview.html = this.html(panel.webview, dist);
		panel.webview.onDidReceiveMessage(msg => {
			if (msg?.type === 'ready') {
				this.ready = true;
				if (this.pending) {
					panel.webview.postMessage({ type: 'load', data: this.pending });
					this.pending = undefined;
				}
			}
		});
		panel.onDidDispose(() => {
			this.panel = undefined;
			this.ready = false;
			this.pending = undefined;
		});
		return panel;
	}

	private html(webview: vscode.Webview, dist: vscode.Uri): string {
		const uri = (...segments: string[]) => webview.asWebviewUri(vscode.Uri.joinPath(dist, ...segments)).toString();
		const dir = (...segments: string[]) => uri(...segments) + '/';
		const nonce = crypto.randomBytes(16).toString('base64');
		const config = {
			workerUrl: uri('pdfjs', 'pdf.worker.min.mjs'),
			cMapUrl: dir('pdfjs', 'cmaps'),
			standardFontDataUrl: dir('pdfjs', 'standard_fonts'),
			wasmUrl: dir('pdfjs', 'wasm'),
		};
		const csp = [
			`default-src 'none'`,
			`img-src ${webview.cspSource} blob: data:`,
			`style-src 'nonce-${nonce}'`,
			`font-src ${webview.cspSource} blob: data:`,
			`script-src 'nonce-${nonce}' 'wasm-unsafe-eval'`,
			`worker-src blob:`,
			`connect-src ${webview.cspSource}`,
		].join('; ');

		return /* html */ `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="${csp}">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style nonce="${nonce}">
	html, body { height: 100%; margin: 0; padding: 0; overflow: hidden; }
	body { display: flex; flex-direction: column; background: var(--vscode-editor-background); color: var(--vscode-foreground); font-family: var(--vscode-font-family); font-size: var(--vscode-font-size); }
	#toolbar { display: flex; align-items: center; gap: 4px; padding: 4px 8px; border-bottom: 1px solid var(--vscode-panel-border, transparent); flex: none; }
	#toolbar button { background: var(--vscode-button-secondaryBackground); color: var(--vscode-button-secondaryForeground); border: none; padding: 2px 8px; min-width: 28px; cursor: pointer; border-radius: 2px; }
	#toolbar button:hover { background: var(--vscode-button-secondaryHoverBackground); }
	#zoom { min-width: 4em; text-align: center; }
	#page { margin-left: auto; opacity: 0.8; }
	#viewer { flex: 1; overflow: auto; padding: 12px 0; }
	.page { margin: 0 auto 12px; background: white; box-shadow: 0 1px 4px rgba(0, 0, 0, 0.4); }
	.page canvas { display: block; width: 100%; height: 100%; }
	#message { padding: 16px; }
</style>
<title>mPDF Preview</title>
</head>
<body>
<div id="toolbar">
	<button id="zoom-out" title="Zoom out">−</button>
	<span id="zoom"></span>
	<button id="zoom-in" title="Zoom in">+</button>
	<button id="zoom-fit" title="Fit to width">Fit</button>
	<span id="page"></span>
</div>
<div id="message">Loading…</div>
<div id="viewer"></div>
<script type="application/json" id="config">${JSON.stringify(config)}</script>
<script nonce="${nonce}" src="${uri('webview', 'preview.js')}"></script>
</body>
</html>`;
	}
}

// Runs inside the mPDF Preview webview. Renders a PDF received via postMessage with pdf.js.
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import type { PDFDocumentLoadingTask, PDFPageProxy, RenderTask } from 'pdfjs-dist';

interface Config {
	workerUrl: string;
	cMapUrl: string;
	standardFontDataUrl: string;
	wasmUrl: string;
}

type Zoom = 'fit' | number;

interface State {
	zoom: Zoom;
}

declare function acquireVsCodeApi(): {
	postMessage(message: unknown): void;
	getState(): State | undefined;
	setState(state: State): void;
};

const vscode = acquireVsCodeApi();
const config: Config = JSON.parse(document.getElementById('config')!.textContent!);

const viewer = document.getElementById('viewer')!;
const pageLabel = document.getElementById('page')!;
const zoomLabel = document.getElementById('zoom')!;
const message = document.getElementById('message')!;

const minZoom = 0.25;
const maxZoom = 5;
const pageGap = 12;

interface PageView {
	page: PDFPageProxy;
	container: HTMLDivElement;
	canvas?: HTMLCanvasElement;
	task?: RenderTask;
	scale: number;
}

let loadingTask: PDFDocumentLoadingTask | undefined;
let pages: PageView[] = [];
let zoom: Zoom = vscode.getState()?.zoom ?? 'fit';

// Webview resources are served from another origin, so the worker is started from a blob URL.
const workerReady = (async () => {
	const source = await (await fetch(config.workerUrl)).text();
	const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
	pdfjs.GlobalWorkerOptions.workerPort = new Worker(url, { type: 'module' });
})();

function fitScale(page: PDFPageProxy): number {
	const width = page.getViewport({ scale: 1 }).width;
	return Math.max(minZoom, (viewer.clientWidth - 2 * pageGap) / width);
}

function scaleFor(page: PDFPageProxy): number {
	return zoom === 'fit' ? fitScale(page) : zoom;
}

/** Index of the page at the top of the view and how far (0..1) into it the view is scrolled. */
function anchor(): { index: number; offset: number } {
	const top = viewer.scrollTop;
	for (let i = 0; i < pages.length; i++) {
		const { offsetTop, offsetHeight } = pages[i].container;
		if (offsetTop + offsetHeight > top) {
			return { index: i, offset: Math.max(0, (top - offsetTop) / offsetHeight) };
		}
	}
	return { index: 0, offset: 0 };
}

function restore(position: { index: number; offset: number }): void {
	const view = pages[Math.min(position.index, pages.length - 1)];
	if (view) {
		viewer.scrollTop = view.container.offsetTop + position.offset * view.container.offsetHeight;
	}
}

function clearCanvas(view: PageView): void {
	view.task?.cancel();
	view.task = undefined;
	view.canvas?.remove();
	view.canvas = undefined;
}

/** Sizes every page placeholder for the current zoom; pages are drawn lazily when visible. */
function layout(): void {
	for (const view of pages) {
		view.scale = scaleFor(view.page);
		const viewport = view.page.getViewport({ scale: view.scale });
		view.container.style.width = `${Math.floor(viewport.width)}px`;
		view.container.style.height = `${Math.floor(viewport.height)}px`;
		clearCanvas(view);
	}
	const shown = zoom === 'fit' ? pages[0]?.scale : zoom;
	zoomLabel.textContent = shown ? `${Math.round(shown * 100)}%` : '';
	renderVisible();
}

function render(view: PageView): void {
	if (view.canvas) {
		return;
	}
	const ratio = window.devicePixelRatio || 1;
	const viewport = view.page.getViewport({ scale: view.scale });
	const canvas = document.createElement('canvas');
	canvas.width = Math.floor(viewport.width * ratio);
	canvas.height = Math.floor(viewport.height * ratio);
	view.canvas = canvas;
	view.container.appendChild(canvas);
	view.task = view.page.render({
		canvas,
		viewport,
		transform: ratio === 1 ? undefined : [ratio, 0, 0, ratio, 0, 0],
	});
	view.task.promise.catch(error => {
		if (error?.name !== 'RenderingCancelledException') {
			console.error(error);
		}
	});
}

function renderVisible(): void {
	const top = viewer.scrollTop - viewer.clientHeight;
	const bottom = viewer.scrollTop + 2 * viewer.clientHeight;
	for (const view of pages) {
		const { offsetTop, offsetHeight } = view.container;
		if (offsetTop + offsetHeight >= top && offsetTop <= bottom) {
			render(view);
		} else if (view.canvas) {
			clearCanvas(view); // free memory of pages far away
		}
	}
	if (pages.length > 0) {
		pageLabel.textContent = `${anchor().index + 1} / ${pages.length}`;
	}
}

async function load(data: Uint8Array): Promise<void> {
	await workerReady;
	const position = anchor();
	const task = pdfjs.getDocument({
		data,
		cMapUrl: config.cMapUrl,
		cMapPacked: true,
		standardFontDataUrl: config.standardFontDataUrl,
		wasmUrl: config.wasmUrl,
	});
	const next = await task.promise;
	const nextPages = await Promise.all(
		Array.from({ length: next.numPages }, (_, i) => next.getPage(i + 1)),
	);

	// Swap only after the new document is ready so the view does not flash.
	pages.forEach(clearCanvas);
	await loadingTask?.destroy();
	loadingTask = task;
	viewer.replaceChildren();
	pages = nextPages.map(page => {
		const container = document.createElement('div');
		container.className = 'page';
		viewer.appendChild(container);
		return { page, container, scale: 1 };
	});
	message.hidden = pages.length > 0;
	message.textContent = pages.length > 0 ? '' : 'No pages.';
	layout();
	restore(position);
	renderVisible();
}

function setZoom(next: Zoom): void {
	const position = anchor();
	zoom = typeof next === 'number' ? Math.min(maxZoom, Math.max(minZoom, next)) : next;
	vscode.setState({ zoom });
	layout();
	restore(position);
	renderVisible();
}

function currentScale(): number {
	return zoom === 'fit' ? (pages[0]?.scale ?? 1) : zoom;
}

document.getElementById('zoom-out')!.addEventListener('click', () => setZoom(Math.round(currentScale() * 10 - 1) / 10));
document.getElementById('zoom-in')!.addEventListener('click', () => setZoom(Math.round(currentScale() * 10 + 1) / 10));
document.getElementById('zoom-fit')!.addEventListener('click', () => setZoom('fit'));

let scrollFrame = 0;
viewer.addEventListener('scroll', () => {
	cancelAnimationFrame(scrollFrame);
	scrollFrame = requestAnimationFrame(renderVisible);
});

let resizeTimer: ReturnType<typeof setTimeout> | undefined;
window.addEventListener('resize', () => {
	clearTimeout(resizeTimer);
	resizeTimer = setTimeout(() => {
		if (zoom === 'fit') {
			setZoom('fit');
		} else {
			renderVisible();
		}
	}, 100);
});

window.addEventListener('message', (event: MessageEvent) => {
	const msg = event.data;
	if (msg?.type === 'load') {
		load(msg.data).catch(error => {
			message.hidden = false;
			message.textContent = `Failed to display PDF: ${error?.message ?? error}`;
		});
	}
});

vscode.postMessage({ type: 'ready' });

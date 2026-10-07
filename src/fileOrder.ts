import * as vscode from 'vscode';
import * as fs from 'fs/promises';
import * as path from 'path';
import * as exceljs from 'exceljs';
import { PDFDocument } from 'pdf-lib';

import { Entry } from './fileExplorer';
import { PdfConverter } from './saveAsPdf';
import { getWorkspaceFolder, intermediatePdfPath, mergedPdfPath, settingFile } from './paths';
import { PdfPreview } from './preview';

export class WorkSheet {
	visible: boolean;
	printable: boolean;

	constructor(public name: string, state: string) {
		// hidden or veryHidden sheets are not printable
		this.visible = state === 'visible';
		this.printable = state === 'visible';
	}

	static fromJSON(json: Partial<WorkSheet>): WorkSheet {
		const sheet = new WorkSheet(String(json.name), 'visible');
		sheet.visible = json.visible ?? true;
		sheet.printable = json.printable ?? true;
		return sheet;
	}
}

export class Document {
	worksheets: WorkSheet[] = [];

	constructor(public name: string) { }

	static fromJSON(json: { name?: unknown; worksheets?: Partial<WorkSheet>[] }): Document {
		const doc = new Document(String(json.name));
		doc.worksheets = (json.worksheets ?? []).map(WorkSheet.fromJSON);
		return doc;
	}
}

export type Node = Document | WorkSheet;

/** Restores the contents of `.mPDF.json`. */
export function parseSetting(text: string): Document[] {
	const json: unknown = JSON.parse(text);
	return Array.isArray(json) ? json.map(Document.fromJSON) : [];
}

/** Moves `source` in front of `target`, or to the end when there is no target. */
export function moveDocument(documents: Document[], source: Document, target: Document | undefined): void {
	if (source === target) {
		return;
	}
	const from = documents.indexOf(source);
	if (from < 0) {
		return;
	}
	documents.splice(from, 1);
	const to = target ? documents.indexOf(target) : -1;
	documents.splice(to < 0 ? documents.length : to, 0, source);
}

const mimeType = 'application/vnd.code.tree.fileorder';

export class FileOrderProvider implements vscode.TreeDataProvider<Node>, vscode.TreeDragAndDropController<Node> {
	private documents: Document[] = [];
	private readonly converter: PdfConverter;
	private readonly preview: PdfPreview;
	private merging?: Promise<void>;
	private mergeRequested = false;

	readonly dropMimeTypes = [mimeType];
	readonly dragMimeTypes = ['text/uri-list'];

	private readonly _onDidChangeTreeData = new vscode.EventEmitter<Node | undefined>();
	readonly onDidChangeTreeData = this._onDidChangeTreeData.event;

	constructor(context: vscode.ExtensionContext) {
		this.converter = new PdfConverter(context.extensionPath);
		this.preview = new PdfPreview(context.extensionUri);
		context.subscriptions.push(
			this.preview,
			vscode.commands.registerCommand('fileOrder.preview', () => this.showPreview()),
			vscode.commands.registerCommand('fileOrder.add', (entry: Entry) => this.add(entry)),
			vscode.commands.registerCommand('fileOrder.update', (uri: vscode.Uri) => this.update(uri)),
			vscode.commands.registerCommand('fileOrder.delete', (item: Document) => this.delete(item)),
			vscode.commands.registerCommand('fileOrder.select', (sheet: WorkSheet) => this.select(sheet)),
			vscode.commands.registerCommand('fileOrder.merge', () => this.merge()),
			vscode.commands.registerCommand('fileOrder.publish', (item: Document) => this.publish(item)),
		);
		this.loadSetting();
	}

	getChildren(element?: Node): Node[] {
		if (element instanceof Document) {
			return element.worksheets;
		}
		return element ? [] : this.documents;
	}

	getTreeItem(element: Node): vscode.TreeItem {
		if (element instanceof Document) {
			const treeItem = new vscode.TreeItem(element.name, element.worksheets.length > 0 ? vscode.TreeItemCollapsibleState.Expanded : vscode.TreeItemCollapsibleState.None);
			treeItem.contextValue = 'file';
			return treeItem;
		}

		const treeItem = new vscode.TreeItem(element.name);
		treeItem.contextValue = 'worksheet';
		if (element.printable) {
			treeItem.iconPath = new vscode.ThemeIcon(element.visible ? 'check' : 'clear');
			treeItem.command = { command: 'fileOrder.select', title: 'select', arguments: [element] };
		} else {
			treeItem.iconPath = new vscode.ThemeIcon('circle-slash');
		}
		return treeItem;
	}

	handleDrag(source: readonly Node[], dataTransfer: vscode.DataTransfer): void {
		if (source[0] instanceof Document) {
			dataTransfer.set(mimeType, new vscode.DataTransferItem(source[0]));
		}
	}

	handleDrop(target: Node | undefined, dataTransfer: vscode.DataTransfer): void {
		const source: unknown = dataTransfer.get(mimeType)?.value;
		if (!(source instanceof Document)) {
			return;
		}
		const targetDocument = target instanceof WorkSheet ? this.findWorkbook(target) : target;
		moveDocument(this.documents, source, targetDocument);
		this.changed();
		this.merge();
	}

	async publish(item: Document): Promise<void> {
		const folder = getWorkspaceFolder();
		if (!folder) {
			return;
		}
		if (PdfConverter.isSupported(item.name)) {
			const sheets = item.worksheets.filter(s => s.visible).map(s => s.name);
			try {
				await vscode.window.withProgress(
					{ location: vscode.ProgressLocation.Window, title: `mPDF: converting ${item.name}` },
					() => this.converter.convert(folder.uri.fsPath, { name: item.name, sheets }),
				);
			} catch (error) {
				vscode.window.showErrorMessage(`mPDF: failed to convert ${item.name}: ${errorMessage(error)}`);
				return;
			}
		}
		await this.merge();
	}

	async add(entry: Entry): Promise<void> {
		const name = vscode.workspace.asRelativePath(entry.uri, false);
		const existing = this.documents.find(d => d.name === name);
		if (existing) {
			await this.publish(existing);
			return;
		}

		const node = new Document(name);
		if (path.extname(name).toLowerCase() === '.xlsx') {
			try {
				const wb = await new exceljs.Workbook().xlsx.readFile(entry.uri.fsPath);
				node.worksheets = wb.worksheets.map(sheet => new WorkSheet(sheet.name, sheet.state));
			} catch (error) {
				vscode.window.showErrorMessage(`mPDF: failed to read ${name}: ${errorMessage(error)}`);
				return;
			}
		}
		this.documents.push(node);
		this.changed();
		await this.publish(node);
	}

	async update(uri: vscode.Uri): Promise<void> {
		const name = vscode.workspace.asRelativePath(uri, false);
		const document = this.documents.find(d => d.name === name);
		if (document) {
			await this.publish(document);
		}
	}

	async delete(item: Document): Promise<void> {
		const index = this.documents.indexOf(item);
		if (index > -1) {
			this.documents.splice(index, 1);
		}
		this.changed();
		await this.merge();
	}

	private findWorkbook(worksheet: WorkSheet): Document | undefined {
		return this.documents.find(d => d.worksheets.includes(worksheet));
	}

	async select(sheet: WorkSheet): Promise<void> {
		const workbook = this.findWorkbook(sheet);
		if (!workbook) {
			return;
		}
		// Keep at least one sheet selected.
		if (sheet.visible && workbook.worksheets.filter(s => s.visible).length === 1) {
			return;
		}
		sheet.visible = !sheet.visible;
		this.changed();
		await this.publish(workbook);
	}

	/** Refreshes the tree and persists the order. */
	private changed(): void {
		this._onDidChangeTreeData.fire(undefined);
		this.saveSetting().catch(error => {
			vscode.window.showErrorMessage(`mPDF: failed to save ${settingFile}: ${errorMessage(error)}`);
		});
	}

	private async saveSetting(): Promise<void> {
		const folder = getWorkspaceFolder();
		if (folder) {
			await fs.writeFile(path.join(folder.uri.fsPath, settingFile), JSON.stringify(this.documents, null, 2));
		}
	}

	private async loadSetting(): Promise<void> {
		const folder = getWorkspaceFolder();
		if (!folder) {
			return;
		}
		let text: string;
		try {
			text = await fs.readFile(path.join(folder.uri.fsPath, settingFile), 'utf8');
		} catch {
			return; // not created yet
		}
		try {
			this.documents = parseSetting(text);
			this._onDidChangeTreeData.fire(undefined);
		} catch (error) {
			vscode.window.showWarningMessage(`mPDF: ignored broken ${settingFile}: ${errorMessage(error)}`);
		}
	}

	/** Merges the intermediate PDFs. Requests during a running merge are coalesced into one more run. */
	merge(): Promise<void> {
		if (this.merging) {
			this.mergeRequested = true;
			return this.merging;
		}
		this.merging = (async () => {
			do {
				this.mergeRequested = false;
				await this.mergeOnce();
			} while (this.mergeRequested);
		})().finally(() => {
			this.merging = undefined;
		});
		return this.merging;
	}

	private async mergeOnce(): Promise<void> {
		const folder = getWorkspaceFolder();
		if (!folder || this.documents.length === 0) {
			return;
		}
		const mergedPath = mergedPdfPath(folder);
		const missing: string[] = [];
		let bytes: Uint8Array;
		try {
			const merged = await PDFDocument.create();
			for (const doc of this.documents) {
				let source: Buffer;
				try {
					source = await fs.readFile(intermediatePdfPath(folder.uri.fsPath, doc.name));
				} catch {
					missing.push(doc.name);
					continue;
				}
				const pdf = await PDFDocument.load(source);
				const pages = await merged.copyPages(pdf, pdf.getPageIndices());
				pages.forEach(page => merged.addPage(page));
			}
			bytes = await merged.save();
			await fs.writeFile(mergedPath, bytes);
		} catch (error) {
			vscode.window.showErrorMessage(`mPDF: failed to merge into ${path.basename(mergedPath)}: ${errorMessage(error)}`);
			return;
		}
		if (missing.length > 0) {
			vscode.window.showWarningMessage(`mPDF: skipped documents without PDF: ${missing.join(', ')}`);
		}
		this.preview.show(path.basename(mergedPath), bytes);
	}

	/** Opens the preview; merges first if nothing has been shown in this session. */
	private async showPreview(): Promise<void> {
		if (!this.preview.reveal()) {
			await this.merge();
		}
	}
}

function errorMessage(error: unknown): string {
	return error instanceof Error ? error.message : String(error);
}

export class FileOrder {
	constructor(context: vscode.ExtensionContext) {
		const treeDataProvider = new FileOrderProvider(context);
		context.subscriptions.push(vscode.window.createTreeView('fileOrder', {
			treeDataProvider,
			showCollapseAll: true,
			canSelectMany: false,
			dragAndDropController: treeDataProvider,
		}));
	}
}

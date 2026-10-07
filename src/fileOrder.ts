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
	/** Whether the document is included in the merged PDF. */
	enabled = true;

	constructor(public name: string) { }

	static fromJSON(json: { name?: unknown; enabled?: boolean; worksheets?: Partial<WorkSheet>[] }): Document {
		const doc = new Document(String(json.name));
		doc.enabled = json.enabled ?? true;
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

/**
 * Moves `source` in front of `target`, or to the end when there is no target.
 * A `source` that is not in the list yet is inserted there.
 */
export function moveDocument(documents: Document[], source: Document, target: Document | undefined): void {
	if (source === target) {
		return;
	}
	const from = documents.indexOf(source);
	if (from >= 0) {
		documents.splice(from, 1);
	}
	const to = target ? documents.indexOf(target) : -1;
	documents.splice(to < 0 ? documents.length : to, 0, source);
}

/** Parses a `text/uri-list` (RFC 2483): one URI per line, `#` lines are comments. */
export function parseUriList(text: string): vscode.Uri[] {
	return text
		.split(/\r?\n/)
		.map(line => line.trim())
		.filter(line => line.length > 0 && !line.startsWith('#'))
		.map(line => vscode.Uri.parse(line));
}

/**
 * Shows or hides `sheet` in the PDF of `workbook`.
 * Returns false (and changes nothing) when that would leave no sheet to print.
 */
export function setSheetVisible(workbook: Document, sheet: WorkSheet, visible: boolean): boolean {
	if (!sheet.printable) {
		return false;
	}
	if (!visible && workbook.worksheets.every(s => s === sheet || !s.visible)) {
		return false;
	}
	sheet.visible = visible;
	return true;
}

function checkbox(checked: boolean, tooltip: string): vscode.TreeItem['checkboxState'] {
	return {
		state: checked ? vscode.TreeItemCheckboxState.Checked : vscode.TreeItemCheckboxState.Unchecked,
		tooltip,
	};
}

const mimeType = 'application/vnd.code.tree.fileorder';

export class FileOrderProvider implements vscode.TreeDataProvider<Node>, vscode.TreeDragAndDropController<Node> {
	private documents: Document[] = [];
	private readonly converter: PdfConverter;
	private readonly preview: PdfPreview;
	private merging?: Promise<void>;
	private mergeRequested = false;

	// text/uri-list: files dragged from the built-in Explorer or the mPDF File Explorer.
	readonly dropMimeTypes = [mimeType, 'text/uri-list'];
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
			treeItem.checkboxState = checkbox(element.enabled, 'Include in the merged PDF');
			const printable = element.worksheets.filter(s => s.printable);
			if (!element.enabled) {
				treeItem.description = 'excluded';
			} else if (printable.length > 0) {
				treeItem.description = `${printable.filter(s => s.visible).length}/${printable.length} sheets`;
			}
			return treeItem;
		}

		const treeItem = new vscode.TreeItem(element.name);
		treeItem.contextValue = 'worksheet';
		if (element.printable) {
			treeItem.checkboxState = checkbox(element.visible, 'Print this sheet');
			// Clicking the label toggles as well as clicking the checkbox.
			treeItem.command = { command: 'fileOrder.select', title: 'select', arguments: [element] };
		} else {
			treeItem.iconPath = new vscode.ThemeIcon('circle-slash');
			treeItem.description = 'hidden';
			treeItem.tooltip = 'Hidden sheets are not printed';
		}
		return treeItem;
	}

	/** Applies checkbox clicks: documents are included/excluded, sheets are shown/hidden. */
	async onDidChangeCheckboxState(event: vscode.TreeCheckboxChangeEvent<Node>): Promise<void> {
		const republish = new Set<Document>();
		let remerge = false;
		for (const [node, state] of event.items) {
			const checked = state === vscode.TreeItemCheckboxState.Checked;
			if (node instanceof Document) {
				node.enabled = checked;
				if (checked) {
					republish.add(node); // the source may have changed while excluded
				} else {
					remerge = true;
				}
			} else {
				const workbook = this.findWorkbook(node);
				if (workbook && setSheetVisible(workbook, node, checked)) {
					republish.add(workbook);
				}
			}
		}
		// Also re-renders checkboxes whose change was refused.
		this.changed();
		for (const doc of republish) {
			if (doc.enabled) {
				await this.publish(doc);
			} else {
				remerge = true;
			}
		}
		if (remerge) {
			await this.merge();
		}
	}

	handleDrag(source: readonly Node[], dataTransfer: vscode.DataTransfer): void {
		if (source[0] instanceof Document) {
			dataTransfer.set(mimeType, new vscode.DataTransferItem(source[0]));
		}
	}

	async handleDrop(target: Node | undefined, dataTransfer: vscode.DataTransfer): Promise<void> {
		const targetDocument = target instanceof WorkSheet ? this.findWorkbook(target) : target;

		const source: unknown = dataTransfer.get(mimeType)?.value;
		if (source instanceof Document) {
			moveDocument(this.documents, source, targetDocument);
			this.changed();
			await this.merge();
			return;
		}

		const uriList = await dataTransfer.get('text/uri-list')?.asString();
		if (uriList) {
			await this.addFiles(parseUriList(uriList), targetDocument);
		}
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
		await this.addFiles([entry.uri], undefined);
	}

	/**
	 * Adds files in front of `target` (or at the end) and converts them.
	 * Files already in the list are re-enabled, and moved when a target is given.
	 */
	private async addFiles(uris: vscode.Uri[], target: Document | undefined): Promise<void> {
		const folder = getWorkspaceFolder();
		if (!folder) {
			return;
		}
		const added: Document[] = [];
		const skipped: string[] = [];
		for (const uri of uris) {
			if (uri.scheme !== 'file' || vscode.workspace.getWorkspaceFolder(uri)?.uri.toString() !== folder.uri.toString()) {
				skipped.push(`${path.basename(uri.fsPath)} (outside the workspace folder)`);
				continue;
			}
			const name = vscode.workspace.asRelativePath(uri, false);
			if (!PdfConverter.isSupported(name)) {
				skipped.push(`${name} (only .xlsx and .docx are supported)`);
				continue;
			}
			let doc = this.documents.find(d => d.name === name);
			if (doc) {
				doc.enabled = true;
				if (target) {
					moveDocument(this.documents, doc, target);
				}
			} else {
				try {
					doc = await this.readDocument(uri, name);
				} catch (error) {
					vscode.window.showErrorMessage(`mPDF: failed to read ${name}: ${errorMessage(error)}`);
					continue;
				}
				moveDocument(this.documents, doc, target);
			}
			added.push(doc);
		}
		if (skipped.length > 0) {
			vscode.window.showWarningMessage(`mPDF: skipped ${skipped.join(', ')}`);
		}
		if (added.length === 0) {
			return;
		}
		this.changed();
		for (const doc of added) {
			await this.publish(doc);
		}
	}

	private async readDocument(uri: vscode.Uri, name: string): Promise<Document> {
		const doc = new Document(name);
		if (path.extname(name).toLowerCase() === '.xlsx') {
			const wb = await new exceljs.Workbook().xlsx.readFile(uri.fsPath);
			doc.worksheets = wb.worksheets.map(sheet => new WorkSheet(sheet.name, sheet.state));
		}
		return doc;
	}

	async update(uri: vscode.Uri): Promise<void> {
		const name = vscode.workspace.asRelativePath(uri, false);
		const document = this.documents.find(d => d.name === name);
		if (document?.enabled) {
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
		if (!workbook || !setSheetVisible(workbook, sheet, !sheet.visible)) {
			return;
		}
		this.changed();
		if (workbook.enabled) {
			await this.publish(workbook);
		}
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
		if (!folder) {
			return;
		}
		const mergedPath = mergedPdfPath(folder);
		const documents = this.documents.filter(d => d.enabled);
		if (documents.length === 0) {
			// Leave the last merged file alone, but do not keep showing stale pages.
			this.preview.showMessage(path.basename(mergedPath), 'No documents are checked.');
			return;
		}
		const missing: string[] = [];
		let bytes: Uint8Array;
		try {
			const merged = await PDFDocument.create();
			for (const doc of documents) {
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
		const tree = vscode.window.createTreeView('fileOrder', {
			treeDataProvider,
			showCollapseAll: true,
			canSelectMany: false,
			dragAndDropController: treeDataProvider,
			// A document's checkbox (include in PDF) and its sheets' checkboxes (print sheet) are independent.
			manageCheckboxStateManually: true,
		});
		context.subscriptions.push(
			tree,
			tree.onDidChangeCheckboxState(event => treeDataProvider.onDidChangeCheckboxState(event)),
		);
	}
}

import * as vscode from 'vscode';
import * as fs from 'fs/promises';
import * as path from 'path';

import { getWorkspaceFolder, intermediateDir } from './paths';

export interface Entry {
	uri: vscode.Uri;
	type: vscode.FileType;
}

const debounceMs = 1000;

/** Office lock files (`~$book.xlsx`) and temp files written while saving. */
function isTemporary(uri: vscode.Uri): boolean {
	const name = path.basename(uri.fsPath);
	return name.startsWith('~$') || name.toLowerCase().endsWith('.tmp');
}

function isIntermediate(folder: vscode.WorkspaceFolder, uri: vscode.Uri): boolean {
	const relative = path.relative(folder.uri.fsPath, uri.fsPath);
	return relative === intermediateDir || relative.startsWith(intermediateDir + path.sep);
}

export class FileTreeProvider implements vscode.TreeDataProvider<Entry>, vscode.TreeDragAndDropController<Entry> {

	// Files can be dragged into the file Order view, like from the built-in Explorer.
	readonly dragMimeTypes = ['text/uri-list'];
	readonly dropMimeTypes: string[] = [];

	handleDrag(source: readonly Entry[], dataTransfer: vscode.DataTransfer): void {
		const files = source.filter(entry => entry.type === vscode.FileType.File);
		if (files.length > 0) {
			dataTransfer.set('text/uri-list', new vscode.DataTransferItem(files.map(entry => entry.uri.toString()).join('\r\n')));
		}
	}

	private readonly _onDidChangeTreeData = new vscode.EventEmitter<Entry | undefined>();
	readonly onDidChangeTreeData = this._onDidChangeTreeData.event;

	refresh(): void {
		this._onDidChangeTreeData.fire(undefined);
	}

	async getChildren(element?: Entry): Promise<Entry[]> {
		const dir = element?.uri ?? getWorkspaceFolder()?.uri;
		if (!dir) {
			return [];
		}

		const children = await fs.readdir(dir.fsPath, { withFileTypes: true });
		const entries = children
			.filter(child => child.isFile() || child.isDirectory())
			.map(child => ({
				uri: vscode.Uri.file(path.join(dir.fsPath, child.name)),
				type: child.isDirectory() ? vscode.FileType.Directory : vscode.FileType.File,
			}));
		entries.sort((a, b) => {
			if (a.type === b.type) {
				return a.uri.fsPath.localeCompare(b.uri.fsPath);
			}
			return a.type === vscode.FileType.Directory ? -1 : 1;
		});
		return entries;
	}

	getTreeItem(element: Entry): vscode.TreeItem {
		const treeItem = new vscode.TreeItem(element.uri, element.type === vscode.FileType.Directory ? vscode.TreeItemCollapsibleState.Collapsed : vscode.TreeItemCollapsibleState.None);
		if (element.type === vscode.FileType.File) {
			treeItem.contextValue = 'file';
		}
		return treeItem;
	}
}

export class FileExplorer {
	private readonly pending = new Map<string, NodeJS.Timeout>();

	constructor(context: vscode.ExtensionContext) {
		const treeDataProvider = new FileTreeProvider();
		context.subscriptions.push(vscode.window.createTreeView('fileExplorer', {
			treeDataProvider,
			dragAndDropController: treeDataProvider,
			canSelectMany: true,
		}));

		const folder = getWorkspaceFolder();
		if (!folder) {
			return;
		}

		const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(folder, '**/*'));
		const onChange = (uri: vscode.Uri) => {
			if (isTemporary(uri) || isIntermediate(folder, uri)) {
				return;
			}
			this.notifyUpdate(uri);
		};
		const onCreateOrDelete = (uri: vscode.Uri) => {
			if (!isIntermediate(folder, uri)) {
				treeDataProvider.refresh();
			}
			onChange(uri);
		};
		context.subscriptions.push(
			watcher,
			watcher.onDidChange(onChange),
			watcher.onDidCreate(onCreateOrDelete),
			watcher.onDidDelete(uri => {
				if (!isIntermediate(folder, uri)) {
					treeDataProvider.refresh();
				}
			}),
			{ dispose: () => this.pending.forEach(timer => clearTimeout(timer)) },
		);
	}

	/** Saving an Office file fires several events in a row; coalesce them. */
	private notifyUpdate(uri: vscode.Uri): void {
		const key = uri.fsPath;
		clearTimeout(this.pending.get(key));
		this.pending.set(key, setTimeout(() => {
			this.pending.delete(key);
			vscode.commands.executeCommand('fileOrder.update', uri).then(undefined, err => {
				console.error('fileOrder.update failed', err);
			});
		}, debounceMs));
	}
}

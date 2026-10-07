import * as vscode from 'vscode';

import { FileExplorer } from './fileExplorer';
import { FileOrder } from './fileOrder';

export function activate(context: vscode.ExtensionContext) {
	new FileExplorer(context);
	new FileOrder(context);
}

export function deactivate() { }

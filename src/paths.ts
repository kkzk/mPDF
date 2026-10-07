import * as vscode from 'vscode';
import * as path from 'path';

/** Directory (relative to the workspace root) that holds intermediate PDFs. */
export const intermediateDir = '.mPDF';

/** File (relative to the workspace root) that persists the document order. */
export const settingFile = '.mPDF.json';

export function getWorkspaceFolder(): vscode.WorkspaceFolder | undefined {
	return vscode.workspace.workspaceFolders?.find(folder => folder.uri.scheme === 'file');
}

/** `<workspace>/.mPDF/<dir>/<name>.pdf` for a document given by its workspace-relative path. */
export function intermediatePdfPath(workspaceDir: string, relativePath: string): string {
	const { dir, name } = path.parse(relativePath);
	return path.join(workspaceDir, intermediateDir, dir, `${name}.pdf`);
}

/** `<workspace>/<workspace name>.pdf` */
export function mergedPdfPath(folder: vscode.WorkspaceFolder): string {
	return path.join(folder.uri.fsPath, `${folder.name}.pdf`);
}

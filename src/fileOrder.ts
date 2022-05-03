import * as vscode from 'vscode';
import * as path from 'path';
import * as exceljs from 'exceljs';
import * as fs from "fs";

import replaceExt = require('replace-ext');
import PDFMerger = require('pdf-merger-js');

import { Entry } from "./fileExplorer";

type NodeType = "worksheet" | "workbook" | "misc";

interface Node {
    name: string;
    worksheets?: WorkSheet[];
}

class WorkSheet implements Node {
    name: string;
    visible: boolean = true;
    printable: boolean = true;

    constructor(name: string, state: string){
        this.name = name;
        this.visible = true ? state === "visible": false;
        this.printable = true ? state === "visible": false;  // hidden or veryHidden is not printable
    }
}

class Document implements Node {
    name: string;
    worksheets?: WorkSheet[];

    constructor (name: string) {
        this.name = name;
        this.worksheets = [];
    }
}

export class FileOrderProvidor implements vscode.TreeDataProvider<Node>, vscode.TreeDragAndDropController<Node> {
    private documents: Document[] = [];
    private terminal?: vscode.Terminal;

    constructor(context: vscode.ExtensionContext) {
        this.loadSetting();

        context.subscriptions.push(vscode.commands.registerCommand("fileOrder.add", (name) => this.add(name)));
        context.subscriptions.push(vscode.commands.registerCommand("fileOrder.delete", (name) => this.delete(name)));
        context.subscriptions.push(vscode.commands.registerCommand("fileOrder.select", (element) => this.select(element)));
        context.subscriptions.push(vscode.commands.registerCommand("fileOrder.merge", () => this.merge()));

        context.subscriptions.push(vscode.commands.registerCommand('fileOrder.publish', (item: Node) => {
            if (!this.terminal) {
                this.terminal = vscode.window.createTerminal(`mPDF`, "powershell.exe");
            }
            const scriptPath = path.join(context.extensionPath, "script", "saveAsPdf.ps1");
            if (vscode.workspace.workspaceFolders !== undefined){
                const workspaceDir = path.dirname(vscode.workspace.workspaceFolders[0].uri.fsPath);
                const parentName = path.basename(workspaceDir);
                const pdfFilename = path.join(workspaceDir, parentName);
                this.terminal.show(true);
                this.terminal.sendText(`. "${scriptPath}"; Save-pdf .mpdf.json "${item.name}" `);
            }
        }));
    }

	dropMimeTypes = ['application/vnd.code.tree.fileOrderProvidor'];
	dragMimeTypes = ['text/uri-list'];

    private _onDidChangeTreeData: vscode.EventEmitter<Node | undefined> = new vscode.EventEmitter<Node | undefined>();

    readonly onDidChangeTreeData: vscode.Event<Node | undefined> = this._onDidChangeTreeData.event;

    getChildren(element?: Node): vscode.ProviderResult<Node[]> {
        if(element?.worksheets) {
            return element.worksheets;
        }
        return this.documents;
    }

    getTreeItem(element: Node): vscode.TreeItem | Thenable<vscode.TreeItem> {
        if (element.worksheets) {
            // Top node allways has worksheets property even thou that is not excel.
            var treeItem = new vscode.TreeItem(element.name, vscode.TreeItemCollapsibleState.Expanded);
            treeItem.contextValue = "file";
        }
        else {
            var treeItem = new vscode.TreeItem(element.name);
            treeItem.contextValue = "worksheet";
            if ((element as WorkSheet).printable) {
                treeItem.iconPath = (element as WorkSheet).visible ? new vscode.ThemeIcon("check"): new vscode.ThemeIcon("clear");
                treeItem.command = { command: 'fileOrder.select', title: "select", arguments: [element] };
            }
            else {
                treeItem.iconPath = new vscode.ThemeIcon("circle-slash");
            }

        }
        return treeItem;
    }

    public async handleDrop(target: Node | undefined, sources: vscode.DataTransfer, token: vscode.CancellationToken): Promise<void> {
		const transferItem = sources.get('application/vnd.code.tree.fileOrderProvidor');
		if (!transferItem) {
			return;
		}
        const treeItems: Node[] = transferItem.value;
        console.log(target);
        console.log(treeItems[0].name);
        this.documents.splice(this.documents.indexOf(treeItems[0]), 1);
        this.documents.splice(this.documents.indexOf(target as Node), 0, treeItems[0]);
        this._onDidChangeTreeData.fire(undefined);
        vscode.commands.executeCommand("fileOrder.merge");
    }

    public async handleDrag(source: Node[], treeDataTransfer: vscode.DataTransfer, token: vscode.CancellationToken): Promise<void> {
        if (source[0] instanceof Document) {
            // console.log(source);
            treeDataTransfer.set('application/vnd.code.tree.fileOrderProvidor', new vscode.DataTransferItem(source));
        }
    }

    add(item: Entry): void {
        var node = new Document(vscode.workspace.asRelativePath(item.uri, false));
        if (item.uri.fsPath.endsWith("xlsx")) {
            var wb = new exceljs.Workbook();
            wb.xlsx.readFile(item.uri.fsPath).then(wb => {
                node.worksheets = wb.worksheets.map((sheet) => new WorkSheet(sheet.name, sheet.state));
                this.documents.push(node);
                this.saveSetting();
                this._onDidChangeTreeData.fire(undefined);
            });
        }
        else {
            this.documents.push(node);
            this.saveSetting();
            this._onDidChangeTreeData.fire(undefined);
        }
        vscode.commands.executeCommand("fileOrder.publish", node);
    }

    delete(item: Node) {
        const index = this.documents.indexOf(item, 0);
        if (index > -1) {
            this.documents.splice(index, 1);
        }
        this._onDidChangeTreeData.fire(undefined);
        this.saveSetting();
        vscode.commands.executeCommand("fileOrder.merge");
    }

    private findWorkbook(worksheet: WorkSheet): Document | undefined {
        for (let workbook of this.documents){
            if (workbook.worksheets?.filter(s => s === worksheet).length !== 0) {
                return workbook;
            }
        }
        return undefined;
    }

    select(element: Node){
        // # console.log(element);
        const workbook = this.findWorkbook((element as WorkSheet));
        (element as WorkSheet).visible = !(element as WorkSheet).visible;
        if (workbook?.worksheets?.filter(s => s.visible).length !== 0) {
            this._onDidChangeTreeData.fire(undefined);
            this.saveSetting();
            vscode.commands.executeCommand("fileOrder.publish", workbook);
        }
        else {
            (element as WorkSheet).visible = !(element as WorkSheet).visible;
        }
    }

    saveSetting(){
        var data = JSON.stringify(this.documents, null, 2);
		const workspaceFolder = vscode.workspace.workspaceFolders?.filter(folder => folder.uri.scheme === 'file')[0];
        if (workspaceFolder) {
            fs.writeFileSync(path.join(workspaceFolder.uri.fsPath, "./.mPDF.json"), data);
        }
    }

    loadSetting(){
		const workspaceFolder = vscode.workspace.workspaceFolders?.filter(folder => folder.uri.scheme === 'file')[0];
        if (workspaceFolder) {
            try {
                this.documents = JSON.parse((fs.readFileSync(path.join(workspaceFolder.uri.fsPath, "./.mPDF.json")).toString()));                
            } catch (error) {
                // pass                
            }
        }
    }

    private merge() {
        const merger = new PDFMerger();
		const workspaceFolder = vscode.workspace.workspaceFolders?.filter(folder => folder.uri.scheme === 'file')[0];
        if (workspaceFolder) {
            const sources = this.documents.map((wb) => {
                return replaceExt(path.join(workspaceFolder.uri.fsPath, ".mPDF", wb.name), ".pdf");
            });
            const mergedPath = path.join(workspaceFolder.uri.fsPath, workspaceFolder.name + ".pdf");
            console.log(sources);
            (async (sources: string[], mergedPath: string) => {
                sources.forEach(element => {
                    merger.add(element);
                });
                await merger.save(mergedPath);
                // vscode.window.showTextDocument(vscode.Uri.file(mergedPath));
                vscode.commands.executeCommand("vscode.open", vscode.Uri.file(mergedPath));
            })(sources, mergedPath).catch(() => {
                console.log("error");
            });
        }
    }
}

export class FileOrder {
    constructor(context: vscode.ExtensionContext) {
        const treeDataProvider = new FileOrderProvidor(context);
        const tree = vscode.window.createTreeView('fileOrder', { treeDataProvider: treeDataProvider, showCollapseAll: true, canSelectMany: false, dragAndDropController: treeDataProvider });
        context.subscriptions.push(tree);
    }
}
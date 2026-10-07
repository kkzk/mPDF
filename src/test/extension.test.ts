import * as assert from 'assert';
import * as path from 'path';

import * as vscode from 'vscode';

import { Document, WorkSheet, moveDocument, parseSetting, parseUriList, setSheetVisible } from '../fileOrder';
import { intermediatePdfPath } from '../paths';

suite('paths', () => {
	test('intermediatePdfPath replaces the extension under .mPDF', () => {
		const root = path.resolve('/ws');
		assert.strictEqual(intermediatePdfPath(root, 'a.xlsx'), path.join(root, '.mPDF', 'a.pdf'));
		assert.strictEqual(intermediatePdfPath(root, 'folder/b.docx'), path.join(root, '.mPDF', 'folder', 'b.pdf'));
		assert.strictEqual(intermediatePdfPath(root, 'ファイル名・中黒.xlsx'), path.join(root, '.mPDF', 'ファイル名・中黒.pdf'));
	});
});

suite('setting', () => {
	test('parseSetting restores class instances', () => {
		const docs = parseSetting(JSON.stringify([
			{ name: 'a.xlsx', worksheets: [{ name: 'S1', visible: false, printable: true }] },
			{ name: 'b.docx', worksheets: [] },
		]));
		assert.strictEqual(docs.length, 2);
		assert.ok(docs[0] instanceof Document);
		assert.ok(docs[0].worksheets[0] instanceof WorkSheet);
		assert.strictEqual(docs[0].worksheets[0].visible, false);
		assert.strictEqual(docs[0].worksheets[0].printable, true);
	});

	test('parseSetting round-trips JSON.stringify output', () => {
		const doc = new Document('a.xlsx');
		doc.enabled = false;
		doc.worksheets = [new WorkSheet('S1', 'visible'), new WorkSheet('S2', 'hidden')];
		const [restored] = parseSetting(JSON.stringify([doc]));
		assert.deepStrictEqual(restored, doc);
	});

	test('documents saved by 0.1.x (no "enabled") are included', () => {
		const [doc] = parseSetting(JSON.stringify([{ name: 'a.docx', worksheets: [] }]));
		assert.strictEqual(doc.enabled, true);
	});
});

suite('setSheetVisible', () => {
	function workbook(...states: string[]): Document {
		const doc = new Document('a.xlsx');
		doc.worksheets = states.map((state, i) => new WorkSheet(`S${i + 1}`, state));
		return doc;
	}

	test('hides a sheet while another stays visible', () => {
		const doc = workbook('visible', 'visible');
		assert.strictEqual(setSheetVisible(doc, doc.worksheets[0], false), true);
		assert.strictEqual(doc.worksheets[0].visible, false);
	});

	test('refuses to hide the last visible sheet', () => {
		const doc = workbook('visible', 'visible');
		setSheetVisible(doc, doc.worksheets[0], false);
		assert.strictEqual(setSheetVisible(doc, doc.worksheets[1], false), false);
		assert.strictEqual(doc.worksheets[1].visible, true);
	});

	test('hidden sheets in Excel do not count as visible', () => {
		const doc = workbook('visible', 'hidden');
		assert.strictEqual(setSheetVisible(doc, doc.worksheets[0], false), false);
		assert.strictEqual(setSheetVisible(doc, doc.worksheets[1], true), false);
		assert.strictEqual(doc.worksheets[1].visible, false);
	});
});

suite('moveDocument', () => {
	const names = (docs: Document[]) => docs.map(d => d.name);

	test('moves in front of the target', () => {
		const [a, b, c] = ['a', 'b', 'c'].map(n => new Document(n));
		const docs = [a, b, c];
		moveDocument(docs, c, a);
		assert.deepStrictEqual(names(docs), ['c', 'a', 'b']);
	});

	test('moves to the end without a target', () => {
		const [a, b, c] = ['a', 'b', 'c'].map(n => new Document(n));
		const docs = [a, b, c];
		moveDocument(docs, a, undefined);
		assert.deepStrictEqual(names(docs), ['b', 'c', 'a']);
	});

	test('inserts a new document in front of the target', () => {
		const [a, b, c] = ['a', 'b', 'c'].map(n => new Document(n));
		const docs = [a, b];
		moveDocument(docs, c, b);
		assert.deepStrictEqual(names(docs), ['a', 'c', 'b']);
	});

	test('appends a new document without a target', () => {
		const [a, b] = ['a', 'b'].map(n => new Document(n));
		const docs = [a];
		moveDocument(docs, b, undefined);
		assert.deepStrictEqual(names(docs), ['a', 'b']);
	});

	test('dropping on itself keeps the order', () => {
		const [a, b] = ['a', 'b'].map(n => new Document(n));
		const docs = [a, b];
		moveDocument(docs, b, b);
		assert.deepStrictEqual(names(docs), ['a', 'b']);
	});
});

suite('parseUriList', () => {
	test('reads one URI per line and skips comments and blank lines', () => {
		const a = vscode.Uri.file(path.resolve('/ws/a.xlsx'));
		const b = vscode.Uri.file(path.resolve('/ws/ファイル名・中黒.xlsx'));
		const uris = parseUriList(`# comment\r\n${a.toString()}\r\n\r\n${b.toString()}\n`);
		assert.deepStrictEqual(uris.map(u => u.fsPath), [a.fsPath, b.fsPath]);
	});
});

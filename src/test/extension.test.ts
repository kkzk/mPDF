import * as assert from 'assert';
import * as path from 'path';

import { Document, WorkSheet, moveDocument, parseSetting } from '../fileOrder';
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
		doc.worksheets = [new WorkSheet('S1', 'visible'), new WorkSheet('S2', 'hidden')];
		const [restored] = parseSetting(JSON.stringify([doc]));
		assert.deepStrictEqual(restored, doc);
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

	test('dropping on itself keeps the order', () => {
		const [a, b] = ['a', 'b'].map(n => new Document(n));
		const docs = [a, b];
		moveDocument(docs, b, b);
		assert.deepStrictEqual(names(docs), ['a', 'b']);
	});
});

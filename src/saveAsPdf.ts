import { execFile } from 'child_process';
import * as fs from 'fs/promises';
import * as os from 'os';
import * as path from 'path';

import { intermediatePdfPath } from './paths';

type JobKind = 'excel' | 'word';

interface Job {
	kind: JobKind;
	source: string;
	output: string;
	sheets: string[];
}

export interface ConvertTarget {
	/** Workspace-relative path of the source document. */
	name: string;
	/** Sheet names to export (Excel only). */
	sheets: string[];
}

const timeoutMs = 5 * 60 * 1000;

function kindOf(name: string): JobKind | undefined {
	switch (path.extname(name).toLowerCase()) {
		case '.xlsx':
			return 'excel';
		case '.docx':
			return 'word';
		default:
			return undefined;
	}
}

/**
 * Converts Office documents to PDF by running `script/saveAsPdf.ps1` in a hidden
 * PowerShell process. Conversions run one at a time so that rapid successive saves
 * do not start several Excel/Word instances at once.
 */
export class PdfConverter {
	private queue: Promise<void> = Promise.resolve();
	private readonly scriptPath: string;

	constructor(extensionPath: string) {
		this.scriptPath = path.join(extensionPath, 'script', 'saveAsPdf.ps1');
	}

	static isSupported(name: string): boolean {
		return kindOf(name) !== undefined;
	}

	/** Resolves once the conversion finishes; rejects with the script's error message on failure. */
	convert(workspaceDir: string, target: ConvertTarget): Promise<void> {
		const kind = kindOf(target.name);
		if (!kind) {
			return Promise.resolve();
		}
		const job: Job = {
			kind,
			source: path.resolve(workspaceDir, target.name),
			output: intermediatePdfPath(workspaceDir, target.name),
			sheets: kind === 'excel' ? target.sheets : [],
		};
		const result = this.queue.then(() => this.run(job));
		this.queue = result.catch(() => undefined);
		return result;
	}

	private async run(job: Job): Promise<void> {
		console.log(`convert "${job.source}" -> "${job.output}"`);
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'mpdf-'));
		const jobPath = path.join(dir, 'job.json');
		try {
			await fs.writeFile(jobPath, JSON.stringify(job), 'utf8');
			await new Promise<void>((resolve, reject) => {
				execFile(
					'powershell.exe',
					['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', this.scriptPath, jobPath],
					{ windowsHide: true, timeout: timeoutMs, encoding: 'utf8' },
					(error, _stdout, stderr) => {
						if (error) {
							reject(new Error(stderr.trim() || error.message));
						} else {
							resolve();
						}
					},
				);
			});
		} finally {
			await fs.rm(dir, { recursive: true, force: true });
		}
	}
}

import { link, mkdir, readFile, realpath, rm, stat, writeFile } from "node:fs/promises";
import { basename, dirname, extname, join, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { randomUUID } from "node:crypto";
import { defineTool } from "@deepseek-ai/dsh-tools";
import { AlignmentType, Document, HeadingLevel, LevelFormat, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from "docx";
//#region src/office-file-route.ts
/** Same-origin, workspace-confined Office file reader for the browser SDK. */
const OFFICE_FILE_ROUTE_PREFIX = "/fylar-office-files";
const MAX_OFFICE_FILE_BYTES = 128 * 1024 * 1024;
const OFFICE_MIME = {
	".doc": "application/msword",
	".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
	".xls": "application/vnd.ms-excel",
	".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
};
function finish(res, status) {
	res.writeHead(status);
	res.end();
}
function within(root, target) {
	return target === root || target.startsWith(root + sep);
}
/** Create a GET/HEAD handler restricted to registered DSH workspace roots. */
function createOfficeFileHandler(workspaces) {
	return async (req, res) => {
		if (req.method !== "GET" && req.method !== "HEAD") {
			res.setHeader("allow", "GET, HEAD");
			finish(res, 405);
			return;
		}
		let url;
		try {
			url = new URL(req.url ?? "/", "http://dsh.local");
		} catch {
			finish(res, 400);
			return;
		}
		if (url.pathname !== "/fylar-office-files") {
			finish(res, 404);
			return;
		}
		const workspacePath = url.searchParams.get("workspace");
		const requestedPath = url.searchParams.get("path");
		if (workspacePath === null || requestedPath === null || requestedPath.trim() === "") {
			finish(res, 400);
			return;
		}
		const registered = workspaces.list().find((workspace) => workspace.path === workspacePath);
		if (registered === void 0) {
			finish(res, 403);
			return;
		}
		const canonicalRoot = await realpath(registered.path).catch(() => null);
		if (canonicalRoot === null) {
			finish(res, 404);
			return;
		}
		const candidate = resolve(canonicalRoot, requestedPath);
		if (!within(canonicalRoot, candidate)) {
			finish(res, 403);
			return;
		}
		const canonicalTarget = await realpath(candidate).catch(() => null);
		if (canonicalTarget === null || !within(canonicalRoot, canonicalTarget)) {
			finish(res, canonicalTarget === null ? 404 : 403);
			return;
		}
		const extension = extname(canonicalTarget).toLowerCase();
		const contentType = OFFICE_MIME[extension];
		if (contentType === void 0) {
			finish(res, 415);
			return;
		}
		const info = await stat(canonicalTarget).catch(() => null);
		if (info === null || !info.isFile()) {
			finish(res, 404);
			return;
		}
		if (info.size > 134217728) {
			finish(res, 413);
			return;
		}
		const body = await readFile(canonicalTarget);
		res.writeHead(200, {
			"cache-control": "private, no-store",
			"content-length": String(body.byteLength),
			"content-type": contentType,
			"cross-origin-resource-policy": "same-origin",
			"x-content-type-options": "nosniff"
		});
		res.end(req.method === "HEAD" ? void 0 : body);
	};
}
//#endregion
//#region src/docx-generator.ts
/** Markdown-to-DOCX conversion used by the model-facing Office tool. */
const HEADING_LEVELS = [
	HeadingLevel.HEADING_1,
	HeadingLevel.HEADING_2,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_4,
	HeadingLevel.HEADING_5,
	HeadingLevel.HEADING_6
];
/** Convert small, common Markdown inline spans into Word text runs. */
function inlineRuns(text) {
	const runs = [];
	const pattern = /(\*\*[^*]+\*\*|__[^_]+__|`[^`]+`|\*[^*]+\*|_[^_]+_)/gu;
	let at = 0;
	for (const match of text.matchAll(pattern)) {
		const index = match.index ?? 0;
		if (index > at) runs.push(new TextRun(text.slice(at, index)));
		const token = match[0];
		if (token.startsWith("**") || token.startsWith("__")) runs.push(new TextRun({
			text: token.slice(2, -2),
			bold: true
		}));
		else if (token.startsWith("`")) runs.push(new TextRun({
			text: token.slice(1, -1),
			font: "Courier New",
			shading: { fill: "F2F2F2" }
		}));
		else runs.push(new TextRun({
			text: token.slice(1, -1),
			italics: true
		}));
		at = index + token.length;
	}
	if (at < text.length) runs.push(new TextRun(text.slice(at)));
	return runs.length === 0 ? [new TextRun("")] : runs;
}
function tableCells(line) {
	return line.trim().replace(/^\|/u, "").replace(/\|$/u, "").split("|").map((value) => value.trim());
}
function isTableSeparator(line) {
	const cells = tableCells(line);
	return cells.length > 0 && cells.every((cell) => /^:?-{3,}:?$/u.test(cell));
}
function makeTable(rows) {
	return new Table({
		width: {
			size: 100,
			type: WidthType.PERCENTAGE
		},
		rows: rows.map((row, rowIndex) => new TableRow({ children: row.map((cell) => new TableCell({ children: [new Paragraph({ children: rowIndex === 0 ? [new TextRun({
			text: cell,
			bold: true
		})] : inlineRuns(cell) })] })) }))
	});
}
function markdownBlocks(markdown) {
	const children = [];
	const lines = markdown.replace(/\r\n?/gu, "\n").split("\n");
	let index = 0;
	let inFence = false;
	let fenceLines = [];
	while (index < lines.length) {
		const line = lines[index] ?? "";
		if (/^```/u.test(line.trim())) {
			if (inFence) {
				children.push(new Paragraph({
					children: [new TextRun({
						text: fenceLines.join("\n"),
						font: "Courier New"
					})],
					shading: { fill: "F6F8FA" },
					spacing: {
						before: 120,
						after: 120
					}
				}));
				fenceLines = [];
				inFence = false;
			} else inFence = true;
			index += 1;
			continue;
		}
		if (inFence) {
			fenceLines.push(line);
			index += 1;
			continue;
		}
		if (index + 1 < lines.length && line.includes("|") && isTableSeparator(lines[index + 1] ?? "")) {
			const rows = [tableCells(line)];
			index += 2;
			while (index < lines.length && (lines[index] ?? "").includes("|") && (lines[index] ?? "").trim() !== "") {
				rows.push(tableCells(lines[index] ?? ""));
				index += 1;
			}
			children.push(makeTable(rows));
			continue;
		}
		const heading = /^(#{1,6})\s+(.+)$/u.exec(line);
		if (heading !== null) {
			const marks = heading[1];
			const text = heading[2];
			children.push(new Paragraph({
				heading: HEADING_LEVELS[marks.length - 1],
				children: inlineRuns(text)
			}));
			index += 1;
			continue;
		}
		const bullet = /^\s*[-*+]\s+(.+)$/u.exec(line);
		if (bullet !== null) {
			children.push(new Paragraph({
				children: inlineRuns(bullet[1]),
				bullet: { level: 0 }
			}));
			index += 1;
			continue;
		}
		const ordered = /^\s*\d+[.)]\s+(.+)$/u.exec(line);
		if (ordered !== null) {
			children.push(new Paragraph({
				children: inlineRuns(ordered[1]),
				numbering: {
					reference: "fylar-numbering",
					level: 0
				}
			}));
			index += 1;
			continue;
		}
		if (line.trim() === "") children.push(new Paragraph(""));
		else children.push(new Paragraph({
			children: inlineRuns(line),
			spacing: { after: 120 }
		}));
		index += 1;
	}
	if (inFence && fenceLines.length > 0) children.push(new Paragraph({
		children: [new TextRun({
			text: fenceLines.join("\n"),
			font: "Courier New"
		})],
		shading: { fill: "F6F8FA" }
	}));
	return children;
}
/** Build a standards-compliant DOCX buffer from a title and Markdown body. */
async function createDocxBuffer(title, markdown) {
	const content = markdownBlocks(markdown);
	const document = new Document({
		creator: "DSH Fylar Office Editor",
		title: title ?? "",
		description: "Generated by dsh-fylar-office-editor",
		numbering: { config: [{
			reference: "fylar-numbering",
			levels: [{
				level: 0,
				format: LevelFormat.DECIMAL,
				text: "%1.",
				alignment: AlignmentType.START,
				style: { paragraph: { indent: {
					left: 720,
					hanging: 360
				} } }
			}]
		}] },
		sections: [{
			properties: {},
			children: [...title === void 0 || title.trim() === "" ? [] : [new Paragraph({
				heading: HeadingLevel.TITLE,
				children: inlineRuns(title.trim())
			})], ...content]
		}]
	});
	return Packer.toBuffer(document);
}
//#endregion
//#region src/office-tool.ts
/** Model-facing Office creation and delivery tools for the active DSH workspace. */
const MAX_MARKDOWN_CHARS = 1e6;
const PRESENTABLE_OFFICE_TYPES = [
	"doc",
	"docx",
	"xls",
	"xlsx",
	"pptx"
];
/** Validate constraints that the DSH schema DSL cannot express. */
function validateCreateDocxArgs(args) {
	if (args.file_path.trim() === "") throw new Error("file_path must be a non-empty path");
	if (extname(args.file_path).toLowerCase() !== ".docx") throw new Error("file_path must end with .docx");
	if (args.content.length > MAX_MARKDOWN_CHARS) throw new Error(`content exceeds the ${MAX_MARKDOWN_CHARS}-character limit`);
}
/** Validate the path shape before resolving it against the active Workspace. */
function validatePresentOfficeFileArgs(args) {
	if (args.file_path.trim() === "") throw new Error("file_path must be a non-empty path");
	const extension = extname(args.file_path).toLowerCase().slice(1);
	if (!PRESENTABLE_OFFICE_TYPES.includes(extension)) throw new Error("file_path must end with .doc, .docx, .xls, .xlsx, or .pptx");
	return extension;
}
async function writeNewFile(path, bytes, signal) {
	await mkdir(dirname(path), { recursive: true });
	const temporary = join(dirname(path), `.${basename(path)}.${randomUUID()}.tmp`);
	try {
		await writeFile(temporary, bytes, {
			flag: "wx",
			signal
		});
		signal.throwIfAborted();
		await link(temporary, path);
	} finally {
		await rm(temporary, { force: true }).catch(() => void 0);
	}
}
/** Register Office creation/presentation tools and their stable model guidance. */
function registerOfficeTools(ctx) {
	ctx.systemPrompt.section({
		name: "tool:office-create-docx",
		order: 106,
		text: "When the user asks for a Word document, use office_create_docx to create a real .docx file in the current workspace. Pass the complete document body as Markdown. When an existing .doc, .docx, .xls, .xlsx, or .pptx file was created through bash, Python, or another tool, call office_present_file exactly once with its Workspace-relative path so the Web UI can render a clickable Office file card. Do not call office_present_file after office_create_docx because that tool already publishes its result."
	});
	ctx.tools.register(defineTool({
		name: "office_create_docx",
		description: "Create a real Microsoft Word .docx file in the current workspace from Markdown content. The result opens in the DSH Office side panel.",
		parameters: {
			file_path: {
				type: "string",
				required: true,
				description: "Workspace-relative output path ending in .docx, for example reports/summary.docx. The file must not already exist."
			},
			title: {
				type: "string",
				description: "Optional document title placed before the Markdown body."
			},
			content: {
				type: "string",
				required: true,
				description: "Complete document body in Markdown. Headings, paragraphs, bullets, numbered lists, fenced code, simple tables, bold, italic, and inline code are supported."
			}
		},
		output: {
			schema: {
				type: "object",
				additionalProperties: false,
				properties: {
					path: {
						type: "string",
						required: true
					},
					fileName: {
						type: "string",
						required: true
					},
					bytes: {
						type: "integer",
						required: true
					},
					type: {
						type: "string",
						required: true,
						enum: ["docx"]
					}
				}
			},
			render: (_args, value) => [{
				type: "text",
				text: `<path>${value.path}</path>\n<type>file</type>\n<content>Created DOCX file (${value.bytes} bytes)</content>`
			}]
		},
		async execute(args, exec) {
			validateCreateDocxArgs(args);
			const session = exec.agent?.session;
			if (session === void 0) throw new Error("office_create_docx requires an agent workspace");
			const cwd = session.header.cwd;
			if (cwd === void 0) throw new Error("office_create_docx requires an agent workspace path");
			if (ctx.sandboxPolicy.resolve({ session }).mode === "read-only") throw new Error("[sandbox: read-only] office_create_docx cannot create files in the current permission mode");
			const resolveOptions = {
				cwd,
				signal: exec.signal
			};
			const workspace = await ctx.fs.resolve(".", resolveOptions);
			let target = await ctx.fs.resolve(args.file_path, resolveOptions);
			if (!ctx.fs.contains(workspace, target)) throw new Error("file_path must stay inside the current workspace");
			if (await ctx.fs.stat(target, exec.signal) !== void 0) throw new Error(`refusing to overwrite existing file: ${args.file_path}`);
			const buffer = await createDocxBuffer(args.title, args.content);
			exec.signal.throwIfAborted();
			await writeNewFile(ctx.fs.processPath(target), buffer, exec.signal);
			target = await ctx.fs.resolve(args.file_path, resolveOptions);
			if (!ctx.fs.contains(workspace, target)) throw new Error("created file resolved outside the current workspace");
			return {
				path: args.file_path,
				fileName: basename(args.file_path),
				bytes: buffer.byteLength,
				type: "docx"
			};
		},
		presentCall: (args) => ({
			card: "generic",
			title: `Create Word document ${args.file_path}`,
			kind: "other",
			locations: [{ path: args.file_path }]
		}),
		presentResult(args, result) {
			if (result.isError) return void 0;
			return {
				card: "generic",
				title: `Created Word document ${args.file_path}`,
				kind: "other",
				locations: [{ path: args.file_path }]
			};
		}
	}));
	ctx.tools.register(defineTool({
		name: "office_present_file",
		description: "Publish an existing Office file from the current workspace as a clickable preview/edit card. Use this after another tool creates DOC, DOCX, XLS, XLSX, or PPTX.",
		parameters: { file_path: {
			type: "string",
			required: true,
			description: "Workspace-relative path to an existing .doc, .docx, .xls, .xlsx, or .pptx file."
		} },
		output: {
			schema: {
				type: "object",
				additionalProperties: false,
				properties: {
					path: {
						type: "string",
						required: true
					},
					fileName: {
						type: "string",
						required: true
					},
					bytes: {
						type: "integer",
						required: true
					},
					type: {
						type: "string",
						required: true,
						enum: [...PRESENTABLE_OFFICE_TYPES]
					}
				}
			},
			render: (_args, value) => [{
				type: "text",
				text: `<path>${value.path}</path>\n<type>file</type>\n<content>Office file is ready to preview (${value.bytes} bytes)</content>`
			}]
		},
		async execute(args, exec) {
			const type = validatePresentOfficeFileArgs(args);
			const session = exec.agent?.session;
			if (session === void 0) throw new Error("office_present_file requires an agent workspace");
			const cwd = session.header.cwd;
			if (cwd === void 0) throw new Error("office_present_file requires an agent workspace path");
			const resolveOptions = {
				cwd,
				signal: exec.signal
			};
			const workspace = await ctx.fs.resolve(".", resolveOptions);
			const target = await ctx.fs.resolve(args.file_path, resolveOptions);
			if (!ctx.fs.contains(workspace, target)) throw new Error("file_path must stay inside the current workspace");
			const info = await ctx.fs.stat(target, exec.signal);
			if (info === void 0) throw new Error(`Office file does not exist: ${args.file_path}`);
			if (info.type !== "file") throw new Error(`Office path is not a regular file: ${args.file_path}`);
			if (info.size !== void 0 && info.size > 134217728) throw new Error(`Office file exceeds the ${MAX_OFFICE_FILE_BYTES}-byte preview limit`);
			return {
				path: args.file_path,
				fileName: basename(args.file_path),
				bytes: info.size ?? 0,
				type
			};
		},
		presentCall: (args) => ({
			card: "generic",
			title: `Present Office file ${args.file_path}`,
			kind: "other",
			locations: [{ path: args.file_path }]
		}),
		presentResult(args, result) {
			if (result.isError) return void 0;
			return {
				card: "generic",
				title: `Office file ready ${args.file_path}`,
				kind: "other",
				locations: [{ path: args.file_path }]
			};
		}
	}));
}
//#endregion
//#region src/office-sdk.generated.ts
/** Generated by scripts/sync-office-sdk.mjs. */
const OFFICE_SDK_VERSION = "1.0.0";
const OFFICE_SDK_REVISION = "f24b12f6b857";
const OFFICE_SDK_ROUTE_PREFIX = "/fylar-office-sdk";
const OFFICE_SDK_BASE_PATH = `${OFFICE_SDK_ROUTE_PREFIX}/${OFFICE_SDK_VERSION}-${OFFICE_SDK_REVISION}`;
//#endregion
//#region src/index.ts
/** DeepSeek Harness host entry for the Fylar Office SDK static runtime. */
/** Stable Cordis plugin name. */
const name = "dsh-fylar-office-editor";
/** Host services required by the SDK route, workspace file route, and Agent tool. */
const inject = [
	"webServer",
	"tools",
	"fs",
	"workspaceRegistry",
	"systemPrompt",
	"sandboxPolicy"
];
const SDK_ROOT = fileURLToPath(new URL("../vendor/office-sdk/lib/", import.meta.url));
const CACHE_CONTROL = "public, max-age=31536000, immutable";
const MIME = {
	".js": "text/javascript; charset=utf-8",
	".css": "text/css; charset=utf-8",
	".json": "application/json; charset=utf-8",
	".wasm": "application/wasm",
	".zip": "application/zip",
	".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
	".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
	".txt": "text/plain; charset=utf-8",
	".ts": "text/plain; charset=utf-8"
};
function end(res, status) {
	res.writeHead(status);
	res.end();
}
function isWithin(root, target) {
	return target === root || target.startsWith(root + sep);
}
/**
* Create the versioned Office static-asset handler.
* @param sdkRoot - Absolute vendored lib directory; overridable by tests.
* @returns A node:http request handler.
*/
function createOfficeAssetHandler(sdkRoot = SDK_ROOT) {
	return async (req, res) => {
		if (req.method !== "GET" && req.method !== "HEAD") {
			end(res, 405);
			return;
		}
		const rawPath = (req.url ?? "/").split(/[?#]/u, 1)[0] ?? "/";
		let pathname;
		try {
			pathname = decodeURIComponent(rawPath);
		} catch {
			end(res, 400);
			return;
		}
		if (!pathname.startsWith(`/fylar-office-sdk/`)) {
			end(res, 404);
			return;
		}
		if (!pathname.startsWith(`${OFFICE_SDK_BASE_PATH}/`)) {
			end(res, 404);
			return;
		}
		const relativePath = pathname.slice(OFFICE_SDK_BASE_PATH.length + 1);
		const segments = relativePath.split("/");
		if (relativePath === "" || pathname.includes("\\") || segments.some((segment) => segment === "" || segment === "." || segment === "..")) {
			end(res, 403);
			return;
		}
		const canonicalRoot = await realpath(sdkRoot);
		const target = resolve(canonicalRoot, ...segments);
		if (!isWithin(canonicalRoot, target)) {
			end(res, 403);
			return;
		}
		const targetStat = await stat(target).catch(() => null);
		if (targetStat === null || !targetStat.isFile()) {
			end(res, 404);
			return;
		}
		const canonicalTarget = await realpath(target);
		if (!isWithin(canonicalRoot, canonicalTarget)) {
			end(res, 403);
			return;
		}
		const body = await readFile(canonicalTarget);
		res.writeHead(200, {
			"cache-control": CACHE_CONTROL,
			"content-length": String(body.byteLength),
			"content-type": MIME[extname(canonicalTarget).toLowerCase()] ?? "application/octet-stream",
			"cross-origin-resource-policy": "same-origin",
			"x-content-type-options": "nosniff"
		});
		res.end(req.method === "HEAD" ? void 0 : body);
	};
}
/** Register the Office runtime under the DSH Web origin. */
function apply(ctx) {
	const handler = createOfficeAssetHandler();
	const fileHandler = createOfficeFileHandler(ctx.workspaceRegistry);
	ctx.effect(() => ctx.webServer.register({
		kind: "prefix",
		path: OFFICE_SDK_ROUTE_PREFIX,
		handler
	}), "dsh-fylar-office-editor: static runtime");
	ctx.effect(() => ctx.webServer.register({
		kind: "prefix",
		path: OFFICE_FILE_ROUTE_PREFIX,
		handler: fileHandler
	}), "dsh-fylar-office-editor: workspace files");
	registerOfficeTools(ctx);
}
//#endregion
export { apply, createOfficeAssetHandler, inject, name };

//# sourceMappingURL=index.js.map
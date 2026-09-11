window.__ModuleLoader__.load({
	id: "@fylar/dsh-fylar-office-editor",
	factory: (require) => {
		var module = { exports: {} };
		var exports = module.exports;
		Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });
		let react_jsx_runtime = require("react/jsx-runtime");
		let react = require("react");
		//#region src/client/office-errors.ts
		var OfficePluginFailure = class extends Error {
			kind;
			constructor(kind, message, options) {
				super(message, options);
				this.kind = kind;
				this.name = "OfficePluginFailure";
			}
		};
		const RETRYABLE = {
			"unsupported-file": false,
			"unsupported-browser": false,
			"runtime-load": true,
			asset: true,
			open: true,
			create: true,
			export: true,
			"pdf-environment": false,
			cancelled: false,
			unknown: true
		};
		function issueFor(kind) {
			return {
				kind,
				retryable: RETRYABLE[kind]
			};
		}
		function classifyOfficeError(error, fallback) {
			if (error instanceof OfficePluginFailure) return issueFor(error.kind);
			const message = error instanceof Error ? error.message : String(error ?? "");
			if (/OPEN_CANCELLED|cancelled|canceled/iu.test(message)) return issueFor("cancelled");
			if (/worker|sharedworker|wasm|webassembly|404|cors|mime|network|fetch/iu.test(message)) return issueFor("asset");
			return issueFor(fallback);
		}
		//#endregion
		//#region src/client/office-files.ts
		const OFFICE_FILE_ACCEPT = ".doc,.docx,.xls,.xlsx,.pptx";
		const SUPPORTED_EXTENSIONS = /* @__PURE__ */ new Set([
			"doc",
			"docx",
			"xls",
			"xlsx",
			"pptx"
		]);
		function officeFileExtension(name) {
			const index = name.lastIndexOf(".");
			return index < 0 ? "" : name.slice(index + 1).toLowerCase();
		}
		function officeBaseName(path) {
			const index = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
			return index < 0 ? path : path.slice(index + 1);
		}
		function isSupportedOfficeFile(name) {
			return SUPPORTED_EXTENSIONS.has(officeFileExtension(name));
		}
		function assertSupportedOfficeFile(file) {
			if (!isSupportedOfficeFile(file.name)) throw new OfficePluginFailure("unsupported-file", `Unsupported Office file: ${file.name}`);
		}
		function inferOfficeDocType(name) {
			switch (officeFileExtension(name)) {
				case "doc":
				case "docx": return 1;
				case "xls":
				case "xlsx": return 2;
				case "pptx": return 3;
				default: throw new OfficePluginFailure("unsupported-file", `Unsupported Office file: ${name}`);
			}
		}
		function createdFileName(docType) {
			if (docType === 1) return "Untitled.docx";
			if (docType === 2) return "Untitled.xlsx";
			return "Untitled.pptx";
		}
		function describeOfficeDocument(name, fallbackType) {
			const extension = officeFileExtension(name);
			const docType = extension === "" && fallbackType !== void 0 ? fallbackType : inferOfficeDocType(name);
			if (docType === 1) return {
				docType,
				product: "Word",
				extension: extension === "doc" ? "DOC" : "DOCX",
				supportsPdf: true,
				accent: "word"
			};
			if (docType === 2) return {
				docType,
				product: "Excel",
				extension: extension === "xls" ? "XLS" : "XLSX",
				supportsPdf: true,
				accent: "excel"
			};
			return {
				docType,
				product: "PowerPoint",
				extension: "PPTX",
				supportsPdf: false,
				accent: "powerpoint"
			};
		}
		function officeDirectoryLabel(path, rootLabel) {
			const normalized = path.replaceAll("\\", "/");
			const index = normalized.lastIndexOf("/");
			return index <= 0 ? rootLabel : normalized.slice(0, index);
		}
		//#endregion
		//#region src/client/office-artifact-controller.ts
		/** Global browser session linking Chat artifacts, the Office tab, and the right Sidebar. */
		const OFFICE_FILE_ROUTE_PREFIX = "/fylar-office-files";
		function isSameWorkspaceSource(source, target) {
			return source.kind === "workspace" && source.cwd === target.cwd && source.path === target.path;
		}
		function workspaceOfficeUrl(cwd, path) {
			return `${OFFICE_FILE_ROUTE_PREFIX}?${new URLSearchParams({
				workspace: cwd,
				path
			}).toString()}`;
		}
		/** Owns the one live Office document and every place where its DOM may appear. */
		var OfficeSessionController = class {
			workspace;
			detailsPanel;
			snapshot = { current: null };
			listeners = /* @__PURE__ */ new Set();
			disposed = false;
			source = {
				getSnapshot: () => this.snapshot,
				subscribe: (listener) => {
					this.listeners.add(listener);
					return () => {
						this.listeners.delete(listener);
					};
				}
			};
			constructor(workspace, detailsPanel) {
				this.workspace = workspace;
				this.detailsPanel = detailsPanel;
			}
			publish(change) {
				this.snapshot = {
					...this.snapshot,
					...change
				};
				for (const listener of this.listeners) listener();
			}
			present(selection) {
				this.publish({ current: selection });
				if (selection.presentation === "details") {
					this.detailsPanel.open();
					return;
				}
				this.detailsPanel.close();
			}
			async openWorkspaceNow(target, presentation) {
				const fileName = officeBaseName(target.path);
				const source = {
					kind: "workspace",
					path: target.path,
					cwd: target.cwd,
					fileName,
					url: workspaceOfficeUrl(target.cwd, target.path),
					docType: inferOfficeDocType(fileName)
				};
				this.present({
					source,
					presentation
				});
				await this.workspace.openUrl(fileName, source.url);
			}
			async openWorkspace(target, presentation) {
				if (this.disposed || target.cwd === void 0) return;
				const current = this.snapshot.current;
				if (current !== null && isSameWorkspaceSource(current.source, target)) {
					this.present({
						...current,
						presentation
					});
					return;
				}
				await this.openWorkspaceNow({
					path: target.path,
					cwd: target.cwd
				}, presentation);
			}
			async openDetails(target) {
				await this.openWorkspace(target, "details");
			}
			async openInline(target) {
				await this.openWorkspace(target, "inline");
			}
			async openLocal(file) {
				if (this.disposed) return;
				let docType;
				try {
					docType = inferOfficeDocType(file.name);
				} catch {
					await this.workspace.open(file);
					return;
				}
				const source = {
					kind: "local",
					fileName: file.name,
					docType
				};
				this.present({
					source,
					presentation: "office-view"
				});
				await this.workspace.open(file);
			}
			async create(docType) {
				if (this.disposed) return;
				const fileName = createdFileName(docType);
				this.present({
					source: {
						kind: "created",
						fileName,
						docType
					},
					presentation: "office-view"
				});
				await this.workspace.create(docType);
			}
			activateOfficeView() {
				const current = this.snapshot.current;
				if (this.disposed || current === null || current.presentation === "office-view") return;
				this.present({
					...current,
					presentation: "office-view"
				});
			}
			parkOfficeView() {
				const current = this.snapshot.current;
				if (this.disposed || current?.presentation !== "office-view") return;
				this.publish({ current: {
					...current,
					presentation: "parked"
				} });
			}
			collapseInline() {
				const current = this.snapshot.current;
				if (this.disposed || current?.presentation !== "inline") return;
				this.publish({ current: {
					...current,
					presentation: "parked"
				} });
			}
			/** Keep the document alive when DSH closes its native right-Sidebar tab. */
			detailsPanelClosed() {
				const current = this.snapshot.current;
				if (this.disposed || current?.presentation !== "details") return;
				this.publish({ current: {
					...current,
					presentation: "parked"
				} });
			}
			async retry() {
				await this.workspace.retry();
			}
			async close() {
				if (this.disposed) return;
				this.publish({ current: null });
				this.detailsPanel.close();
				await this.workspace.close();
			}
			async dispose() {
				if (this.disposed) return;
				this.disposed = true;
				this.detailsPanel.close();
				await this.workspace.dispose();
				this.snapshot = { current: null };
				this.listeners.clear();
			}
		};
		//#endregion
		//#region src/client/locales.ts
		const OFFICE_LOCALE_NS = "fylar-office";
		const zh = {
			"view.label": "Office",
			"document.fallback": "Office 文档",
			"document.kind": "{product} 文档",
			"workspace.root": "工作区根目录",
			"tool.present.running": "正在准备 Office 文件预览…",
			"tool.present.error": "Office 文件发布失败",
			"tool.present.stopped": "Office 文件发布已停止",
			"tool.present.ready": "Office 文件已可预览",
			"tool.create.running": "正在生成 Word 文档…",
			"tool.create.error": "Word 文档生成失败",
			"tool.create.stopped": "Word 文档生成已停止",
			"tool.create.ready": "Word 文档已生成",
			"artifact.workspaceUnavailable": "当前会话工作区不可用",
			"artifact.parked": "临时编辑已保留 · 继续编辑",
			"artifact.details": "正在侧边栏编辑 · 点击移至对话",
			"artifact.officeView": "正在 Office 页签编辑 · 点击移至对话",
			"artifact.openInline": "{product} 文档 · 点击在对话中打开",
			"artifact.more": "更多操作：{name}",
			"artifact.group": "本轮生成的 Office 文件",
			"mode.openFailed": "打开失败",
			"mode.preparing": "正在准备",
			"mode.switching": "正在切换…",
			"mode.readOnly": "只读",
			"mode.dirty": "有临时修改",
			"mode.editing": "临时编辑",
			"help.readOnly": "文档当前为只读状态，可在更多菜单中启用编辑。",
			"help.editing": "当前修改不会覆盖原文件，请通过导出菜单下载副本保存。",
			"action.export": "导出",
			"action.downloading": "正在下载…",
			"action.exporting": "正在导出…",
			"action.download": "下载并保存 {extension}",
			"action.downloadDescription": "保存当前编辑内容的本地副本",
			"action.exportPdf": "导出 PDF",
			"action.exportPdfDescription": "生成便于分享的 PDF 副本",
			"action.enableEditing": "启用编辑",
			"action.enableEditingDescription": "允许修改当前文档",
			"action.readOnly": "切换为只读",
			"action.readOnlyDescription": "防止误触修改文档",
			"action.openDetails": "在侧边栏打开",
			"action.closeSession": "关闭编辑会话",
			"action.closeSessionDescription": "未下载的临时修改将不再保留",
			"action.fullscreen": "全屏编辑",
			"action.exitFullscreen": "退出全屏编辑",
			"action.fullscreenUnsupported": "当前浏览器不支持全屏",
			"action.collapse": "收起",
			"action.collapseAria": "收起对话内编辑器",
			"action.moreDocument": "更多文档操作",
			"action.closeDocument": "关闭 Office 文档",
			"action.retry": "重试",
			"action.dismiss": "关闭提示",
			"editor.inlineAria": "对话内 Office 文档编辑器",
			"editor.aria": "Office 文档编辑器",
			"editor.opening": "正在打开 {name}…",
			"empty.description": "打开或新建 Word、Excel、PowerPoint 文档，在浏览器中临时编辑并下载保存。",
			"empty.loading": "正在加载 Office 运行时…",
			"empty.openFile": "打开 Office 文件",
			"empty.newDocument": "新建文档",
			"empty.newWord": "新建 Word",
			"empty.newExcel": "新建 Excel",
			"empty.newPowerPoint": "新建 PowerPoint",
			"pdf.insecureContext": "PDF 导出需要在 localhost 或 HTTPS 安全上下文中运行。",
			"pdf.localFontsUnavailable": "当前浏览器不支持读取本地字体，无法导出 PDF。",
			"error.unsupported-file.title": "不支持的文件",
			"error.unsupported-file.detail": "请选择 DOC、DOCX、XLS、XLSX 或 PPTX 文件。",
			"error.unsupported-browser.title": "浏览器能力不可用",
			"error.unsupported-browser.detail": "Fylar Office 需要浏览器支持 Worker、SharedWorker、WebAssembly、IndexedDB、Blob 和 ArrayBuffer。",
			"error.runtime-load.title": "Office 运行时加载失败",
			"error.runtime-load.detail": "请重试；如果仍然失败，请检查已安装的 Office SDK 资源。",
			"error.asset.title": "Office 资源加载失败",
			"error.asset.detail": "Worker、WebAssembly、语言包或运行时资源无法加载。",
			"error.open.title": "无法打开文档",
			"error.open.detail": "请确认文件是有效且受支持的 Office 文档，然后重试。",
			"error.create.title": "无法新建文档",
			"error.create.detail": "Office SDK 无法新建文档，请重试。",
			"error.export.title": "导出失败",
			"error.export.detail": "Office SDK 无法完成导出，请重试。",
			"error.pdf-environment.title": "当前环境无法导出 PDF",
			"error.pdf-environment.detail": "PDF 导出需要 localhost 或 HTTPS，并且浏览器必须支持访问本地字体。",
			"error.cancelled.title": "操作已取消",
			"error.cancelled.detail": "当前操作已被更新的 Office 操作替代。",
			"error.unknown.title": "Office 操作失败",
			"error.unknown.detail": "发生了意外错误，请重试。"
		};
		const en = {
			"view.label": "Office",
			"document.fallback": "Office document",
			"document.kind": "{product} document",
			"workspace.root": "Workspace root",
			"tool.present.running": "Preparing Office file preview…",
			"tool.present.error": "Failed to publish Office file",
			"tool.present.stopped": "Office file publishing stopped",
			"tool.present.ready": "Office file ready to preview",
			"tool.create.running": "Creating Word document…",
			"tool.create.error": "Failed to create Word document",
			"tool.create.stopped": "Word document creation stopped",
			"tool.create.ready": "Word document created",
			"artifact.workspaceUnavailable": "The current conversation workspace is unavailable",
			"artifact.parked": "Temporary edits retained · Continue editing",
			"artifact.details": "Editing in the side panel · Click to move into chat",
			"artifact.officeView": "Editing in the Office tab · Click to move into chat",
			"artifact.openInline": "{product} document · Click to open in chat",
			"artifact.more": "More actions: {name}",
			"artifact.group": "Office files created in this turn",
			"mode.openFailed": "Open failed",
			"mode.preparing": "Preparing",
			"mode.switching": "Switching…",
			"mode.readOnly": "Read only",
			"mode.dirty": "Temporary changes",
			"mode.editing": "Temporary editing",
			"help.readOnly": "This document is read only. Enable editing from the More menu.",
			"help.editing": "Changes do not overwrite the original file. Download a copy from the Export menu to save them.",
			"action.export": "Export",
			"action.downloading": "Downloading…",
			"action.exporting": "Exporting…",
			"action.download": "Download and save {extension}",
			"action.downloadDescription": "Save a local copy of the current edits",
			"action.exportPdf": "Export PDF",
			"action.exportPdfDescription": "Create a PDF copy for sharing",
			"action.enableEditing": "Enable editing",
			"action.enableEditingDescription": "Allow changes to the current document",
			"action.readOnly": "Switch to read only",
			"action.readOnlyDescription": "Prevent accidental document changes",
			"action.openDetails": "Open in side panel",
			"action.closeSession": "Close editing session",
			"action.closeSessionDescription": "Temporary changes that were not downloaded will be lost",
			"action.fullscreen": "Edit in full screen",
			"action.exitFullscreen": "Exit full screen",
			"action.fullscreenUnsupported": "Full screen is not supported by this browser",
			"action.collapse": "Collapse",
			"action.collapseAria": "Collapse the inline document editor",
			"action.moreDocument": "More document actions",
			"action.closeDocument": "Close Office document",
			"action.retry": "Retry",
			"action.dismiss": "Dismiss",
			"editor.inlineAria": "Inline Office document editor",
			"editor.aria": "Office document editor",
			"editor.opening": "Opening {name}…",
			"empty.description": "Open or create Word, Excel, and PowerPoint documents for temporary browser editing and download.",
			"empty.loading": "Loading Office runtime…",
			"empty.openFile": "Open Office file",
			"empty.newDocument": "Create document",
			"empty.newWord": "New Word",
			"empty.newExcel": "New Excel",
			"empty.newPowerPoint": "New PowerPoint",
			"pdf.insecureContext": "PDF export requires localhost or a secure HTTPS context.",
			"pdf.localFontsUnavailable": "This browser cannot access local fonts, so PDF export is unavailable.",
			"error.unsupported-file.title": "Unsupported file",
			"error.unsupported-file.detail": "Choose a DOC, DOCX, XLS, XLSX, or PPTX file.",
			"error.unsupported-browser.title": "Browser capabilities unavailable",
			"error.unsupported-browser.detail": "Fylar Office requires Worker, SharedWorker, WebAssembly, IndexedDB, Blob, and ArrayBuffer support.",
			"error.runtime-load.title": "Failed to load the Office runtime",
			"error.runtime-load.detail": "Retry. If it still fails, check the installed Office SDK assets.",
			"error.asset.title": "Failed to load Office assets",
			"error.asset.detail": "A Worker, WebAssembly module, locale bundle, or runtime asset could not be loaded.",
			"error.open.title": "Unable to open document",
			"error.open.detail": "Confirm that the file is a valid, supported Office document, then retry.",
			"error.create.title": "Unable to create document",
			"error.create.detail": "The Office SDK could not create the document. Please retry.",
			"error.export.title": "Export failed",
			"error.export.detail": "The Office SDK could not complete the export. Please retry.",
			"error.pdf-environment.title": "PDF export is unavailable in this environment",
			"error.pdf-environment.detail": "PDF export requires localhost or HTTPS and browser access to local fonts.",
			"error.cancelled.title": "Operation cancelled",
			"error.cancelled.detail": "A newer Office operation replaced this operation.",
			"error.unknown.title": "Office operation failed",
			"error.unknown.detail": "An unexpected error occurred. Please retry."
		};
		function translator(dictionary) {
			return (key, params) => {
				const template = dictionary[key] ?? key;
				if (params === void 0) return template;
				return template.replace(/\{(\w+)\}/gu, (match, name) => Object.prototype.hasOwnProperty.call(params, name) ? String(params[name]) : match);
			};
		}
		/** Chinese fallback for direct embedding and isolated component tests. */
		const zhT = translator(zh);
		function officeIssueCopy(t, kind) {
			return {
				title: t(`error.${kind}.title`),
				detail: t(`error.${kind}.detail`)
			};
		}
		//#endregion
		//#region src/client/OfficeArtifactRow.tsx
		const LABEL_KEYS = {
			present: {
				running: "tool.present.running",
				ready: "tool.present.ready",
				error: "tool.present.error",
				stopped: "tool.present.stopped"
			},
			create: {
				running: "tool.create.running",
				ready: "tool.create.ready",
				error: "tool.create.error",
				stopped: "tool.create.stopped"
			}
		};
		function pathFromArgs(raw) {
			try {
				const value = JSON.parse(raw);
				if (typeof value === "object" && value !== null) {
					const path = value.file_path;
					if (typeof path === "string" && path !== "") return path;
				}
			} catch {}
			return "document.docx";
		}
		function officeArtifactRowModel(block) {
			const settled = "kind" in block;
			const raw = (settled ? block.call?.argsRaw : block.argsRaw) ?? "";
			const state = !settled ? "running" : block.error?.code === "interrupted" ? "stopped" : block.isError ? "error" : "ready";
			return {
				path: pathFromArgs(raw),
				state
			};
		}
		/** Render generation progress only; the assistant-tail card is the file entry. */
		function OfficeArtifactRow({ block, toolName, t = zhT }) {
			const model = officeArtifactRowModel(block);
			const label = t(LABEL_KEYS[toolName === "office_present_file" ? "present" : "create"][model.state]);
			return /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
				className: "fylar-office-tool-status",
				"data-state": model.state,
				role: "status",
				children: [
					/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", {
						className: "fylar-office-tool-status-dot",
						"aria-hidden": true
					}),
					/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", { children: label }),
					/* @__PURE__ */ (0, react_jsx_runtime.jsx)("small", {
						title: model.path,
						children: officeBaseName(model.path)
					})
				]
			});
		}
		//#endregion
		//#region node_modules/@deepseek-ai/dsh-session/lib/types/surface.js
		/** Runtime counterpart of the message-producing event union. */
		const SURFACE_EVENT_TYPES = /* @__PURE__ */ new Set([
			"system/message",
			"user/message",
			"assistant/message",
			"tool/result"
		]);
		/**
		* Narrow an event to a surface-eligible event carrying its required marker.
		* @param event - event to test.
		* @returns true when both the type and marker identify a surface event.
		*/
		function isSurfaceEvent(event) {
			if (!SURFACE_EVENT_TYPES.has(event.type)) return false;
			return event.surfaceOp !== void 0;
		}
		/**
		* Narrow an event to an append-origin surface event: one that entered the
		* surface at its own log position and was never itself a replacement copy.
		*
		* The model-visible surface deliberately shadows replaced ranges, so it is the
		* wrong source for a human transcript — a landed replacement would erase
		* conversation the user already saw. Append-origin events are that transcript's
		* durable source material; replacement copies stay model-only.
		* @param event - event to test.
		* @returns true when the event appended to the surface tail.
		*/
		function isAppendSurfaceEvent(event) {
			return isSurfaceEvent(event) && event.surfaceOp === "append";
		}
		//#endregion
		//#region src/client/office-capabilities.ts
		const REQUIRED_BROWSER_CAPABILITIES = [
			"Worker",
			"SharedWorker",
			"WebAssembly",
			"indexedDB",
			"Blob",
			"ArrayBuffer"
		];
		/** Return missing Office runtime browser capabilities. */
		function missingOfficeCapabilities(subject = globalThis) {
			return REQUIRED_BROWSER_CAPABILITIES.filter((name) => !(name in subject));
		}
		/** Fail before runtime initialization when the browser cannot host the SDK. */
		function assertOfficeCapabilities(subject = globalThis) {
			const missing = missingOfficeCapabilities(subject);
			if (missing.length > 0) throw new OfficePluginFailure("unsupported-browser", `Missing browser capabilities: ${missing.join(", ")}`);
		}
		/** Map the host language to one of the two locale bundles shipped by the vendored SDK. */
		function resolveOfficeLocale(language) {
			return language?.toLowerCase().startsWith("zh") === true ? "zh-CN" : "en-US";
		}
		function pdfEnvironmentProblem(subject = window) {
			const browser = subject;
			if (!browser.isSecureContext) return "insecure-context";
			if (typeof browser.queryLocalFonts !== "function") return "local-fonts-unavailable";
			return null;
		}
		//#endregion
		//#region src/client/OfficeActionMenu.tsx
		function MenuIcon({ icon }) {
			return icon === "export" ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("svg", {
				viewBox: "0 0 16 16",
				width: "16",
				height: "16",
				"aria-hidden": true,
				children: /* @__PURE__ */ (0, react_jsx_runtime.jsx)("path", {
					d: "M8 2.5v7m0 0 2.5-2.5M8 9.5 5.5 7M3 11v2h10v-2",
					fill: "none",
					stroke: "currentColor",
					strokeWidth: "1.35",
					strokeLinecap: "round",
					strokeLinejoin: "round"
				})
			}) : /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("svg", {
				viewBox: "0 0 16 16",
				width: "17",
				height: "17",
				"aria-hidden": true,
				children: [
					/* @__PURE__ */ (0, react_jsx_runtime.jsx)("circle", {
						cx: "3.25",
						cy: "8",
						r: "1.1",
						fill: "currentColor"
					}),
					/* @__PURE__ */ (0, react_jsx_runtime.jsx)("circle", {
						cx: "8",
						cy: "8",
						r: "1.1",
						fill: "currentColor"
					}),
					/* @__PURE__ */ (0, react_jsx_runtime.jsx)("circle", {
						cx: "12.75",
						cy: "8",
						r: "1.1",
						fill: "currentColor"
					})
				]
			});
		}
		/** Shared DSH-like action menu used by editor headers and delivery cards. */
		function OfficeActionMenu({ label, icon, variant = "icon", disabled = false, items }) {
			const [open, setOpen] = (0, react.useState)(false);
			const menuId = (0, react.useId)();
			const rootRef = (0, react.useRef)(null);
			const triggerRef = (0, react.useRef)(null);
			const menuRef = (0, react.useRef)(null);
			(0, react.useEffect)(() => {
				if (!open) return;
				menuRef.current?.querySelector("button:not(:disabled)")?.focus();
				const closeOutside = (event) => {
					if (rootRef.current?.contains(event.target) !== true) setOpen(false);
				};
				const closeWithEscape = (event) => {
					if (event.key !== "Escape") return;
					event.preventDefault();
					setOpen(false);
					triggerRef.current?.focus();
				};
				document.addEventListener("pointerdown", closeOutside);
				document.addEventListener("keydown", closeWithEscape);
				return () => {
					document.removeEventListener("pointerdown", closeOutside);
					document.removeEventListener("keydown", closeWithEscape);
				};
			}, [open]);
			const navigateMenu = (event) => {
				if (![
					"ArrowDown",
					"ArrowUp",
					"Home",
					"End"
				].includes(event.key)) return;
				const controls = [...menuRef.current?.querySelectorAll("button:not(:disabled)") ?? []];
				if (controls.length === 0) return;
				event.preventDefault();
				const current = controls.indexOf(document.activeElement);
				controls[event.key === "Home" ? 0 : event.key === "End" ? controls.length - 1 : event.key === "ArrowUp" ? (current <= 0 ? controls.length : current) - 1 : (current + 1) % controls.length]?.focus();
			};
			return /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
				ref: rootRef,
				className: "fylar-office-action-menu-wrap",
				children: [/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("button", {
					ref: triggerRef,
					type: "button",
					className: "fylar-office-action-trigger",
					"data-variant": variant,
					"aria-label": label,
					"aria-haspopup": "menu",
					"aria-expanded": open,
					"aria-controls": open ? menuId : void 0,
					disabled,
					onClick: () => {
						setOpen((value) => !value);
					},
					children: [
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)(MenuIcon, { icon }),
						variant === "text" ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", { children: label }) : null,
						variant === "text" ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("svg", {
							className: "fylar-office-action-chevron",
							viewBox: "0 0 12 12",
							width: "12",
							height: "12",
							"aria-hidden": true,
							children: /* @__PURE__ */ (0, react_jsx_runtime.jsx)("path", {
								d: "m3 4.5 3 3 3-3",
								fill: "none",
								stroke: "currentColor",
								strokeWidth: "1.3",
								strokeLinecap: "round",
								strokeLinejoin: "round"
							})
						}) : null
					]
				}), open ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("div", {
					ref: menuRef,
					id: menuId,
					className: "fylar-office-action-menu",
					role: "menu",
					"aria-label": label,
					onKeyDown: navigateMenu,
					children: items.map((item) => /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("button", {
						type: "button",
						role: "menuitem",
						"data-danger": item.danger === true ? "true" : void 0,
						disabled: item.disabled,
						title: item.disabled ? item.description : void 0,
						onClick: () => {
							setOpen(false);
							item.onSelect();
						},
						children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", { children: item.label }), item.description === void 0 ? null : /* @__PURE__ */ (0, react_jsx_runtime.jsx)("small", { children: item.description })]
					}, item.id))
				}) : null]
			});
		}
		//#endregion
		//#region src/client/OfficeFileIcon.tsx
		function OfficeFileIcon({ descriptor, compact = false }) {
			const letter = descriptor.docType === 1 ? "W" : descriptor.docType === 2 ? "X" : "P";
			return /* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", {
				className: "fylar-office-artifact-icon",
				"data-accent": descriptor.accent,
				"data-compact": compact ? "true" : void 0,
				"aria-hidden": true,
				children: /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("svg", {
					viewBox: "0 0 24 24",
					width: "22",
					height: "22",
					children: [
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("path", {
							d: "M6 2.75h8l4 4V21.25H6z",
							fill: "none",
							stroke: "currentColor",
							strokeWidth: "1.5"
						}),
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("path", {
							d: "M14 2.75v4h4",
							fill: "none",
							stroke: "currentColor",
							strokeWidth: "1.5"
						}),
						/* @__PURE__ */ (0, react_jsx_runtime.jsx)("text", {
							x: "8",
							y: "17",
							fill: "currentColor",
							fontSize: "8",
							fontWeight: "700",
							children: letter
						})
					]
				})
			});
		}
		//#endregion
		//#region src/client/OfficeEditorSurface.tsx
		/** Shared Office editor chrome for the details column and inline Chat card. */
		/** Render one live SDK host; callers decide whether it occupies Chat or details. */
		function OfficeEditorSurface({ snapshot, current, presentation, attach, detach, onRetry, setReadOnly, exportDocument, exportPdf, clearIssue, markPotentialEdits, onClose, onCollapse, onOpenDetails, t = zhT }) {
			const surfaceRef = (0, react.useRef)(null);
			const viewerRef = (0, react.useRef)(null);
			const [fullscreen, setFullscreen] = (0, react.useState)(false);
			const fullscreenSupported = typeof Element !== "undefined" && typeof Element.prototype.requestFullscreen === "function" && typeof document.exitFullscreen === "function";
			(0, react.useEffect)(() => {
				const container = viewerRef.current;
				if (container === null) return;
				attach(container);
				return () => {
					detach(container);
				};
			}, [attach, detach]);
			(0, react.useEffect)(() => {
				const container = viewerRef.current;
				if (container === null || snapshot.status !== "ready" || snapshot.readOnly) return;
				const mark = () => {
					markPotentialEdits();
				};
				container.addEventListener("pointerdown", mark, true);
				container.addEventListener("keydown", mark, true);
				container.addEventListener("input", mark, true);
				return () => {
					container.removeEventListener("pointerdown", mark, true);
					container.removeEventListener("keydown", mark, true);
					container.removeEventListener("input", mark, true);
				};
			}, [markPotentialEdits, snapshot]);
			(0, react.useEffect)(() => {
				const surface = surfaceRef.current;
				if (presentation !== "inline" || snapshot.status !== "ready" || surface === null || typeof surface.scrollIntoView !== "function") return;
				surface.scrollIntoView({
					block: "start",
					inline: "nearest",
					behavior: "smooth"
				});
			}, [presentation, snapshot.status]);
			(0, react.useEffect)(() => {
				const update = () => {
					setFullscreen(document.fullscreenElement === surfaceRef.current);
				};
				document.addEventListener("fullscreenchange", update);
				return () => {
					document.removeEventListener("fullscreenchange", update);
				};
			}, []);
			const toggleFullscreen = () => {
				const surface = surfaceRef.current;
				if (surface === null || !fullscreenSupported) return;
				(document.fullscreenElement === surface ? document.exitFullscreen() : surface.requestFullscreen()).catch((error) => {
					console.error("[dsh-fylar-office-editor] fullscreen transition failed", error);
				});
			};
			const ready = snapshot.status === "ready";
			const busy = snapshot.status === "loading" || ready && snapshot.operation !== null;
			const issue = snapshot.status === "error" ? snapshot.issue : ready ? snapshot.issue : void 0;
			const localizedIssue = issue === void 0 ? void 0 : officeIssueCopy(t, issue.kind);
			const fileName = ready ? snapshot.fileName : current?.source.fileName ?? t("document.fallback");
			const descriptor = describeOfficeDocument(fileName, ready ? snapshot.docType : current?.source.docType ?? 1);
			const pdfProblemCode = ready && descriptor.supportsPdf ? pdfEnvironmentProblem() : null;
			const pdfProblem = pdfProblemCode === "insecure-context" ? t("pdf.insecureContext") : pdfProblemCode === "local-fonts-unavailable" ? t("pdf.localFontsUnavailable") : null;
			const modeLabel = snapshot.status === "error" ? t("mode.openFailed") : !ready ? t("mode.preparing") : snapshot.operation === "readonly" ? t("mode.switching") : snapshot.readOnly ? t("mode.readOnly") : snapshot.hasPotentialEdits ? t("mode.dirty") : t("mode.editing");
			const modeTone = snapshot.status === "error" ? "error" : !ready ? "loading" : snapshot.readOnly ? "readonly" : snapshot.hasPotentialEdits ? "dirty" : "editing";
			const modeHelp = ready && snapshot.readOnly ? t("help.readOnly") : t("help.editing");
			const exportLabel = ready && snapshot.operation === "export" ? t("action.downloading") : ready && snapshot.operation === "pdf" ? t("action.exporting") : t("action.export");
			const exportItems = ready ? [{
				id: "office",
				label: t("action.download", { extension: descriptor.extension }),
				description: t("action.downloadDescription"),
				onSelect: () => {
					exportDocument();
				}
			}] : [];
			if (ready && descriptor.supportsPdf) exportItems.push({
				id: "pdf",
				label: t("action.exportPdf"),
				description: pdfProblem ?? t("action.exportPdfDescription"),
				disabled: pdfProblem !== null,
				onSelect: () => {
					exportPdf();
				}
			});
			const moreItems = [];
			if (ready) moreItems.push({
				id: "readonly",
				label: snapshot.readOnly ? t("action.enableEditing") : t("action.readOnly"),
				description: snapshot.readOnly ? t("action.enableEditingDescription") : t("action.readOnlyDescription"),
				onSelect: () => {
					setReadOnly(!snapshot.readOnly);
				}
			});
			if (onOpenDetails !== void 0) moreItems.push({
				id: "details",
				label: t("action.openDetails"),
				onSelect: onOpenDetails
			});
			const directClose = presentation === "details";
			if (!directClose) moreItems.push({
				id: "close",
				label: t("action.closeSession"),
				description: t("action.closeSessionDescription"),
				danger: true,
				onSelect: onClose
			});
			return /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("section", {
				ref: surfaceRef,
				className: `fylar-office-details fylar-office-editor-${presentation}`,
				"data-fullscreen": fullscreen ? "true" : void 0,
				"aria-label": presentation === "inline" ? t("editor.inlineAria") : t("editor.aria"),
				children: [
					/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("header", {
						className: "fylar-office-details-header",
						children: [/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
							className: "fylar-office-document-identity",
							children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)(OfficeFileIcon, {
								descriptor,
								compact: true
							}), /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
								className: "fylar-office-details-heading",
								children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("strong", { children: fileName }), /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("span", {
									className: "fylar-office-document-meta",
									children: [
										/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", {
											className: "fylar-office-document-product",
											children: t("document.kind", { product: descriptor.product })
										}),
										/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", {
											className: "fylar-office-document-separator",
											"aria-hidden": true,
											children: "·"
										}),
										/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("span", {
											className: "fylar-office-mode-badge",
											"data-tone": modeTone,
											title: modeHelp,
											"aria-label": `${modeLabel}: ${modeHelp}`,
											children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", { "aria-hidden": true }), modeLabel]
										})
									]
								})]
							})]
						}), /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
							className: "fylar-office-details-header-actions",
							children: [
								/* @__PURE__ */ (0, react_jsx_runtime.jsx)(OfficeActionMenu, {
									label: exportLabel,
									icon: "export",
									variant: "text",
									disabled: !ready || busy,
									items: exportItems
								}),
								/* @__PURE__ */ (0, react_jsx_runtime.jsx)("button", {
									type: "button",
									className: "fylar-office-details-close",
									disabled: busy || !fullscreenSupported,
									"aria-label": fullscreen ? t("action.exitFullscreen") : t("action.fullscreen"),
									title: fullscreenSupported ? fullscreen ? t("action.exitFullscreen") : t("action.fullscreen") : t("action.fullscreenUnsupported"),
									onClick: toggleFullscreen,
									children: fullscreen ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("svg", {
										viewBox: "0 0 16 16",
										width: "16",
										height: "16",
										"aria-hidden": true,
										children: /* @__PURE__ */ (0, react_jsx_runtime.jsx)("path", {
											d: "M6.5 2.5v4h-4M9.5 2.5v4h4M6.5 13.5v-4h-4M9.5 13.5v-4h4",
											fill: "none",
											stroke: "currentColor",
											strokeWidth: "1.4",
											strokeLinecap: "round",
											strokeLinejoin: "round"
										})
									}) : /* @__PURE__ */ (0, react_jsx_runtime.jsx)("svg", {
										viewBox: "0 0 16 16",
										width: "16",
										height: "16",
										"aria-hidden": true,
										children: /* @__PURE__ */ (0, react_jsx_runtime.jsx)("path", {
											d: "M6 2.5H2.5V6M10 2.5h3.5V6M6 13.5H2.5V10M10 13.5h3.5V10",
											fill: "none",
											stroke: "currentColor",
											strokeWidth: "1.4",
											strokeLinecap: "round",
											strokeLinejoin: "round"
										})
									})
								}),
								onCollapse === void 0 ? null : /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("button", {
									type: "button",
									className: "fylar-office-details-close",
									disabled: busy,
									"aria-label": t("action.collapseAria"),
									onClick: onCollapse,
									children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("svg", {
										viewBox: "0 0 16 16",
										width: "15",
										height: "15",
										"aria-hidden": true,
										children: /* @__PURE__ */ (0, react_jsx_runtime.jsx)("path", {
											d: "M3.5 10l4.5-4 4.5 4",
											fill: "none",
											stroke: "currentColor",
											strokeWidth: "1.5",
											strokeLinecap: "round",
											strokeLinejoin: "round"
										})
									}), /* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", { children: t("action.collapse") })]
								}),
								moreItems.length === 0 ? null : /* @__PURE__ */ (0, react_jsx_runtime.jsx)(OfficeActionMenu, {
									label: t("action.moreDocument"),
									icon: "more",
									disabled: snapshot.status === "loading",
									items: moreItems
								}),
								directClose ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("button", {
									type: "button",
									className: "fylar-office-details-close",
									disabled: busy,
									"aria-label": t("action.closeDocument"),
									title: t("action.closeDocument"),
									onClick: onClose,
									children: /* @__PURE__ */ (0, react_jsx_runtime.jsx)("svg", {
										viewBox: "0 0 16 16",
										width: "16",
										height: "16",
										"aria-hidden": true,
										children: /* @__PURE__ */ (0, react_jsx_runtime.jsx)("path", {
											d: "M4 4l8 8M12 4l-8 8",
											fill: "none",
											stroke: "currentColor",
											strokeWidth: "1.5",
											strokeLinecap: "round"
										})
									})
								}) : null
							]
						})]
					}),
					issue !== void 0 ? /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
						className: "fylar-office-details-error",
						role: "alert",
						children: [
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("strong", { children: localizedIssue?.title }),
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", { children: localizedIssue?.detail }),
							/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", { children: [issue.retryable ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("button", {
								type: "button",
								onClick: () => void onRetry(),
								children: t("action.retry")
							}) : null, ready ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("button", {
								type: "button",
								onClick: clearIssue,
								children: t("action.dismiss")
							}) : null] })
						]
					}) : null,
					/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
						className: "fylar-office-details-body",
						children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("div", {
							ref: viewerRef,
							className: "fylar-office-details-viewer"
						}), snapshot.status === "loading" ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("div", {
							className: "fylar-office-details-loading",
							role: "status",
							children: t("editor.opening", { name: snapshot.label })
						}) : null]
					})
				]
			});
		}
		//#endregion
		//#region src/client/office-deliverables-model.ts
		/** Return only file outputs that existed when the closing message was rendered. */
		function officeFilesForClosing(data, seq) {
			if (data === void 0) return [];
			return data.produced.filter((file) => file.seq <= seq).map((file) => file.path);
		}
		//#endregion
		//#region src/client/OfficeDeliverables.tsx
		/** Conversation Definition kind and the Turn-data key it exclusively owns. */
		const OFFICE_DELIVERABLES_KIND = "fylar-office-deliverables";
		const OFFICE_DELIVERY_TOOL_NAMES = /* @__PURE__ */ new Set(["office_create_docx", "office_present_file"]);
		function pathFromArguments(raw) {
			if (typeof raw !== "string") return void 0;
			try {
				const value = JSON.parse(raw);
				if (typeof value !== "object" || value === null) return void 0;
				const path = value.file_path;
				return typeof path === "string" && path !== "" ? path : void 0;
			} catch {
				return;
			}
		}
		/** Accumulate successful Office creation calls into their engine-owned Turn. */
		const officeDeliverablesDefinition = {
			kind: OFFICE_DELIVERABLES_KIND,
			match: (event) => {
				if (event.type === "turn/start") return {
					id: String(event.data.turn),
					role: "start"
				};
				if (event.type === "tool/call" && OFFICE_DELIVERY_TOOL_NAMES.has(event.data.name)) return {
					id: String(event.data.turn),
					role: "update"
				};
				if (event.type === "tool/result" && isAppendSurfaceEvent(event)) return {
					id: String(event.data.turn),
					role: "update"
				};
				return null;
			},
			start: (_context, match) => {
				if (match.event.type !== "turn/start") throw new Error("Office deliverables require turn/start");
				return {
					turn: match.event.data.turn,
					calls: /* @__PURE__ */ new Map(),
					produced: []
				};
			},
			update: (context, match) => {
				if (match.event.type === "tool/call") {
					const path = pathFromArguments(match.event.data.arguments);
					if (path === void 0) return context.state;
					const calls = new Map(context.state.calls);
					calls.set(String(match.event.data.callId), path);
					return {
						...context.state,
						calls
					};
				}
				if (match.event.type !== "tool/result") return context.state;
				if (match.event.data.message.content[0]?.isError === true) return context.state;
				const path = context.state.calls.get(String(match.event.data.message.source.callId));
				if (path === void 0 || context.state.produced.some((item) => item.path === path)) return context.state;
				return {
					...context.state,
					produced: [...context.state.produced, {
						path,
						seq: match.event.seq
					}]
				};
			},
			buildLocationData: (context, scope) => scope !== "turn" || context.state === void 0 ? null : {
				kind: "turn",
				turn: context.state.turn,
				key: OFFICE_DELIVERABLES_KIND,
				value: { produced: context.state.produced }
			}
		};
		function selectOfficeDeliverables(owner) {
			const paths = officeFilesForClosing(owner.turn.data.get(OFFICE_DELIVERABLES_KIND), owner.seq);
			return paths.length === 0 ? null : paths;
		}
		function isCurrentArtifact(current, path, cwd) {
			return current?.source.kind === "workspace" && current.source.path === path && current.source.cwd === cwd;
		}
		function ArtifactCard({ path, cwd, workspace, current, openInline, openDetails, download, close, t }) {
			const descriptor = describeOfficeDocument(path);
			const active = isCurrentArtifact(current, path, cwd);
			const busy = workspace.status === "loading" || workspace.status === "ready" && workspace.operation !== null;
			const status = cwd === void 0 ? t("artifact.workspaceUnavailable") : active && current?.presentation === "parked" ? t("artifact.parked") : active && current?.presentation === "details" ? t("artifact.details") : active && current?.presentation === "office-view" ? t("artifact.officeView") : t("artifact.openInline", { product: descriptor.product });
			const menuItems = [{
				id: "details",
				label: t("action.openDetails"),
				onSelect: openDetails
			}];
			if (active && workspace.status === "ready") menuItems.push({
				id: "download",
				label: t("action.download", { extension: descriptor.extension }),
				onSelect: () => {
					download();
				}
			});
			if (active) menuItems.push({
				id: "close",
				label: t("action.closeSession"),
				danger: true,
				onSelect: close
			});
			return /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("article", {
				className: "fylar-office-artifact-card",
				children: [/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("button", {
					type: "button",
					className: "fylar-office-artifact-primary",
					disabled: cwd === void 0 || busy,
					onClick: openInline,
					children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)(OfficeFileIcon, { descriptor }), /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("span", {
						className: "fylar-office-artifact-copy",
						children: [
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("strong", { children: officeBaseName(path) }),
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", { children: status }),
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("small", { children: officeDirectoryLabel(path, t("workspace.root")) })
						]
					})]
				}), /* @__PURE__ */ (0, react_jsx_runtime.jsx)(OfficeActionMenu, {
					label: t("artifact.more", { name: officeBaseName(path) }),
					icon: "more",
					disabled: cwd === void 0 || busy,
					items: menuItems
				})]
			});
		}
		/** Render successful files after the closing assistant response. */
		function OfficeDeliverables({ matched, useSessions, sessionId, useWorkspace, useOfficeSession, openDetails, openInlineArtifact, collapseInline, closeDocument, retrySession, attach, detach, setReadOnly, exportDocument, exportPdf, clearIssue, markPotentialEdits, t = zhT }) {
			const cwd = useSessions((list) => list.byId[sessionId]?.cwd);
			const workspace = useWorkspace((value) => value);
			const session = useOfficeSession((value) => value);
			return /* @__PURE__ */ (0, react_jsx_runtime.jsx)("div", {
				className: "fylar-office-deliverables",
				"aria-label": t("artifact.group"),
				children: matched.map((path) => {
					const target = {
						path,
						cwd
					};
					const current = session.current;
					return current?.presentation === "inline" && isCurrentArtifact(current, path, cwd) ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)(OfficeEditorSurface, {
						snapshot: workspace,
						current,
						presentation: "inline",
						attach,
						detach,
						onRetry: retrySession,
						setReadOnly,
						exportDocument,
						exportPdf,
						clearIssue,
						markPotentialEdits,
						t,
						onOpenDetails: () => {
							openDetails(target);
						},
						onCollapse: collapseInline,
						onClose: closeDocument
					}, path) : /* @__PURE__ */ (0, react_jsx_runtime.jsx)(ArtifactCard, {
						path,
						cwd,
						workspace,
						current,
						openInline: () => {
							openInlineArtifact(target);
						},
						openDetails: () => {
							openDetails(target);
						},
						download: exportDocument,
						close: closeDocument,
						t
					}, path);
				})
			});
		}
		//#endregion
		//#region src/client/OfficeDetailsPanel.tsx
		/** Fylar Office editor hosted in a DSH right-Sidebar tab. */
		/** Render the persistent SDK root plus compact edit/export controls. */
		function OfficeDetailsPanel({ useWorkspace, useOfficeSession, attach, detach, retrySession, setReadOnly, exportDocument, exportPdf, clearIssue, markPotentialEdits, closePanel, panelOpened, panelClosed, useTabInfo, t = zhT }) {
			const { tab } = useTabInfo();
			const snapshot = useWorkspace((value) => value);
			const current = useOfficeSession((value) => value.current);
			(0, react.useEffect)(() => {
				panelOpened(tab.id);
				const handleClose = () => {
					panelClosed(tab.id);
				};
				if (tab.signal.aborted) handleClose();
				else tab.signal.addEventListener("abort", handleClose, { once: true });
				return () => {
					tab.signal.removeEventListener("abort", handleClose);
				};
			}, [
				panelClosed,
				panelOpened,
				tab.id,
				tab.signal
			]);
			return /* @__PURE__ */ (0, react_jsx_runtime.jsx)(OfficeEditorSurface, {
				snapshot,
				current,
				presentation: "details",
				attach,
				detach,
				onRetry: retrySession,
				setReadOnly,
				exportDocument,
				exportPdf,
				clearIssue,
				markPotentialEdits,
				t,
				onClose: closePanel
			});
		}
		const OFFICE_SDK_BASE_PATH = `/fylar-office-sdk/1.0.0-f24b12f6b857`;
		//#endregion
		//#region src/client/office-sdk-loader.ts
		const dynamicImport = (url) => import(
			/* @vite-ignore */
			url
);
		/** Deduplicated same-origin loader for the unbundled Office SDK runtime. */
		var OfficeSdkLoader = class {
			basePath;
			importer;
			preloadPromise = null;
			sdkPromise = null;
			constructor(basePath = OFFICE_SDK_BASE_PATH, importer = dynamicImport) {
				this.basePath = basePath;
				this.importer = importer;
			}
			preload() {
				if (this.preloadPromise !== null) return this.preloadPromise;
				const request = this.importer(`${this.basePath}/preload.js`).then(() => void 0);
				this.preloadPromise = request.catch((error) => {
					this.preloadPromise = null;
					throw new OfficePluginFailure("runtime-load", "Office preload failed", { cause: error });
				});
				return this.preloadPromise;
			}
			load() {
				if (this.sdkPromise !== null) return this.sdkPromise;
				const request = this.importer(`${this.basePath}/UI.js`).then((value) => {
					const module = value;
					const sdk = module.default ?? (typeof module.openfile === "function" && typeof module.createfile === "function" ? {
						openfile: module.openfile,
						createfile: module.createfile
					} : void 0);
					if (sdk === void 0 || typeof sdk.openfile !== "function" || typeof sdk.createfile !== "function") throw new Error("UI.js does not expose openfile and createfile");
					return sdk;
				});
				this.sdkPromise = request.catch((error) => {
					this.sdkPromise = null;
					throw new OfficePluginFailure("runtime-load", "Office UI runtime failed to load", { cause: error });
				});
				return this.sdkPromise;
			}
			resetFailure() {
				this.preloadPromise = null;
				this.sdkPromise = null;
			}
		};
		//#endregion
		//#region src/client/office-workspace.ts
		function defaultEnvironment() {
			return {
				window,
				document,
				createResizeObserver: (callback) => typeof ResizeObserver === "undefined" ? null : new ResizeObserver(callback)
			};
		}
		function isDocType(value) {
			return value === 1 || value === 2 || value === 3;
		}
		/** Browser-wide owner of the single Office SDK instance and persistent DOM host. */
		var OfficeWorkspace = class {
			loader;
			snapshot = { status: "empty" };
			listeners = /* @__PURE__ */ new Set();
			environment;
			root;
			parking;
			observer;
			widget = null;
			app = null;
			generation = 0;
			queue = Promise.resolve();
			disposed = false;
			lastRetry = null;
			timers = [];
			frames = [];
			source = {
				getSnapshot: () => this.snapshot,
				subscribe: (listener) => {
					this.listeners.add(listener);
					return () => {
						this.listeners.delete(listener);
					};
				}
			};
			actions = {
				preload: () => this.preload(),
				retry: () => this.retry(),
				attach: (container) => {
					this.attach(container);
				},
				detach: (container) => {
					this.detach(container);
				},
				open: (file) => this.open(file),
				openUrl: (fileName, url) => this.openUrl(fileName, url),
				create: (docType) => this.create(docType),
				setReadOnly: (readOnly) => this.setReadOnly(readOnly),
				exportDocument: () => this.exportDocument(),
				exportPdf: () => this.exportPdf(),
				close: () => this.close(),
				clearIssue: () => {
					this.clearIssue();
				},
				markPotentialEdits: () => {
					this.markPotentialEdits();
				}
			};
			constructor(loader, environment = defaultEnvironment()) {
				this.loader = loader;
				this.environment = environment;
				this.root = environment.document.createElement("div");
				this.root.className = "fylar-office-sdk-host";
				this.root.dataset.fylarOfficeSdkHost = "true";
				this.parking = environment.document.createElement("div");
				this.parking.hidden = true;
				this.parking.dataset.fylarOfficeParking = "true";
				this.parking.append(this.root);
				(environment.document.body ?? environment.document.documentElement).append(this.parking);
				this.observer = environment.createResizeObserver(() => {
					this.refresh();
				});
				this.observer?.observe(this.root);
				environment.window.addEventListener("beforeunload", this.beforeUnload);
			}
			beforeUnload = (event) => {
				if (this.snapshot.status !== "ready" || !this.snapshot.hasPotentialEdits) return;
				event.preventDefault();
				event.returnValue = "";
			};
			publish(snapshot) {
				this.snapshot = snapshot;
				for (const listener of this.listeners) listener();
			}
			publishIssue(issue) {
				const current = this.ready();
				this.publish(current === null ? {
					status: "error",
					issue
				} : {
					...current,
					issue
				});
			}
			enqueue(operation) {
				const result = this.queue.then(operation, operation);
				this.queue = result.catch(() => void 0);
				return result;
			}
			assertCurrent(generation) {
				if (generation !== this.generation || this.disposed) throw new OfficePluginFailure("cancelled", "OPEN_CANCELLED");
			}
			clearScheduledRefresh() {
				for (const timer of this.timers) this.environment.window.clearTimeout(timer);
				for (const frame of this.frames) this.environment.window.cancelAnimationFrame?.(frame);
				this.timers = [];
				this.frames = [];
			}
			refresh() {
				this.app?.Document.updateVisibleArea();
			}
			scheduleRefresh() {
				this.clearScheduledRefresh();
				const run = () => {
					this.refresh();
				};
				if (typeof this.environment.window.requestAnimationFrame === "function") {
					this.frames.push(this.environment.window.requestAnimationFrame(run));
					this.frames.push(this.environment.window.requestAnimationFrame(() => {
						this.frames.push(this.environment.window.requestAnimationFrame(run));
					}));
				}
				this.timers.push(this.environment.window.setTimeout(run, 80));
				this.timers.push(this.environment.window.setTimeout(run, 220));
			}
			options() {
				return {
					uiOptions: {
						showTopBar: true,
						showBottomBar: true
					},
					mode: {
						readOnly: false,
						lang: resolveOfficeLocale(this.environment.document.documentElement.lang || this.environment.window.navigator.language)
					}
				};
			}
			async cleanupWidget() {
				this.clearScheduledRefresh();
				const current = this.widget;
				this.widget = null;
				this.app = null;
				try {
					await current?.close();
				} catch (error) {
					console.warn("[dsh-fylar-office-editor] widget cleanup failed", error);
				} finally {
					this.root.replaceChildren();
				}
			}
			async mount(action, label, origin, fallbackType, createWidget, retry) {
				const generation = ++this.generation;
				this.lastRetry = retry;
				this.publish({
					status: "loading",
					action,
					label
				});
				await this.enqueue(async () => {
					let openedWidget = null;
					try {
						this.assertCurrent(generation);
						await this.cleanupWidget();
						this.assertCurrent(generation);
						assertOfficeCapabilities(this.environment.window);
						await this.loader.preload();
						this.assertCurrent(generation);
						await this.loader.load();
						this.assertCurrent(generation);
						const result = await createWidget();
						openedWidget = result.widget;
						this.assertCurrent(generation);
						const rendered = await openedWidget.mount(this.root).render();
						this.assertCurrent(generation);
						this.widget = openedWidget;
						this.app = rendered;
						openedWidget = null;
						const docType = isDocType(result.docType) ? result.docType : fallbackType;
						this.publish({
							status: "ready",
							fileName: label,
							docType,
							origin,
							readOnly: rendered.Document.isReadOnly(),
							operation: null,
							hasPotentialEdits: false
						});
						this.scheduleRefresh();
					} catch (error) {
						await openedWidget?.close().catch((cleanupError) => {
							console.warn("[dsh-fylar-office-editor] cancelled widget cleanup failed", cleanupError);
						});
						const issue = classifyOfficeError(error, action);
						if (issue.kind !== "cancelled" && generation === this.generation && !this.disposed) {
							console.error(`[dsh-fylar-office-editor] ${action} failed`, error);
							this.publish({
								status: "error",
								issue
							});
						}
					}
				});
			}
			async preload() {
				if (this.disposed || this.widget !== null) return;
				try {
					assertOfficeCapabilities(this.environment.window);
					if (this.snapshot.status === "empty") this.publish({
						status: "loading",
						action: "preload",
						label: "Office"
					});
					await this.loader.preload();
					if (this.snapshot.status === "loading" && this.snapshot.action === "preload") this.publish({ status: "empty" });
				} catch (error) {
					console.error("[dsh-fylar-office-editor] preload failed", error);
					if (this.widget === null) this.publish({
						status: "error",
						issue: classifyOfficeError(error, "runtime-load")
					});
				}
			}
			async retry() {
				this.loader.resetFailure();
				if (this.lastRetry !== null) await this.lastRetry();
				else await this.preload();
			}
			attach(container) {
				if (this.disposed) return;
				container.append(this.root);
				this.scheduleRefresh();
			}
			detach(container) {
				if (this.disposed || this.root.parentElement !== container) return;
				this.parking.append(this.root);
			}
			async open(file) {
				try {
					assertSupportedOfficeFile(file);
				} catch (error) {
					this.publishIssue(classifyOfficeError(error, "unsupported-file"));
					return;
				}
				const retry = () => this.open(file);
				await this.mount("open", file.name, "opened", inferOfficeDocType(file.name), async () => (await this.loader.load()).openfile({
					docId: `local-${Date.now()}`,
					fileName: file.name,
					file
				}, this.options()), retry);
			}
			/** Fetch a workspace Office file from the same-origin Host route and open it. */
			async openUrl(fileName, url) {
				try {
					assertSupportedOfficeFile({ name: fileName });
				} catch (error) {
					this.publishIssue(classifyOfficeError(error, "unsupported-file"));
					return;
				}
				const retry = () => this.openUrl(fileName, url);
				await this.mount("open", fileName, "workspace", inferOfficeDocType(fileName), async () => {
					const response = await this.environment.window.fetch(url, { credentials: "same-origin" });
					if (!response.ok) throw new OfficePluginFailure("asset", `WORKSPACE_FILE_${response.status}`);
					const file = await response.blob();
					return (await this.loader.load()).openfile({
						docId: `workspace-${encodeURIComponent(url)}`,
						fileName,
						file
					}, this.options());
				}, retry);
			}
			async create(docType) {
				const fileName = createdFileName(docType);
				const retry = () => this.create(docType);
				await this.mount("create", fileName, "created", docType, async () => (await this.loader.load()).createfile(docType, this.options()), retry);
			}
			ready() {
				return this.snapshot.status === "ready" ? this.snapshot : null;
			}
			async readyOperation(operation, fallback, run, complete) {
				const ready = this.ready();
				const app = this.app;
				if (ready === null || app === null || ready.operation !== null) return false;
				this.publish({
					...ready,
					operation,
					issue: void 0
				});
				try {
					const result = await run(app);
					const current = this.ready();
					if (current === null || this.app !== app) return false;
					const completed = {
						...current,
						operation: null,
						issue: void 0
					};
					this.publish(complete?.(completed, result) ?? completed);
					return true;
				} catch (error) {
					console.error(`[dsh-fylar-office-editor] ${operation} failed`, error);
					const current = this.ready();
					if (current !== null && this.app === app) this.publish({
						...current,
						operation: null,
						issue: classifyOfficeError(error, fallback)
					});
					return false;
				}
			}
			async setReadOnly(readOnly) {
				await this.readyOperation("readonly", "unknown", (app) => {
					if (readOnly) app.Document.startReadOnly();
					else app.Document.endReadOnly();
					return app.Document.isReadOnly();
				}, (current, actualReadOnly) => ({
					...current,
					readOnly: actualReadOnly,
					hasPotentialEdits: current.hasPotentialEdits || !readOnly
				}));
			}
			async exportDocument() {
				await this.readyOperation("export", "export", (app) => app.Document.exportDocument(), (current) => ({
					...current,
					hasPotentialEdits: false
				}));
			}
			async exportPdf() {
				const ready = this.ready();
				if (ready === null || ready.docType === 3) return;
				if (pdfEnvironmentProblem(this.environment.window) !== null) {
					this.publish({
						...ready,
						issue: issueFor("pdf-environment")
					});
					return;
				}
				await this.readyOperation("pdf", "export", (app) => app.Document.exportPdf());
			}
			clearIssue() {
				const ready = this.ready();
				if (ready !== null && ready.issue !== void 0) {
					const { issue: _issue, ...withoutIssue } = ready;
					this.publish(withoutIssue);
				}
			}
			markPotentialEdits() {
				const current = this.ready();
				if (current === null || current.readOnly || current.hasPotentialEdits) return;
				this.publish({
					...current,
					hasPotentialEdits: true
				});
			}
			async close() {
				const generation = ++this.generation;
				this.lastRetry = null;
				const ready = this.ready();
				if (ready !== null) this.publish({
					...ready,
					operation: "close",
					issue: void 0
				});
				await this.enqueue(async () => {
					if (generation !== this.generation) return;
					await this.cleanupWidget();
					if (!this.disposed) this.publish({ status: "empty" });
				});
			}
			async dispose() {
				if (this.disposed) return;
				this.disposed = true;
				++this.generation;
				this.environment.window.removeEventListener("beforeunload", this.beforeUnload);
				this.observer?.disconnect();
				await this.enqueue(async () => {
					await this.cleanupWidget();
				});
				this.parking.remove();
				this.root.remove();
				this.listeners.clear();
			}
		};
		//#endregion
		//#region src/client/OfficeView.tsx
		function OfficeView({ useWorkspace, useOfficeSession, preload, attach, detach, setReadOnly, exportDocument, exportPdf, clearIssue, markPotentialEdits, openLocal, createDocument, retrySession, activateOfficeView, parkOfficeView, closeDocument, t = zhT }) {
			const snapshot = useWorkspace((value) => value);
			const current = useOfficeSession((value) => value.current);
			const inputRef = (0, react.useRef)(null);
			(0, react.useEffect)(() => {
				activateOfficeView();
				preload();
				return () => {
					parkOfficeView();
				};
			}, [
				activateOfficeView,
				parkOfficeView,
				preload
			]);
			const chooseFile = () => {
				inputRef.current?.click();
			};
			const onFileChange = (event) => {
				const file = event.currentTarget.files?.[0];
				event.currentTarget.value = "";
				if (file !== void 0) openLocal(file);
			};
			if (current !== null) return /* @__PURE__ */ (0, react_jsx_runtime.jsx)(OfficeEditorSurface, {
				snapshot,
				current,
				presentation: "office-view",
				attach,
				detach,
				onRetry: retrySession,
				setReadOnly,
				exportDocument,
				exportPdf,
				clearIssue,
				markPotentialEdits,
				t,
				onClose: () => {
					closeDocument();
				}
			});
			const issue = snapshot.status === "error" ? snapshot.issue : void 0;
			const issueCopy = issue === void 0 ? void 0 : officeIssueCopy(t, issue.kind);
			const busy = snapshot.status === "loading";
			return /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("section", {
				className: "fylar-office-view",
				"aria-label": "Fylar Office",
				children: [/* @__PURE__ */ (0, react_jsx_runtime.jsx)("input", {
					ref: inputRef,
					type: "file",
					accept: OFFICE_FILE_ACCEPT,
					hidden: true,
					onChange: onFileChange
				}), /* @__PURE__ */ (0, react_jsx_runtime.jsx)("div", {
					className: "fylar-office-empty",
					children: /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
						className: "fylar-office-empty-card",
						children: [
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("div", {
								className: "fylar-office-brand-mark",
								"aria-hidden": true,
								children: "FO"
							}),
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("h2", { children: "Fylar Office" }),
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("p", { children: t("empty.description") }),
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("small", { children: "Powered by Fylar Office SDK" }),
							busy ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("div", {
								className: "fylar-office-empty-status",
								role: "status",
								children: t("empty.loading")
							}) : null,
							issue === void 0 ? null : /* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
								className: "fylar-office-empty-error",
								role: "alert",
								children: [
									/* @__PURE__ */ (0, react_jsx_runtime.jsx)("strong", { children: issueCopy?.title }),
									/* @__PURE__ */ (0, react_jsx_runtime.jsx)("span", { children: issueCopy?.detail }),
									issue.retryable ? /* @__PURE__ */ (0, react_jsx_runtime.jsx)("button", {
										className: "fylar-office-button",
										onClick: () => void retrySession(),
										children: t("action.retry")
									}) : null
								]
							}),
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("div", {
								className: "fylar-office-actions",
								children: /* @__PURE__ */ (0, react_jsx_runtime.jsx)("button", {
									className: "fylar-office-button fylar-office-button-primary",
									disabled: busy,
									onClick: chooseFile,
									children: t("empty.openFile")
								})
							}),
							/* @__PURE__ */ (0, react_jsx_runtime.jsx)("div", {
								className: "fylar-office-create-label",
								children: t("empty.newDocument")
							}),
							/* @__PURE__ */ (0, react_jsx_runtime.jsxs)("div", {
								className: "fylar-office-actions",
								children: [
									/* @__PURE__ */ (0, react_jsx_runtime.jsx)("button", {
										className: "fylar-office-button",
										disabled: busy,
										onClick: () => void createDocument(1),
										children: t("empty.newWord")
									}),
									/* @__PURE__ */ (0, react_jsx_runtime.jsx)("button", {
										className: "fylar-office-button",
										disabled: busy,
										onClick: () => void createDocument(2),
										children: t("empty.newExcel")
									}),
									/* @__PURE__ */ (0, react_jsx_runtime.jsx)("button", {
										className: "fylar-office-button",
										disabled: busy,
										onClick: () => void createDocument(3),
										children: t("empty.newPowerPoint")
									})
								]
							})
						]
					})
				})]
			});
		}
		//#endregion
		//#region src/client/styles.ts
		/** Compose and install plugin-owned CSS without emitting a separate runtime asset. */
		const OFFICE_STYLES = [
			`
.fylar-office-view {
  position: relative;
  display: flex;
  flex-direction: column;
  min-height: 0;
  height: 100%;
  background: var(--dsw-alias-bg-base, #f7f7f5);
  color: var(--dsw-alias-fg-default, #1f2328);
}

.fylar-office-button {
  appearance: none;
  border: 1px solid var(--dsw-alias-border-default, #c8cdd3);
  border-radius: 7px;
  background: var(--dsw-alias-bg-elevated, #fff);
  color: inherit;
  min-height: 44px;
  padding: 7px 11px;
  font: inherit;
  cursor: pointer;
}

.fylar-office-button:hover:not(:disabled) {
  background: var(--dsw-alias-bg-hover, #f0f2f4);
}

.fylar-office-button:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}

.fylar-office-button-primary {
  background: #1769e0;
  color: #fff;
  border-color: #1769e0;
}

.fylar-office-button:focus-visible {
  outline: 2px solid var(--dsw-alias-brand-primary, #1769e0);
  outline-offset: 2px;
}

.fylar-office-sdk-host {
  width: 100%;
  height: 100%;
  min-height: 0;
  overflow: hidden;
}

.fylar-office-empty {
  position: absolute;
  inset: 0;
  display: grid;
  place-items: center;
  padding: 32px;
}

.fylar-office-empty-card {
  width: min(560px, 100%);
  padding: 30px;
  border: 1px solid var(--dsw-alias-border-subtle, #d8dee4);
  border-radius: 14px;
  background: var(--dsw-alias-bg-elevated, #fff);
  text-align: center;
  box-shadow: 0 8px 30px rgba(0, 0, 0, 0.06);
}

.fylar-office-brand-mark {
  display: grid;
  width: 48px;
  height: 48px;
  margin: 0 auto 12px;
  place-items: center;
  border-radius: 12px;
  background: var(--dsw-alias-brand-primary, #1769e0);
  color: #fff;
  font-weight: 700;
}

.fylar-office-empty h2 {
  margin: 0 0 8px;
  font-size: 24px;
}

.fylar-office-empty p {
  margin: 0 0 22px;
  color: var(--dsw-alias-fg-muted, #667085);
}

.fylar-office-empty-card > small {
  display: block;
  margin: -12px 0 22px;
  color: var(--dsw-alias-fg-subtle, #98a2b3);
}

.fylar-office-empty-error {
  display: flex;
  flex-direction: column;
  gap: 4px;
  margin: 0 0 16px;
  padding: 10px;
  border-radius: 8px;
  background: color-mix(in srgb, #d92d20 10%, var(--dsw-alias-bg-base, #fff));
  text-align: left;
}

.fylar-office-empty-status {
  margin: 0 0 16px;
  color: var(--dsw-alias-fg-muted, #667085);
  font-size: 13px;
}

.fylar-office-actions {
  display: flex;
  justify-content: center;
  gap: 8px;
  flex-wrap: wrap;
}

.fylar-office-create-label {
  margin: 22px 0 8px;
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: var(--dsw-alias-fg-muted, #667085);
}

@media (max-width: 639px) {
  .fylar-office-empty { padding: 16px; }
  .fylar-office-empty-card { padding: 22px 16px; border-radius: 10px; box-shadow: none; }
}
`,
			`
.fylar-office-artifact-icon {
  display: grid;
  flex: 0 0 38px;
  width: 38px;
  height: 38px;
  place-items: center;
  border-radius: 9px;
  background: color-mix(in srgb, #1769e0 11%, transparent);
  color: #1769e0;
}

.fylar-office-artifact-icon[data-compact='true'] {
  flex-basis: 32px;
  width: 32px;
  height: 32px;
  border-radius: 7px;
}

.fylar-office-artifact-icon[data-compact='true'] svg {
  width: 19px;
  height: 19px;
}

.fylar-office-artifact-icon[data-accent='excel'] {
  background: color-mix(in srgb, #178a52 11%, transparent);
  color: #178a52;
}

.fylar-office-artifact-icon[data-accent='powerpoint'] {
  background: color-mix(in srgb, #d6572a 11%, transparent);
  color: #d6572a;
}

.fylar-office-action-menu-wrap {
  position: relative;
  display: flex;
  flex: 0 0 auto;
  align-items: center;
  justify-content: center;
}

.fylar-office-action-trigger {
  display: inline-flex;
  min-width: 44px;
  height: 44px;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 0 9px;
  border: 0;
  border-radius: 8px;
  background: transparent;
  color: var(--dsw-alias-fg-default, #344054);
  font: inherit;
  font-size: 13px;
  line-height: 1;
  cursor: pointer;
}

.fylar-office-action-trigger[data-variant='text'] {
  min-width: auto;
  padding-inline: 10px 8px;
}

.fylar-office-action-trigger:hover:not(:disabled),
.fylar-office-action-trigger[aria-expanded='true'] {
  background: var(--dsw-alias-bg-hover, #f1f3f5);
}

.fylar-office-action-trigger:disabled {
  cursor: not-allowed;
  opacity: 0.45;
}

.fylar-office-action-chevron {
  color: var(--dsw-alias-fg-muted, #667085);
}

.fylar-office-action-menu {
  position: absolute;
  z-index: 40;
  top: calc(100% + 4px);
  right: 0;
  display: flex;
  width: max-content;
  min-width: 220px;
  max-width: min(280px, calc(100vw - 24px));
  flex-direction: column;
  gap: 2px;
  padding: 6px;
  border: 1px solid var(--dsw-alias-border-subtle, #d8dee4);
  border-radius: 10px;
  background: var(--dsw-alias-bg-elevated, #fff);
  box-shadow: 0 12px 32px color-mix(in srgb, currentColor 15%, transparent);
}

.fylar-office-action-menu button {
  display: flex;
  min-height: 44px;
  flex-direction: column;
  justify-content: center;
  gap: 2px;
  padding: 7px 10px;
  border: 0;
  border-radius: 7px;
  background: transparent;
  color: inherit;
  text-align: left;
  font: inherit;
  font-size: 13px;
  cursor: pointer;
}

.fylar-office-action-menu button:hover:not(:disabled) {
  background: var(--dsw-alias-bg-hover, #eef1f4);
}

.fylar-office-action-menu button:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}

.fylar-office-action-menu button[data-danger='true'] {
  color: var(--dsw-alias-fg-danger, #b42318);
}

.fylar-office-action-menu button small {
  max-width: 240px;
  color: var(--dsw-alias-fg-muted, #667085);
  font-size: 11px;
  font-weight: 400;
  line-height: 1.35;
  white-space: normal;
}

.fylar-office-action-trigger:focus-visible,
.fylar-office-action-menu button:focus-visible {
  outline: 2px solid var(--dsw-alias-brand-primary, #1769e0);
  outline-offset: 2px;
}

@media (max-width: 639px) {
  .fylar-office-action-menu {
    max-width: min(260px, calc(100vw - 16px));
  }
}
`,
			`
.fylar-office-tool-status {
  display: flex;
  width: min(520px, 100%);
  min-height: 36px;
  align-items: center;
  gap: 8px;
  margin: 6px 0;
  color: var(--dsw-alias-fg-muted, #667085);
  font-size: 13px;
}

.fylar-office-tool-status small {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.fylar-office-tool-status-dot {
  width: 8px;
  height: 8px;
  flex: 0 0 auto;
  border-radius: 50%;
  background: var(--dsw-alias-fg-muted, #667085);
}

.fylar-office-tool-status[data-state='running'] .fylar-office-tool-status-dot {
  background: var(--dsw-alias-brand-primary, #1769e0);
  animation: fylar-office-pulse 1.2s ease-in-out infinite;
}

.fylar-office-tool-status[data-state='ready'] .fylar-office-tool-status-dot { background: #178a52; }
.fylar-office-tool-status[data-state='error'] .fylar-office-tool-status-dot { background: #d92d20; }

@keyframes fylar-office-pulse {
  50% { opacity: 0.35; }
}

.fylar-office-artifact-card {
  display: flex;
  width: min(560px, 100%);
  align-items: stretch;
  margin: 8px 0;
  border: 1px solid var(--dsw-alias-border-subtle, #d8dee4);
  border-radius: 12px;
  background: var(--dsw-alias-bg-elevated, #fff);
  color: inherit;
  box-shadow: 0 1px 2px color-mix(in srgb, currentColor 5%, transparent);
}

.fylar-office-artifact-primary {
  display: flex;
  min-width: 0;
  flex: 1;
  align-items: center;
  gap: 12px;
  padding: 13px 14px;
  border: 0;
  border-radius: 11px 0 0 11px;
  background: transparent;
  color: inherit;
  text-align: left;
  font: inherit;
  cursor: pointer;
}

.fylar-office-artifact-primary:hover:not(:disabled) {
  background: var(--dsw-alias-bg-hover, #f8fafc);
}

.fylar-office-artifact-primary:disabled { cursor: not-allowed; opacity: 0.55; }

.fylar-office-artifact-copy {
  display: flex;
  min-width: 0;
  flex: 1;
  flex-direction: column;
  gap: 2px;
}

.fylar-office-artifact-copy strong,
.fylar-office-artifact-copy span,
.fylar-office-artifact-copy small {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.fylar-office-artifact-copy strong { font-size: 14px; }
.fylar-office-artifact-copy span { color: var(--dsw-alias-fg-muted, #667085); font-size: 12px; }
.fylar-office-artifact-copy small { color: var(--dsw-alias-fg-subtle, #98a2b3); font-size: 11px; }

.fylar-office-artifact-card > .fylar-office-action-menu-wrap {
  flex: 0 0 52px;
}

.fylar-office-artifact-card > .fylar-office-action-menu-wrap .fylar-office-action-menu {
  right: 6px;
}

.fylar-office-deliverables {
  display: flex;
  width: 100%;
  min-width: 0;
  flex-direction: column;
  align-items: stretch;
  margin-top: 6px;
}

.fylar-office-artifact-primary:focus-visible,
.fylar-office-artifact-card .fylar-office-action-trigger:focus-visible {
  outline: 2px solid var(--dsw-alias-brand-primary, #1769e0);
  outline-offset: 2px;
}
`,
			`
.fylar-office-details {
  position: relative;
  display: flex;
  height: 100%;
  min-height: 0;
  flex-direction: column;
  isolation: isolate;
  z-index: 0;
  container-type: inline-size;
  background: var(--dsw-alias-bg-base, #fff);
  color: var(--dsw-alias-fg-default, #1f2328);
}

.fylar-office-details:fullscreen {
  width: 100%;
  max-width: none;
  height: 100%;
  min-height: 0;
  margin: 0;
  border: 0;
  border-radius: 0;
  background: var(--dsw-alias-bg-base, #fff);
  box-shadow: none;
}

.fylar-office-details:fullscreen::backdrop {
  background: var(--dsw-alias-bg-base, #fff);
}

.fylar-office-details:fullscreen .fylar-office-details-header {
  border-radius: 0;
}

.fylar-office-editor-inline {
  --fylar-office-inline-width: min(
    900px,
    calc(var(--dsh-conversation-column-width, 964px) - 64px)
  );

  width: var(--fylar-office-inline-width);
  max-width: none;
  height: clamp(
    420px,
    calc(
      var(--dsh-conversation-viewport-height, 100dvh) -
      var(--dsh-composer-height, 152px) -
      72px
    ),
    600px
  );
  box-sizing: border-box;
  margin: 8px 0 12px;
  margin-left: calc((100% - var(--fylar-office-inline-width)) / 2);
  overflow: visible;
  scroll-margin-top: 16px;
  scroll-margin-bottom: calc(var(--dsh-composer-height, 152px) + 16px);
  border: 1px solid var(--dsw-alias-border-subtle, #d8dee4);
  border-radius: 12px;
  box-shadow: 0 8px 24px rgba(15, 23, 42, 0.07);
}

.fylar-office-details-header {
  position: relative;
  z-index: 10;
  display: flex;
  min-height: 52px;
  box-sizing: border-box;
  align-items: center;
  gap: 8px;
  padding: 4px 8px 4px 10px;
  border-bottom: 1px solid var(--dsw-alias-border-subtle, #d8dee4);
  background: var(--dsw-alias-bg-elevated, #fff);
}

.fylar-office-editor-inline .fylar-office-details-header {
  border-radius: 11px 11px 0 0;
}

.fylar-office-document-identity {
  display: flex;
  min-width: 0;
  flex: 1;
  align-items: center;
  gap: 8px;
}

.fylar-office-details-heading {
  display: flex;
  min-width: 0;
  flex: 1;
  align-items: center;
  gap: 8px;
}

.fylar-office-details-heading strong {
  min-width: 0;
  flex: 1 1 auto;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: 14px;
  font-weight: 600;
  line-height: 1.25;
}

.fylar-office-document-meta {
  display: flex;
  min-width: 0;
  flex: 0 1 auto;
  align-items: center;
  gap: 5px;
  font-size: 11px;
  color: var(--dsw-alias-fg-muted, #667085);
}

.fylar-office-document-product,
.fylar-office-document-separator {
  flex: 0 0 auto;
}

.fylar-office-mode-badge {
  display: inline-flex;
  min-width: 0;
  align-items: center;
  gap: 5px;
  overflow: hidden;
  color: var(--dsw-alias-fg-muted, #667085);
  text-overflow: ellipsis;
  white-space: nowrap;
}

.fylar-office-mode-badge > span {
  width: 6px;
  height: 6px;
  flex: 0 0 auto;
  border-radius: 50%;
  background: #1769e0;
}

.fylar-office-mode-badge[data-tone='dirty'] > span { background: #e78b18; }
.fylar-office-mode-badge[data-tone='readonly'] > span { background: #98a2b3; }
.fylar-office-mode-badge[data-tone='loading'] > span { background: #1769e0; animation: fylar-office-pulse 1.2s ease-in-out infinite; }
.fylar-office-mode-badge[data-tone='error'] > span { background: #d92d20; }

.fylar-office-details-header-actions {
  display: flex;
  flex: 0 0 auto;
  align-items: center;
  gap: 4px;
}

.fylar-office-details-close {
  display: inline-flex;
  min-width: 44px;
  height: 44px;
  align-items: center;
  justify-content: center;
  gap: 5px;
  padding: 0 9px;
  border: 0;
  border-radius: 8px;
  background: transparent;
  color: inherit;
  font: inherit;
  font-size: 13px;
  cursor: pointer;
}

.fylar-office-details-close:hover {
  background: var(--dsw-alias-bg-hover, #eef1f4);
}

.fylar-office-details-close:disabled,
.fylar-office-details-error button:disabled {
  cursor: not-allowed;
  opacity: 0.45;
}

.fylar-office-details-error button {
  white-space: nowrap;
  border: 1px solid var(--dsw-alias-border-default, #c8cdd3);
  border-radius: 6px;
  background: var(--dsw-alias-bg-elevated, #fff);
  color: inherit;
  padding: 5px 8px;
  font: inherit;
  font-size: 12px;
  cursor: pointer;
}

.fylar-office-details-body {
  position: relative;
  z-index: 0;
  min-height: 0;
  flex: 1;
  overflow: hidden;
}

.fylar-office-details-viewer {
  position: absolute;
  inset: 0;
  min-height: 0;
  overflow: hidden;
}

.fylar-office-details-loading {
  position: absolute;
  inset: 0;
  z-index: 2;
  display: grid;
  place-items: center;
  padding: 20px;
  background: color-mix(in srgb, var(--dsw-alias-bg-base, #fff) 88%, transparent);
  font-size: 13px;
}

.fylar-office-details-error {
  display: flex;
  flex-direction: column;
  gap: 4px;
  padding: 9px 11px;
  border-bottom: 1px solid color-mix(in srgb, #d92d20 35%, transparent);
  background: color-mix(in srgb, #d92d20 10%, var(--dsw-alias-bg-base, #fff));
  color: var(--dsw-alias-fg-default, #7a271a);
  font-size: 12px;
}

.fylar-office-details-error div {
  display: flex;
  gap: 6px;
  margin-top: 4px;
}

.fylar-office-details-error strong,
.fylar-office-details-error span {
  display: block;
}

.fylar-office-details-close:focus-visible,
.fylar-office-details-error button:focus-visible {
  outline: 2px solid var(--dsw-alias-brand-primary, #1769e0);
  outline-offset: 2px;
}

@container (max-width: 720px) {
  .fylar-office-document-product,
  .fylar-office-document-separator { display: none; }
  .fylar-office-details-header { gap: 5px; padding-inline: 7px 5px; }
  .fylar-office-document-identity,
  .fylar-office-details-heading { gap: 6px; }
}

@container (max-width: 460px) {
  .fylar-office-details-header-actions .fylar-office-action-trigger[data-variant='text'] > span,
  .fylar-office-details-header-actions .fylar-office-action-chevron { display: none; }
  .fylar-office-details-header-actions .fylar-office-action-trigger[data-variant='text'] {
    min-width: 44px;
    padding-inline: 0;
  }
  .fylar-office-details-close span { display: none; }
  .fylar-office-details-close { padding-inline: 0; }
}

@container (max-width: 360px) {
  .fylar-office-document-meta { display: none; }
  .fylar-office-document-identity { gap: 6px; }
}

@media (max-width: 1023px) {
  .fylar-office-editor-inline {
    width: 100%;
    height: clamp(
      360px,
      calc(
        var(--dsh-conversation-viewport-height, 100dvh) -
        var(--dsh-composer-height, 152px) -
        64px
      ),
      560px
    );
    margin: 8px 0 12px;
  }
}

@media (max-width: 639px) {
  .fylar-office-editor-inline {
    width: 100%;
    height: clamp(
      320px,
      calc(
        var(--dsh-conversation-viewport-height, 100svh) -
        var(--dsh-composer-height, 152px) -
        48px
      ),
      480px
    );
    min-height: 320px;
    margin-inline: 0;
    border-radius: 0;
    box-shadow: none;
  }
}
`
		].join("\n");
		/** Install plugin-owned styles and return their disposer. */
		function installOfficeStyles(documentTarget = document) {
			if (documentTarget.querySelector("style[data-dsh-fylar-office-editor]") !== null) return () => void 0;
			const style = documentTarget.createElement("style");
			style.dataset.dshFylarOfficeEditor = "true";
			style.textContent = OFFICE_STYLES;
			documentTarget.head.append(style);
			return () => {
				style.remove();
			};
		}
		//#endregion
		//#region src/client/index.ts
		const OFFICE_RIGHTBAR_ID = "@fylar/dsh-fylar-office-editor";
		const OFFICE_RIGHTBAR_KIND = "fylar-office";
		/** Browser services required for Tool status, Chat delivery, Office, and the right Sidebar. */
		const inject = [
			"slots",
			"locale",
			"uiConversation",
			"sidebarRight",
			"sidebarRightTabs"
		];
		/** Register one Office session across Chat, the Office view, and the right Sidebar. */
		function apply(ctx) {
			const workspace = new OfficeWorkspace(new OfficeSdkLoader());
			let officeTabId = null;
			const controller = new OfficeSessionController(workspace, {
				open() {
					ctx.sidebarRight.openTab(OFFICE_RIGHTBAR_KIND);
					const active = ctx.sidebarRight.active();
					if (active?.kind === "fylar-office") officeTabId = active.id;
				},
				close() {
					const tabId = officeTabId;
					officeTabId = null;
					if (tabId !== null) ctx.sidebarRight.close(tabId);
				}
			});
			const hooks = {
				workspace: workspace.source,
				officeSession: controller.source
			};
			const editorActions = {
				attach: workspace.actions.attach,
				detach: workspace.actions.detach,
				setReadOnly: workspace.actions.setReadOnly,
				exportDocument: workspace.actions.exportDocument,
				exportPdf: workspace.actions.exportPdf,
				clearIssue: workspace.actions.clearIssue,
				markPotentialEdits: workspace.actions.markPotentialEdits
			};
			const detailsInjected = {
				...editorActions,
				hooks,
				retrySession: () => controller.retry(),
				closePanel: () => {
					controller.close();
				},
				panelOpened: (tabId) => {
					officeTabId = tabId;
				},
				panelClosed: (tabId) => {
					if (officeTabId === tabId) officeTabId = null;
					controller.detailsPanelClosed();
				}
			};
			const deliverablesInjected = {
				...editorActions,
				hooks,
				openDetails: (target) => {
					controller.openDetails(target);
				},
				openInlineArtifact: (target) => {
					controller.openInline(target);
				},
				collapseInline: () => {
					controller.collapseInline();
				},
				closeDocument: () => {
					controller.close();
				},
				retrySession: () => controller.retry()
			};
			const officeViewInjected = {
				...editorActions,
				preload: workspace.actions.preload,
				hooks,
				openLocal: (file) => controller.openLocal(file),
				createDocument: (docType) => controller.create(docType),
				retrySession: () => controller.retry(),
				activateOfficeView: () => {
					controller.activateOfficeView();
				},
				parkOfficeView: () => {
					controller.parkOfficeView();
				},
				closeDocument: () => controller.close()
			};
			ctx.effect(() => ctx.locale.register(OFFICE_LOCALE_NS, {
				zh,
				en
			}), "dsh-fylar-office-editor: dictionaries");
			const t = ctx.locale.bind(OFFICE_LOCALE_NS);
			ctx.effect(() => ctx.sidebarRightTabs.register({
				id: OFFICE_RIGHTBAR_ID,
				kind: OFFICE_RIGHTBAR_KIND,
				title: () => t("view.label")
			}), "dsh-fylar-office-editor: right-Sidebar tab type");
			ctx.effect(() => ctx.slots.inject("sidebar.right.pane.tab", () => ctx.slots.register({
				name: "sidebar.right.pane.tab",
				key: OFFICE_RIGHTBAR_ID,
				locale: OFFICE_LOCALE_NS,
				inject: () => detailsInjected
			}, OfficeDetailsPanel)), "dsh-fylar-office-editor: right-Sidebar tab body");
			ctx.effect(() => installOfficeStyles(), "dsh-fylar-office-editor: styles");
			ctx.effect(() => () => {
				controller.dispose();
			}, "dsh-fylar-office-editor: workspace linkage");
			ctx.uiConversation.events.register(officeDeliverablesDefinition);
			for (const key of ["office_create_docx", "office_present_file"]) ctx.slots.inject("tool.call.toolview", () => ctx.slots.register({
				name: "tool.call.toolview",
				key,
				locale: OFFICE_LOCALE_NS
			}, OfficeArtifactRow));
			ctx.slots.inject("conversation.chat.turnTail", () => ctx.slots.register({
				name: "conversation.chat.turnTail",
				select: selectOfficeDeliverables,
				priority: -20,
				locale: OFFICE_LOCALE_NS,
				inject: () => deliverablesInjected
			}, OfficeDeliverables));
			ctx.slots.inject("conversation.view", () => ctx.slots.register({
				name: "conversation.view",
				id: "office",
				order: 5,
				label: "Office",
				locale: OFFICE_LOCALE_NS,
				inject: () => officeViewInjected
			}, OfficeView));
		}
		//#endregion
		exports.OFFICE_RIGHTBAR_ID = OFFICE_RIGHTBAR_ID;
		exports.OFFICE_RIGHTBAR_KIND = OFFICE_RIGHTBAR_KIND;
		exports.apply = apply;
		exports.inject = inject;
		return module.exports;
	}
});

//# sourceMappingURL=client.js.map
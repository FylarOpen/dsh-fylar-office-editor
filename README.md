# Fylar Office Editor for DeepSeek Harness

> Turn Office files created by an agent into documents you can open, review, edit, and download without leaving DeepSeek Harness.

English · [简体中文](./README.zh-CN.md)

`@fylar/dsh-fylar-office-editor` connects Office documents to the DeepSeek Harness (DSH) Web experience. It adds a single delivery card to the conversation, an Office workspace tab, a side-panel editor, and a fullscreen editor, all powered by the Fylar Office SDK.

The plugin can create a real DOCX file from the agent's Markdown response. It can also present existing DOC, DOCX, XLS, XLSX, and PPTX files created by Python, shell commands, or other tools in the current DSH workspace.

## Highlights

- **Office files in the conversation** — open the delivered file inline, move the same editing session to the side panel, or enter fullscreen.
- **Real DOCX generation** — the agent can convert headings, paragraphs, lists, simple tables, code blocks, bold, italic, and inline code into an OOXML Word document.
- **Word, Excel, and PowerPoint editing** — open supported local or workspace files, or create blank documents from the Office tab.
- **One persistent editing session** — moving between Chat, the Office tab, the side panel, and fullscreen does not reopen the document.
- **Export instead of silent overwrite** — browser edits stay in the SDK session until you download the edited Office file.
- **English and Simplified Chinese UI** — the plugin shell follows the DSH locale; the SDK loads its matching language pack.
- **Workspace-confined file access** — the Host only serves supported files from registered DSH workspace roots and blocks traversal and symlink escapes.

## The Fylar Office SDK

This plugin is powered by the [Fylar Office SDK](https://fylar.ai/) — a fully client-side, serverless Office runtime that brings native Office capabilities to any web application.

- **Web-native and serverless** — no server deployment, third-party conversion tools, or desktop Office suites; files are parsed, rendered, edited, and exported directly in the browser.
- **Document content stays local** — Office files are processed on the client instead of being uploaded for cloud document processing. The SDK may still contact the Fylar licensing service to validate credentials and meter usage as described in the commercial license.
- **No iframe, just a DOM node** — integration is a programmable DOM node instead of an embedded external viewer.
- **High-fidelity compatibility** — built around the Office (OOXML) specification to preserve layout, formatting, and document behavior, with pixel-perfect rendering wherever possible.
- **High-performance rendering** — documents open in seconds: first paint in under 500 ms and typical opening in 1–2 s, all processed 100% locally.
- **Everything is programmable** — all Office capabilities are exposed in an API-first way, and you can use the API to operate the editor as needed.

In this plugin, the SDK runs entirely inside the DSH browser session: after the browser receives a file as a Blob, the SDK owns document rendering and editing (see [Editing and saving](#editing-and-saving)).

The bundled SDK is proprietary commercial software. Each License Credential or associated customer account includes an aggregate Free Allowance of up to 3,000 Metered Uses. Continued use after that allowance requires an active paid entitlement or another written commercial license. See the [Fylar Office SDK Commercial License](./vendor/office-sdk/legal.txt) for the controlling definitions and terms.

## Get started

### Requirements

- DeepSeek Harness from the npm `latest` channel. This release's exact development baseline is `0.1.2-rc.1`; alpha builds are evaluated separately for upcoming changes and are not a supported user environment.
- Node.js `^22.19.0 || >=24.0.0`.
- The DSH Web profile.
- A modern browser with Worker, SharedWorker, WebAssembly, IndexedDB, Blob, and ArrayBuffer support.
- A valid Fylar Office SDK License Credential or other authorization. The bundled commercial license includes the Free Allowance described above.

### 1. Install from npm

Stop the currently running DSH process, then install the plugin package:

```bash
npx @deepseek-ai/dsh plugin --profile web add \
  @fylar/dsh-fylar-office-editor
```

If running DSH from its source repository, use the repository script instead:

```bash
pnpm dsh plugin --profile web add \
  @fylar/dsh-fylar-office-editor
```

The npm package includes its verified `lib/` output and SDK runtime, so users do not need to clone or build the plugin.

### 2. Restart DSH

```bash
npx @deepseek-ai/dsh web
```

When using the DSH source repository:

```bash
pnpm dsh web
```

Refresh an already open DSH page after the Web process restarts.

### 3. Ask for a document

Try one of these requests in a workspace-write session:

```text
Turn this conversation into a Word project summary and save it as reports/project-summary.docx.

Create an XLSX comparison table in the workspace, then open it for preview.

Open the existing slides/company-introduction.pptx in the Office editor.
```

The built-in agent tool directly generates DOCX files. XLSX and PPTX content must currently be created by another agent tool or script; this plugin then validates and presents the resulting file.

## How it works

There are two delivery paths:

1. `office_create_docx` converts Markdown to a real DOCX file, writes it to the active DSH workspace, and publishes the result to the conversation.
2. `office_present_file` validates an existing Office file in the active workspace and publishes it without generating or modifying the file.

DSH selects these tools from the user's request and the plugin's model guidance. Users normally do not need to mention the tool names.

After publication, the assistant turn contains one Office delivery card. Open it inline, move it to the side panel, switch to the Office tab, or enter fullscreen. Every presentation reuses the same browser-side SDK widget.

## Supported formats

| Format | Open | Edit | Create from Office tab | Download edited file | Export PDF |
| ------ | ---- | ---- | ---------------------- | -------------------- | ---------- |
| DOC    | Yes  | Yes  | No                     | DOC                  | Yes        |
| DOCX   | Yes  | Yes  | Yes                    | DOCX                 | Yes        |
| XLS    | Yes  | Yes  | No                     | XLS                  | Yes        |
| XLSX   | Yes  | Yes  | Yes                    | XLSX                 | Yes        |
| PPTX   | Yes  | Yes  | Yes                    | PPTX                 | No         |

The Office tab also lets you select a supported local file. Local files remain browser-side and are not written into the DSH workspace by opening them.

## Editing and saving

Document rendering and editing are owned by the Fylar Office SDK after the browser receives the file as a Blob. They do not write through the DSH filesystem while you type.

- Editing does not automatically overwrite the original workspace file.
- Use **Export** to download and retain the edited Office document.
- Closing or replacing the current document ends the browser editing session.
- Refreshing or leaving the page triggers a browser warning when the plugin has detected possible edits.
- A successful Office-format download clears that warning state; PDF export does not.

## PDF export

PDF export is available for Word and Excel formats when all of the following are true:

- the page runs on `localhost` or HTTPS;
- the browser supports `window.queryLocalFonts`;
- the user grants local-font access.

When the environment does not qualify, the PDF action is disabled and the plugin explains why. PowerPoint PDF export will be introduced in a future release.

## Upgrade and uninstall

To replace an existing installation, stop DSH, remove the current plugin row, install the current npm package version, and restart:

```bash
npx @deepseek-ai/dsh plugin --profile web remove @fylar/dsh-fylar-office-editor
npx @deepseek-ai/dsh plugin --profile web add \
  @fylar/dsh-fylar-office-editor
npx @deepseek-ai/dsh web
```

## Current limits

- Public compatibility follows the DSH npm `latest` channel; the repository lockfile pins the exact development baseline, while alpha evaluation is only an early warning for a future plugin update.
- Only DOCX has a dedicated agent-side content generator in this release.
- Browser edits must be downloaded; they are not synchronized back to the workspace file.
- The plugin keeps one active document per page rather than one document per DSH task.
- PDF export depends on browser and font-access capabilities.

## Distribution and notices

This plugin is distributed through the npm registry and includes the proprietary Fylar Office SDK runtime. The SDK license service endpoint is fixed at `https://openlicense.fylar.ai`. The SDK is governed by the [Fylar Office SDK Commercial License](./vendor/office-sdk/legal.txt), any applicable order form, and any separate written commercial agreement. It is not open-source software. The bundled license provides an aggregate maximum of 3,000 Metered Uses per License Credential or associated customer account; continued use requires a paid entitlement or another written commercial license after that allowance is exhausted.

Third-party notices are preserved in [THIRD_PARTY_NOTICES](./vendor/office-sdk/THIRD_PARTY_NOTICES). See [NOTICE.md](./NOTICE.md) before using or redistributing a packaged build.

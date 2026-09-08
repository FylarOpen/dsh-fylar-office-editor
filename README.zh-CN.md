# DeepSeek Harness 的 Fylar Office Editor

> 让 Agent 生成的 Office 文件直接进入 DeepSeek Harness，在对话中打开、预览、编辑并下载。

[English](./README.md) · 简体中文

`dsh-fylar-office-editor` 将 Office 文档接入 DeepSeek Harness（DSH）Web。插件提供对话交付卡片、Office 工作区页签、右侧栏编辑器和全屏编辑器，底层由 Fylar Office SDK 驱动。

插件可以把 Agent 的 Markdown 回复转换为真实 DOCX，也可以展示由 Python、Shell 或其他工具在当前 DSH 工作区生成的 DOC、DOCX、XLS、XLSX 和 PPTX 文件。

## 主要能力

- **在对话中交付 Office 文件**：可以在消息内展开、把同一个编辑会话移到右侧栏，或进入全屏。
- **生成真实 DOCX**：支持把标题、段落、列表、简单表格、代码块、粗体、斜体和行内代码转换为 OOXML Word 文档。
- **编辑 Word、Excel 和 PowerPoint**：打开受支持的本地文件或工作区文件，也可以从 Office 页签新建空白文档。
- **复用同一个编辑会话**：在 Chat、Office 页签、右侧栏和全屏之间切换时不会重新打开文档。
- **通过下载保留编辑结果**：浏览器内修改保留在 SDK 会话中，下载编辑后的 Office 文件后才能长期保存。
- **中英文界面**：插件外层跟随 DSH 语言；SDK 加载对应的语言包。
- **限制在工作区内读取**：Host 只提供已注册 DSH 工作区中的受支持文件，并阻止目录穿越和符号链接逃逸。

## Fylar Office SDK

本插件由 [Fylar Office SDK](https://fylar.ai/) 驱动。它是一个完全运行在浏览器端、无服务器（serverless）的 Office 运行时，为任意 Web 应用带来原生 Office 能力。

- **Web 原生、无服务器**：无需部署服务器、第三方转换工具或桌面 Office 套件；文件直接在浏览器内完成解析、渲染、编辑与导出。
- **文档内容留在本地处理**：Office 文件在客户端完成处理，不会上传到云端进行文档处理。SDK 仍可能按照商业许可证的说明连接 Fylar 授权服务，以校验凭据和统计用量。
- **无需 iframe，只是一个 DOM 节点**：以可编程的 DOM 节点方式集成，而不是内嵌外部查看器。
- **高保真兼容**：围绕 Office（OOXML）规范构建，尽可能保留版式、格式与文档行为，做到像素级渲染。
- **高性能渲染**：文档秒级打开——首屏渲染低于 500ms，常规打开约 1~2 秒，全部在本地处理。
- **一切皆可编程**：以 API-first 的方式暴露全部 Office 能力，根据需要使用 API 来操作编辑器。

在本插件中，SDK 完全运行于 DSH 浏览器会话内：浏览器拿到文件 Blob 后，文档的渲染与编辑均由 SDK 管理（见下文“编辑与保存”一节）。

随插件提供的 SDK 是专有商用软件。每个 License Credential 或关联客户账户享有合计最多 3,000 次 Metered Uses 的免费额度；额度用尽后，继续使用需要有效的付费权益或其他书面商业许可。具体定义和条款以 [Fylar Office SDK 商业许可证](./vendor/office-sdk/legal.txt)为准。

## 快速开始

### 环境要求

- 使用 npm `latest` 渠道的 DeepSeek Harness。本版本的开发依赖精确固定为 `0.1.1-rc.2`；alpha 版本单独用于评估后续改动，不属于正式支持的用户环境。
- Node.js `^22.19.0 || >=24.0.0`。
- DSH Web profile。
- 浏览器支持 Worker、SharedWorker、WebAssembly、IndexedDB、Blob 和 ArrayBuffer。
- 具备有效的 Fylar Office SDK License Credential 或其他授权；随附商业许可证包含上述免费额度。

### 1. 从 npm 安装

先停止正在运行的 DSH，再安装插件包：

```bash
npx @deepseek-ai/dsh plugin --profile web add \
  dsh-fylar-office-editor
```

如果从 DSH 源码仓库运行，则在 DSH 源码根目录执行：

```bash
pnpm dsh plugin --profile web add \
  dsh-fylar-office-editor
```

npm 包已经包含验证过的 `lib/` 构建产物和 SDK runtime，普通用户不需要克隆或构建插件。

### 2. 重启 DSH

```bash
npx @deepseek-ai/dsh web
```

使用 DSH 源码仓库时：

```bash
pnpm dsh web
```

Web 进程重启后，刷新已经打开的 DSH 页面。

### 3. 在对话中提出需求

请在允许写工作区的会话中尝试：

```text
把本轮对话整理成 Word 项目总结，保存为 reports/project-summary.docx。

在工作区生成一份 XLSX 对比表，然后打开预览。

用 Office 编辑器打开已有的 slides/company-introduction.pptx。
```

插件内置的 Agent 工具可以直接生成 DOCX。XLSX 和 PPTX 内容目前需要由其他 Agent 工具或脚本生成，再由本插件校验并发布。

## 工作原理

插件提供两条文件交付路径：

1. `office_create_docx` 把 Markdown 转换为真实 DOCX，写入当前 DSH 工作区，并把结果发布到对话。
2. `office_present_file` 校验当前工作区中已经存在的 Office 文件并发布，但不会生成或修改文件。

DSH 会根据用户请求和插件加入的模型规则选择这些工具。普通用户不需要在提示词中指定工具名。

文件发布成功后，助手消息中只显示一张 Office 交付卡片。可以在对话内打开、移到右侧栏、切换到 Office 页签或进入全屏；这些位置始终复用同一个浏览器 SDK widget。

## 格式支持

| 格式 | 打开 | 编辑 | Office 页签新建 | 下载编辑结果 | 导出 PDF |
| ---- | ---- | ---- | --------------- | ------------ | -------- |
| DOC  | 支持 | 支持 | 不支持          | DOC          | 支持     |
| DOCX | 支持 | 支持 | 支持            | DOCX         | 支持     |
| XLS  | 支持 | 支持 | 不支持          | XLS          | 支持     |
| XLSX | 支持 | 支持 | 支持            | XLSX         | 支持     |
| PPTX | 支持 | 支持 | 支持            | PPTX         | 不支持   |

Office 页签也可以选择受支持的本地文件。本地文件打开后保留在浏览器侧，不会因为打开操作而写入 DSH 工作区。

## 编辑与保存

浏览器取得文件 Blob 后，文档渲染和内部编辑状态均由 Fylar Office SDK 管理。输入内容时不会通过 DSH 文件系统回写原文件。

- 页面内编辑不会自动覆盖工作区原文件。
- 使用“导出”下载并保留编辑后的 Office 文档。
- 关闭或替换当前文档会结束浏览器编辑会话。
- 插件检测到可能存在修改时，刷新或离开页面会触发浏览器提醒。
- 成功下载 Office 格式会清除提醒状态；PDF 导出不会清除。

## PDF 导出

Word 和 Excel 格式在以下条件全部满足时可以导出 PDF：

- 页面运行在 `localhost` 或 HTTPS；
- 浏览器支持 `window.queryLocalFonts`；
- 用户允许读取本地字体。

环境不满足时，PDF 操作会被禁用并显示原因。PowerPoint PDF 导出将在后续版本推出。

## 升级与卸载

替换已安装版本时，先停止 DSH，移除当前插件，再安装当前 npm 包版本并重启：

```bash
npx @deepseek-ai/dsh plugin --profile web remove dsh-fylar-office-editor
npx @deepseek-ai/dsh plugin --profile web add \
  dsh-fylar-office-editor
npx @deepseek-ai/dsh web
```

## 当前限制

- 对外兼容跟随 DSH npm `latest` 渠道；仓库 lockfile 精确固定开发基线，alpha 评估只用于提前发现下一版插件需要处理的问题。
- 当前只有 DOCX 提供专用的 Agent 内容生成工具。
- 浏览器内编辑必须下载保存，不会同步覆盖工作区原文件。
- 每个页面只保留一个活动文档，而不是每个 DSH 任务分别保留一个文档。
- PDF 导出依赖浏览器能力和本地字体权限。

## 分发与声明

本插件通过 npm registry 分发，并包含专有的 Fylar Office SDK runtime。SDK 受 [Fylar Office SDK 商业许可证](./vendor/office-sdk/legal.txt)、适用的订单以及其他书面商业协议约束，不属于开源软件。随附许可证为每个 License Credential 或关联客户账户提供合计最多 3,000 次 Metered Uses 的免费额度；额度用尽后，继续使用需要付费权益或其他书面商业许可。

第三方声明保留在 [THIRD_PARTY_NOTICES](./vendor/office-sdk/THIRD_PARTY_NOTICES)。使用或分发构建包前请阅读 [NOTICE.md](./NOTICE.md)。

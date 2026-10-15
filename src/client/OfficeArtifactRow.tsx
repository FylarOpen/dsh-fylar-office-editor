/** Compact progress row for Office creation and existing-file publication. */
import type { ToolCallViewProps } from '@deepseek-ai/dsh-client-ui-tool/client'
import type { PropsLocale } from '@deepseek-ai/dsh-client-ui-slots'
import { officeBaseName } from './office-files.ts'
import { OFFICE_LOCALE_NS, zhT } from './locales.ts'

export type OfficeArtifactRowProps = ToolCallViewProps & PropsLocale<typeof OFFICE_LOCALE_NS>

interface ToolRowModel {
  readonly path: string | undefined
  readonly state: 'running' | 'ready' | 'error' | 'stopped'
}

const LABEL_KEYS = {
  present: {
    running: 'tool.present.running', ready: 'tool.present.ready',
    error: 'tool.present.error', stopped: 'tool.present.stopped',
  },
  create: {
    running: 'tool.create.running', ready: 'tool.create.ready',
    error: 'tool.create.error', stopped: 'tool.create.stopped',
  },
} as const

function pathFromArgs(raw: string): string {
  try {
    const value = JSON.parse(raw) as unknown
    if (typeof value === 'object' && value !== null) {
      const path = (value as Record<string, unknown>).file_path
      if (typeof path === 'string' && path !== '') return path
    }
  } catch {
    // Streaming calls may expose an incomplete JSON prefix.
  }
  return 'document.docx'
}

export function officeArtifactRowModel(block: ToolCallViewProps['block']): ToolRowModel {
  const settled = 'kind' in block
  const raw = settled ? block.call?.argsRaw ?? '' : block.phase === 'start' ? block.argsRaw : undefined
  const state = !settled
    ? 'running'
    : block.error?.code === 'interrupted'
      ? 'stopped'
      : block.isError ? 'error' : 'ready'
  return { path: raw === undefined ? undefined : pathFromArgs(raw), state }
}

/** Render generation progress only; the assistant-tail card is the file entry. */
export function OfficeArtifactRow({ block, toolName, t = zhT }: OfficeArtifactRowProps) {
  const model = officeArtifactRowModel(block)
  const publishing = toolName === 'office_present_file'
  const label = t(LABEL_KEYS[publishing ? 'present' : 'create'][model.state])
  return (
    <div
      className="fylar-office-tool-status"
      data-state={model.state}
      role="status"
    >
      <span className="fylar-office-tool-status-dot" aria-hidden />
      <span>{label}</span>
      {model.path === undefined ? null : <small title={model.path}>{officeBaseName(model.path)}</small>}
    </div>
  )
}

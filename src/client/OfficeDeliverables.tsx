/** Turn-scoped Office outputs and their assistant-tail file cards. */
import type {
  ConversationNodeDefinition,
  ConversationTurnDataMap,
} from '@deepseek-ai/dsh-client-ui-conversation/client'
import type { TurnTailOwnerProps } from '@deepseek-ai/dsh-client-ui-chat/client'
import { isAppendSurfaceEvent } from '@deepseek-ai/dsh-session/surface'
import type { InjectFace, PropsLocale, PropsRuntime } from '@deepseek-ai/dsh-client-ui-slots'
import type { OfficeArtifactTarget, OfficeSessionSelection, OfficeSessionSnapshot } from './office-artifact-controller.ts'
import { OfficeEditorSurface } from './OfficeEditorSurface.tsx'
import type { ObservableSnapshot, OfficeWorkspaceActions, OfficeWorkspaceSnapshot } from './office-workspace.ts'
import { describeOfficeDocument, officeBaseName, officeDirectoryLabel } from './office-files.ts'
import { OfficeFileIcon } from './OfficeFileIcon.tsx'
import { OfficeActionMenu, type OfficeActionMenuItem } from './OfficeActionMenu.tsx'
import {
  officeFilesForClosing,
  type OfficeDeliverablesData,
} from './office-deliverables-model.ts'
import { OFFICE_LOCALE_NS, type OfficeTranslate, zhT } from './locales.ts'

export interface OfficeDeliverablesTurnData extends OfficeDeliverablesData {}

/** Conversation Definition kind and the Turn-data key it exclusively owns. */
export const OFFICE_DELIVERABLES_KIND = 'fylar-office-deliverables'
const OFFICE_DELIVERY_TOOL_NAMES = new Set(['office_create_docx', 'office_present_file'])

declare module '@deepseek-ai/dsh-client-ui-conversation/client' {
  interface ConversationTurnDataMap {
    /** Office files created or published by a dsh-fylar-office-editor Tool in this turn. */
    'fylar-office-deliverables': OfficeDeliverablesTurnData
  }
}

interface OfficeDeliverablesState extends OfficeDeliverablesTurnData {
  readonly turn: number
  readonly calls: ReadonlyMap<string, string>
}

function pathFromArguments(raw: unknown): string | undefined {
  if (typeof raw !== 'string') return undefined
  try {
    const value = JSON.parse(raw) as unknown
    if (typeof value !== 'object' || value === null) return undefined
    const path = (value as Record<string, unknown>).file_path
    return typeof path === 'string' && path !== '' ? path : undefined
  } catch {
    return undefined
  }
}

/** Accumulate successful Office creation calls into their engine-owned Turn. */
export const officeDeliverablesDefinition: ConversationNodeDefinition<OfficeDeliverablesState> = {
  kind: OFFICE_DELIVERABLES_KIND,
  match: event => {
    if (event.type === 'turn/start') return { id: String(event.data.turn), role: 'start' }
    if (event.type === 'tool/call' && OFFICE_DELIVERY_TOOL_NAMES.has(event.data.name)) {
      return { id: String(event.data.turn), role: 'update' }
    }
    if (event.type === 'tool/result' && isAppendSurfaceEvent(event)) {
      return { id: String(event.data.turn), role: 'update' }
    }
    return null
  },
  start: (_context, match) => {
    if (match.event.type !== 'turn/start') throw new Error('Office deliverables require turn/start')
    return { turn: match.event.data.turn, calls: new Map(), produced: [] }
  },
  update: (context, match) => {
    if (match.event.type === 'tool/call') {
      const path = pathFromArguments(match.event.data.arguments)
      if (path === undefined) return context.state
      const calls = new Map(context.state.calls)
      calls.set(String(match.event.data.callId), path)
      return { ...context.state, calls }
    }
    if (match.event.type !== 'tool/result') return context.state
    const content = match.event.data.message.content[0]
    if (content?.isError === true) return context.state
    const path = context.state.calls.get(String(match.event.data.message.source.callId))
    if (path === undefined || context.state.produced.some(item => item.path === path)) return context.state
    return { ...context.state, produced: [...context.state.produced, { path, seq: match.event.seq }] }
  },
  buildLocationData: (context, scope) => scope !== 'turn' || context.state === undefined
    ? null
    : {
      kind: 'turn',
      turn: context.state.turn,
      key: OFFICE_DELIVERABLES_KIND,
      value: { produced: context.state.produced },
    },
}

export function selectOfficeDeliverables(owner: TurnTailOwnerProps): readonly string[] | null {
  const paths = officeFilesForClosing(owner.turn.data.get(OFFICE_DELIVERABLES_KIND), owner.seq)
  return paths.length === 0 ? null : paths
}

export interface OfficeDeliverablesInjected extends Pick<
  OfficeWorkspaceActions,
  'attach' | 'detach' | 'setReadOnly' | 'exportDocument' | 'exportPdf' | 'clearIssue' | 'markPotentialEdits'
> {
  hooks: {
    workspace: ObservableSnapshot<OfficeWorkspaceSnapshot>
    officeSession: ObservableSnapshot<OfficeSessionSnapshot>
  }
  openDetails: (target: OfficeArtifactTarget) => void
  openInlineArtifact: (target: OfficeArtifactTarget) => void
  collapseInline: () => void
  closeDocument: () => void
  retrySession: () => Promise<void>
}

type OfficeDeliverablesProps = PropsRuntime<'conversation.chat.turnTail'>
  & { matched: readonly string[] }
  & PropsLocale<typeof OFFICE_LOCALE_NS>
  & InjectFace<OfficeDeliverablesInjected>

function isCurrentArtifact(current: OfficeSessionSelection | null, path: string, cwd: string | undefined): boolean {
  return current?.source.kind === 'workspace'
    && current.source.path === path
    && current.source.cwd === cwd
}

interface ArtifactCardProps {
  readonly path: string
  readonly cwd: string | undefined
  readonly workspace: OfficeWorkspaceSnapshot
  readonly current: OfficeSessionSelection | null
  readonly openInline: () => void
  readonly openDetails: () => void
  readonly download: () => Promise<void>
  readonly close: () => void
  readonly t: OfficeTranslate
}

function ArtifactCard({
  path, cwd, workspace, current, openInline, openDetails, download, close, t,
}: ArtifactCardProps) {
  const descriptor = describeOfficeDocument(path)
  const active = isCurrentArtifact(current, path, cwd)
  const busy = workspace.status === 'loading' || (workspace.status === 'ready' && workspace.operation !== null)
  const status = cwd === undefined
    ? t('artifact.workspaceUnavailable')
    : active && current?.presentation === 'parked'
      ? t('artifact.parked')
      : active && current?.presentation === 'details'
        ? t('artifact.details')
        : active && current?.presentation === 'office-view'
          ? t('artifact.officeView')
          : t('artifact.openInline', { product: descriptor.product })

  const menuItems: OfficeActionMenuItem[] = [
    { id: 'details', label: t('action.openDetails'), onSelect: openDetails },
  ]
  if (active && workspace.status === 'ready') {
    menuItems.push({
      id: 'download',
      label: t('action.download', { extension: descriptor.extension }),
      onSelect: () => { void download() },
    })
  }
  if (active) {
    menuItems.push({ id: 'close', label: t('action.closeSession'), danger: true, onSelect: close })
  }

  return (
    <article className="fylar-office-artifact-card">
      <button type="button" className="fylar-office-artifact-primary" disabled={cwd === undefined || busy} onClick={openInline}>
        <OfficeFileIcon descriptor={descriptor} />
        <span className="fylar-office-artifact-copy">
          <strong>{officeBaseName(path)}</strong>
          <span>{status}</span>
          <small>{officeDirectoryLabel(path, t('workspace.root'))}</small>
        </span>
      </button>
      <OfficeActionMenu
        label={t('artifact.more', { name: officeBaseName(path) })}
        icon="more"
        disabled={cwd === undefined || busy}
        items={menuItems}
      />
    </article>
  )
}

/** Render successful files after the closing assistant response. */
export function OfficeDeliverables({
  matched, useSessions, sessionId, useWorkspace, useOfficeSession,
  openDetails, openInlineArtifact, collapseInline, closeDocument, retrySession,
  attach, detach, setReadOnly, exportDocument, exportPdf, clearIssue, markPotentialEdits,
  t = zhT,
}: OfficeDeliverablesProps) {
  const cwd = useSessions(list => list.byId[sessionId]?.cwd)
  const workspace = useWorkspace(value => value)
  const session = useOfficeSession(value => value)
  return (
    <div className="fylar-office-deliverables" aria-label={t('artifact.group')}>
      {matched.map(path => {
        const target = { path, cwd }
        const current = session.current
        const inline = current?.presentation === 'inline' && isCurrentArtifact(current, path, cwd)
        return inline ? (
          <OfficeEditorSurface
            key={path}
            snapshot={workspace}
            current={current}
            presentation="inline"
            attach={attach}
            detach={detach}
            onRetry={retrySession}
            setReadOnly={setReadOnly}
            exportDocument={exportDocument}
            exportPdf={exportPdf}
            clearIssue={clearIssue}
            markPotentialEdits={markPotentialEdits}
            t={t}
            onOpenDetails={() => { openDetails(target) }}
            onCollapse={collapseInline}
            onClose={closeDocument}
          />
        ) : (
          <ArtifactCard
            key={path}
            path={path}
            cwd={cwd}
            workspace={workspace}
            current={current}
            openInline={() => { openInlineArtifact(target) }}
            openDetails={() => { openDetails(target) }}
            download={exportDocument}
            close={closeDocument}
            t={t}
          />
        )
      })}
    </div>
  )
}

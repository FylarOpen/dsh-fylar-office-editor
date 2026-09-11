/** Fylar Office editor hosted in a DSH right-Sidebar tab. */
import { useEffect } from 'react'
import type { TabId } from '@deepseek-ai/dsh-client-ui-sidebar-right/client'
import type { InjectFace, PropsLocale, PropsRuntime } from '@deepseek-ai/dsh-client-ui-slots'
import type { ObservableSnapshot, OfficeWorkspaceActions, OfficeWorkspaceSnapshot } from './office-workspace.ts'
import type { OfficeSessionSnapshot } from './office-artifact-controller.ts'
import { OfficeEditorSurface } from './OfficeEditorSurface.tsx'
import { OFFICE_LOCALE_NS, zhT } from './locales.ts'

export interface OfficeDetailsInjected extends Pick<
  OfficeWorkspaceActions,
  'attach' | 'detach' | 'setReadOnly' | 'exportDocument' | 'exportPdf' | 'clearIssue' | 'markPotentialEdits'
> {
  hooks: {
    workspace: ObservableSnapshot<OfficeWorkspaceSnapshot>
    officeSession: ObservableSnapshot<OfficeSessionSnapshot>
  }
  retrySession: () => Promise<void>
  closePanel: () => void
  panelOpened: (tabId: TabId) => void
  panelClosed: (tabId: TabId) => void
}

export type OfficeDetailsPanelProps = PropsRuntime<'sidebar.right.pane.tab'>
  & PropsLocale<typeof OFFICE_LOCALE_NS>
  & InjectFace<OfficeDetailsInjected>

/** Render the persistent SDK root plus compact edit/export controls. */
export function OfficeDetailsPanel({
  useWorkspace, useOfficeSession, attach, detach, retrySession, setReadOnly,
  exportDocument, exportPdf, clearIssue, markPotentialEdits, closePanel,
  panelOpened, panelClosed, useTabInfo, t = zhT,
}: OfficeDetailsPanelProps) {
  const { tab } = useTabInfo()
  const snapshot = useWorkspace(value => value)
  const current = useOfficeSession(value => value.current)
  useEffect(() => {
    panelOpened(tab.id)
    const handleClose = (): void => { panelClosed(tab.id) }
    if (tab.signal.aborted) handleClose()
    else tab.signal.addEventListener('abort', handleClose, { once: true })
    return () => { tab.signal.removeEventListener('abort', handleClose) }
  }, [panelClosed, panelOpened, tab.id, tab.signal])
  return (
    <OfficeEditorSurface
      snapshot={snapshot}
      current={current}
      presentation="details"
      attach={attach}
      detach={detach}
      onRetry={retrySession}
      setReadOnly={setReadOnly}
      exportDocument={exportDocument}
      exportPdf={exportPdf}
      clearIssue={clearIssue}
      markPotentialEdits={markPotentialEdits}
      t={t}
      onClose={closePanel}
    />
  )
}

/** Browser plugin linking Agent-created Office files to Chat, Office, and the right Sidebar. */
import type { Context as ClientContext } from '@deepseek-ai/cordis'
import type {} from '@deepseek-ai/dsh-client-locale/client'
import type {} from '@deepseek-ai/dsh-client-ui-conversation/client'
import type {} from '@deepseek-ai/dsh-client-ui-renderer/client'
import type { TabId } from '@deepseek-ai/dsh-client-ui-sidebar-right/client'
import type {} from '@deepseek-ai/dsh-client-ui-tool/client'
import { OfficeSessionController } from './office-artifact-controller.ts'
import { OfficeArtifactRow } from './OfficeArtifactRow.tsx'
import {
  OfficeDeliverables, type OfficeDeliverablesInjected,
  OFFICE_DELIVERABLES_KIND, officeDeliverablesDefinition,
} from './OfficeDeliverables.tsx'
import { OfficeDetailsPanel, type OfficeDetailsInjected } from './OfficeDetailsPanel.tsx'
import { OfficeSdkLoader } from './office-sdk-loader.ts'
import { OfficeWorkspace } from './office-workspace.ts'
import { OfficeView, type OfficeViewInjected } from './OfficeView.tsx'
import { installOfficeStyles } from './styles.ts'
import { en, OFFICE_LOCALE_NS, zh } from './locales.ts'

export const OFFICE_RIGHTBAR_ID = '@fylar/dsh-fylar-office-editor'
export const OFFICE_RIGHTBAR_KIND = 'fylar-office'

/** Browser services required for Tool status, Chat delivery, Office, and the right Sidebar. */
export const inject = ['slots', 'locale', 'uiConversation', 'sidebarRight', 'sidebarRightTabs']

/** Register one Office session across Chat, the Office view, and the right Sidebar. */
export function apply(ctx: ClientContext): void {
  const loader = new OfficeSdkLoader()
  const workspace = new OfficeWorkspace(loader)
  let officeTabId: TabId | null = null
  const detailsPanel = {
    open(): void {
      ctx.sidebarRight.openTab(OFFICE_RIGHTBAR_KIND)
      const active = ctx.sidebarRight.active()
      if (active?.kind === OFFICE_RIGHTBAR_KIND) officeTabId = active.id
    },
    close(): void {
      const tabId = officeTabId
      officeTabId = null
      if (tabId !== null) ctx.sidebarRight.close(tabId)
    },
  }

  const controller = new OfficeSessionController(workspace, detailsPanel)
  const hooks = { workspace: workspace.source, officeSession: controller.source }
  const editorActions = {
    attach: workspace.actions.attach,
    detach: workspace.actions.detach,
    setReadOnly: workspace.actions.setReadOnly,
    exportDocument: workspace.actions.exportDocument,
    exportPdf: workspace.actions.exportPdf,
    clearIssue: workspace.actions.clearIssue,
    markPotentialEdits: workspace.actions.markPotentialEdits,
  }
  const detailsInjected: OfficeDetailsInjected = {
    ...editorActions,
    hooks,
    retrySession: () => controller.retry(),
    closePanel: () => { void controller.close() },
    panelOpened: tabId => { officeTabId = tabId },
    panelClosed: tabId => {
      if (officeTabId === tabId) officeTabId = null
      controller.detailsPanelClosed()
    },
  }
  const deliverablesInjected: OfficeDeliverablesInjected = {
    ...editorActions,
    hooks,
    openDetails: target => { void controller.openDetails(target) },
    openInlineArtifact: target => { void controller.openInline(target) },
    collapseInline: () => { controller.collapseInline() },
    closeDocument: () => { void controller.close() },
    retrySession: () => controller.retry(),
  }
  const officeViewInjected: OfficeViewInjected = {
    ...editorActions,
    preload: workspace.actions.preload,
    hooks,
    openLocal: file => controller.openLocal(file),
    createDocument: docType => controller.create(docType),
    retrySession: () => controller.retry(),
    activateOfficeView: () => { controller.activateOfficeView() },
    parkOfficeView: () => { controller.parkOfficeView() },
    closeDocument: () => controller.close(),
  }
  ctx.effect(() => ctx.locale.register(OFFICE_LOCALE_NS, { zh, en }), 'dsh-fylar-office-editor: dictionaries')
  const t = ctx.locale.bind(OFFICE_LOCALE_NS)
  ctx.effect(() => ctx.sidebarRightTabs.register({
    id: OFFICE_RIGHTBAR_ID,
    kind: OFFICE_RIGHTBAR_KIND,
    title: () => t('view.label'),
  }), 'dsh-fylar-office-editor: right-Sidebar tab type')
  ctx.effect(() => ctx.slots.inject('sidebar.right.pane.tab', () => ctx.slots.register({
    name: 'sidebar.right.pane.tab',
    key: OFFICE_RIGHTBAR_ID,
    locale: OFFICE_LOCALE_NS,
    inject: () => detailsInjected,
  }, OfficeDetailsPanel)), 'dsh-fylar-office-editor: right-Sidebar tab body')
  ctx.effect(() => installOfficeStyles(), 'dsh-fylar-office-editor: styles')
  ctx.effect(() => () => { void controller.dispose() }, 'dsh-fylar-office-editor: workspace linkage')
  ctx.uiConversation.events.register(officeDeliverablesDefinition)
  for (const key of ['office_create_docx', 'office_present_file']) {
    ctx.slots.inject('tool.call.toolview', () => ctx.slots.register({
      name: 'tool.call.toolview',
      key,
      locale: OFFICE_LOCALE_NS,
    }, OfficeArtifactRow))
  }
  ctx.slots.inject('conversation.chat.turnTail', () => ctx.slots.register({
    name: 'conversation.chat.turnTail',
    id: OFFICE_DELIVERABLES_KIND,
    order: -20,
    locale: OFFICE_LOCALE_NS,
    inject: () => deliverablesInjected,
  }, OfficeDeliverables))
  ctx.slots.inject('conversation.view', () => ctx.slots.register({
    name: 'conversation.view',
    id: 'office',
    order: 5,
    label: 'Office',
    locale: OFFICE_LOCALE_NS,
    inject: () => officeViewInjected,
  }, OfficeView))
}

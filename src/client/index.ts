/** Browser plugin linking Agent-created Office files to Chat, Office, and details. */
import type { Context as ClientContext } from '@deepseek-ai/cordis'
import type {} from '@deepseek-ai/dsh-client-locale/client'
import type {} from '@deepseek-ai/dsh-client-ui-conversation/client'
import type {} from '@deepseek-ai/dsh-client-ui-layout/client'
import type {} from '@deepseek-ai/dsh-client-ui-renderer/client'
import type {} from '@deepseek-ai/dsh-client-ui-tool/client'
import { OfficeSessionController } from './office-artifact-controller.ts'
import { OfficeArtifactRow } from './OfficeArtifactRow.tsx'
import {
  OfficeDeliverables, type OfficeDeliverablesInjected,
  officeDeliverablesDefinition, selectOfficeDeliverables,
} from './OfficeDeliverables.tsx'
import { OfficeDetailsPanel, type OfficeDetailsInjected } from './OfficeDetailsPanel.tsx'
import { OfficeSdkLoader } from './office-sdk-loader.ts'
import { OfficeWorkspace } from './office-workspace.ts'
import { OfficeView, type OfficeViewInjected } from './OfficeView.tsx'
import { installOfficeStyles } from './styles.ts'
import { OfficeLayoutLease } from './layout-interop.ts'
import { en, OFFICE_LOCALE_NS, zh } from './locales.ts'

/** Browser services required for Tool status, Chat delivery, Office, and details. */
export const inject = ['slots', 'layout', 'locale', 'uiConversation']

/** Register one Office session across Chat, the Office view, and on-demand details. */
export function apply(ctx: ClientContext): void {
  const loader = new OfficeSdkLoader()
  const workspace = new OfficeWorkspace(loader)
  const layoutLease = new OfficeLayoutLease(ctx.layout)
  let disposePanel: (() => void) | null = null
  let detailsInjected: OfficeDetailsInjected

  const releasePanel = (): void => {
    const dispose = disposePanel
    disposePanel = null
    dispose?.()
  }
  const showPanel = (): void => {
    if (disposePanel !== null) return
    disposePanel = ctx.slots.register({
      name: 'details',
      priority: -100,
      locale: OFFICE_LOCALE_NS,
      inject: () => detailsInjected,
    }, OfficeDetailsPanel)
  }

  const controller = new OfficeSessionController(workspace, ctx.layout, layoutLease, showPanel, releasePanel)
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
  detailsInjected = {
    ...editorActions,
    hooks,
    retrySession: () => controller.retry(),
    closePanel: () => { void controller.close() },
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
    select: selectOfficeDeliverables,
    priority: -20,
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

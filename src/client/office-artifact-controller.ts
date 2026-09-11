/** Global browser session linking Chat artifacts, the Office tab, and the right Sidebar. */
import { createdFileName, inferOfficeDocType, officeBaseName } from './office-files.ts'
import type { ObservableSnapshot, OfficeWorkspace } from './office-workspace.ts'
import type { OfficeDocType } from './office-types.ts'

const OFFICE_FILE_ROUTE_PREFIX = '/fylar-office-files'

export type OfficePresentation = 'inline' | 'details' | 'office-view' | 'parked'

export type OfficeDocumentSource =
  | {
    readonly kind: 'workspace'
    readonly path: string
    readonly cwd: string
    readonly fileName: string
    readonly url: string
    readonly docType: OfficeDocType
  }
  | { readonly kind: 'local'; readonly fileName: string; readonly docType: OfficeDocType }
  | { readonly kind: 'created'; readonly fileName: string; readonly docType: OfficeDocType }

export interface OfficeSessionSelection {
  readonly source: OfficeDocumentSource
  readonly presentation: OfficePresentation
}

export interface OfficeSessionSnapshot {
  readonly current: OfficeSessionSelection | null
}

export interface OfficeArtifactTarget {
  readonly path: string
  readonly cwd?: string | undefined
}

export interface OfficeDetailsPanelLike {
  open(): void
  close(): void
}

function isSameWorkspaceSource(source: OfficeDocumentSource, target: OfficeArtifactTarget): boolean {
  return source.kind === 'workspace' && source.cwd === target.cwd && source.path === target.path
}

export function workspaceOfficeUrl(cwd: string, path: string): string {
  return `${OFFICE_FILE_ROUTE_PREFIX}?${new URLSearchParams({ workspace: cwd, path }).toString()}`
}

/** Owns the one live Office document and every place where its DOM may appear. */
export class OfficeSessionController {
  private snapshot: OfficeSessionSnapshot = { current: null }
  private readonly listeners = new Set<() => void>()
  private disposed = false

  readonly source: ObservableSnapshot<OfficeSessionSnapshot> = {
    getSnapshot: () => this.snapshot,
    subscribe: listener => {
      this.listeners.add(listener)
      return () => { this.listeners.delete(listener) }
    },
  }

  constructor(
    private readonly workspace: OfficeWorkspace,
    private readonly detailsPanel: OfficeDetailsPanelLike,
  ) {}

  private publish(change: Partial<OfficeSessionSnapshot>): void {
    this.snapshot = { ...this.snapshot, ...change }
    for (const listener of this.listeners) listener()
  }

  private present(selection: OfficeSessionSelection): void {
    this.publish({ current: selection })
    if (selection.presentation === 'details') {
      this.detailsPanel.open()
      return
    }
    this.detailsPanel.close()
  }

  private async openWorkspaceNow(target: { path: string; cwd: string }, presentation: 'details' | 'inline'): Promise<void> {
    const fileName = officeBaseName(target.path)
    const source: Extract<OfficeDocumentSource, { kind: 'workspace' }> = {
      kind: 'workspace',
      path: target.path,
      cwd: target.cwd,
      fileName,
      url: workspaceOfficeUrl(target.cwd, target.path),
      docType: inferOfficeDocType(fileName),
    }
    this.present({ source, presentation })
    await this.workspace.openUrl(fileName, source.url)
  }

  private async openWorkspace(target: OfficeArtifactTarget, presentation: 'details' | 'inline'): Promise<void> {
    if (this.disposed || target.cwd === undefined) return
    const current = this.snapshot.current
    if (current !== null && isSameWorkspaceSource(current.source, target)) {
      this.present({ ...current, presentation })
      return
    }
    await this.openWorkspaceNow({ path: target.path, cwd: target.cwd }, presentation)
  }

  async openDetails(target: OfficeArtifactTarget): Promise<void> {
    await this.openWorkspace(target, 'details')
  }

  async openInline(target: OfficeArtifactTarget): Promise<void> {
    await this.openWorkspace(target, 'inline')
  }

  async openLocal(file: File): Promise<void> {
    if (this.disposed) return
    let docType: OfficeDocType
    try {
      docType = inferOfficeDocType(file.name)
    } catch {
      await this.workspace.open(file)
      return
    }
    const source: OfficeDocumentSource = {
      kind: 'local', fileName: file.name, docType,
    }
    this.present({ source, presentation: 'office-view' })
    await this.workspace.open(file)
  }

  async create(docType: OfficeDocType): Promise<void> {
    if (this.disposed) return
    const fileName = createdFileName(docType)
    this.present({ source: { kind: 'created', fileName, docType }, presentation: 'office-view' })
    await this.workspace.create(docType)
  }

  activateOfficeView(): void {
    const current = this.snapshot.current
    if (this.disposed || current === null || current.presentation === 'office-view') return
    this.present({ ...current, presentation: 'office-view' })
  }

  parkOfficeView(): void {
    const current = this.snapshot.current
    if (this.disposed || current?.presentation !== 'office-view') return
    this.publish({ current: { ...current, presentation: 'parked' } })
  }

  collapseInline(): void {
    const current = this.snapshot.current
    if (this.disposed || current?.presentation !== 'inline') return
    this.publish({ current: { ...current, presentation: 'parked' } })
  }

  /** Keep the document alive when DSH closes its native right-Sidebar tab. */
  detailsPanelClosed(): void {
    const current = this.snapshot.current
    if (this.disposed || current?.presentation !== 'details') return
    this.publish({ current: { ...current, presentation: 'parked' } })
  }

  async retry(): Promise<void> {
    await this.workspace.retry()
  }

  async close(): Promise<void> {
    if (this.disposed) return
    this.publish({ current: null })
    this.detailsPanel.close()
    await this.workspace.close()
  }

  async dispose(): Promise<void> {
    if (this.disposed) return
    this.disposed = true
    this.detailsPanel.close()
    await this.workspace.dispose()
    this.snapshot = { current: null }
    this.listeners.clear()
  }
}

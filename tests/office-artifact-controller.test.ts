// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest'
import { OfficeSessionController, workspaceOfficeUrl } from '../src/client/office-artifact-controller.ts'
import { officeFilesForClosing } from '../src/client/office-deliverables-model.ts'
import { officeArtifactRowModel } from '../src/client/OfficeArtifactRow.tsx'
import type { OfficeWorkspace, OfficeWorkspaceSnapshot } from '../src/client/office-workspace.ts'

function workspace(snapshot: OfficeWorkspaceSnapshot) {
  let current = snapshot
  return {
    source: { getSnapshot: () => current, subscribe: () => () => {} },
    openUrl: vi.fn(async () => {}),
    open: vi.fn(async () => {}),
    create: vi.fn(async () => {}),
    retry: vi.fn(async () => {}),
    close: vi.fn(async () => {}),
    dispose: vi.fn(async () => {}),
    setSnapshot(value: OfficeWorkspaceSnapshot) { current = value },
  } as unknown as OfficeWorkspace & { setSnapshot(value: OfficeWorkspaceSnapshot): void }
}

function dependencies() {
  return { panel: { open: vi.fn(), close: vi.fn() } }
}

const ready = (hasPotentialEdits = false): OfficeWorkspaceSnapshot => ({
  status: 'ready', fileName: 'old.docx', docType: 1, origin: 'workspace',
  readOnly: false, operation: null, hasPotentialEdits,
})

describe('Office session linkage', () => {
  it('opens a generated file in details only after a valid Workspace target', async () => {
    const office = workspace({ status: 'empty' })
    const d = dependencies()
    const controller = new OfficeSessionController(office, d.panel)

    await controller.openDetails({ path: 'reports/summary.docx' })
    expect(d.panel.open).not.toHaveBeenCalled()

    await controller.openDetails({ path: 'reports/summary.docx', cwd: '/workspace' })
    expect(d.panel.open).toHaveBeenCalledOnce()
    expect(office.openUrl).toHaveBeenCalledWith(
      'summary.docx', workspaceOfficeUrl('/workspace', 'reports/summary.docx'),
    )
    expect(controller.source.getSnapshot().current).toMatchObject({
      presentation: 'details', source: { kind: 'workspace', path: 'reports/summary.docx' },
    })
  })

  it('keeps the right-Sidebar tab open after a load failure and retries in place', async () => {
    const office = workspace({ status: 'empty' })
    const d = dependencies()
    vi.mocked(office.openUrl).mockImplementation(async () => {
      office.setSnapshot({
        status: 'error',
        issue: { kind: 'open', retryable: true },
      })
    })
    const controller = new OfficeSessionController(office, d.panel)
    await controller.openDetails({ path: 'broken.docx', cwd: '/workspace' })
    expect(d.panel.open).toHaveBeenCalledOnce()
    expect(d.panel.close).not.toHaveBeenCalled()
    await controller.retry()
    expect(office.retry).toHaveBeenCalledOnce()
    expect(d.panel.open).toHaveBeenCalledOnce()
  })

  it('moves one live Workspace document without reopening the SDK', async () => {
    const office = workspace({ status: 'empty' })
    const d = dependencies()
    const controller = new OfficeSessionController(office, d.panel)
    await controller.openInline({ path: 'same.docx', cwd: '/workspace' })
    office.setSnapshot(ready())
    await controller.openDetails({ path: 'same.docx', cwd: '/workspace' })
    expect(office.openUrl).toHaveBeenCalledOnce()
    expect(controller.source.getSnapshot().current?.presentation).toBe('details')
    controller.activateOfficeView()
    expect(controller.source.getSnapshot().current?.presentation).toBe('office-view')
    expect(d.panel.close).toHaveBeenCalled()
    controller.parkOfficeView()
    expect(controller.source.getSnapshot().current?.presentation).toBe('parked')
    expect(office.close).not.toHaveBeenCalled()
  })

  it('replaces the current document directly even when it may contain edits', async () => {
    const office = workspace({ status: 'empty' })
    const d = dependencies()
    const controller = new OfficeSessionController(office, d.panel)
    await controller.openInline({ path: 'old.docx', cwd: '/workspace' })
    office.setSnapshot(ready(true))

    await controller.openInline({ path: 'new.docx', cwd: '/workspace' })
    expect(office.openUrl).toHaveBeenCalledTimes(2)
    expect(controller.source.getSnapshot().current).toMatchObject({
      presentation: 'inline', source: { kind: 'workspace', path: 'new.docx' },
    })
  })

  it('opens local files and creates all document types in the Office view', async () => {
    const office = workspace({ status: 'empty' })
    const d = dependencies()
    const controller = new OfficeSessionController(office, d.panel)
    const file = new File(['x'], 'sheet.xlsx')
    await controller.openLocal(file)
    expect(office.open).toHaveBeenCalledWith(file)
    expect(controller.source.getSnapshot().current).toMatchObject({
      presentation: 'office-view', source: { kind: 'local', docType: 2 },
    })
    await controller.create(3)
    expect(office.create).toHaveBeenCalledWith(3)
    expect(controller.source.getSnapshot().current).toMatchObject({ source: { kind: 'created', docType: 3 } })
  })

  it('closes the native right-Sidebar tab and the Office workspace directly', async () => {
    const office = workspace(ready(true))
    const d = dependencies()
    const controller = new OfficeSessionController(office, d.panel)
    await controller.close()
    expect(office.close).toHaveBeenCalledOnce()
    expect(d.panel.close).toHaveBeenCalledOnce()
    expect(controller.source.getSnapshot().current).toBeNull()
  })

  it('parks the document when DSH closes its native right-Sidebar tab', async () => {
    const office = workspace({ status: 'empty' })
    const d = dependencies()
    const controller = new OfficeSessionController(office, d.panel)
    await controller.openDetails({ path: 'notes.docx', cwd: '/workspace' })

    controller.detailsPanelClosed()

    expect(controller.source.getSnapshot().current?.presentation).toBe('parked')
    expect(office.close).not.toHaveBeenCalled()
  })

  it('derives stable running, success, and failure status states', () => {
    const running = { callId: 'c1', name: 'office_create_docx', argsRaw: '{"file_path":"a.docx"}', subCalls: [] }
    expect(officeArtifactRowModel(running as never)).toEqual({ path: 'a.docx', state: 'running' })
    expect(officeArtifactRowModel({
      kind: 'tool-result', callId: 'c1', isError: false,
      call: { argsRaw: '{"file_path":"a.docx"}' }, content: [], subCalls: [],
    } as never)).toEqual({ path: 'a.docx', state: 'ready' })
    expect(officeArtifactRowModel({
      kind: 'tool-result', callId: 'c1', isError: true,
      call: { argsRaw: '{"file_path":"a.docx"}' }, content: [], subCalls: [],
    } as never).state).toBe('error')
  })

  it('only exposes successfully produced files visible at the closing event', () => {
    expect(officeFilesForClosing({
      produced: [{ path: 'first.docx', seq: 7 }, { path: 'later.docx', seq: 12 }],
    }, 10)).toEqual(['first.docx'])
  })
})

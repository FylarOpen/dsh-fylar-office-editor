// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react'
import type { ComponentProps } from 'react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { OfficeArtifactRow } from '../src/client/OfficeArtifactRow.tsx'
import { OfficeDeliverables, type OfficeDeliverablesTurnData } from '../src/client/OfficeDeliverables.tsx'
import type { OfficeSessionSnapshot } from '../src/client/office-artifact-controller.ts'
import type { OfficeWorkspaceSnapshot } from '../src/client/office-workspace.ts'

vi.mock('@deepseek-ai/dsh-session/surface', () => ({
  isAppendSurfaceEvent: () => true,
}))

afterEach(cleanup)

describe('Office delivery presentation', () => {
  it('keeps the Tool row as non-clickable generation status', () => {
    render(<OfficeArtifactRow {...({
      cwd: '/workspace',
      phase: 'result',
      block: {
        kind: 'tool-result', callId: 'c1', isError: false,
        call: { argsRaw: '{"file_path":"reports/a.docx"}' }, content: [], subCalls: [],
      },
    } as unknown as ComponentProps<typeof OfficeArtifactRow>)} />)
    expect(screen.getByRole('status').textContent).toContain('Word 文档已生成')
    expect(screen.queryByRole('button')).toBeNull()
  })

  it('shows existing XLSX publication as preview-ready status', () => {
    render(<OfficeArtifactRow {...({
      toolName: 'office_present_file',
      cwd: '/workspace',
      phase: 'result',
      block: {
        kind: 'tool-result', callId: 'c2', isError: false,
        call: { argsRaw: '{"file_path":"reports/table.xlsx"}' }, content: [], subCalls: [],
      },
    } as unknown as ComponentProps<typeof OfficeArtifactRow>)} />)
    expect(screen.getByRole('status').textContent).toContain('Office 文件已可预览')
    expect(screen.getByRole('status').textContent).toContain('table.xlsx')
  })

  it('hides the file name until the Tool call has dispatched its arguments', () => {
    const view = render(<OfficeArtifactRow {...({
      phase: 'preparing',
      toolName: 'office_create_docx',
      block: { phase: 'preparing', callId: 'c1', name: 'office_create_docx', subCalls: [] },
    } as unknown as ComponentProps<typeof OfficeArtifactRow>)} />)
    expect(screen.getByRole('status').textContent).toContain('正在生成 Word 文档')
    expect(view.container.querySelector('small')).toBeNull()
  })

  it('subscribes to delivered files and renders their primary action and secondary menu', () => {
    const workspace: OfficeWorkspaceSnapshot = { status: 'empty' }
    const session: OfficeSessionSnapshot = { current: null }
    const openInline = vi.fn()
    const openDetails = vi.fn()
    let data: OfficeDeliverablesTurnData | undefined
    const listeners = new Set<() => void>()
    const source = {
      getSnapshot: () => data,
      subscribe: (listener: () => void) => {
        listeners.add(listener)
        return () => { listeners.delete(listener) }
      },
    }
    const props = {
      turn: { data: { source: () => source } },
      seq: 10,
      sessionId: 's1',
      useSessions: (selector: (value: unknown) => unknown) => selector({ byId: { s1: { cwd: '/workspace' } } }),
      useWorkspace: (selector: (value: OfficeWorkspaceSnapshot) => unknown) => selector(workspace),
      useOfficeSession: (selector: (value: OfficeSessionSnapshot) => unknown) => selector(session),
      openDetails,
      openInlineArtifact: openInline,
      collapseInline: vi.fn(),
      closeDocument: vi.fn(),
      retrySession: vi.fn(async () => {}),
      attach: vi.fn(), detach: vi.fn(), setReadOnly: vi.fn(), exportDocument: vi.fn(), exportPdf: vi.fn(),
      clearIssue: vi.fn(), markPotentialEdits: vi.fn(),
    } as unknown as ComponentProps<typeof OfficeDeliverables>
    const view = render(<OfficeDeliverables {...props} />)
    expect(view.container.childElementCount).toBe(0)

    act(() => {
      data = { produced: [
        { path: 'reports/weekly/a.docx', seq: 7 },
        { path: 'reports/later.docx', seq: 12 },
      ] }
      for (const listener of listeners) listener()
    })

    expect(screen.getByText('reports/weekly')).toBeTruthy()
    expect(screen.queryByText('later.docx')).toBeNull()
    fireEvent.click(screen.getByRole('button', { name: /^a\.docx/ }))
    expect(openInline).toHaveBeenCalledOnce()
    const menuTrigger = screen.getByRole('button', { name: '更多操作：a.docx' })
    menuTrigger.focus()
    fireEvent.click(menuTrigger)
    expect(screen.getByRole('menu')).toBeTruthy()
    fireEvent.keyDown(document, { key: 'Escape' })
    expect(screen.queryByRole('menu')).toBeNull()
    expect(document.activeElement).toBe(menuTrigger)

    fireEvent.click(menuTrigger)
    fireEvent.pointerDown(document.body)
    expect(screen.queryByRole('menu')).toBeNull()

    fireEvent.click(menuTrigger)
    fireEvent.click(screen.getByRole('menuitem', { name: '在侧边栏打开' }))
    expect(openDetails).toHaveBeenCalledOnce()
  })
})

// @vitest-environment jsdom
import { render } from '@testing-library/react'
import type { ComponentProps } from 'react'
import { describe, expect, it, vi } from 'vitest'
import { OfficeDetailsPanel } from '../src/client/OfficeDetailsPanel.tsx'

describe('OfficeDetailsPanel right-Sidebar lifecycle', () => {
  it('binds the native tab and reports when DSH closes it', () => {
    const controller = new AbortController()
    const panelOpened = vi.fn()
    const panelClosed = vi.fn()
    const props = {
      useTabInfo: () => ({ tab: { id: 'office-tab', signal: controller.signal } }),
      useWorkspace: (select: (value: unknown) => unknown) => select({
        status: 'ready', fileName: 'notes.docx', docType: 1, origin: 'workspace',
        readOnly: false, operation: null, hasPotentialEdits: false,
      }),
      useOfficeSession: (select: (value: unknown) => unknown) => select({
        current: {
          source: {
            kind: 'workspace', path: 'notes.docx', cwd: '/workspace', fileName: 'notes.docx',
            url: '/fylar-office-files', docType: 1,
          },
          presentation: 'details',
        },
      }),
      attach: vi.fn(), detach: vi.fn(), retrySession: vi.fn(), setReadOnly: vi.fn(),
      exportDocument: vi.fn(), exportPdf: vi.fn(), clearIssue: vi.fn(), markPotentialEdits: vi.fn(),
      closePanel: vi.fn(), panelOpened, panelClosed,
    } as unknown as ComponentProps<typeof OfficeDetailsPanel>

    const view = render(<OfficeDetailsPanel {...props} />)
    expect(panelOpened).toHaveBeenCalledWith('office-tab')

    controller.abort()
    expect(panelClosed).toHaveBeenCalledWith('office-tab')
    view.unmount()
  })
})

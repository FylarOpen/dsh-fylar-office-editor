import { describe, expect, it, vi } from 'vitest'
import {
  OFFICE_DELIVERABLES_KIND,
  officeDeliverablesDefinition,
} from '../src/client/OfficeDeliverables.tsx'

vi.mock('@deepseek-ai/dsh-session/surface', () => ({
  isAppendSurfaceEvent: () => true,
}))

describe('Office deliverables conversation definition', () => {
  it('tracks both generated DOCX and existing Office-file publication calls', () => {
    const match = officeDeliverablesDefinition.match
    expect(match?.({
      type: 'tool/call',
      data: { turn: 3, name: 'office_create_docx' },
    } as never)).toEqual({ id: '3', role: 'update' })
    expect(match?.({
      type: 'tool/call',
      data: { turn: 3, name: 'office_present_file' },
    } as never)).toEqual({ id: '3', role: 'update' })
    expect(match?.({
      type: 'tool/call',
      data: { turn: 3, name: 'bash' },
    } as never)).toBeNull()
  })

  it('publishes Turn data under its owned kind', () => {
    const data = officeDeliverablesDefinition.buildLocationData?.({
      key: `${OFFICE_DELIVERABLES_KIND}\u00001`,
      kind: OFFICE_DELIVERABLES_KIND,
      id: '1',
      matches: [],
      start: undefined,
      state: { turn: 1, calls: new Map(), produced: [] },
      current: new Map(),
    }, 'turn', null)

    expect(data).toEqual({
      kind: 'turn',
      turn: 1,
      key: OFFICE_DELIVERABLES_KIND,
      value: { produced: [] },
    })
    expect(data?.key).toBe(officeDeliverablesDefinition.kind)
  })

  it('uses the Tool result message outcome to exclude failed files', () => {
    const state = { turn: 1, calls: new Map([['c1', 'report.docx']]), produced: [] }
    const result = (isError: boolean) => ({
      event: {
        type: 'tool/result', seq: 7,
        data: { message: { source: { callId: 'c1' }, content: [{ type: 'text', text: 'result' }], isError } },
      },
    })

    expect(officeDeliverablesDefinition.update({ state } as never, result(true) as never)).toBe(state)
    expect(officeDeliverablesDefinition.update({ state } as never, result(false) as never).produced)
      .toEqual([{ path: 'report.docx', seq: 7 }])
  })
})

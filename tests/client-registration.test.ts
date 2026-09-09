// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest'
import { apply, inject } from '../src/client/index.ts'
import { en, OFFICE_LOCALE_NS, zh } from '../src/client/locales.ts'

vi.mock('@deepseek-ai/dsh-session/surface', () => ({
  isAppendSurfaceEvent: () => true,
}))

interface SlotRegistration {
  readonly name: string
  readonly key?: string
  readonly id?: string
  readonly order?: number
  readonly label?: string
  readonly locale?: string
  readonly inject?: () => unknown
}

describe('Office client registration', () => {
  it('registers the fixed Office view and shares one Workspace source with Chat delivery', async () => {
    const registrations: SlotRegistration[] = []
    const cleanups: Array<() => void> = []
    const slots = {
      register: vi.fn((registration: SlotRegistration) => {
        registrations.push(registration)
        return () => undefined
      }),
      inject: vi.fn((_name: string, register: () => void) => { register() }),
    }
    const locale = { register: vi.fn(() => () => undefined) }
    const conversationRegister = vi.fn(() => () => undefined)
    const ctx = {
      slots,
      locale,
      layout: { openDetails: vi.fn(), closeDetails: vi.fn(), toggleSidebar: vi.fn() },
      uiConversation: { events: { register: conversationRegister } },
      effect: vi.fn((setup: () => void | (() => void)) => {
        const cleanup = setup()
        if (typeof cleanup === 'function') cleanups.push(cleanup)
      }),
    }

    apply(ctx as never)

    expect(inject).toContain('locale')
    expect(inject).toContain('uiConversation')
    expect(inject).not.toContain('conversationEvents')
    expect(locale.register).toHaveBeenCalledWith(OFFICE_LOCALE_NS, { zh, en })
    const officeView = registrations.find(item => item.name === 'conversation.view')
    expect(officeView).toMatchObject({ id: 'office', order: 5, label: 'Office', locale: OFFICE_LOCALE_NS })
    const delivery = registrations.find(item => item.name === 'conversation.chat.turnTail')
    if (officeView?.inject === undefined || delivery?.inject === undefined) {
      throw new Error('Office view or Chat delivery injection is missing')
    }
    const viewInjected = officeView.inject() as { hooks: { workspace: unknown } }
    const deliveryInjected = delivery.inject() as { hooks: { workspace: unknown }; preload?: unknown; open?: unknown }
    expect(viewInjected.hooks.workspace).toBe(deliveryInjected.hooks.workspace)
    expect(deliveryInjected.preload).toBeUndefined()
    expect(deliveryInjected.open).toBeUndefined()
    expect(conversationRegister).toHaveBeenCalledOnce()
    expect(registrations.filter(item => item.name === 'tool.call.toolview').map(item => item.key)).toEqual([
      'office_create_docx',
      'office_present_file',
    ])
    expect(registrations.filter(item => item.name === 'tool.call.toolview').every(item => item.locale === OFFICE_LOCALE_NS)).toBe(true)
    expect(delivery).toMatchObject({ locale: OFFICE_LOCALE_NS })
    expect(registrations.some(item => item.name === 'shell.overlay')).toBe(false)

    for (const cleanup of cleanups.reverse()) cleanup()
    await Promise.resolve()
  })

})

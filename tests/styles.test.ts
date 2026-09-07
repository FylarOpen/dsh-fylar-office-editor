// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { installOfficeStyles, OFFICE_STYLES } from '../src/client/styles.ts'

describe('Office styles', () => {
  it('installs one style element and disposes its owner', () => {
    const dispose = installOfficeStyles(document)
    const style = document.querySelector<HTMLStyleElement>('style[data-dsh-fylar-office-editor]')
    expect(style?.textContent).toBe(OFFICE_STYLES)

    const disposeDuplicate = installOfficeStyles(document)
    expect(document.querySelectorAll('style[data-dsh-fylar-office-editor]')).toHaveLength(1)
    disposeDuplicate()
    expect(style?.isConnected).toBe(true)

    dispose()
    expect(style?.isConnected).toBe(false)
  })
})

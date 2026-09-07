import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  OFFICE_SDK_LICENSE_ENDPOINT, sdkZipLibRelativePath, sdkZipRootPrefix,
  validateLicenseDefaults, verifyNoticeCopies,
} from '../scripts/lib/sdk-integrity.mjs'

describe('vendored Office SDK integrity', () => {
  it('rejects drift in separately packed root legal notices', async () => {
    const vendorRoot = await mkdtemp(join(tmpdir(), 'fylar-office-notices-'))
    try {
      await mkdir(resolve(vendorRoot, 'lib'))
      for (const path of ['legal.txt', 'THIRD_PARTY_NOTICES']) {
        await writeFile(resolve(vendorRoot, path), `original ${path}`)
        await writeFile(resolve(vendorRoot, 'lib', path), `original ${path}`)
      }
      await expect(verifyNoticeCopies(vendorRoot)).resolves.toBeUndefined()
      await writeFile(resolve(vendorRoot, 'legal.txt'), 'changed')
      await expect(verifyNoticeCopies(vendorRoot)).rejects.toThrow(/root notice differs/)
    } finally {
      await rm(vendorRoot, { recursive: true, force: true })
    }
  })

  it('allows only the public license endpoint and empty auxiliary values', () => {
    expect(validateLicenseDefaults(`
      const a = "${OFFICE_SDK_LICENSE_ENDPOINT}"
      const b = ""
      const c = ''
      const d = ""
      export default { a, b, c, d }
    `).get('a')).toBe(OFFICE_SDK_LICENSE_ENDPOINT)
    expect(() => validateLicenseDefaults('const a = "https://example.com"; export default { a }'))
      .toThrow(/license endpoint/)
    expect(() => validateLicenseDefaults(`
      const a = "${OFFICE_SDK_LICENSE_ENDPOINT}"
      const b = "embedded-credential"
      export default { a, b }
    `)).toThrow(/must not embed/)
    expect(() => validateLicenseDefaults(`
      const a = "${OFFICE_SDK_LICENSE_ENDPOINT}"
      const secret = getCredential()
      export default { a, secret }
    `)).toThrow(/unsupported code/)
  })

  it('normalizes zip separators and rejects traversal entries', () => {
    expect(sdkZipLibRelativePath('lib/worker/index.js')).toBe('worker/index.js')
    expect(sdkZipLibRelativePath('external(1)/lib/worker/index.js', 'external(1)/'))
      .toBe('worker/index.js')
    expect(sdkZipLibRelativePath('unrelated/readme.txt')).toBeNull()
    expect(() => sdkZipLibRelativePath('lib/worker\\..\\escape.js')).toThrow(/Unsafe SDK zip path/)
    expect(() => sdkZipLibRelativePath('lib/./escape.js')).toThrow(/Unsafe SDK zip path/)
  })

  it('discovers a rootless or singly wrapped SDK zip and rejects ambiguous roots', () => {
    expect(sdkZipRootPrefix(['lib/package.json', 'legal.txt'])).toBe('')
    expect(sdkZipRootPrefix(['external(1)/lib/package.json', 'external(1)/legal.txt']))
      .toBe('external(1)/')
    expect(() => sdkZipRootPrefix(['a/lib/package.json', 'b/lib/package.json']))
      .toThrow(/multiple candidate/)
    expect(() => sdkZipRootPrefix(['../lib/package.json'])).toThrow(/Unsafe SDK zip path/)
  })
})

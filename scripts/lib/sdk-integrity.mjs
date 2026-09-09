import { createHash } from 'node:crypto'
import { lstat, readFile, readdir } from 'node:fs/promises'
import { relative, resolve, sep } from 'node:path'

export const REQUIRED_SDK_PATHS = [
  'UI.js',
  'UI.d.ts',
  'style.css',
  'package.json',
  'manifest.json',
  'UI.runtime.js',
  'ApiBase.js',
  'licenseDefault.js',
  'mountDocxApp.js',
  'mountXlsxApp.js',
  'mountPptxApp.js',
  'worker/index.js',
  'worker/sharedWorkerStorage.worker.js',
  'worker/wasm/bootloader.wasm',
  'worker/wasm/canvaskit.wasm',
  'worker/wasm/cfb2ooxml.wasm',
  'worker/wasm/emfjs.wasm',
  'worker/wasm/omath.wasm',
  'worker/wasm/ooxmlcrypt.wasm',
  'worker/wasm/vml2dml.wasm',
  'worker/wasm/wordart_generate.wasm',
  'locale-assets/locale.js',
  'locale-assets/deploy/sample/blank.docx',
  'locale-assets/deploy/sample/blank.xlsx',
  'locale-assets/deploy/sample/blank.pptx',
  'vendor/xregexp-all-min.js',
  'vendor/zepto.min.js',
  'legal.txt',
  'THIRD_PARTY_NOTICES',
]

export const OFFICE_SDK_LICENSE_ENDPOINT = 'https://openlicense.fylar.ai'

export function sha256(data) {
  return createHash('sha256').update(data).digest('hex')
}

/** Find the optional wrapper directory containing the archive's unique `lib/` tree. */
export function sdkZipRootPrefix(rawNames) {
  const suffix = 'lib/package.json'
  const candidates = new Set()
  for (const rawName of rawNames) {
    const normalized = rawName.replaceAll('\\', '/')
    if (!normalized.endsWith(suffix)) continue
    const prefix = normalized.slice(0, -suffix.length)
    const segments = prefix === ''
      ? []
      : (prefix.endsWith('/') ? prefix.slice(0, -1).split('/') : prefix.split('/'))
    if (
      prefix.startsWith('/')
      || /^[A-Za-z]:\//.test(prefix)
      || segments.some(segment => segment === '' || segment === '.' || segment === '..')
    ) throw new Error(`Unsafe SDK zip path: ${rawName}`)
    candidates.add(prefix)
  }
  if (candidates.size === 0) throw new Error('SDK archive does not contain lib/package.json')
  if (candidates.size > 1) throw new Error('SDK archive contains multiple candidate lib/ roots')
  return candidates.values().next().value
}

/** Normalize one zip entry and return its safe path inside the selected `lib/`, or null when unrelated. */
export function sdkZipLibRelativePath(rawName, rootPrefix = '') {
  const normalized = rawName.replaceAll('\\', '/')
  const libPrefix = `${rootPrefix}lib/`
  if (!normalized.startsWith(libPrefix) || normalized.endsWith('/')) return null
  const path = normalized.slice(libPrefix.length)
  if (path.startsWith('/') || path.split('/').some(segment => segment === '' || segment === '.' || segment === '..')) {
    throw new Error(`Unsafe SDK zip path: ${rawName}`)
  }
  return path
}

export async function listRegularFiles(root) {
  const files = []
  async function visit(directory) {
    for (const entry of await readdir(directory, { withFileTypes: true })) {
      const absolute = resolve(directory, entry.name)
      if (entry.isSymbolicLink()) throw new Error(`SDK tree must not contain symlinks: ${absolute}`)
      if (entry.isDirectory()) await visit(absolute)
      else if (entry.isFile()) files.push(relative(root, absolute).split(sep).join('/'))
      else throw new Error(`SDK tree contains unsupported entry: ${absolute}`)
    }
  }
  await visit(root)
  return files.sort()
}

export async function buildFileInventory(root) {
  const files = await listRegularFiles(root)
  return Promise.all(files.map(async (path) => {
    const body = await readFile(resolve(root, path))
    return { path, size: body.byteLength, sha256: sha256(body) }
  }))
}

/**
 * Keep browser-shipped license defaults free of embedded credentials.
 * The endpoint is public configuration; every other string constant must be empty.
 */
export function validateLicenseDefaults(source) {
  const declarationPattern = /\bconst\s+([A-Za-z_$][\w$]*)\s*=\s*(["'])(.*?)\2\s*;?/g
  const exportPattern = /\bexport\s+default\s*\{\s*([A-Za-z_$][\w$]*(?:\s*,\s*[A-Za-z_$][\w$]*)*\s*,?)\s*\}\s*;?/g
  const assignments = [...source.matchAll(declarationPattern)]
  const values = new Map(assignments.map(([, name, , value]) => [name, value]))
  const exports = [...source.matchAll(exportPattern)]
  const remainder = source.replace(declarationPattern, '').replace(exportPattern, '').trim()
  if (assignments.length !== values.size || exports.length !== 1 || remainder !== '') {
    throw new Error('Office SDK license defaults contain unsupported code or duplicate declarations')
  }
  if (values.get('a') !== OFFICE_SDK_LICENSE_ENDPOINT) {
    throw new Error(`Office SDK license endpoint must be ${OFFICE_SDK_LICENSE_ENDPOINT}`)
  }
  for (const [name, value] of values) {
    if (!['a', 'b', 'c', 'd'].includes(name)) {
      throw new Error(`Office SDK license defaults contain unexpected field: ${name}`)
    }
    if (name !== 'a' && value !== '') {
      throw new Error(`Office SDK license defaults must not embed a non-empty ${name} value`)
    }
  }
  const exportedNames = exports[0][1].split(',').map(name => name.trim()).filter(Boolean)
  if (
    exportedNames.length !== values.size
    || exportedNames.some(name => !values.has(name))
  ) throw new Error('Office SDK license default export does not match its declarations')
  return values
}

export async function validateRequiredFiles(root) {
  for (const path of REQUIRED_SDK_PATHS) {
    const stat = await lstat(resolve(root, path)).catch(() => null)
    if (stat === null || !stat.isFile()) throw new Error(`Office SDK is missing required file: ${path}`)
  }
  const manifest = JSON.parse(await readFile(resolve(root, 'manifest.json'), 'utf8'))
  const metadata = JSON.parse(await readFile(resolve(root, 'package.json'), 'utf8'))
  if (metadata.version !== manifest.version) {
    throw new Error(`Office SDK version mismatch: package=${metadata.version}, manifest=${manifest.version}`)
  }
  if (metadata.name !== '@office/office-sdk') {
    throw new Error(`Unexpected Office SDK package name: ${String(metadata.name)}`)
  }
  validateLicenseDefaults(await readFile(resolve(root, 'licenseDefault.js'), 'utf8'))
  return { metadata, manifest }
}

/** Ensure the separately packed root notices remain byte-identical to the locked SDK copies. */
export async function verifyNoticeCopies(vendorRoot) {
  for (const path of ['legal.txt', 'THIRD_PARTY_NOTICES']) {
    const [rootCopy, lockedCopy] = await Promise.all([
      readFile(resolve(vendorRoot, path)),
      readFile(resolve(vendorRoot, 'lib', path)),
    ])
    if (!rootCopy.equals(lockedCopy)) {
      throw new Error(`Office SDK root notice differs from locked lib copy: ${path}`)
    }
  }
}

export async function verifyLock(vendorRoot) {
  const libRoot = resolve(vendorRoot, 'lib')
  const { metadata } = await validateRequiredFiles(libRoot)
  await verifyNoticeCopies(vendorRoot)
  const lock = JSON.parse(await readFile(resolve(vendorRoot, 'LOCK.json'), 'utf8'))
  if (lock.sdkVersion !== metadata.version) {
    throw new Error(`Office SDK lock version mismatch: lock=${lock.sdkVersion}, package=${metadata.version}`)
  }
  const actual = await buildFileInventory(libRoot)
  if (JSON.stringify(actual) !== JSON.stringify(lock.files)) {
    throw new Error('Office SDK file inventory does not match LOCK.json')
  }
  return { lock, metadata }
}

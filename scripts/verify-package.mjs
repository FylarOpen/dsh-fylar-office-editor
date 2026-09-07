import { execFileSync } from 'node:child_process'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const npmCache = mkdtempSync(join(tmpdir(), 'fylar-office-npm-cache-'))
let output
try {
  output = execFileSync('npm', ['pack', '--dry-run', '--json', '--ignore-scripts'], {
    cwd: new URL('..', import.meta.url),
    encoding: 'utf8',
    env: { ...process.env, npm_config_cache: npmCache },
  })
} finally {
  rmSync(npmCache, { recursive: true, force: true })
}
const [result] = JSON.parse(output)
if (result === undefined || !Array.isArray(result.files)) throw new Error('npm pack did not return a file inventory')

const paths = result.files.map(file => file.path).sort()
const required = [
  'NOTICE.md',
  'README.md',
  'README.zh-CN.md',
  'cordis.patch.yml',
  'lib/client.js',
  'lib/index.js',
  'lib/client.js.map',
  'lib/index.js.map',
  'lib/types/client/index.d.ts',
  'lib/types/index.d.ts',
  'package.json',
  'vendor/office-sdk/LOCK.json',
  'vendor/office-sdk/THIRD_PARTY_NOTICES',
  'vendor/office-sdk/legal.txt',
  'vendor/office-sdk/lib/UI.js',
  'vendor/office-sdk/lib/worker/index.js',
  'vendor/office-sdk/lib/worker/wasm/canvaskit.wasm',
]
for (const path of required) {
  if (!paths.includes(path)) throw new Error(`Packed artifact is missing: ${path}`)
}
for (const path of paths) {
  const allowed = path === 'package.json'
    || path === 'README.md'
    || path === 'README.zh-CN.md'
    || path === 'NOTICE.md'
    || path === 'cordis.patch.yml'
    || path.startsWith('lib/')
    || path.startsWith('vendor/office-sdk/')
  if (!allowed) throw new Error(`Unexpected packed file: ${path}`)
  if (path.endsWith('.tgz')) throw new Error(`Nested tarball in package: ${path}`)
}
console.log(`Verified npm pack allowlist: ${paths.length} files, ${result.size} bytes packed`)

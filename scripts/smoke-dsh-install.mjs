import { execFileSync } from 'node:child_process'
import { mkdtempSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

const projectRoot = resolve(new URL('..', import.meta.url).pathname)
const manifest = JSON.parse(readFileSync(resolve(projectRoot, 'package.json'), 'utf8'))
const tarballName = `${manifest.name.replace(/^@/u, '').replace('/', '-')}-${manifest.version}.tgz`
const tarball = resolve(projectRoot, tarballName)
const temporaryRoot = mkdtempSync(join(tmpdir(), 'fylar-office-dsh-smoke-'))
const dshHome = join(temporaryRoot, 'home')
const npmCache = join(temporaryRoot, 'npm-cache')
mkdirSync(dshHome, { recursive: true })

const executable = process.platform === 'win32' ? 'npx.cmd' : 'npx'
const channel = process.argv[2] ?? 'latest'
const sourceRoot = process.env.DSH_SOURCE_ROOT
const localSource = sourceRoot !== undefined && sourceRoot !== ''
if (!localSource && !['latest', 'alpha'].includes(channel)) {
  throw new Error(`Unsupported DSH npm channel: ${channel}`)
}
const command = localSource ? (process.platform === 'win32' ? 'pnpm.cmd' : 'pnpm') : executable
const prefix = localSource ? ['dsh'] : ['--yes', '--package', `@deepseek-ai/dsh@${channel}`, 'dsh']
const environment = {
  ...process.env,
  DSH_HOME: dshHome,
  npm_config_cache: npmCache,
  npm_config_registry: process.env.npm_config_registry ?? 'https://registry.npmjs.org/',
}

function run(arguments_) {
  return execFileSync(command, [...prefix, ...arguments_], {
    cwd: localSource ? sourceRoot : projectRoot,
    env: environment,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
  })
}

try {
  run(['plugin', '--profile', 'web', 'add', tarball])
  const config = run(['--profile', 'web', '--dump-config'])
  if (!config.includes('dsh-fylar-office-editor')) {
    throw new Error('DSH composed config does not contain the Fylar Office bundle and plugin row')
  }
  console.log(localSource
    ? `Verified local tgz install with DSH source at ${sourceRoot}`
    : `Verified local tgz install with the DSH npm ${channel} channel`)
} finally {
  rmSync(temporaryRoot, { recursive: true, force: true })
}

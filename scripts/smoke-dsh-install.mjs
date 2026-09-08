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
if (!['latest', 'alpha'].includes(channel)) {
  throw new Error(`Unsupported DSH npm channel: ${channel}`)
}
const prefix = ['--yes', '--package', `@deepseek-ai/dsh@${channel}`, 'dsh']
const environment = { ...process.env, DSH_HOME: dshHome, npm_config_cache: npmCache }

function run(arguments_) {
  return execFileSync(executable, [...prefix, ...arguments_], {
    cwd: projectRoot,
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
  console.log(`Verified local tgz install with the DSH npm ${channel} channel`)
} finally {
  rmSync(temporaryRoot, { recursive: true, force: true })
}

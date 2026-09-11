import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { createRequire } from 'node:module'
import { runInNewContext } from 'node:vm'
import * as host from '../lib/index.js'

const PACKAGE_NAME = '@fylar/dsh-fylar-office-editor'
const PLUGIN_ID = 'dsh-fylar-office-editor'
const client = await readFile(new URL('../lib/client.js', import.meta.url), 'utf8')
const manifest = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'))
const require = createRequire(import.meta.url)
const platformModules = new Set([
  'react',
  'react/jsx-runtime',
  'react-dom',
  'react-dom/client',
  '@deepseek-ai/cordis',
  '@deepseek-ai/dsh-client-store',
  '@deepseek-ai/dsh-client-ui-slots',
  '@deepseek-ai/dsh-client-ui-primitives',
])

const registrations = []
runInNewContext(client, {
  window: { __ModuleLoader__: { load: registration => registrations.push(registration) } },
}, { filename: 'lib/client.js' })
assert.equal(registrations.length, 1, 'Client must register one DSH module')
const [registration] = registrations
assert.equal(manifest.name, PACKAGE_NAME)
assert.equal(registration.id, PACKAGE_NAME)
const plugin = registration.factory(request => {
  assert.ok(platformModules.has(request), `Client requests an undeclared DSH module: ${request}`)
  return require(request)
})
assert.equal(typeof plugin.apply, 'function', 'Client must export apply')
assert.ok(Array.isArray(plugin.inject), 'Client must declare its injected services')
assert.equal(host.name, PLUGIN_ID)
assert.equal(typeof host.apply, 'function', 'Host must export apply')

assert.equal(manifest.dsh?.bundle?.patch, './cordis.patch.yml')
assert.equal(manifest.dsh?.client?.external?.length ?? 0, 0, 'Non-platform Client dependencies must be bundled')
assert.ok(manifest.dsh?.client?.inject?.includes('@deepseek-ai/dsh-client-locale'), 'Client requires the DSH locale service')
assert.ok(manifest.dsh?.client?.inject?.includes('@deepseek-ai/dsh-client-ui-chat'), 'Client requires the DSH Chat slots')
assert.ok(manifest.dsh?.client?.inject?.includes('@deepseek-ai/dsh-client-ui-conversation'), 'Client requires the DSH Conversation service')
assert.ok(manifest.dsh?.client?.inject?.includes('@deepseek-ai/dsh-client-ui-sidebar-right'), 'Client requires the DSH right-Sidebar service')
assert.ok(plugin.inject.includes('sidebarRight'), 'Client must inject the right-Sidebar controller')
assert.ok(plugin.inject.includes('sidebarRightTabs'), 'Client must inject the right-Sidebar tab registry')
for (const lifecycle of ['prepare', 'prepack', 'preinstall', 'install', 'postinstall']) {
  assert.equal(manifest.scripts?.[lifecycle], undefined, `Repository installs must not execute ${lifecycle}`)
}
for (const forbidden of ['office.core.js', 'mountDocxApp.js', 'canvaskit.wasm']) {
  assert.ok(!client.includes(forbidden), `Client bundled Office runtime content: ${forbidden}`)
}

console.log('Verified Host and Client entry points and DSH module loading')

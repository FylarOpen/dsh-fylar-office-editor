import assert from 'node:assert/strict'
import { fileURLToPath } from 'node:url'
import {
  OFFICE_SDK_VERSION, OFFICE_SDK_REVISION, OFFICE_SDK_ROUTE_PREFIX, OFFICE_SDK_BASE_PATH,
} from '../src/office-sdk.generated.ts'
import { verifyLock } from './lib/sdk-integrity.mjs'

const vendorRoot = fileURLToPath(new URL('../vendor/office-sdk', import.meta.url))
const { lock, metadata } = await verifyLock(vendorRoot)
const revision = lock.sourceSha256.slice(0, 12)
assert.equal(OFFICE_SDK_VERSION, metadata.version, 'Generated Office SDK version is stale')
assert.equal(OFFICE_SDK_REVISION, revision, 'Generated Office SDK revision is stale')
assert.equal(OFFICE_SDK_ROUTE_PREFIX, '/fylar-office-sdk')
assert.equal(OFFICE_SDK_BASE_PATH, `${OFFICE_SDK_ROUTE_PREFIX}/${metadata.version}-${revision}`)
console.log(`Verified Office SDK ${metadata.version}: ${lock.files.length} files`)

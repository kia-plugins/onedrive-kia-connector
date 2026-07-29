/**
 * Smoke test for the bundled dist/index.js — proves the CJS/ESM dual export
 * in src/index.ts (`export default mod; module.exports = mod;`) against
 * actual esbuild output. Build + require + activate() are the SDK kit's
 * shared `bundleLoadSmoke`; only the ids below are this connector's.
 */
import { join } from 'node:path';
import { bundleLoadSmoke } from '@kiagent/connector-sdk/testing';

describe('dist bundle loads standalone', () => {
  it('require()s dist/index.js and activate() returns the onedrive source', async () => {
    await bundleLoadSmoke({
      root: join(__dirname, '..', '..'),
      selfId: 'kia.onedrive',
      sourceIds: ['onedrive'],
    });
  }, 30_000);
});

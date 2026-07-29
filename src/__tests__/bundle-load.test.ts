/**
 * Smoke tests for the bundled dist/index.js — proves the CJS/ESM dual export
 * in src/index.ts (`export default mod; module.exports = mod;`) against
 * actual esbuild output.
 *
 * Two `it`s on purpose. The first is the SDK kit's shared `bundleLoadSmoke`
 * (build + require + activate, contributed source ids). The second keeps the
 * two properties the kit does not model, both guarded here before the SDK
 * migration: that the bundle loads in a BARE node process — outside jest's
 * module registry, the way the extension-host child require()s it — and that
 * the descriptor still declares `auth: 'oauth'`, the mode manifest.json's
 * `{ id: 'onedrive', oauth: 'microsoft' }` contribution is paired with.
 */
import { execSync } from 'node:child_process';
import { join } from 'node:path';
import { bundleLoadSmoke } from '@kiagent/connector-sdk/testing';
import type { HostFor, Query } from '@kiagent/connector-sdk';

describe('dist bundle loads standalone', () => {
  it('require()s dist/index.js and activate() returns the onedrive source', async () => {
    await bundleLoadSmoke({
      root: join(__dirname, '..', '..'),
      selfId: 'kia.onedrive',
      sourceIds: ['onedrive'],
    });
  }, 30_000);

  it('loads in a bare node process and keeps the oauth auth mode on the descriptor', async () => {
    const root = join(__dirname, '..', '..');
    execSync('npm run build', { cwd: root });

    // Bare Node child-process require — proves the bundle loads outside
    // jest's module registry, the way the extension host child loads it.
    const out = execSync(
      'node -e "const m=require(\'./dist/index.js\');const e=m.default??m;' +
        "if(typeof e.activate!=='function')throw new Error('no activate');" +
        'console.log(\'activate:\'+typeof e.activate)"',
      { cwd: root },
    ).toString();
    expect(out).toContain('activate:function');

    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const mod = require(join(root, 'dist', 'index.js'));
    const entry = mod.default ?? mod;
    expect(typeof entry.activate).toBe('function');

    const unused = () => {
      throw new Error('unused in this smoke test');
    };
    const host: HostFor<'net' | 'query'> = {
      self: { id: 'kia.onedrive', dataDir: '/tmp' },
      log: () => {},
      net: { fetch: unused },
      query: {
        document: unused,
        children: unused,
        byExternalId: unused,
        search: unused,
        count: unused,
        accounts: unused,
      } as unknown as Query,
    };
    const result = await entry.activate(host);

    expect(result.sources?.[0]?.descriptor.auth).toBe('oauth');
  }, 30_000);
});

/**
 * reauthenticate(account, auth): re-auth THIS account, verified against the
 * Graph identity. Replaces connect()'s deleted needsReauth lookup heuristic,
 * which could only guess its target by (source, identifier, status) and
 * silently hijacked the first `.find()` match.
 */
import { FILES_SCOPES, createOneDriveSource } from '../source';
import { sourceErrorCode } from '@kiagent/connector-sdk';
import { fakeAccount, graphFetch, instantClock, makeAuth, makeHost } from '../testing/harness';

function makeSource(world: Parameters<typeof graphFetch>[0] = {}) {
  const { fetchFn, calls } = graphFetch(world);
  return { source: createOneDriveSource(makeHost(fetchFn), instantClock), calls };
}

describe('reauthenticate', () => {
  it('accepts the matching identity, requests FILES_SCOPES, and resolves with nothing', async () => {
    const { source } = makeSource({ about: { mail: 'ed@example.com' } });
    const { auth, statuses, getScopes } = makeAuth();

    await expect(
      source.reauthenticate!(fakeAccount({ identifier: 'ed@example.com' }), auth),
    ).resolves.toBeUndefined();

    expect(getScopes()).toEqual(FILES_SCOPES);
    expect(statuses).toEqual(['Waiting for Microsoft sign-in…', 'Fetching Microsoft profile…']);
  });

  it('matches on EITHER Graph identity field, trimmed and case-folded', async () => {
    const { source } = makeSource({ about: { userPrincipalName: 'ED@Tenant.onmicrosoft.com' } });
    const { auth } = makeAuth();

    await expect(
      source.reauthenticate!(fakeAccount({ identifier: '  ed@tenant.onmicrosoft.com ' }), auth),
    ).resolves.toBeUndefined();
  });

  it('rejects a different Microsoft identity as a PERMANENT error, not a retryable one', async () => {
    const { source } = makeSource({ about: { mail: 'someone-else@example.com' } });
    const { auth } = makeAuth();

    const err = await source
      .reauthenticate!(fakeAccount({ identifier: 'ed@example.com' }), auth)
      .then(() => null)
      .catch((e: unknown) => e);

    expect(String(err)).toMatch(/signed in as someone-else@example\.com/);
    expect(String(err)).toMatch(/this account is ed@example\.com/);
    expect(sourceErrorCode(err)).toBe('permanent');
  });

  it('throws before any fetch when oauth returns no access token', async () => {
    const { source, calls } = makeSource({});
    const { auth } = makeAuth({ creds: { refreshToken: 'ms-test-refresh-deadbeef' } });

    await expect(source.reauthenticate!(fakeAccount(), auth)).rejects.toThrow(/no access token/);
    expect(calls).toHaveLength(0);
  });

  it('throws when Graph /me is missing both mail and userPrincipalName', async () => {
    const { source } = makeSource({ about: {} });
    const { auth } = makeAuth();

    await expect(source.reauthenticate!(fakeAccount(), auth)).rejects.toThrow(
      /missing mail and userPrincipalName/,
    );
  });

  it('never opens the folder picker — reconnect can only lose scope, never change it', async () => {
    const { source } = makeSource({ about: { mail: 'ed@example.com' } });
    const { auth, getPickerSpec } = makeAuth();

    await source.reauthenticate!(fakeAccount({ identifier: 'ed@example.com' }), auth);

    expect(getPickerSpec()).toBeUndefined();
  });
});

/**
 * connect(auth) suite: platform-owned Microsoft OAuth (auth.oauth with
 * FILES_SCOPES), the Graph /me profile fetch, and the shared folder-picker
 * (auth.pickFolders) — multi-root selection, spec wiring, empty selection,
 * and cancel propagation.
 */
import { createOneDriveSource, FILES_SCOPES } from '../source';
import {
  fakeAccount,
  fakeQuery,
  graphFetch,
  instantClock,
  makeAuth,
  makeHost,
} from '../testing/harness';

describe('connect', () => {
  it('oauth happy path: scope, statuses, identifier = mail, picked folders → folderRoots config', async () => {
    const { fetchFn, calls } = graphFetch({ about: { mail: 'ed@example.com' } });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth, statuses, getScopes } = makeAuth({
      picked: [
        { id: 'FOLD1', name: 'Projects', hasChildren: true },
        { id: 'SH1', name: 'Shared specs', hasChildren: true },
      ],
    });

    const res = await source.connect(auth);

    expect(getScopes()).toEqual(FILES_SCOPES);
    expect(statuses).toEqual(['Waiting for Microsoft sign-in…', 'Fetching Microsoft profile…']);
    expect(res).toEqual({
      identifier: 'ed@example.com',
      config: {
        folderRoots: [
          { id: 'FOLD1', name: 'Projects' },
          { id: 'SH1', name: 'Shared specs' },
        ],
      },
    });
    expect(calls).toHaveLength(1);
    expect(calls[0]).toContain('/v1.0/me?');
  });

  it('falls back to userPrincipalName when mail is absent', async () => {
    const { fetchFn } = graphFetch({ about: { userPrincipalName: 'ed@tenant.onmicrosoft.com' } });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth();

    const res = await source.connect(auth);

    expect(res.identifier).toBe('ed@tenant.onmicrosoft.com');
  });

  it('passes the pinned picker spec: My files + Shared tabs, multiSelect', async () => {
    const { fetchFn } = graphFetch({});
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth, getPickerSpec } = makeAuth();

    await source.connect(auth);

    const spec = getPickerSpec()!;
    expect(spec.modes).toEqual([
      { key: 'my-files', label: 'My files' },
      { key: 'shared', label: 'Shared with me' },
    ]);
    expect(spec.multiSelect).toBe(true);
    expect(spec.purpose).toBe('connect');
    // Connect has nothing to preselect — that is manageFolders' job.
    expect(spec.selected).toBeUndefined();
  });

  it("spec.roots('my-files') is the static OneDrive-root node — no API call", async () => {
    const { fetchFn, calls } = graphFetch({});
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth, getPickerSpec } = makeAuth();
    await source.connect(auth);
    const before = calls.length;

    await expect(getPickerSpec()!.roots('my-files')).resolves.toEqual([
      { id: 'root', name: 'OneDrive', hasChildren: true },
    ]);
    expect(calls).toHaveLength(before);
  });

  it("spec.roots('shared') / children / count run against Graph with the connect token", async () => {
    const { fetchFn } = graphFetch({
      sharedRoots: [{ id: 'SH1', name: 'Shared specs', folder: { childCount: 2 } }],
      children: {
        SH1: [
          { id: 'SUB1', name: 'Sub', folder: { childCount: 1 } },
          { id: 'f1', name: 'a.txt' }, // not a folder — filtered out
        ],
        SUB1: [{ id: 'f2', name: 'b.txt' }],
      },
      items: { SH1: { id: 'SH1', name: 'Shared specs', folder: { childCount: 2 } } },
    });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth, getPickerSpec } = makeAuth();
    await source.connect(auth);
    const spec = getPickerSpec()!;

    await expect(spec.roots('shared')).resolves.toEqual([
      { id: 'SH1', name: 'Shared specs', hasChildren: true },
    ]);
    await expect(spec.children('SH1')).resolves.toEqual([
      { id: 'SUB1', name: 'Sub', hasChildren: true },
    ]);
    await expect(spec.count!('SH1')).resolves.toEqual({ count: 2, capped: false });
  });

  it('throws when the user confirms an empty selection', async () => {
    const { fetchFn } = graphFetch({});
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth({ picked: [] });

    await expect(source.connect(auth)).rejects.toThrow(/no folders selected/);
  });

  it('propagates a pickFolders rejection (user cancelled) out of connect', async () => {
    const { fetchFn } = graphFetch({});
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth({
      picked: async () => {
        throw new Error('picker cancelled');
      },
    });

    await expect(source.connect(auth)).rejects.toThrow(/picker cancelled/);
  });

  it('throws when oauth returns no accessToken, before any fetch', async () => {
    const { fetchFn, calls } = graphFetch({});
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth({ creds: { refreshToken: 'ms-test-refresh-deadbeef' } });
    await expect(source.connect(auth)).rejects.toThrow(/no access token/);
    expect(calls).toHaveLength(0);
  });

  it('throws when Graph /me is missing both mail and userPrincipalName', async () => {
    const { fetchFn } = graphFetch({ about: {} });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth();
    await expect(source.connect(auth)).rejects.toThrow(/missing mail and userPrincipalName/);
  });
});

describe('connect no longer restores by identifier', () => {
  it('a needsReauth account with the same identifier does NOT skip the picker — reconnect is reauthenticate()', async () => {
    const { fetchFn } = graphFetch({ about: { mail: 'ed@example.com' } });
    const prior = fakeAccount({ config: { folderRoots: [{ id: 'OLD1', name: 'Kept' }] } });
    const source = createOneDriveSource(makeHost(fetchFn, fakeQuery([], [prior])), instantClock);
    const { auth, statuses, getPickerSpec } = makeAuth({
      picked: [{ id: 'NEW1', name: 'New', hasChildren: true }],
    });

    const res = await source.connect(auth);

    expect(getPickerSpec()).toBeDefined();
    expect(statuses).toEqual(['Waiting for Microsoft sign-in…', 'Fetching Microsoft profile…']);
    expect(res).toEqual({
      identifier: 'ed@example.com',
      config: { folderRoots: [{ id: 'NEW1', name: 'New' }] },
    });
  });
});

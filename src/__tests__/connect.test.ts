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
  it('oauth happy path: scope, statuses, identifier = mail, picked folders → roots config', async () => {
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
        roots: [
          { rootFolderId: 'FOLD1', rootName: 'Projects' },
          { rootFolderId: 'SH1', rootName: 'Shared specs' },
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

describe('reconnect after auth failure keeps the stored selection', () => {
  const about = { mail: 'ed@example.com' };

  it('skips the picker and returns the prior config verbatim', async () => {
    const { fetchFn, calls } = graphFetch({ about });
    const prior = fakeAccount({
      config: { roots: [{ rootFolderId: 'OLD1', rootName: 'Kept' }] },
    });
    const source = createOneDriveSource(makeHost(fetchFn, fakeQuery([], [prior])), instantClock);
    const { auth, statuses, getPickerSpec } = makeAuth();

    const res = await source.connect(auth);

    expect(res).toEqual({ identifier: 'ed@example.com', config: prior.config });
    expect(getPickerSpec()).toBeUndefined();
    expect(statuses).toEqual([
      'Waiting for Microsoft sign-in…',
      'Fetching Microsoft profile…',
      'Restoring previous folder selection…',
    ]);
    // Only the /me call — no folder lookups on the restore path.
    expect(calls).toHaveLength(1);
  });

  it('restores when the identifier came from the userPrincipalName fallback', async () => {
    const { fetchFn } = graphFetch({ about: { userPrincipalName: 'ed@tenant.onmicrosoft.com' } });
    const prior = fakeAccount({
      identifier: 'ed@tenant.onmicrosoft.com',
      config: { roots: [{ rootFolderId: 'OLD1', rootName: 'Kept' }] },
    });
    const source = createOneDriveSource(makeHost(fetchFn, fakeQuery([], [prior])), instantClock);
    const { auth, getPickerSpec } = makeAuth();

    const res = await source.connect(auth);

    expect(res).toEqual({ identifier: 'ed@tenant.onmicrosoft.com', config: prior.config });
    expect(getPickerSpec()).toBeUndefined();
  });

  it('a healthy account with the same identifier still runs the picker', async () => {
    const { fetchFn } = graphFetch({ about });
    const prior = fakeAccount({
      status: 'live',
      config: { roots: [{ rootFolderId: 'OLD1', rootName: 'Kept' }] },
    });
    const source = createOneDriveSource(makeHost(fetchFn, fakeQuery([], [prior])), instantClock);
    const { auth, getPickerSpec } = makeAuth({
      picked: [{ id: 'NEW1', name: 'New', hasChildren: true }],
    });

    const res = await source.connect(auth);

    expect(getPickerSpec()).toBeDefined();
    expect(res.config).toEqual({ roots: [{ rootFolderId: 'NEW1', rootName: 'New' }] });
  });

  it('a needsReauth account under a different identity does not hijack the flow', async () => {
    const { fetchFn } = graphFetch({ about });
    const prior = fakeAccount({ identifier: 'other@example.com' });
    const source = createOneDriveSource(makeHost(fetchFn, fakeQuery([], [prior])), instantClock);
    const { auth, getPickerSpec } = makeAuth({
      picked: [{ id: 'NEW1', name: 'New', hasChildren: true }],
    });

    const res = await source.connect(auth);

    expect(getPickerSpec()).toBeDefined();
    expect(res.identifier).toBe('ed@example.com');
    expect(res.config).toEqual({ roots: [{ rootFolderId: 'NEW1', rootName: 'New' }] });
  });

  it("another source's needsReauth account with the same identifier is ignored", async () => {
    const { fetchFn } = graphFetch({ about });
    const prior = fakeAccount({ source: 'ms365' });
    const source = createOneDriveSource(makeHost(fetchFn, fakeQuery([], [prior])), instantClock);
    const { auth, getPickerSpec } = makeAuth({
      picked: [{ id: 'NEW1', name: 'New', hasChildren: true }],
    });

    await source.connect(auth);

    expect(getPickerSpec()).toBeDefined();
  });
});

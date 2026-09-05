/**
 * manageFolders(session, channel): edit an existing account's tracked roots
 * with its OWN session credentials — canonical config out (with any legacy
 * `roots` mirror passed through untouched, A-2), a pruned and backfill-free
 * cursor out, the connector-computed archive set (DECISIONS R8) including the
 * shared-ancestor carve-out (spec-reality-diff A5b), and A-3's
 * archiveNullScoped pairing. This connector has no reconcile(), so the archive
 * set computed here is the ONLY removal path OneDrive will ever have:
 * over-archiving is unrecoverable, so every ambiguous case must resolve to
 * "archive nothing".
 */
import {
  GRAPH_BASE,
  ONEDRIVE_DRIVE_ROOT_ID,
  createOneDriveSource,
  type OneDriveCursor,
  type OneDriveItem,
} from '../source';
import type {
  Batch,
  DocumentInput,
  FolderNode,
  FolderScopeUpdate,
} from '@kiagent/connector-sdk';
import {
  collect,
  deltaUrl,
  driveFile,
  driveFolder,
  graphFetch,
  instantClock,
  jsonRes,
  makeFolderChannel,
  makeHost,
  makeSession,
  type GraphWorld,
} from '../testing/harness';

const fr = (...pairs: Array<[string, string]>) => pairs.map(([id, name]) => ({ id, name }));
const node = (id: string, name: string): FolderNode => ({ id, name, hasChildren: true });

/** One stored document as core holds it: an `externalId`, the `scope_root_id`
 *  stamp the emitting pull left on it, and whether it is archived. */
interface Row {
  externalId: string;
  scopeRootId: string;
  archived: boolean;
}
const row = (externalId: string, scopeRootId: string): Row => ({
  externalId,
  scopeRootId,
  archived: false,
});

/**
 * A model of core's `applyFolderScope` write transaction, faithful to the
 * `FolderScopeUpdate` contract: `reattributeScopeRoots` re-stamps
 * `scope_root_id` from → to FIRST, then `archiveScopeRootIds` archives by an
 * `IN`-list over the resulting stamps — one transaction — and a `from` that
 * ALSO appears in `archiveScopeRootIds` THROWS rather than being silently
 * ordered (contracts.ts, C-46/D5). The throw lives here so a contradictory
 * emission fails a test instead of quietly working.
 *
 * The suites drive this instead of a second `pull()` between edits on
 * purpose: C-46 addendum #6 MEASURED that a re-walk does not re-stamp a live
 * row — `hashSkip` (`src/source.ts:296`, gating inside `buildItem` at `:403`)
 * returns before attribution is refreshed — so re-attribution really is the
 * only thing that can move a live row's stamp.
 */
function applyFolderScope(rows: Row[], out: FolderScopeUpdate<OneDriveCursor>): Row[] {
  const archive = new Set(out.archiveScopeRootIds);
  const reattribute = out.reattributeScopeRoots ?? [];
  for (const { from } of reattribute) {
    if (archive.has(from)) {
      throw new Error(`applyFolderScope: '${from}' is in BOTH reattribute and archive`);
    }
  }
  const moves = new Map(reattribute.map((m) => [m.from, m.to]));
  return rows.map((r) => {
    const scopeRootId = moves.get(r.scopeRootId) ?? r.scopeRootId;
    return { ...r, scopeRootId, archived: r.archived || archive.has(scopeRootId) };
  });
}

/** The documents this save was supposed to remove but left searchable: live,
 *  yet stamped with a root the new selection no longer contains. */
const liveOutsideScope = (rows: Row[], config: FolderScopeUpdate<OneDriveCursor>['config']) => {
  const inScope = new Set((config.folderRoots as Array<{ id: string }>).map((r) => r.id));
  return rows.filter((r) => !r.archived && !inScope.has(r.scopeRootId)).map((r) => r.externalId);
};

/** `/me/drive` item fixtures that `resolveRootLocation` reads for containment.
 *  FB = `/Beta/`, FBSUB = `/Beta/Sub/`, FG = `/Gamma/`, SH1 = a shared root
 *  with no `parentReference` at all, i.e. INCOMPARABLE. */
const worldItems: GraphWorld['items'] = {
  FB: driveFolder('FB', 'Beta'),
  FBSUB: driveFolder('FBSUB', 'Sub', {
    parentReference: { driveId: 'drive-1', path: '/drive/root:/Beta' },
  }),
  FG: driveFolder('FG', 'Gamma'),
  // C-46/D4: a SIBLING of FB whose name EXTENDS it — `/BetaBackup/` beside
  // `/Beta/`. Nothing under it is contained by FB.
  FBB: driveFolder('FBB', 'BetaBackup'),
  SH1: { id: 'SH1', name: 'Shared specs', folder: { childCount: 1 } },
};

function makeSource(items: GraphWorld['items'] = worldItems) {
  const { fetchFn, calls } = graphFetch({ items });
  return { source: createOneDriveSource(makeHost(fetchFn), instantClock), calls };
}

describe('manageFolders', () => {
  it('the descriptor advertises folderScope and both new members exist', () => {
    const { source } = makeSource();
    expect(source.descriptor.folderScope).toBe(true);
    expect(typeof source.manageFolders).toBe('function');
    expect(typeof source.reauthenticate).toBe('function');
  });

  it('pre-selects the current roots, returns canonical folderRoots, and costs no Graph lookups on a pure ADD', async () => {
    const { source, calls } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta']) },
      cursor: { delta_tokens: { FB: 'TB' }, scope_roots: ['FB'] } as OneDriveCursor,
    });
    const { channel, statuses, getPickerSpec } = makeFolderChannel({
      picked: [node('FB', 'Beta'), node('FG', 'Gamma')],
    });

    const out = await source.manageFolders!(session, channel);

    const spec = getPickerSpec()!;
    expect(spec.purpose).toBe('manage');
    expect(spec.multiSelect).toBe(true);
    expect(spec.modes).toEqual([
      { key: 'my-files', label: 'My files' },
      { key: 'shared', label: 'Shared with me' },
    ]);
    expect(spec.selected).toEqual([node('FB', 'Beta')]);
    expect(statuses).toEqual(['Loading your OneDrive folders…']);
    expect(out.config).toEqual({ folderRoots: fr(['FB', 'Beta'], ['FG', 'Gamma']) });
    expect(out.cursor).toEqual({ delta_tokens: { FB: 'TB' }, scope_roots: ['FB', 'FG'] });
    expect(out.archiveScopeRootIds).toEqual([]);
    expect(out.archiveNullScoped).toBe(false);
    expect(calls).toEqual([]);
  });

  it('keeps retained roots in their prior config order and appends new ones — order is semantic', async () => {
    const { source } = makeSource();
    const { session } = makeSession({ config: { folderRoots: fr(['FB', 'Beta'], ['FG', 'Gamma']) } });
    const { channel } = makeFolderChannel({
      picked: [node('FG', 'Gamma'), node('FBSUB', 'Sub'), node('FB', 'Beta')],
    });

    const out = await source.manageFolders!(session, channel);

    expect(out.config.folderRoots).toEqual(fr(['FB', 'Beta'], ['FG', 'Gamma'], ['FBSUB', 'Sub']));
    expect(out.cursor!.scope_roots).toEqual(['FB', 'FG', 'FBSUB']);
  });

  it('removing a root that no retained root covers archives it and prunes its delta token', async () => {
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta'], ['FG', 'Gamma']) },
      cursor: { delta_tokens: { FB: 'TB', FG: 'TG' }, scope_roots: ['FB', 'FG'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta')] });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual(['FG']);
    expect(out.cursor).toEqual({ delta_tokens: { FB: 'TB' }, scope_roots: ['FB'] });
    expect(out.archiveNullScoped).toBe(false);
  });

  it('a removed root still covered by a retained ANCESTOR archives NOTHING and drops the ancestor token (A5b)', async () => {
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta'], ['FBSUB', 'Sub']) },
      cursor: { delta_tokens: { FB: 'TB', FBSUB: 'TS' }, scope_roots: ['FB', 'FBSUB'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta')] });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual([]);
    expect(out.cursor).toEqual({ delta_tokens: {}, scope_roots: ['FB'] });
    // Every root re-establishes on the next pull, so the NULL repair is free
    // and recoverable here — A-3's pairing (Step 16's derivation).
    expect(out.archiveNullScoped).toBe(true);
  });

  it('a removed ANCESTOR of a retained root IS archived, and the retained child loses its token so the re-walk un-archives what is still in scope', async () => {
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta'], ['FBSUB', 'Sub']) },
      cursor: { delta_tokens: { FB: 'TB', FBSUB: 'TS' }, scope_roots: ['FB', 'FBSUB'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FBSUB', 'Sub')] });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual(['FB']);
    expect(out.cursor).toEqual({ delta_tokens: {}, scope_roots: ['FBSUB'] });
    expect(out.archiveNullScoped).toBe(true);
  });

  it('the drive-root catch-all covers every My-files root: removing a subfolder of it archives nothing', async () => {
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: [{ id: ONEDRIVE_DRIVE_ROOT_ID, name: 'OneDrive' }, ...fr(['FB', 'Beta'])] },
      cursor: {
        delta_tokens: { [ONEDRIVE_DRIVE_ROOT_ID]: 'TR', FB: 'TB' },
        scope_roots: [ONEDRIVE_DRIVE_ROOT_ID, 'FB'],
      } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node(ONEDRIVE_DRIVE_ROOT_ID, 'OneDrive')] });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual([]);
    expect(out.cursor).toEqual({ delta_tokens: {}, scope_roots: [ONEDRIVE_DRIVE_ROOT_ID] });
    expect(out.archiveNullScoped).toBe(true);
  });

  it('a removed "Shared with me" root is INCOMPARABLE to a My-files root: archived, and no retained token is dropped', async () => {
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta'], ['SH1', 'Shared specs']) },
      cursor: { delta_tokens: { FB: 'TB', SH1: 'TSH' }, scope_roots: ['FB', 'SH1'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta')] });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual(['SH1']);
    expect(out.cursor).toEqual({ delta_tokens: { FB: 'TB' }, scope_roots: ['FB'] });
    expect(out.archiveNullScoped).toBe(false);
  });

  it('a removed root that is GONE upstream (404) is archived rather than failing the save', async () => {
    const { source } = makeSource({ ...worldItems, FG: jsonRes(404, { error: { code: 'itemNotFound' } }) });
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta'], ['FG', 'Gamma']) },
      cursor: { delta_tokens: { FB: 'TB', FG: 'TG' }, scope_roots: ['FB', 'FG'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta')] });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual(['FG']);
  });

  it('a TRANSIENT Graph failure while checking coverage rejects instead of guessing "incomparable"', async () => {
    const { source } = makeSource({ ...worldItems, FG: jsonRes(500, { error: { message: 'boom' } }) });
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta'], ['FG', 'Gamma']) },
      cursor: { delta_tokens: { FB: 'TB', FG: 'TG' }, scope_roots: ['FB', 'FG'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta')] });

    await expect(source.manageFolders!(session, channel)).rejects.toThrow(/^graph 500 /);
  });

  it('never writes a backfill key, even when the account was mid-backfill', async () => {
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta'], ['FG', 'Gamma']) },
      cursor: {
        delta_tokens: { FB: 'TB' },
        scope_roots: ['FB', 'FG'],
        backfill: { root_index: 1, next_link: 'https://graph.microsoft.com/page' },
      } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta')] });

    const out = await source.manageFolders!(session, channel);

    expect(out.cursor).toEqual({ delta_tokens: { FB: 'TB' }, scope_roots: ['FB'] });
    expect('backfill' in (out.cursor as object)).toBe(false);
    expect(out.archiveNullScoped).toBe(false);
  });

  it('archiveNullScoped is EXACTLY "this cursor re-establishes every root" — A-3 pairing as an invariant', async () => {
    const cases: Array<{
      label: string;
      config: Record<string, unknown>;
      cursor: OneDriveCursor;
      picked: FolderNode[];
    }> = [
      {
        label: 'pure add — FB keeps its token',
        config: { folderRoots: fr(['FB', 'Beta']) },
        cursor: { delta_tokens: { FB: 'TB' }, scope_roots: ['FB'] },
        picked: [node('FB', 'Beta'), node('FG', 'Gamma')],
      },
      {
        label: 'uncovered removal — FB keeps its token',
        config: { folderRoots: fr(['FB', 'Beta'], ['FG', 'Gamma']) },
        cursor: { delta_tokens: { FB: 'TB', FG: 'TG' }, scope_roots: ['FB', 'FG'] },
        picked: [node('FB', 'Beta')],
      },
      {
        label: 'ancestor carve-out — FB loses its token',
        config: { folderRoots: fr(['FB', 'Beta'], ['FBSUB', 'Sub']) },
        cursor: { delta_tokens: { FB: 'TB', FBSUB: 'TS' }, scope_roots: ['FB', 'FBSUB'] },
        picked: [node('FB', 'Beta')],
      },
      {
        label: 'mid-backfill account, uncovered removal',
        config: { folderRoots: fr(['FB', 'Beta'], ['FG', 'Gamma']) },
        cursor: { delta_tokens: { FB: 'TB' }, scope_roots: ['FB', 'FG'], backfill: { root_index: 1 } },
        picked: [node('FB', 'Beta')],
      },
    ];

    const seen: Array<{ label: string; flag: boolean; fullReestablish: boolean }> = [];
    for (const c of cases) {
      const { source } = makeSource();
      const { session } = makeSession({ config: c.config, cursor: c.cursor });
      const { channel } = makeFolderChannel({ picked: c.picked });

      const out = await source.manageFolders!(session, channel);
      const cur = out.cursor!;

      seen.push({
        label: c.label,
        flag: out.archiveNullScoped === true,
        fullReestablish:
          Object.keys(cur.delta_tokens).length === 0 && !('backfill' in (cur as object)),
      });
    }

    // Both directions in one assertion: the flag is NEVER true without the
    // full re-establish that makes archiving NULL rows recoverable (A-3), and
    // never false when that re-walk is already happening anyway.
    expect(seen).toEqual([
      { label: 'pure add — FB keeps its token', flag: false, fullReestablish: false },
      { label: 'uncovered removal — FB keeps its token', flag: false, fullReestablish: false },
      { label: 'ancestor carve-out — FB loses its token', flag: true, fullReestablish: true },
      { label: 'mid-backfill account, uncovered removal', flag: false, fullReestablish: false },
    ]);
  });

  it('the picker callbacks run on the SESSION client, not a frozen connect-time token (D9)', async () => {
    const { source } = makeSource();
    const { session } = makeSession({ creds: null, config: { folderRoots: fr(['FB', 'Beta']) } });
    const { channel } = makeFolderChannel({
      picked: async (spec) => {
        await expect(spec.children('FB')).rejects.toThrow(/no credentials available/);
        return [node('FB', 'Beta')];
      },
    });

    const out = await source.manageFolders!(session, channel);

    expect(out.config).toEqual({ folderRoots: fr(['FB', 'Beta']) });
  });

  it('rejects an empty selection — the last root goes away by removing the source (R3)', async () => {
    const { source } = makeSource();
    const { session } = makeSession({ config: { folderRoots: fr(['FB', 'Beta']) } });
    const { channel } = makeFolderChannel({ picked: [] });

    await expect(source.manageFolders!(session, channel)).rejects.toThrow(/no folders selected/);
  });

  it('leaves a stale legacy roots mirror UNTOUCHED — core owns that mirror (R1/A-2)', async () => {
    const { source } = makeSource();
    const { session } = makeSession({
      config: {
        folderRoots: fr(['FB', 'Beta']),
        roots: [{ rootFolderId: 'OLD', rootName: 'Stale mirror' }],
        cadenceOverride: '30m',
      },
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta')] });

    const out = await source.manageFolders!(session, channel);

    // This connector neither writes nor strips `roots` (A-2). The stale copy
    // rides through and core's applyFolderScope re-derives the mirror from
    // `folderRoots` in the SAME transaction, so it never reaches disk. An
    // earlier draft asserted the strip; that was the hole that left the
    // installed 2.0.5 build with no `roots` at all after the first Save.
    expect(out.config).toEqual({
      cadenceOverride: '30m',
      roots: [{ rootFolderId: 'OLD', rootName: 'Stale mirror' }],
      folderRoots: fr(['FB', 'Beta']),
    });
  });
});

describe('C-46/D4 — a name-extending SIBLING is not covered and its documents are archived', () => {
  it('removing /BetaBackup while /Beta is retained archives it end to end', async () => {
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta'], ['FBB', 'BetaBackup']) },
      cursor: { delta_tokens: { FB: 'TB', FBB: 'TBB' }, scope_roots: ['FB', 'FBB'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta')] });

    const out = await source.manageFolders!(session, channel);

    // `/BetaBackup/` is NOT under `/Beta/`, so it leaves scope: archived, not
    // exempted and not re-attributed.
    expect(out.archiveScopeRootIds).toEqual(['FBB']);
    expect(out.reattributeScopeRoots ?? []).toEqual([]);
    // …and FB is untouched: no overlap, so it keeps its delta token.
    expect(out.cursor).toEqual({ delta_tokens: { FB: 'TB' }, scope_roots: ['FB'] });

    const after = applyFolderScope([row('doc-b', 'FB'), row('doc-bb', 'FBB')], out);
    expect(after).toEqual([
      { externalId: 'doc-b', scopeRootId: 'FB', archived: false },
      { externalId: 'doc-bb', scopeRootId: 'FBB', archived: true },
    ]);
    expect(liveOutsideScope(after, out.config)).toEqual([]);
  });
});

describe('invariant 15 — a newly added root does not override file-indexability policy', () => {
  it('the first walk of a root added by manageFolders still drops a policy-ineligible file before any I/O', async () => {
    const { fetchFn, calls } = graphFetch({
      items: worldItems,
      deltaPages: {
        [deltaUrl('FG')]: {
          value: [
            driveFile('mp3-1', 'song.mp3', { file: { mimeType: 'audio/mpeg' } }),
            driveFile('p1', 'spec.pdf'),
          ],
          '@odata.deltaLink': `${GRAPH_BASE}/me/drive/items/FG/delta?token=TOK_G`,
        },
      },
      downloads: { 'https://download.example/p1': new Uint8Array([1, 2, 3]) },
    });
    const source = createOneDriveSource(makeHost(fetchFn), instantClock);

    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta']) },
      cursor: { delta_tokens: { FB: 'TB' }, scope_roots: ['FB'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta'), node('FG', 'Gamma')] });

    const out = await source.manageFolders!(session, channel);
    expect(out.config.folderRoots).toEqual(fr(['FB', 'Beta'], ['FG', 'Gamma']));

    // Now sync under the NEW scope exactly as the engine would: FB keeps its
    // token and is skipped, FG is established from scratch.
    const next = makeSession({ config: out.config }).session;
    const batches = (await collect(source.pull(next, out.cursor))) as Array<
      Batch<OneDriveCursor, OneDriveItem>
    >;
    const items = batches.flatMap((b) => b.items);

    expect(items.map((i) => i.file.id)).toEqual(['p1']);
    expect(calls.some((u) => u.includes('download.example/mp3-1'))).toBe(false);
    expect(calls.some((u) => /\/items\/mp3-1(\?|$)/.test(u))).toBe(false);
    // …and the file that DID survive is attributed to the newly added root,
    // which is what ties the policy gate to this specific root rather than to
    // the account as a whole.
    expect((source.toDocument(items[0]) as DocumentInput).scopeRootId).toBe('FG');
  });
});

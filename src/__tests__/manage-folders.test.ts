/**
 * manageFolders(session, channel): edit an existing account's tracked roots
 * with its OWN session credentials — canonical config out (with any legacy
 * `roots` mirror passed through untouched, A-2), a pruned and backfill-free
 * cursor out, the connector-computed archive set (DECISIONS R8), the
 * shared-ancestor carve-out (spec-reality-diff A5b) — which since C-46/D3 is
 * a RE-ATTRIBUTION rather than a silence — and A-3's archiveNullScoped
 * pairing. This connector has no reconcile(), so what is computed here is the
 * ONLY scope-removal path OneDrive will ever have, in both directions:
 * over-archiving is unrecoverable, so an ambiguous case must never archive;
 * and under-archiving leaks, so a removed root that is NOT archived must be
 * re-attributed to the retained root that covers it. Every removed root lands
 * in exactly one of the two lists.
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
  // C-46/D3's two-edit scenario. FP = `/Papers/`, FA = `/Papers/Alpha/` (it
  // was picked as a top-level root and has since MOVED under Papers
  // upstream), FZ = `/Zeta/`, unrelated to both.
  FP: driveFolder('FP', 'Papers'),
  FA: driveFolder('FA', 'Alpha', {
    parentReference: { driveId: 'drive-1', path: '/drive/root:/Papers' },
  }),
  FZ: driveFolder('FZ', 'Zeta'),
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

  it('pre-selects the current roots, returns canonical folderRoots, and costs no SAVE-path Graph lookups on a pure ADD', async () => {
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
    // A pure ADD still resolves NOTHING on the save path — the containment
    // walk is what this pins, and it never runs. The one call here is C-50's
    // reveal, which walks each PRESELECTED root's parents once before the
    // modal opens so the picker does not open collapsed.
    expect(calls).toEqual([
      `${GRAPH_BASE}/me/drive/items/FB?$select=id,name,parentReference`,
    ]);
  });

  /**
   * C-50 — the picker must OPEN revealed down to the tracked folders. Graph
   * item ids are opaque to the renderer, so the source walks
   * `parentReference` and the modal matches the ids by equality.
   */
  describe('spec.expand (C-50 reveal)', () => {
    /** Fixtures that carry `parentReference.id` — the parent ITEM id, which
     *  the pre-C-50 fixtures never needed and so never set. */
    const nested: GraphWorld['items'] = {
      DEEP: driveFolder('DEEP', 'Deep', {
        parentReference: { driveId: 'drive-1', id: 'MID', path: '/drive/root:/Top/Mid' },
      }),
      MID: driveFolder('MID', 'Mid', {
        parentReference: { driveId: 'drive-1', id: 'TOP', path: '/drive/root:/Top' },
      }),
      TOP: driveFolder('TOP', 'Top', {
        parentReference: { driveId: 'drive-1', id: 'root', path: '/drive/root:' },
      }),
      SHARED: driveFolder('SHARED', 'Team specs', {
        parentReference: { driveId: 'other-drive', id: 'X1', path: '/drives/other-drive/root:/Specs' },
      }),
    };

    it('maps a top-level root to the id the "My files" tab actually lists', async () => {
      // TOP's parent IS the drive root, so the chain ends at the literal id
      // the tab lists. Emitting anything else would match no row.
      const { source } = makeSource(nested);
      const { session } = makeSession({ config: { folderRoots: fr(['TOP', 'Top']) } });
      const { channel, getPickerSpec } = makeFolderChannel({ picked: [node('TOP', 'Top')] });

      await source.manageFolders!(session, channel);

      expect(getPickerSpec()!.expand).toEqual([ONEDRIVE_DRIVE_ROOT_ID]);
    });

    it('walks a nested root to its full chain and never lists the root itself', async () => {
      const { source } = makeSource(nested);
      const { session } = makeSession({ config: { folderRoots: fr(['DEEP', 'Deep']) } });
      const { channel, getPickerSpec } = makeFolderChannel({ picked: [node('DEEP', 'Deep')] });

      await source.manageFolders!(session, channel);

      // Nearest-first, ending at the tab's root id. DEEP is absent:
      // expanding a selected row buys nothing and costs a listing.
      expect(getPickerSpec()!.expand).toEqual(['MID', 'TOP', ONEDRIVE_DRIVE_ROOT_ID]);
    });

    it('stops at a shared drive, whose ancestors the picker never lists', async () => {
      // "Shared with me" lists those roots DIRECTLY; there is no row above
      // one, so emitting its foreign ancestors would match nothing.
      const { source } = makeSource(nested);
      const { session } = makeSession({ config: { folderRoots: fr(['SHARED', 'Team specs']) } });
      const { channel, getPickerSpec } = makeFolderChannel({
        picked: [node('SHARED', 'Team specs')],
      });

      await source.manageFolders!(session, channel);

      expect(getPickerSpec()!.expand).toEqual([]);
    });

    it('an unreadable root reveals less instead of keeping the user out of the editor', async () => {
      // Removing a vanished root is the main reason to open Manage folders,
      // so the reveal must never throw past the picker. GONE is absent from
      // the world; DEEP still reveals.
      const { source } = makeSource(nested);
      const { session } = makeSession({
        config: { folderRoots: fr(['GONE', 'Deleted'], ['DEEP', 'Deep']) },
      });
      const { channel, getPickerSpec } = makeFolderChannel({
        picked: [node('GONE', 'Deleted'), node('DEEP', 'Deep')],
      });

      await source.manageFolders!(session, channel);

      expect(getPickerSpec()!.expand).toEqual(['MID', 'TOP', ONEDRIVE_DRIVE_ROOT_ID]);
    });

    it("the drive-root catch-all expands nothing — it IS a root row", async () => {
      const { source } = makeSource(nested);
      const { session } = makeSession({
        config: { folderRoots: fr([ONEDRIVE_DRIVE_ROOT_ID, 'OneDrive']) },
      });
      const { channel, getPickerSpec } = makeFolderChannel({
        picked: [node(ONEDRIVE_DRIVE_ROOT_ID, 'OneDrive')],
      });

      await source.manageFolders!(session, channel);

      expect(getPickerSpec()!.expand).toEqual([]);
    });
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

  it('a removed root still covered by a retained ANCESTOR is RE-ATTRIBUTED, not archived, and drops the ancestor token (A5b + C-46/D3)', async () => {
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta'], ['FBSUB', 'Sub']) },
      cursor: { delta_tokens: { FB: 'TB', FBSUB: 'TS' }, scope_roots: ['FB', 'FBSUB'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FB', 'Beta')] });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual([]);
    // …and the removed root's live rows are re-stamped onto its coverer
    // rather than left carrying a stamp no later save can match (C-46/D3).
    // Silence here is NOT the right answer: the re-walk the dropped token
    // triggers cannot move a LIVE row's stamp (hashSkip, `source.ts:296`).
    expect(out.reattributeScopeRoots).toEqual([{ from: 'FBSUB', to: 'FB' }]);
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
    // A removed root is never in both lists — nothing covers FB here.
    expect(out.reattributeScopeRoots).toEqual([]);
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
    expect(out.reattributeScopeRoots).toEqual([{ from: 'FB', to: ONEDRIVE_DRIVE_ROOT_ID }]);
    expect(out.cursor).toEqual({ delta_tokens: {}, scope_roots: [ONEDRIVE_DRIVE_ROOT_ID] });
    expect(out.archiveNullScoped).toBe(true);
  });

  it('C-46 addendum: a coverer ADDED IN THE SAME SAVE re-attributes — widening must not archive', async () => {
    // The case the first D3 fix missed. Coverage was keyed on `retained`,
    // i.e. roots present in BOTH the prior config and the pick, so a coverer
    // the user ADDS in this very save could never cover anything and the
    // removed root fell through to `archiveScopeRootIds` — a full archive and
    // re-download of the subtree on the single most likely edit there is
    // ("stop tracking /Beta, track the whole drive"). Coverage is now decided
    // against the NEW selection.
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FB', 'Beta']) },
      cursor: { delta_tokens: { FB: 'TB' }, scope_roots: ['FB'] } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({
      picked: [node(ONEDRIVE_DRIVE_ROOT_ID, 'OneDrive')],
    });

    const out = await source.manageFolders!(session, channel);

    expect(out.archiveScopeRootIds).toEqual([]);
    expect(out.reattributeScopeRoots).toEqual([
      { from: 'FB', to: ONEDRIVE_DRIVE_ROOT_ID },
    ]);
    expect(out.cursor).toEqual({
      delta_tokens: {},
      scope_roots: [ONEDRIVE_DRIVE_ROOT_ID],
    });
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
    expect(out.reattributeScopeRoots).toEqual([]);
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

describe('C-46/D3 — stale attribution must not leak documents across successive edits', () => {
  it('two edits — remove a covered root, then remove its coverer — leave nothing live outside scope', async () => {
    // Edit 0 state: FA and FP are both tracked, and a pull has stamped one
    // document under each. FA has since moved under FP upstream.
    let rows = [row('doc-a', 'FA'), row('doc-p', 'FP')];

    // ── Edit 1: drop FA, keep FP. FP covers FA, so nothing leaves scope —
    // but the live rows stamped 'FA' must not KEEP that stamp: `hashSkip`
    // freezes attribution on an unchanged live row, so no later pull ever
    // refreshes it (C-46 addendum #6, measured). The save must re-attribute.
    const { source: s1 } = makeSource();
    const { session: sess1 } = makeSession({
      config: { folderRoots: fr(['FA', 'Alpha'], ['FP', 'Papers']) },
      cursor: { delta_tokens: { FA: 'TA', FP: 'TP' }, scope_roots: ['FA', 'FP'] } as OneDriveCursor,
    });
    const { channel: ch1 } = makeFolderChannel({ picked: [node('FP', 'Papers')] });

    const out1 = await s1.manageFolders!(sess1, ch1);

    expect(out1.archiveScopeRootIds).toEqual([]);
    expect(out1.reattributeScopeRoots).toEqual([{ from: 'FA', to: 'FP' }]);
    rows = applyFolderScope(rows, out1);
    expect(rows.map((r) => r.scopeRootId)).toEqual(['FP', 'FP']);
    expect(liveOutsideScope(rows, out1.config)).toEqual([]);

    // ── Edit 2: replace FP with the unrelated FZ. FP leaves scope with no
    // retained coverer, so it is archived — and every row the first edit
    // handled must be archived with it, because they are all in FP's subtree.
    const { source: s2 } = makeSource();
    const { session: sess2 } = makeSession({ config: out1.config, cursor: out1.cursor });
    const { channel: ch2 } = makeFolderChannel({ picked: [node('FZ', 'Zeta')] });

    const out2 = await s2.manageFolders!(sess2, ch2);

    expect(out2.archiveScopeRootIds).toEqual(['FP']);
    expect(out2.reattributeScopeRoots).toEqual([]);
    rows = applyFolderScope(rows, out2);

    // The whole point: no document survives LIVE outside the selection. Left
    // stamped 'FA', doc-a matches neither archive list and is searchable
    // forever — this connector has no reconcile() to notice it later.
    expect(liveOutsideScope(rows, out2.config)).toEqual([]);
    expect(rows).toEqual([
      { externalId: 'doc-a', scopeRootId: 'FP', archived: true },
      { externalId: 'doc-p', scopeRootId: 'FP', archived: true },
    ]);
  });

  it('re-attribution and archival are DISJOINT — core throws when a root is in both', async () => {
    // FA is covered by FP and FG is not, in ONE save: the covered root goes
    // to reattribute, the uncovered one to archive, and neither appears twice.
    const { source } = makeSource();
    const { session } = makeSession({
      config: { folderRoots: fr(['FA', 'Alpha'], ['FG', 'Gamma'], ['FP', 'Papers']) },
      cursor: {
        delta_tokens: { FA: 'TA', FG: 'TG', FP: 'TP' },
        scope_roots: ['FA', 'FG', 'FP'],
      } as OneDriveCursor,
    });
    const { channel } = makeFolderChannel({ picked: [node('FP', 'Papers')] });

    const out = await source.manageFolders!(session, channel);

    expect(out.reattributeScopeRoots).toEqual([{ from: 'FA', to: 'FP' }]);
    expect(out.archiveScopeRootIds).toEqual(['FG']);
    const froms = new Set((out.reattributeScopeRoots ?? []).map((m) => m.from));
    expect(out.archiveScopeRootIds.filter((id) => froms.has(id))).toEqual([]);

    // …and the model of core's transaction accepts it (it throws on overlap).
    const after = applyFolderScope([row('doc-a', 'FA'), row('doc-g', 'FG')], out);
    expect(after).toEqual([
      { externalId: 'doc-a', scopeRootId: 'FP', archived: false },
      { externalId: 'doc-g', scopeRootId: 'FG', archived: true },
    ]);
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

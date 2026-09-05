/**
 * C-46/D4 — `covers` / `overlaps` are the containment predicates the
 * folder-scope archive set is computed from, and they must respect the path
 * SEPARATOR boundary: a sibling whose NAME EXTENDS a retained root
 * (`/DocsBackup` beside `/Docs`) is NOT contained by it, and archiving must
 * not be silently skipped for it.
 *
 * These are unit tests of the pure predicates on purpose. Every path
 * `resolveRootLocation` actually builds already ends in `/`, so the defect is
 * NOT reachable through `manageFolders` today — see the end-to-end sibling
 * case in `manage-folders.test.ts`, which is a guard on that undocumented
 * convention, not a reproduction. The predicates must be correct on their own
 * terms regardless: they are exported and the convention is one edit away
 * from changing.
 *
 * Reference shape: core's `@shared/folder-paths.ts:isUnder`.
 */
import { covers, overlaps, type RootLocation } from '../source';

const D = 'drive-1';
const at = (path: string): RootLocation => ({ drive: D, path });

describe('covers — separator boundary (C-46/D4)', () => {
  it('a sibling whose name EXTENDS the retained root is NOT covered', () => {
    expect(covers(at('/Docs'), at('/DocsBackup'))).toBe(false);
    expect(covers(at('/Docs/'), at('/DocsBackup/'))).toBe(false);
    expect(covers(at('/Beta'), at('/BetaBackup/2026'))).toBe(false);
  });

  it('the root itself and anything genuinely under it ARE covered', () => {
    expect(covers(at('/Docs'), at('/Docs'))).toBe(true);
    expect(covers(at('/Docs'), at('/Docs/Sub'))).toBe(true);
    expect(covers(at('/Docs/'), at('/Docs/Sub/'))).toBe(true);
    expect(covers(at('/Docs'), at('/Docs/'))).toBe(true);
  });

  it('the drive root covers everything in its own drive, and nothing in another', () => {
    expect(covers(at('/'), at('/'))).toBe(true);
    expect(covers(at('/'), at('/Docs/'))).toBe(true);
    expect(covers({ drive: 'other', path: '/' }, at('/Docs/'))).toBe(false);
  });

  it('a null location on either side is incomparable, never covered', () => {
    expect(covers(null, at('/Docs/'))).toBe(false);
    expect(covers(at('/Docs/'), null)).toBe(false);
    expect(covers(null, null)).toBe(false);
  });
});

describe('overlaps — separator boundary in BOTH directions (C-46/D4)', () => {
  it('name-extending siblings do not overlap in either direction', () => {
    expect(overlaps(at('/Docs'), at('/DocsBackup'))).toBe(false);
    expect(overlaps(at('/DocsBackup'), at('/Docs'))).toBe(false);
  });

  it('a genuine ancestor/descendant pair overlaps whichever way round it is given', () => {
    expect(overlaps(at('/Docs'), at('/Docs/Sub'))).toBe(true);
    expect(overlaps(at('/Docs/Sub'), at('/Docs'))).toBe(true);
    expect(overlaps(at('/Docs'), at('/Docs'))).toBe(true);
    expect(overlaps(at('/'), at('/Docs/'))).toBe(true);
  });

  it('different drives never overlap, and a null side is incomparable', () => {
    expect(overlaps(at('/Docs/'), { drive: 'other', path: '/Docs/' })).toBe(false);
    expect(overlaps(null, at('/Docs/'))).toBe(false);
  });
});

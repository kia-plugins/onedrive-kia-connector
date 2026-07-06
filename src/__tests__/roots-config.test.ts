/**
 * rootsConfig normalization: the v2 multi-root shape, the OneDrive default
 * (whole drive), per-entry name fallbacks, and dedupe by rootFolderId. No
 * legacy single-root config to migrate here (OneDrive is a NEW v2 connector,
 * unlike google-docs' v2.0.0 → v2.1.0 upgrade path).
 */
import { rootsConfig } from '../source';
import { makeSession } from '../testing/harness';

const roots = (config: Record<string, unknown>) => rootsConfig(makeSession({ config }).session);

describe('rootsConfig', () => {
  it('returns the roots array as-is', () => {
    expect(
      roots({
        roots: [
          { rootFolderId: 'FA', rootName: 'Alpha' },
          { rootFolderId: 'SH1', rootName: 'Shared specs' },
        ],
      }),
    ).toEqual([
      { rootFolderId: 'FA', rootName: 'Alpha' },
      { rootFolderId: 'SH1', rootName: 'Shared specs' },
    ]);
  });

  it('falls back per entry to the raw id when rootName is missing/empty', () => {
    expect(
      roots({ roots: [{ rootFolderId: 'root' }, { rootFolderId: 'FA', rootName: '' }] }),
    ).toEqual([
      { rootFolderId: 'root', rootName: 'root' },
      { rootFolderId: 'FA', rootName: 'FA' },
    ]);
  });

  it('skips invalid entries and keeps the valid ones', () => {
    expect(
      roots({
        roots: [{ rootFolderId: '' }, 42, null, { rootFolderId: 'FA', rootName: 'Alpha' }],
      }),
    ).toEqual([{ rootFolderId: 'FA', rootName: 'Alpha' }]);
  });

  it('an empty or all-invalid roots array falls back to the whole OneDrive default', () => {
    expect(roots({ roots: [] })).toEqual([{ rootFolderId: 'root', rootName: 'OneDrive' }]);
    expect(roots({ roots: [{ rootFolderId: 7 }] })).toEqual([
      { rootFolderId: 'root', rootName: 'OneDrive' },
    ]);
  });

  it('defaults to the whole OneDrive with no config at all', () => {
    expect(roots({})).toEqual([{ rootFolderId: 'root', rootName: 'OneDrive' }]);
  });

  it('dedupes by rootFolderId keeping the FIRST entry', () => {
    expect(
      roots({
        roots: [
          { rootFolderId: 'FA', rootName: 'First' },
          { rootFolderId: 'FB', rootName: 'Other' },
          { rootFolderId: 'FA', rootName: 'Second' },
        ],
      }),
    ).toEqual([
      { rootFolderId: 'FA', rootName: 'First' },
      { rootFolderId: 'FB', rootName: 'Other' },
    ]);
  });
});

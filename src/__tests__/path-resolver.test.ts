/**
 * Direct port of alpha-cent's onedrive-path-resolver.test.ts — same
 * scenarios, `rootName` replacing v1's `display_path` field name (v2's
 * RootConfig shape).
 */
import { buildDisplayPath } from '../path-resolver';

describe('buildDisplayPath', () => {
  it('anchors under the root rootName with " / " separators', () => {
    const root = { rootName: 'My files' };
    const item = {
      name: 'spec.docx',
      parentReference: { path: '/drive/root:/Projects/2026' },
    };
    expect(buildDisplayPath(root, item)).toBe('My files / Projects / 2026 / spec.docx');
  });

  it('handles items at the drive root (empty path)', () => {
    const root = { rootName: 'My files' };
    const item = { name: 'readme.md', parentReference: { path: '/drive/root:' } };
    expect(buildDisplayPath(root, item)).toBe('My files / readme.md');
  });

  it('decodes percent-encoded path segments', () => {
    const root = { rootName: 'My files' };
    const item = {
      name: 'design.pdf',
      parentReference: { path: '/drive/root:/My%20Project/2026' },
    };
    expect(buildDisplayPath(root, item)).toBe('My files / My Project / 2026 / design.pdf');
  });

  it('preserves unicode segments', () => {
    const root = { rootName: 'My files' };
    const item = {
      name: 'notes.txt',
      parentReference: { path: '/drive/root:/Über/Notizen' },
    };
    expect(buildDisplayPath(root, item)).toBe('My files / Über / Notizen / notes.txt');
  });

  it('falls back to "<root> / <name>" when parentReference.path is undefined (shared root)', () => {
    const root = { rootName: 'Shared' };
    const item = { name: 'team-doc.docx', parentReference: { path: undefined } };
    expect(buildDisplayPath(root, item)).toBe('Shared / team-doc.docx');
  });

  it('falls back when parentReference itself is absent', () => {
    const root = { rootName: 'Shared' };
    const item = { name: 'no-parent.txt' };
    expect(buildDisplayPath(root, item)).toBe('Shared / no-parent.txt');
  });
});

/**
 * Binary-mime routing for OneDrive items. v1 (`ingest.ts`) delegated this to
 * `@main/converter`'s `isConvertibleMime` and ran the extraction ITSELF
 * (download bytes, call the local converter, store markdown). v2 drops that:
 * every downloadable file is emitted with `binary: { bytes, mime, filename }`
 * and `markdown: null`, and the v2 ENGINE's convert/vision pipeline does the
 * extraction — this module only decides which mimes are worth downloading at
 * all (`isConvertibleMime` below is the SAME criterion the google-docs-kia-
 * connector template uses for its own binary route, since it is bound to the
 * same v2 engine: kiagent-core src/main/core/engine/convert.ts:43-84's
 * deterministic converters (pdf, docx, xlsx, any text/*) plus the two-pass
 * vision pipeline's image/* candidates, kiagent-core
 * src/main/workers/vision/classify.ts:48-63).
 *
 * v1's `isConvertibleMime` gate (used only to skip a download, never to route
 * differently) is folded into `chooseRoute` below; images were UNSUPPORTED in
 * v1 (no local OCR) and are BINARY here (the v2 vision pipeline OCRs them via
 * fetchBytes) — a deliberate v2 broadening, not a regression.
 */

const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const XLSX_MIME =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** Binary mimes the v2 engine can turn into text — see module doc. */
export function isConvertibleMime(mimeType: string): boolean {
  return (
    mimeType.startsWith('text/') ||
    mimeType.startsWith('image/') ||
    mimeType === 'application/pdf' ||
    mimeType === DOCX_MIME ||
    mimeType === XLSX_MIME
  );
}

export type Route = { kind: 'binary' } | { kind: 'unsupported' };

export function chooseRoute(mimeType: string): Route {
  return isConvertibleMime(mimeType) ? { kind: 'binary' } : { kind: 'unsupported' };
}

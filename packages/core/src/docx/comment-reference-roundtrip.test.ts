import { describe, test, expect } from 'bun:test';
import JSZip from 'jszip';
import { parseDocx } from './parser';
import { repackDocx } from './rezip';

/**
 * Comment reference round-trip, incl. range-less "point" comments.
 *
 * A document can carry two kinds of comment anchor in document.xml:
 *  - a RANGED comment: `commentRangeStart` + `commentRangeEnd` + `commentReference`
 *  - a POINT comment: a bare `commentReference` with NO range markers
 *
 * The model used to represent only the range markers, and emitted the reference
 * as a side-effect of `commentRangeEnd`. So a point comment (no range) had no
 * representation and was dropped — a 2-comment doc round-tripped to 1 reference.
 * The fix models `commentReference` as a first-class node (parsed + serialized),
 * making it the single source of truth for every reference. This guards both
 * comments surviving.
 *
 * Fixture provenance: synthesized here (mirrors mammoth.js test/test-data/comments.docx).
 */

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships';

const CONTENT_TYPES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Types xmlns="${NS_CT}">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
  `</Types>`;

const PACKAGE_RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="${NS_PR}">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`;

const DOCUMENT_RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="${NS_PR}">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>` +
  `</Relationships>`;

// Comment 0 = ranged; comment 2 = point (bare reference, no range markers).
const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>` +
  `<w:p>` +
  `<w:commentRangeStart w:id="0"/>` +
  `<w:r><w:t xml:space="preserve">Ranged text</w:t></w:r>` +
  `<w:commentRangeEnd w:id="0"/>` +
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>` +
  `<w:r><w:t xml:space="preserve"> and a point anchor</w:t></w:r>` +
  `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="2"/></w:r>` +
  `</w:p>` +
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>` +
  `</w:body></w:document>`;

const COMMENTS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:comments xmlns:w="${NS_W}">` +
  `<w:comment w:id="0" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Ranged comment</w:t></w:r></w:p></w:comment>` +
  `<w:comment w:id="2" w:author="Reviewer B" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Point comment</w:t></w:r></w:p></w:comment>` +
  `</w:comments>`;

async function buildCommentsDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES_XML);
  zip.file('_rels/.rels', PACKAGE_RELS_XML);
  zip.file('word/document.xml', DOCUMENT_XML);
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS_XML);
  zip.file('word/comments.xml', COMMENTS_XML);
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }));
}

async function partOf(buf: Uint8Array, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  return zip.file(path)!.async('text');
}

describe('comment reference round-trip', () => {
  test('both a ranged and a point comment keep their references through repack', async () => {
    const doc = await parseDocx(await buildCommentsDocx(), { preloadFonts: false });
    const out = new Uint8Array(await repackDocx(doc));
    const body = await partOf(out, 'word/document.xml');

    // Both references survive (this is what regressed: 2 -> 1).
    expect((body.match(/<w:commentReference/g) ?? []).length).toBe(2);
    expect(body).toContain('w:commentReference w:id="0"');
    expect(body).toContain('w:commentReference w:id="2"');

    // The point comment adds no range, so exactly one range pair remains —
    // and the reference is NOT double-emitted off commentRangeEnd.
    expect((body.match(/<w:commentRangeStart/g) ?? []).length).toBe(1);
    expect((body.match(/<w:commentRangeEnd/g) ?? []).length).toBe(1);

    // comments.xml still defines both comments.
    const comments = await partOf(out, 'word/comments.xml');
    expect((comments.match(/<w:comment\b/g) ?? []).length).toBe(2);
  });
});

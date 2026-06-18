/**
 * Range-less "point" comments must survive the EDITOR save path
 * (toProseDoc → fromProseDoc → serialize), not only the headless
 * parseDocx → repackDocx path (covered by docx/comment-reference-roundtrip).
 *
 * A point comment is a bare `w:commentReference` with NO commentRangeStart/End.
 * It has no text to carry a `comment` mark, so toProseDoc used to drop it,
 * orphaning the comment in comments.xml on the editor path. The `commentRef`
 * PM node now carries the lone reference through; the RANGED comment's reference
 * still comes from its mark (insertCommentRanges), and the two sources are
 * disjoint by id, so neither double-emits. See eigenpal/docx-editor#837.
 */

import { describe, test, expect } from 'bun:test';
import JSZip from 'jszip';
import type { Node as PMNode } from 'prosemirror-model';
import { parseDocx } from '../../../docx/parser';
import { createDocx } from '../../../docx/rezip';
import { toProseDoc } from '../toProseDoc';
import { updateDocumentContent } from '../fromProseDoc';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships';

// One paragraph: a RANGED comment (id 0) over "Ranged text" + a POINT comment
// (id 2: bare reference, no range markers) after " and a point anchor".
const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${NS_W}"><w:body>` +
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
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="${NS_CT}">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
      `</Types>`
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${NS_PR}">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`
  );
  zip.file('word/document.xml', DOCUMENT_XML);
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${NS_PR}">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>` +
      `</Relationships>`
  );
  zip.file('word/comments.xml', COMMENTS_XML);
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }));
}

async function documentXmlOf(buf: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  return zip.file('word/document.xml')!.async('text');
}

function countCommentRefNodes(doc: PMNode): number {
  let n = 0;
  doc.descendants((node) => {
    if (node.type.name === 'commentRef') n += 1;
  });
  return n;
}

describe('point comment round-trips through the PM editor path', () => {
  test('toProseDoc carries a point comment as a commentRef node, ranged comment as a mark', async () => {
    const model = await parseDocx(await buildCommentsDocx(), { preloadFonts: false });
    const pm = toProseDoc(model);

    // The POINT comment (id 2) becomes exactly one commentRef node.
    expect(countCommentRefNodes(pm)).toBe(1);
    let pointId = -1;
    pm.descendants((node) => {
      if (node.type.name === 'commentRef') pointId = node.attrs.commentId as number;
    });
    expect(pointId).toBe(2);

    // The RANGED comment (id 0) becomes a `comment` MARK on the text (NOT a node).
    let rangedMarked = false;
    pm.descendants((node) => {
      if (
        node.isText &&
        node.marks.some((m) => m.type.name === 'comment' && m.attrs.commentId === 0)
      ) {
        rangedMarked = true;
      }
    });
    expect(rangedMarked).toBe(true);
  });

  test('full editor round-trip keeps BOTH references and does not double-emit', async () => {
    const model = await parseDocx(await buildCommentsDocx(), { preloadFonts: false });
    const pm = toProseDoc(model);
    const updated = updateDocumentContent(model, pm); // = fromProseDoc(pm, model)
    const out = new Uint8Array(await createDocx(updated));
    const body = await documentXmlOf(out);

    // Both references survive the editor path (this is what regressed: 2 -> 1).
    expect((body.match(/<w:commentReference/g) ?? []).length).toBe(2);
    expect(body).toContain('w:commentReference w:id="0"');
    expect(body).toContain('w:commentReference w:id="2"');

    // The point comment adds no range; the ranged comment's single pair is
    // regenerated from its mark — not double-counted off the commentRef node.
    expect((body.match(/<w:commentRangeStart/g) ?? []).length).toBe(1);
    expect((body.match(/<w:commentRangeEnd/g) ?? []).length).toBe(1);
  });
});

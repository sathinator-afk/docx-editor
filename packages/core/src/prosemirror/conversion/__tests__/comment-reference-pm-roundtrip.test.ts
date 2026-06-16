/**
 * Regression: a comment created/edited in the EDITOR must serialize a
 * `<w:commentReference>` run on the PM save path (toProseDoc → fromProseDoc →
 * serialize), not only on the `parseDocx → repackDocx` path.
 *
 * Editor comments are ProseMirror `comment` MARKS, so `insertCommentRanges`
 * (fromProseDoc) synthesizes `commentRangeStart`/`End` from those marks. After
 * the reference run moved out of the `commentRangeEnd` serializer case into its
 * own `commentReference` node, the save path had to emit that node explicitly —
 * otherwise every editor-saved comment silently lost its anchor, and an
 * imported point comment vanished on re-save.
 *
 * The existing `comment-reference-roundtrip.test.ts` only exercises
 * `parseDocx → repackDocx`, which never builds the model from PM marks, so it
 * stayed green through the regression. This test closes that gap.
 *
 * See eigenpal/docx-editor#837 (review).
 */

import { describe, test, expect } from 'bun:test';
import { schema } from '../../schema';
import { fromProseDoc } from '../fromProseDoc';
import { toProseDoc } from '../toProseDoc';
import { serializeParagraph } from '../../../docx/serializer/paragraphSerializer';
import type { Paragraph } from '../../../types/document';

/** A paragraph whose middle run carries a `comment` mark (id N). */
function midCommentDoc(commentId: number) {
  const mark = schema.marks.comment.create({ commentId });
  const paragraph = schema.nodes.paragraph.create({}, [
    schema.text('Hello ', []),
    schema.text('commented', [mark]),
    schema.text(' world', []),
  ]);
  return schema.nodes.doc.create({}, [paragraph]);
}

function paragraphOf(doc: ReturnType<typeof midCommentDoc>): Paragraph {
  const pkg = fromProseDoc(doc);
  return pkg.package?.document?.content?.[0] as Paragraph;
}

function refCount(xml: string, id: number): number {
  return (xml.match(new RegExp(`<w:commentReference w:id="${id}"/>`, 'g')) ?? []).length;
}

describe('comment reference round-trips through the PM save path', () => {
  test('fromProseDoc emits a commentReference node right after commentRangeEnd', () => {
    const para = paragraphOf(midCommentDoc(7));
    const types = para.content.map((c) => (c as { type?: string }).type);
    expect(types).toContain('commentRangeStart');
    expect(types).toContain('commentRangeEnd');
    expect(types).toContain('commentReference');
    expect(types.indexOf('commentReference')).toBe(types.indexOf('commentRangeEnd') + 1);
    const ref = para.content.find((c) => (c as { type?: string }).type === 'commentReference') as
      | { id: number }
      | undefined;
    expect(ref?.id).toBe(7);
  });

  test('serialized paragraph carries exactly one <w:commentReference> run', () => {
    const xml = serializeParagraph(paragraphOf(midCommentDoc(7)));
    expect(xml).toContain('<w:commentRangeStart w:id="7"/>');
    expect(xml).toContain('<w:commentRangeEnd w:id="7"/>');
    expect(refCount(xml, 7)).toBe(1);
  });

  test('a comment extending to paragraph end still emits its reference', () => {
    // Exercises the end-of-paragraph close path in insertCommentRanges.
    const mark = schema.marks.comment.create({ commentId: 9 });
    const paragraph = schema.nodes.paragraph.create({}, [
      schema.text('start ', []),
      schema.text('to the very end', [mark]),
    ]);
    const doc = schema.nodes.doc.create({}, [paragraph]);
    const xml = serializeParagraph(fromProseDoc(doc).package?.document?.content?.[0] as Paragraph);
    expect(refCount(xml, 9)).toBe(1);
  });

  test('save → load → save keeps exactly one reference (no accumulation)', () => {
    // toProseDoc must drop the reference node (the mark is the source of truth)
    // so re-saving regenerates exactly one — never zero, never two.
    const doc = midCommentDoc(7);
    const doc2 = toProseDoc(fromProseDoc(doc));

    let markedRuns = 0;
    doc2.descendants((node) => {
      if (node.isText && node.marks.some((m) => m.type.name === 'comment')) markedRuns += 1;
    });
    expect(markedRuns, 'reloaded text keeps its comment mark').toBeGreaterThan(0);

    const xml2 = serializeParagraph(
      fromProseDoc(doc2).package?.document?.content?.[0] as Paragraph
    );
    expect(refCount(xml2, 7)).toBe(1);
  });
});

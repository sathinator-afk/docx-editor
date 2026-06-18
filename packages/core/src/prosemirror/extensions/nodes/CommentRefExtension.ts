/**
 * Comment Reference Extension — carries a range-less "point" comment anchor.
 *
 * Most comments are a `comment` MARK over a range of text; their inline
 * `w:commentReference` glyph is regenerated from those marks on save
 * (insertCommentRanges in fromProseDoc/paragraph.ts). But Word also allows a
 * bare `w:commentReference` with NO `commentRangeStart`/`End` — a "point"
 * comment anchored at a single position. It has no text to mark, so the editor
 * conversion (toProseDoc) had no PM representation for it and dropped it,
 * orphaning the comment in comments.xml on the editor save path.
 *
 * This node carries that lone reference through the PM document invisibly
 * (comment glyphs aren't part of the visible text), so `fromProseDoc` can
 * re-emit the exact `commentReference` run. Mirrors {@link FieldMarkerExtension}
 * — an inline, atomic, non-selectable, display:none marker. No guard is needed
 * (unlike a TOC field there is no begin/end pairing to keep balanced, and a lone
 * dropped anchor is recoverable, not corrupting).
 *
 * See eigenpal/docx-editor#837 (headless point-comment round-trip); this
 * completes the editor-path side.
 */

import { createNodeExtension } from '../create';

export const CommentRefExtension = createNodeExtension({
  name: 'commentRef',
  schemaNodeName: 'commentRef',
  nodeSpec: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: false,
    attrs: {
      /** The `w:id` of the point comment this reference anchors. */
      commentId: { default: 0 },
    },
    parseDOM: [
      {
        tag: 'span.docx-comment-ref',
        getAttrs(dom) {
          const el = dom as HTMLElement;
          return { commentId: parseInt(el.dataset.commentId || '0', 10) || 0 };
        },
      },
    ],
    toDOM(node) {
      const { commentId } = node.attrs as { commentId: number };
      // Comment-reference glyphs are not part of the visible text — carry it
      // invisibly so it survives the edit cycle.
      return [
        'span',
        {
          class: 'docx-comment-ref',
          'data-comment-id': String(commentId),
          style: 'display:none',
        },
      ];
    },
  },
});

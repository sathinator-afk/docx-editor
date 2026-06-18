/**
 * Regression guard: accepting a tracked (suggesting-mode) deletion that overlaps
 * a TOC field must NOT get silently stuck against the fieldMarker firewall, and
 * must NOT orphan the field.
 *
 * A review hypothesised: markRangeAsDeleted could stamp a `deletion` mark onto
 * part of a field (begin/instr/separate but not end); accept would then emit an
 * UNbalanced ReplaceStep that the firewall vetoes; the accept reports handled
 * but the tr is dropped, so the revision is permanently un-acceptable.
 *
 * That premise is FALSE in this schema, and these tests pin down why so a future
 * change can't silently reintroduce it:
 *
 *   1. A fieldMarker is an inline ATOM with NO `marks` opt-in, so its markSet is
 *      empty and `allowsMarkType(deletion) === false`. (Image/Shape DO opt in via
 *      `marks: '_'`, which is why suggesting-mode delete strikes THEM through.)
 *   2. markRangeAsDeleted (handlers/delete.ts) skips any inline node that doesn't
 *      `allowsMarkType(deletion)`, so a fieldMarker never carries a deletion mark.
 *   3. accept (resolveChange in commands/comments.ts) deletes ONLY nodes that
 *      carry the deletion mark, per node — so it never removes a fieldMarker and
 *      never emits a marker-removing ReplaceStep for the firewall to veto.
 *
 * Net: a suggesting-mode delete over a field marks only its TEXT; accept removes
 * that text and leaves all markers intact — no orphan, no veto, no stuck accept.
 */

import { describe, test, expect } from 'bun:test';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { singletonManager, schema } from '../../../schema';
import { createSuggestionModePlugin } from '../../../plugins/suggestionMode';
import { handleSuggestionDelete } from '../../../plugins/suggestionMode/handlers/delete';
import { acceptAllChanges } from '../../../commands';

const TOC = 'TOC \\o "1-3" \\h \\z \\u';

function fm(charType: string): PMNode {
  return schema.node('fieldMarker', { markerType: 'fieldChar', charType });
}
function instr(): PMNode {
  return schema.node('fieldMarker', { markerType: 'instrText', instr: TOC });
}

/** Multi-paragraph TOC: para0 = [begin, instr, separate, 'Entry'], para1 = ['more', end]. */
function multiParaTocDoc(): PMNode {
  const head = schema.nodes.paragraph.create({}, [
    fm('begin'),
    instr(),
    fm('separate'),
    schema.text('Entry'),
  ]);
  const tail = schema.nodes.paragraph.create({}, [schema.text('more'), fm('end')]);
  return schema.nodes.doc.create({}, [head, tail]);
}

/** Production pipeline (incl. the fieldMarker firewall) + an ACTIVE suggesting mode. */
function suggestingState(doc: PMNode): EditorState {
  return EditorState.create({
    doc,
    schema,
    plugins: [...singletonManager.getPlugins(), createSuggestionModePlugin(true, 'Tester')],
  });
}

function countMarkers(doc: PMNode): number {
  let n = 0;
  doc.descendants((node) => {
    if (node.type.name === 'fieldMarker') n += 1;
  });
  return n;
}

describe('fieldMarker is inert to tracked-change marks', () => {
  test('fieldMarker does NOT accept insertion/deletion marks; image DOES (the opt-in contrast)', () => {
    expect(schema.nodes.fieldMarker.allowsMarkType(schema.marks.deletion)).toBe(false);
    expect(schema.nodes.fieldMarker.allowsMarkType(schema.marks.insertion)).toBe(false);
    // Image opts in via `marks: '_'`, which is why deleting an image in suggesting
    // mode strikes it through — fieldMarker deliberately does not.
    expect(schema.nodes.image.allowsMarkType(schema.marks.deletion)).toBe(true);
  });
});

describe('suggesting-mode delete over a TOC field, then accept', () => {
  test('marks only the result text — never the markers — and accept does not orphan the field', () => {
    let state = suggestingState(multiParaTocDoc());
    expect(countMarkers(state.doc)).toBe(4);

    // Select para0's content: [begin, instr, separate, 'Entry'].
    const head = state.doc.firstChild as PMNode;
    state = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, 1, head.nodeSize - 1))
    );

    // Suggesting-mode delete the selection (marks text as deletion, skips atoms
    // that don't allow the mark — i.e. the fieldMarkers).
    handleSuggestionDelete(
      state,
      (tr) => {
        state = state.apply(tr);
      },
      'backward'
    );

    // The result text IS marked deleted (the op ran)...
    let entryMarkedDeleted = false;
    let anyMarkerMarked = false;
    state.doc.descendants((n) => {
      if (
        n.isText &&
        n.text?.includes('Entry') &&
        n.marks.some((m) => m.type.name === 'deletion')
      ) {
        entryMarkedDeleted = true;
      }
      if (n.type.name === 'fieldMarker' && n.marks.some((m) => m.type.name === 'deletion')) {
        anyMarkerMarked = true;
      }
    });
    expect(entryMarkedDeleted).toBe(true);
    // ...but NO fieldMarker carries a deletion mark.
    expect(anyMarkerMarked).toBe(false);
    expect(countMarkers(state.doc)).toBe(4);

    // Accept all changes: must actually APPLY (not be vetoed/stuck) and keep the field.
    const before = state.doc;
    acceptAllChanges()(state, (tr) => {
      state = state.apply(tr);
    });

    expect(state.doc.eq(before)).toBe(false); // accept applied — NOT silently dropped
    expect(countMarkers(state.doc)).toBe(4); // all markers survive
    let hasEntry = false;
    state.doc.descendants((n) => {
      if (n.isText && n.text?.includes('Entry')) hasEntry = true;
    });
    expect(hasEntry).toBe(false); // the marked text was removed
  });
});

/**
 * Hardening regression for the TOC fieldMarker node (Cuneon beta readiness).
 *
 * The fieldMarker atoms are INVISIBLE (display:none), so a careless live edit can
 * silently destroy a document's Table of Contents. FieldMarkerExtension guards it
 * with a single source-agnostic filterTransaction "firewall": it vetoes any
 * transaction whose ReplaceStep/ReplaceAroundStep removes field markers without
 * removing one-or-more WHOLE, properly-nested fields (an order-aware begin/end
 * depth walk). Paragraph JOINS need no guard — they relocate markers in order, so
 * the firewall correctly ALLOWS them.
 *
 * A vetoed root transaction is dropped and EditorState.apply returns the ORIGINAL
 * state, so we assert next.doc.eq(state.doc) / an unchanged marker count to prove
 * a veto. All cases run through the FULL production pipeline
 * (singletonManager.getPlugins()).
 *
 * Harness mirrors conversion/__tests__/toc-fieldmarker-pm-roundtrip.test.ts.
 */

import { describe, test, expect } from 'bun:test';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { wrapIn, joinBackward } from 'prosemirror-commands';
import { singletonManager, schema } from '../../../schema';
import {
  fieldMarkerGuardKey,
  FIELD_MARKER_BYPASS,
  deleteFieldAtCursor,
} from '../FieldMarkerExtension';

const TOC = 'TOC \\o "1-3" \\h \\z \\u';

// --- doc builders -----------------------------------------------------------

function fm(charType: string): PMNode {
  return schema.node('fieldMarker', { markerType: 'fieldChar', charType });
}
function instr(text = TOC): PMNode {
  return schema.node('fieldMarker', { markerType: 'instrText', instr: text });
}

/** One-paragraph TOC: [begin, instr, separate, text, end]. */
function singleParaTocDoc(text = 'Heading 1\t1'): PMNode {
  const para = schema.nodes.paragraph.create({}, [
    fm('begin'),
    instr(),
    fm('separate'),
    schema.text(text),
    fm('end'),
  ]);
  return schema.nodes.doc.create({}, [para]);
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

/** Two complete, adjacent fields back-to-back in one paragraph. */
function twoAdjacentFieldsDoc(): PMNode {
  const para = schema.nodes.paragraph.create({}, [
    fm('begin'),
    instr('TOC'),
    fm('separate'),
    schema.text('one'),
    fm('end'), // field A
    fm('begin'),
    instr('PAGEREF _Toc1'),
    fm('separate'),
    schema.text('two'),
    fm('end'), // field B
  ]);
  return schema.nodes.doc.create({}, [para]);
}

function stateOf(doc: PMNode): EditorState {
  // Full production pipeline: includes the firewall + baseKeymap + history.
  return EditorState.create({ doc, schema, plugins: [...singletonManager.getPlugins()] });
}

function countMarkers(doc: PMNode): number {
  let n = 0;
  doc.descendants((node) => {
    if (node.type.name === 'fieldMarker') n += 1;
  });
  return n;
}

/** Positions of all fieldChar markers of a given charType (begin/separate/end).
 *  Filters markerType==='fieldChar' so instrText markers — whose charType
 *  DEFAULTS to 'begin' — are never mistaken for a field begin. */
function markerPositions(doc: PMNode, charType: string): number[] {
  const out: number[] = [];
  doc.descendants((node, p) => {
    if (
      node.type.name === 'fieldMarker' &&
      node.attrs.markerType === 'fieldChar' &&
      node.attrs.charType === charType
    ) {
      out.push(p);
    }
  });
  return out;
}

const posOfMarker = (doc: PMNode, charType: string): number => {
  const ps = markerPositions(doc, charType);
  if (!ps.length) throw new Error(`marker ${charType} not found`);
  return ps[0];
};

/** A text position inside the (first) occurrence of `text`. */
function posInsideText(doc: PMNode, text: string): number {
  let pos = -1;
  doc.descendants((node, p) => {
    if (pos < 0 && node.isText && node.text?.includes(text)) pos = p + 1;
  });
  if (pos < 0) throw new Error(`text ${text} not found`);
  return pos;
}

// ---------------------------------------------------------------------------

describe('fieldMarker firewall — orphaning deletes are VETOED', () => {
  test('range-delete spanning begin/instr/separate (not the whole field) is VETOED', () => {
    const state = stateOf(singleParaTocDoc());
    expect(countMarkers(state.doc)).toBe(4);
    const from = posOfMarker(state.doc, 'begin');
    const to = posOfMarker(state.doc, 'end'); // exclusive of `end` → unbalanced
    const next = state.apply(state.tr.delete(from, to));
    expect(next.doc.eq(state.doc)).toBe(true); // veto → state unchanged
    expect(countMarkers(next.doc)).toBe(4);
  });

  test('typing a character over a marker-spanning selection is VETOED', () => {
    const state = stateOf(singleParaTocDoc());
    const from = posOfMarker(state.doc, 'begin');
    const to = posOfMarker(state.doc, 'end');
    const withSel = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
    const next = withSel.apply(withSel.tr.replaceSelectionWith(schema.text('Z'), false));
    expect(countMarkers(next.doc)).toBe(4);
  });

  test('removing JUST the separate marker (a stray middle) is VETOED', () => {
    const state = stateOf(singleParaTocDoc());
    const from = posOfMarker(state.doc, 'separate');
    const to = posOfMarker(state.doc, 'end'); // [separate, end): a lone separate at depth 0
    const next = state.apply(state.tr.replaceWith(from, to, schema.text('X')));
    expect(countMarkers(next.doc)).toBe(4);
  });

  test('partial select (one para) over an open field head is VETOED', () => {
    const state = stateOf(multiParaTocDoc());
    const head = state.doc.firstChild as PMNode;
    const next = state.apply(state.tr.delete(1, head.nodeSize - 1)); // begin/instr/separate, no end
    expect(countMarkers(next.doc)).toBe(4);
  });

  test('a bare delete of just `begin` is VETOED', () => {
    const state = stateOf(singleParaTocDoc());
    const beginPos = posOfMarker(state.doc, 'begin');
    const next = state.apply(state.tr.delete(beginPos, beginPos + 1));
    expect(countMarkers(next.doc)).toBe(4);
  });

  test('strand-delete BETWEEN two adjacent fields is VETOED (no cross-field corruption)', () => {
    // The count-only check would WRONGLY allow this: removing [A.end .. B.begin]
    // scores begins===ends===1 yet destroys A.end + B.begin, nesting B inside A.
    const state = stateOf(twoAdjacentFieldsDoc());
    expect(countMarkers(state.doc)).toBe(8);
    const ends = markerPositions(state.doc, 'end');
    const begins = markerPositions(state.doc, 'begin');
    const aEnd = ends[0]; // field A's end
    const bBegin = begins[1]; // field B's begin (immediately after A.end)
    const next = state.apply(state.tr.delete(aEnd, bBegin + 1)); // remove A.end + B.begin
    expect(next.doc.eq(state.doc)).toBe(true); // VETOED — corruption prevented
    expect(countMarkers(next.doc)).toBe(8);
  });
});

describe('fieldMarker firewall — legitimate edits are ALLOWED', () => {
  test('editing FAR from any marker is ALLOWED', () => {
    const doc = schema.nodes.doc.create({}, [
      singleParaTocDoc().firstChild as PMNode,
      schema.nodes.paragraph.create({}, [schema.text('hello')]),
    ]);
    const state = stateOf(doc);
    const helloStart = state.doc.content.size - 'hello'.length - 1;
    const next = state.apply(state.tr.insertText('!', helloStart));
    expect(next.doc.eq(state.doc)).toBe(false); // applied
    expect(countMarkers(next.doc)).toBe(4);
  });

  test('deleting the WHOLE field (balanced begin..end) is ALLOWED', () => {
    const state = stateOf(singleParaTocDoc());
    const from = posOfMarker(state.doc, 'begin');
    let to = 0;
    state.doc.descendants((n, p) => {
      if (n.type.name === 'fieldMarker' && n.attrs.charType === 'end') to = p + n.nodeSize;
    });
    const next = state.apply(state.tr.delete(from, to));
    expect(countMarkers(next.doc)).toBe(0); // whole field removed together → no orphan
  });

  test('Mod-a then delete of a doc containing a complete TOC is ALLOWED', () => {
    const doc = schema.nodes.doc.create({}, [
      singleParaTocDoc().firstChild as PMNode,
      schema.nodes.paragraph.create({}, [schema.text('body')]),
    ]);
    const state = stateOf(doc);
    const next = state.apply(state.tr.delete(0, state.doc.content.size)); // balanced whole-field removal
    expect(countMarkers(next.doc)).toBe(0);
  });

  test('wrapping a TOC paragraph in a blockSdt PRESERVES markers (ReplaceAroundStep gap not vetoed)', () => {
    const state0 = stateOf(singleParaTocDoc());
    expect(countMarkers(state0.doc)).toBe(4);
    const inside = state0.apply(
      state0.tr.setSelection(TextSelection.create(state0.doc, posInsideText(state0.doc, 'Heading')))
    );
    let wrapped: EditorState | null = null;
    const ok = wrapIn(schema.nodes.blockSdt)(inside, (tr) => {
      wrapped = inside.apply(tr);
    });
    expect(ok).toBe(true);
    expect(wrapped).not.toBeNull();
    const w = wrapped as unknown as EditorState;
    expect(w.doc.eq(inside.doc)).toBe(false); // applied → NOT vetoed
    expect(countMarkers(w.doc)).toBe(4); // markers preserved through the gap
    expect((w.doc.firstChild as PMNode).type.name).toBe('blockSdt'); // really wrapped
  });

  test('joining the last TOC result line up is ALLOWED and preserves markers (no over-block)', () => {
    // Removing Layer B: a join relocates markers in order, so it is safe and the
    // firewall allows it. Cursor at the start of the tail paragraph (holds `end`).
    const state = stateOf(multiParaTocDoc());
    const tailStart = (state.doc.firstChild as PMNode).nodeSize + 1;
    const at = state.apply(state.tr.setSelection(TextSelection.create(state.doc, tailStart)));
    let joined = at;
    joinBackward(at, (tr) => {
      joined = at.apply(tr);
    });
    expect(joined.doc.childCount).toBe(1); // merged into one paragraph
    expect(countMarkers(joined.doc)).toBe(4); // all markers survive, in order
  });
});

describe('fieldMarker firewall — deliberate deletion via bypass', () => {
  test('deleteFieldAtCursor removes the whole field via the bypass meta', () => {
    const state = stateOf(singleParaTocDoc());
    const sepPos = posOfMarker(state.doc, 'separate');
    const cursorState = state.apply(
      state.tr.setSelection(TextSelection.create(state.doc, sepPos + 1))
    );
    let applied: EditorState | null = null;
    deleteFieldAtCursor(cursorState, (tr) => {
      applied = cursorState.apply(tr);
    });
    expect(applied).not.toBeNull();
    expect(countMarkers((applied as unknown as EditorState).doc)).toBe(0);
  });

  test('a tr carrying the scoped bypass meta is ALLOWED even when it orphans a marker', () => {
    const state = stateOf(singleParaTocDoc());
    const beginPos = posOfMarker(state.doc, 'begin');
    const tr = state.tr
      .delete(beginPos, beginPos + 1)
      .setMeta(fieldMarkerGuardKey, FIELD_MARKER_BYPASS);
    const next = state.apply(tr);
    expect(countMarkers(next.doc)).toBe(3); // bypass let the orphaning delete through
  });
});

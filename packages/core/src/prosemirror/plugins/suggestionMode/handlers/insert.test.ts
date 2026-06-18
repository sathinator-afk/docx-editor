/**
 * Tests for applyTrackedParagraphInsert — landing multi-paragraph AI output as a
 * single tracked insertion (insertion marks on the new runs + pPrIns on each new
 * paragraph break, all sharing one revision id, accept/reject as one block).
 *
 * The function takes an EditorView but only touches `view.state`/`view.dispatch`
 * (tr.scrollIntoView() is DOM-free), so a fake view that applies dispatched
 * transactions is enough — no jsdom needed, matching the rest of this suite.
 */

import { describe, test, expect } from 'bun:test';
import { Schema, type Node as PMNode } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { applyTrackedParagraphInsert } from './insert';

// Schema with the paragraph-mark revision attrs (pPrIns/pPrDel) the real editor
// carries, plus insertion/deletion marks.
const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: { pPrIns: { default: null }, pPrDel: { default: null } },
      toDOM: () => ['p', 0],
    },
    text: { group: 'inline' },
  },
  marks: {
    insertion: {
      attrs: { revisionId: { default: 0 }, author: { default: '' }, date: { default: '' } },
      toDOM: () => ['ins', 0],
    },
    deletion: {
      attrs: { revisionId: { default: 0 }, author: { default: '' }, date: { default: '' } },
      toDOM: () => ['del', 0],
    },
  },
});

function docOf(...paras: string[]): PMNode {
  return schema.node(
    'doc',
    null,
    paras.map((t) => schema.node('paragraph', null, t ? [schema.text(t)] : []))
  );
}

/** Minimal EditorView stand-in: applies dispatched transactions to its state. */
function fakeView(doc: PMNode) {
  let state = EditorState.create({ schema, doc });
  return {
    get state() {
      return state;
    },
    dispatch(tr: import('prosemirror-state').Transaction) {
      state = state.apply(tr);
    },
    focus() {},
  };
}

/** Paragraphs as [text, pPrIns-revisionId | null]. */
function paragraphs(state: EditorState): Array<[string, number | null]> {
  const out: Array<[string, number | null]> = [];
  state.doc.forEach((node) => {
    if (node.type.name !== 'paragraph') return;
    const ins = node.attrs.pPrIns as { revisionId: number } | null;
    out.push([node.textContent, ins ? ins.revisionId : null]);
  });
  return out;
}

/** Inserted runs: text carrying an insertion mark + its revisionId. */
function insertedRuns(state: EditorState): Array<[string, number]> {
  const out: Array<[string, number]> = [];
  state.doc.descendants((node) => {
    if (!node.isText) return;
    const m = node.marks.find((mk) => mk.type.name === 'insertion');
    if (m) out.push([node.text!, m.attrs.revisionId as number]);
  });
  return out;
}

describe('applyTrackedParagraphInsert', () => {
  test('inserts multiple paragraphs as one tracked revision', () => {
    // Cursor at end of "Intro." (pos 7: <p>(1) Intro.(6) (7=end content)).
    const view = fakeView(docOf('Intro.'));
    const at = view.state.doc.content.size - 1; // inside the paragraph, at its end
    const attrs = applyTrackedParagraphInsert(view as never, at, at, ['Alpha', 'Beta', 'Gamma'], {
      active: true,
      author: 'AI',
    });
    expect(attrs).not.toBeNull();
    const rev = attrs!.revisionId;

    // Structure: [Intro.Alpha][Beta][Gamma] — first two breaks are new (pPrIns),
    // the trailing paragraph keeps the original (null) end mark.
    expect(paragraphs(view.state)).toEqual([
      ['Intro.Alpha', rev],
      ['Beta', rev],
      ['Gamma', null],
    ]);

    // Every inserted run carries the insertion mark with the SAME revision id.
    const runs = insertedRuns(view.state);
    expect(runs.map((r) => r[0])).toEqual(['Alpha', 'Beta', 'Gamma']);
    expect(runs.every((r) => r[1] === rev)).toBe(true);

    // The original "Intro." run is untouched (no insertion mark).
    expect(view.state.doc.firstChild!.firstChild!.marks.length).toBe(0);
  });

  test('inserting at the very END of the document does not throw (regression)', () => {
    // Cursor at the top-level end boundary (insertAt === doc.content.size). This
    // is the position the JS bridge clamps `to` to, and it used to throw
    // `RangeError: There is no position before the top-level node`.
    const view = fakeView(docOf('Intro.'));
    const end = view.state.doc.content.size; // top-level boundary, not inside a paragraph
    const attrs = applyTrackedParagraphInsert(view as never, end, end, ['Alpha', 'Beta'], {
      active: true,
      author: 'AI',
    });
    expect(attrs).not.toBeNull();
    const rev = attrs!.revisionId;
    // Both inserted runs landed as tracked insertions sharing one revision id.
    const runs = insertedRuns(view.state);
    expect(runs.map((r) => r[0])).toEqual(['Alpha', 'Beta']);
    expect(runs.every((r) => r[1] === rev)).toBe(true);
    // One new paragraph break for two inserted paragraphs at end-of-doc; the
    // trailing paragraph must NOT carry pPrIns, and exactly one earlier para must.
    const paras = paragraphs(view.state);
    expect(paras[paras.length - 1][1]).toBeNull();
    expect(paras.filter((p) => p[1] === rev).length).toBe(1);
    // The original "Intro." text survives un-marked.
    expect(view.state.doc.textContent).toContain('Intro.');
  });

  test('single paragraph inserts inline with no new paragraph break', () => {
    const view = fakeView(docOf('Start.'));
    const at = view.state.doc.content.size - 1;
    const attrs = applyTrackedParagraphInsert(view as never, at, at, ['More'], {
      active: true,
      author: 'AI',
    });
    expect(attrs).not.toBeNull();
    // Still one paragraph, no pPrIns; just the inserted run is marked.
    expect(paragraphs(view.state)).toEqual([['Start.More', null]]);
    expect(insertedRuns(view.state)).toEqual([['More', attrs!.revisionId]]);
  });

  test('splits the host paragraph, keeping trailing text in the last block', () => {
    // Cursor BETWEEN "Left" and "Right" in one paragraph.
    const view = fakeView(docOf('LeftRight'));
    const at = 5; // <p>(1)Left(5)Right(10)
    void at;
    const state0 = view.state;
    const sel = TextSelection.create(state0.doc, 5);
    view.dispatch(state0.tr.setSelection(sel));
    const attrs = applyTrackedParagraphInsert(view as never, 5, 5, ['One', 'Two'], {
      active: true,
      author: 'AI',
    });
    const rev = attrs!.revisionId;
    // [LeftOne][Two Right] — One merges after Left, Two prepends Right.
    expect(paragraphs(view.state)).toEqual([
      ['LeftOne', rev],
      ['TwoRight', null],
    ]);
  });

  test('replace mode: a non-empty range is tracked-deleted, sharing the date', () => {
    const view = fakeView(docOf('Keep DROP tail'));
    // Delete "DROP" (positions 6..10) and insert two paragraphs there.
    const attrs = applyTrackedParagraphInsert(view as never, 6, 10, ['New', 'Lines'], {
      active: true,
      author: 'AI',
    });
    const rev = attrs!.revisionId;
    // The deleted run carries a deletion mark sharing the insertion's date.
    let deletedText = '';
    let deletionDate = '';
    view.state.doc.descendants((node) => {
      if (!node.isText) return;
      const d = node.marks.find((m) => m.type.name === 'deletion');
      if (d) {
        deletedText += node.text;
        deletionDate = d.attrs.date as string;
      }
    });
    expect(deletedText).toBe('DROP');
    expect(deletionDate).toBe(attrs!.date);
    // The new paragraphs landed as a tracked insertion.
    expect(insertedRuns(view.state).map((r) => r[0])).toEqual(['New', 'Lines']);
    expect(insertedRuns(view.state).every((r) => r[1] === rev)).toBe(true);
  });
});

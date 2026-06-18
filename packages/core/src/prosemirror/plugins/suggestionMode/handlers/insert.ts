/**
 * Suggesting-mode insertion path. Handles both pure inserts and the
 * replace case (selection non-empty → mark selection deleted, then insert).
 */

import { Fragment, Slice } from 'prosemirror-model';
import type { Node as PMNode, Mark, MarkType } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

import { findAdjacentRevision } from '../adjacency';
import { makeMarkAttrs } from '../markAttrs';
import { SUGGESTION_META, type MarkAttrs, type SuggestionModeState } from '../state';
import { markRangeAsDeleted } from './delete';

/**
 * Add the insertion mark to every inline node in `[from, to)` that can carry
 * it and doesn't already hold a tracked-change mark. Shared by the paste
 * handler and the plugin's append-transaction catch-all so both stamp pasted
 * / dropped content the same way: text is the short-circuit (a leaf text
 * node's own markSet is empty, so `allowsMarkType` is false even though the
 * paragraph permits the mark), inline atoms (image, shape) go through the
 * `allowsMarkType` arm, and anything already carrying an insertion/deletion
 * mark is left alone so we never overwrite another author's change.
 */
export function markRangeAsInserted(
  tr: Transaction,
  doc: PMNode,
  from: number,
  to: number,
  insertionType: MarkType,
  deletionType: MarkType | undefined,
  attrs: MarkAttrs
): void {
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText && !(node.isInline && node.type.allowsMarkType(insertionType))) return;
    const hasTrackedMark = node.marks.some(
      (m) => m.type === insertionType || (deletionType && m.type === deletionType)
    );
    if (hasTrackedMark) return;
    const start = Math.max(pos, from);
    const end = Math.min(pos + node.nodeSize, to);
    if (start >= end) return;
    tr.addMark(start, end, insertionType.create(attrs));
  });
}

/**
 * Insert text as a tracked insertion, optionally marking replaced selection as deletion.
 */
export function applySuggestionInsert(
  view: EditorView,
  from: number,
  to: number,
  text: string,
  pluginState: SuggestionModeState
): boolean {
  const insertionType = view.state.schema.marks.insertion;
  if (!insertionType) return false;

  const tr = view.state.tr;
  tr.setMeta(SUGGESTION_META, true);

  const insertAttrs =
    findAdjacentRevision(view.state.doc, from, 'insertion', pluginState.author) ||
    makeMarkAttrs(pluginState);

  if (from !== to) {
    const deletionType = view.state.schema.marks.deletion;
    if (deletionType) {
      // Replace op: pass the insertion's date down so the deletion shares
      // the (author, date) triple — that's what extractTrackedChanges
      // uses to detect adjacent del+ins and fold them into one
      // 'replacement' card. The `w:id` stays distinct so we don't trip
      // the OOXML move-pair serializer (fromProseDoc/paragraph.ts:340).
      markRangeAsDeleted(
        tr,
        view.state.doc,
        from,
        to,
        insertionType,
        deletionType,
        pluginState,
        insertAttrs.date
      );
    }
  }

  const insertAt = tr.mapping.map(to);
  tr.insertText(text, insertAt, insertAt);

  // Strip inherited deletion marks — new text must never be marked as deleted.
  const deletionType = view.state.schema.marks.deletion;
  if (deletionType) {
    tr.removeMark(insertAt, insertAt + text.length, deletionType);
  }

  // Apply the correct insertion mark. If the cursor was inside an existing
  // insertion by the same author, insertText already inherited that mark and
  // insertAttrs will match — addMark is effectively a no-op that preserves
  // the continuous mark span. We intentionally do NOT removeMark(insertionType)
  // first, because that fragments the mark span and creates a nested change.
  tr.addMark(insertAt, insertAt + text.length, insertionType.create(insertAttrs));

  view.dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Like {@link applySuggestionInsert} but PRESERVES the replaced span's run
 * formatting on the newly inserted text, and works without globally toggling
 * suggestion mode (pass a `{ active: true, author }` state directly).
 *
 * `applySuggestionInsert` uses `tr.insertText`, which inherits only the boundary
 * marks — so an accepted rewrite can lose the original bold/italic/colour. This
 * variant copies the formatting marks from the start of the replaced range onto
 * the inserted text via `schema.text(text, marks)`, dropping only the
 * tracked-change / comment marks (those are stamped fresh, not inherited).
 *
 * The deletion side reuses {@link markRangeAsDeleted} verbatim (text stays in the
 * doc, struck through, sharing the insertion's date so the sidebar folds the
 * pair into one "replacement" card). Designed to be driven programmatically —
 * e.g. an AI assist proposing an edit as a tracked change.
 *
 * @returns the minted insertion mark attrs (`{ revisionId, author, date }`) so
 *   callers can target accept/reject, or `null` if the schema lacks the marks.
 */
export function applyTrackedReplaceKeepingMarks(
  view: EditorView,
  from: number,
  to: number,
  text: string,
  pluginState: SuggestionModeState
): MarkAttrs | null {
  const { schema } = view.state;
  const insertionType = schema.marks.insertion;
  const deletionType = schema.marks.deletion;
  if (!insertionType || !deletionType) return null;

  // Run formatting on the first char of the replaced span. Real Mark objects
  // carry full attrs (colour/theme tints, font slots) losslessly. Drop tracked-
  // change / comment marks — the new text gets a fresh insertion mark and must
  // not inherit a stale ins/del/comment.
  const tracked = new Set(['insertion', 'deletion', 'comment']);
  const keep: readonly Mark[] = view.state.doc
    .resolve(from)
    .marks()
    .filter((m) => !tracked.has(m.type.name));

  const tr = view.state.tr;
  tr.setMeta(SUGGESTION_META, true); // we author the marks; skip the catch-all

  const insertAttrs =
    findAdjacentRevision(view.state.doc, from, 'insertion', pluginState.author) ||
    makeMarkAttrs(pluginState);

  // 1) Tracked-delete the old span (shares the insertion's date → one card).
  if (from !== to) {
    markRangeAsDeleted(
      tr,
      view.state.doc,
      from,
      to,
      insertionType,
      deletionType,
      pluginState,
      insertAttrs.date
    );
  }

  // 2) Insert the new text AFTER the struck-through span, carrying formatting.
  const insertAt = tr.mapping.map(to);
  if (text.length > 0) {
    tr.insert(insertAt, schema.text(text, keep));
    const end = insertAt + text.length;
    tr.removeMark(insertAt, end, deletionType); // new text is never a deletion
    tr.addMark(insertAt, end, insertionType.create(insertAttrs));
  }

  view.dispatch(tr.scrollIntoView());
  return insertAttrs;
}

/**
 * Insert MULTIPLE paragraphs of new text as a single tracked insertion — the
 * structural cousin of {@link applyTrackedReplaceKeepingMarks}. Used to land
 * multi-paragraph AI output (a Draft / extracted key points) as one revision the
 * author can accept or reject inline, rather than as an untracked dump.
 *
 * Mechanics: build an OPEN (openStart=openEnd=1) slice of paragraph nodes and
 * `tr.replace` it at the insertion point — the same merge ProseMirror uses when
 * pasting multi-line text, so the target paragraph splits into `Pleft+first …
 * last+Pright`. Then every new paragraph break (all but the trailing one, which
 * is the original paragraph's pre-existing mark) gets `pPrIns`, and all inserted
 * runs carry the `insertion` mark — ALL sharing one `{revisionId, author, date}`
 * triple so accept/reject and the sidebar treat the block as one change.
 *
 * A non-empty `[from, to)` is tracked-deleted first (replace semantics), sharing
 * the insertion's date so the del+ins fold into one "replacement" card.
 *
 * @returns the minted insertion attrs, or `null` if the schema lacks the marks.
 */
export function applyTrackedParagraphInsert(
  view: EditorView,
  from: number,
  to: number,
  paragraphs: string[],
  pluginState: SuggestionModeState
): MarkAttrs | null {
  const { schema } = view.state;
  const insertionType = schema.marks.insertion;
  const deletionType = schema.marks.deletion;
  const paragraphType = schema.nodes.paragraph;
  if (!insertionType || !deletionType || !paragraphType) return null;

  const paras = paragraphs.filter((p) => p.length > 0);
  if (!paras.length) return null;

  // Run formatting at the insertion point, minus tracked-change / comment marks
  // (the new text gets a fresh insertion mark, never an inherited ins/del/comment).
  const tracked = new Set(['insertion', 'deletion', 'comment']);
  const keep: readonly Mark[] = view.state.doc
    .resolve(from)
    .marks()
    .filter((m) => !tracked.has(m.type.name));

  // Inserted paragraphs inherit the target paragraph's block attrs (style,
  // alignment, indent) so they match their surroundings — minus any paragraph-
  // mark revision, which we stamp explicitly below.
  const targetPara = view.state.doc.resolve(from).parent;
  const baseAttrs = { ...targetPara.attrs, pPrIns: null, pPrDel: null };

  const tr = view.state.tr;
  tr.setMeta(SUGGESTION_META, true); // we author the marks; skip the catch-all

  const insertAttrs =
    findAdjacentRevision(view.state.doc, from, 'insertion', pluginState.author) ||
    makeMarkAttrs(pluginState);

  // 1) Tracked-delete a non-empty selection first (shares the insertion's date).
  if (from !== to) {
    markRangeAsDeleted(
      tr,
      view.state.doc,
      from,
      to,
      insertionType,
      deletionType,
      pluginState,
      insertAttrs.date
    );
  }

  // 2) Build the open paragraph slice and splice it in at the insertion point.
  const insertionMark = insertionType.create(insertAttrs);
  const marks = [...keep, insertionMark];
  const nodes = paras.map((t) => paragraphType.create(baseAttrs, schema.text(t, marks)));
  const slice = new Slice(Fragment.fromArray(nodes), 1, 1);
  const insertAt = tr.mapping.map(to);
  tr.replace(insertAt, insertAt, slice);

  // 3) Stamp pPrIns on every NEW paragraph break. After the open-merge the
  //    affected paragraphs are: the block containing `insertAt`, then the next
  //    (paras.length - 1) blocks. Each break BUT the trailing one is new (the
  //    last block keeps the original paragraph's pre-existing end mark), so the
  //    first (paras.length - 1) affected paragraphs get pPrIns.
  //
  //    Resolve the host paragraph carefully: when the slice is spliced at a
  //    top-level boundary (e.g. the very END of the document, insertAt ===
  //    doc.content.size), resolve(insertAt) returns depth 0 and before() would
  //    throw `RangeError: There is no position before the top-level node`. The
  //    open-start slice merges its first paragraph into the block just BEFORE
  //    that boundary, so step back one position to land inside that host.
  const size = tr.doc.content.size;
  let $at = tr.doc.resolve(Math.min(insertAt, size));
  if ($at.depth === 0 && insertAt > 0) {
    $at = tr.doc.resolve(Math.min(insertAt, size) - 1);
  }
  if ($at.depth > 0) {
    let paraPos = $at.before($at.depth);
    for (let i = 0; i < paras.length - 1; i++) {
      const para = tr.doc.nodeAt(paraPos);
      if (!para || para.type.name !== 'paragraph') break;
      tr.setNodeMarkup(paraPos, undefined, { ...para.attrs, pPrIns: insertAttrs });
      paraPos += para.nodeSize; // setNodeMarkup preserves size
    }
  }

  view.dispatch(tr.scrollIntoView());
  return insertAttrs;
}

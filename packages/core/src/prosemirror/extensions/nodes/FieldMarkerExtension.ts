/**
 * Field Marker Extension — preserves raw field characters and instruction text.
 *
 * A complex field that fits in ONE paragraph becomes a single atomic `field`
 * node (see FieldExtension). But a field whose result spans MULTIPLE paragraphs
 * — most notably a Table of Contents (TOC) — cannot be one node: it is stored in
 * the model as separate `fieldChar` (begin/separate/end) and `instrText` runs.
 *
 * Before this extension those run-content pieces had no PM representation, so
 * `toProseDoc` dropped them and the field's instruction (e.g. the whole TOC)
 * vanished on the editor save path (toProseDoc → fromProseDoc → serialize).
 *
 * This node carries each raw marker through the PM document invisibly (field
 * codes are hidden in Word's default view), so `fromProseDoc` can re-emit the
 * exact `fieldChar`/`instrText` run content and the serializer round-trips it.
 *
 * Because the markers are INVISIBLE (display:none atoms), a careless live edit
 * can silently destroy a document's TOC. A `filterTransaction` firewall guards
 * against that: it vetoes any transaction whose ReplaceStep/ReplaceAroundStep
 * removes field markers WITHOUT removing one-or-more WHOLE, properly-nested
 * fields. Source-agnostic — covers range-delete, type-over-selection, cut,
 * paste-replace, partial select-all, drag-out, and programmatic deletes — using
 * an ORDER-AWARE begin/end depth walk, so a delete spanning the gap between two
 * ADJACENT fields (count-balanced yet corrupting) is still rejected. Paragraph
 * JOINS (Backspace/Delete at a field-paragraph boundary) need no keymap guard:
 * they RELOCATE markers without removing them and preserve begin<end order, so
 * the field stays well-formed and the firewall correctly allows them. Undo/redo
 * and a plugin-key-scoped bypass meta are exempt; `deleteFieldAtCursor` (opt-in)
 * deletes a whole field on purpose via that bypass.
 */

import { Plugin, PluginKey } from 'prosemirror-state';
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { ReplaceStep, ReplaceAroundStep } from 'prosemirror-transform';
import { isHistoryTransaction } from 'prosemirror-history';
import { createNodeExtension } from '../create';
import type { ExtensionRuntime } from '../types';

/** PluginKey for the field-marker firewall. Also scopes the bypass meta so no
 *  unrelated string-keyed meta can accidentally disable the guard. */
export const fieldMarkerGuardKey = new PluginKey('fieldMarkerGuard');

/** Set `tr.setMeta(fieldMarkerGuardKey, FIELD_MARKER_BYPASS)` to opt a
 *  transaction OUT of the firewall (deliberate whole-field delete; conversion
 *  / normalization rewrites). */
export const FIELD_MARKER_BYPASS = 'bypass';

/**
 * A removed range is SAFE iff it removes only WHOLE, properly-nested fields.
 *
 * We walk the fieldMarkers in the range in document order with a begin/end depth
 * counter. The range is UNSAFE if:
 *   - depth ever goes negative (an `end` whose matching `begin` is outside the
 *     range — e.g. the strand BETWEEN two adjacent fields, [A.end .. B.begin]),
 *   - a lone `instrText`/`separate` appears at depth 0 (its field head is
 *     outside the range — e.g. selecting just the TOC instruction), or
 *   - depth is non-zero at the end (a dangling `begin`).
 *
 * A count-only `begins === ends` test is NOT sufficient: deleting
 * [fieldA.end .. fieldB.begin] is count-balanced (one begin, one end) yet
 * destroys A's `end` and B's `begin`, nesting B's runs inside A — corruption.
 */
function rangeRemovalIsSafe(doc: PMNode, from: number, to: number): boolean {
  if (to <= from) return true;
  let depth = 0;
  let safe = true;
  doc.nodesBetween(from, to, (node) => {
    if (!safe) return false;
    if (node.type.name !== 'fieldMarker') return true;
    const { markerType, charType } = node.attrs as { markerType: string; charType: string };
    if (markerType === 'fieldChar') {
      if (charType === 'begin') {
        depth += 1;
      } else if (charType === 'end') {
        depth -= 1;
        if (depth < 0) safe = false; // `end` whose `begin` is outside the range
      } else if (depth < 1) {
        safe = false; // `separate` outside any in-range field
      }
    } else if (depth < 1) {
      safe = false; // `instrText` outside any in-range field
    }
    return false; // atom leaf — no children
  });
  return safe && depth === 0; // depth !== 0 → a dangling `begin`
}

/** Firewall: veto any transaction whose ANY ReplaceStep/ReplaceAroundStep removes
 *  a marker without removing its whole field. Source-agnostic. */
function fieldMarkerGuardPlugin(): Plugin {
  return new Plugin({
    key: fieldMarkerGuardKey,
    filterTransaction(tr: Transaction): boolean {
      if (!tr.docChanged) return true;
      if (tr.getMeta(fieldMarkerGuardKey) === FIELD_MARKER_BYPASS) return true;
      // Undo/redo legitimately restores arbitrary prior docs — never veto it.
      // Uses prosemirror-history's EXPORTED helper (robust; no meta-key guessing).
      if (isHistoryTransaction(tr)) return true;
      for (let i = 0; i < tr.steps.length; i++) {
        const step = tr.steps[i];
        // Each step's from/to are in the coords of the doc BEFORE that step,
        // i.e. tr.docs[i] (tr.docs[0] === the original doc).
        const stepDoc = tr.docs[i];
        if (!stepDoc) continue;
        if (step instanceof ReplaceStep) {
          if (!rangeRemovalIsSafe(stepDoc, step.from, step.to)) return false;
        } else if (step instanceof ReplaceAroundStep) {
          // Only [from,gapFrom) and [gapTo,to) are REMOVED; the gap
          // [gapFrom,gapTo) is preserved (moved into the new structure), so a
          // marker inside the gap — e.g. a wrap/lift around a TOC paragraph — is
          // NOT removed and must not trigger a veto.
          if (!rangeRemovalIsSafe(stepDoc, step.from, step.gapFrom)) return false;
          if (!rangeRemovalIsSafe(stepDoc, step.gapTo, step.to)) return false;
        }
      }
      return true;
    },
  });
}

/** Walk to the field range [begin..end) containing `pos`, pairing the nearest
 *  preceding `begin` with the next `end`. */
function fieldRangeAt(state: EditorState, pos: number): { from: number; to: number } | null {
  let beginPos: number | null = null;
  let endPos: number | null = null;
  let open = false;
  state.doc.descendants((node, p) => {
    if (node.type.name !== 'fieldMarker') return true;
    const a = node.attrs as { markerType: string; charType: string };
    if (a.markerType === 'fieldChar' && a.charType === 'begin') {
      if (p <= pos) {
        beginPos = p;
        open = true;
        endPos = null;
      }
    } else if (a.markerType === 'fieldChar' && a.charType === 'end' && open) {
      if (endPos === null) endPos = p + node.nodeSize;
    }
    return false;
  });
  if (beginPos === null || endPos === null) return null;
  if (pos < beginPos || pos > endPos) return null;
  return { from: beginPos, to: endPos };
}

/** Deliberate whole-field delete (cursor inside a field), bypassing the
 *  firewall. Not bound by default; expose via a chord or menu if desired. */
export const deleteFieldAtCursor: Command = (state, dispatch) => {
  const sel = state.selection;
  if (!sel.empty) return false;
  const range = fieldRangeAt(state, sel.$from.pos);
  if (!range) return false;
  if (dispatch) {
    dispatch(
      state.tr
        .delete(range.from, range.to)
        .setMeta(fieldMarkerGuardKey, FIELD_MARKER_BYPASS)
        .scrollIntoView()
    );
  }
  return true;
};

export const FieldMarkerExtension = createNodeExtension({
  name: 'fieldMarker',
  schemaNodeName: 'fieldMarker',
  nodeSpec: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: false,
    attrs: {
      /** 'fieldChar' or 'instrText' */
      markerType: { default: 'fieldChar' },
      /** fieldChar: 'begin' | 'separate' | 'end' */
      charType: { default: 'begin' },
      /** instrText: the field instruction (e.g. 'TOC \\o "1-3" \\h \\z \\u') */
      instr: { default: '' },
      fldLock: { default: false },
      dirty: { default: false },
    },
    parseDOM: [
      {
        tag: 'span.docx-field-code',
        getAttrs(dom) {
          const el = dom as HTMLElement;
          return {
            markerType: el.dataset.markerType || 'fieldChar',
            charType: el.dataset.charType || 'begin',
            instr: el.dataset.instr || '',
            fldLock: el.dataset.fldLock === 'true',
            dirty: el.dataset.dirty === 'true',
          };
        },
      },
    ],
    toDOM(node) {
      const a = node.attrs as {
        markerType: string;
        charType: string;
        instr: string;
        fldLock: boolean;
        dirty: boolean;
      };
      // Field codes are hidden in the default view — carry the data invisibly.
      return [
        'span',
        {
          class: 'docx-field-code',
          'data-marker-type': a.markerType,
          'data-char-type': a.charType,
          'data-instr': a.instr,
          ...(a.fldLock ? { 'data-fld-lock': 'true' } : {}),
          ...(a.dirty ? { 'data-dirty': 'true' } : {}),
          style: 'display:none',
        },
      ];
    },
  },
  onSchemaReady(): ExtensionRuntime {
    return {
      plugins: [fieldMarkerGuardPlugin()],
    };
  },
});

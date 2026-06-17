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
 */

import { createNodeExtension } from '../create';

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
});

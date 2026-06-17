/**
 * Regression: a multi-paragraph field (notably a Table of Contents) must
 * survive the EDITOR save path (toProseDoc → fromProseDoc → serialize), not only
 * the headless `parseDocx → repackDocx` path.
 *
 * A single-paragraph field is one atomic `field` node. But a field whose result
 * spans paragraphs is stored as raw `fieldChar` (begin/separate/end) + `instrText`
 * RUN content. `toProseDoc` used to drop those (no run-content case), so editing
 * and saving a document silently lost its whole TOC. The `fieldMarker` node now
 * carries each marker through PM so the field round-trips.
 *
 * See the TOC-drop investigation (Cuneon beta readiness).
 */

import { describe, test, expect } from 'bun:test';
import { schema } from '../../schema';
import { fromProseDoc } from '../fromProseDoc';
import { toProseDoc } from '../toProseDoc';
import type { Document, Paragraph, Run } from '../../../types/document';

const TOC = 'TOC \\o "1-3" \\h \\z \\u';

function flatRunContent(para: Paragraph) {
  return para.content.flatMap((c) => (c as Run).type === 'run' ? (c as Run).content : []);
}

describe('multi-paragraph field markers round-trip through the PM save path', () => {
  test('fromProseDoc re-emits fieldChar/instrText run content for fieldMarker nodes', () => {
    const para = schema.nodes.paragraph.create({}, [
      schema.node('fieldMarker', { markerType: 'fieldChar', charType: 'begin' }),
      schema.node('fieldMarker', { markerType: 'instrText', instr: TOC }),
      schema.node('fieldMarker', { markerType: 'fieldChar', charType: 'separate' }),
      schema.text('Heading 1\t1'),
      schema.node('fieldMarker', { markerType: 'fieldChar', charType: 'end' }),
    ]);
    const doc = schema.nodes.doc.create({}, [para]);
    const model = fromProseDoc(doc).package?.document?.content?.[0] as Paragraph;
    const flat = flatRunContent(model) as Array<{ type: string; text?: string; charType?: string }>;

    const instr = flat.find((rc) => rc.type === 'instrText');
    expect(instr?.text).toBe(TOC);
    const chars = flat.filter((rc) => rc.type === 'fieldChar').map((rc) => rc.charType);
    expect(chars).toEqual(['begin', 'separate', 'end']);
  });

  test('toProseDoc represents raw field runs as fieldMarker nodes (no loss)', () => {
    const model = {
      package: {
        document: {
          content: [
            {
              type: 'paragraph',
              content: [
                { type: 'run', content: [{ type: 'fieldChar', charType: 'begin' }] },
                { type: 'run', content: [{ type: 'instrText', text: TOC }] },
                { type: 'run', content: [{ type: 'fieldChar', charType: 'separate' }] },
                { type: 'run', content: [{ type: 'text', text: 'Entry' }] },
                { type: 'run', content: [{ type: 'fieldChar', charType: 'end' }] },
              ],
            },
          ],
        },
      },
    } as unknown as Document;

    const pm = toProseDoc(model);
    let instrSeen = false;
    let charTypes: string[] = [];
    pm.descendants((node) => {
      if (node.type.name === 'fieldMarker') {
        const a = node.attrs as { markerType: string; charType: string; instr: string };
        if (a.markerType === 'instrText' && a.instr === TOC) instrSeen = true;
        if (a.markerType === 'fieldChar') charTypes.push(a.charType);
      }
    });
    expect(instrSeen).toBe(true);
    expect(charTypes).toEqual(['begin', 'separate', 'end']);
  });

  test('full round-trip toProseDoc -> fromProseDoc keeps the instruction exactly once', () => {
    const para = schema.nodes.paragraph.create({}, [
      schema.node('fieldMarker', { markerType: 'fieldChar', charType: 'begin' }),
      schema.node('fieldMarker', { markerType: 'instrText', instr: TOC }),
      schema.node('fieldMarker', { markerType: 'fieldChar', charType: 'separate' }),
      schema.text('x'),
      schema.node('fieldMarker', { markerType: 'fieldChar', charType: 'end' }),
    ]);
    const doc = schema.nodes.doc.create({}, [para]);
    const reloaded = toProseDoc(fromProseDoc(doc));
    let instrCount = 0;
    reloaded.descendants((node) => {
      if (node.type.name === 'fieldMarker' && (node.attrs as { instr: string }).instr === TOC) instrCount += 1;
    });
    expect(instrCount).toBe(1);
  });
});

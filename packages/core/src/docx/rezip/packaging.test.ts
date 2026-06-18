import { describe, test, expect } from 'bun:test';
import JSZip from 'jszip';
import type { Document, Footnote, Endnote } from '../../types/document';
import { serializeFootnotesToZip, serializeEndnotesToZip } from './packaging';

function makeDoc(overrides: Partial<Document['package']> = {}): Document {
  return {
    package: {
      document: { content: [] },
      ...overrides,
    },
  };
}

const SEPARATOR_ENDNOTE: Endnote = {
  type: 'endnote',
  id: -1,
  noteType: 'separator',
  content: [{ type: 'paragraph', content: [{ type: 'run', content: [{ type: 'separator' }] }] }],
};

const NORMAL_ENDNOTE: Endnote = {
  type: 'endnote',
  id: 1,
  noteType: 'normal',
  content: [
    {
      type: 'paragraph',
      content: [
        { type: 'run', content: [{ type: 'endnoteRefMark' }] },
        { type: 'run', content: [{ type: 'text', text: ' Edited endnote body.' }] },
      ],
    },
  ],
};

const SEPARATOR_FOOTNOTE: Footnote = {
  type: 'footnote',
  id: -1,
  noteType: 'separator',
  content: [{ type: 'paragraph', content: [{ type: 'run', content: [{ type: 'separator' }] }] }],
};

const NORMAL_FOOTNOTE: Footnote = {
  type: 'footnote',
  id: 1,
  noteType: 'normal',
  content: [
    {
      type: 'paragraph',
      content: [{ type: 'run', content: [{ type: 'text', text: 'A footnote.' }] }],
    },
  ],
};

async function entryText(zip: JSZip, path: string): Promise<string | null> {
  const f = zip.file(path);
  return f ? await f.async('text') : null;
}

/** Minimal package skeleton (Content_Types + document rels) so part registration
 *  has files to amend — mirrors a real .docx that lacks a notes part. */
function seedPackage(zip: JSZip): void {
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
  );
}

describe('serializeEndnotesToZip', () => {
  test('writes word/endnotes.xml with separators ahead of normal notes', async () => {
    const zip = new JSZip();
    const doc = makeDoc({ endnotes: [NORMAL_ENDNOTE], endnoteSeparators: [SEPARATOR_ENDNOTE] });
    await serializeEndnotesToZip(doc, zip, 6);

    const xml = await entryText(zip, 'word/endnotes.xml');
    expect(xml).not.toBeNull();
    expect(xml!).toContain('Edited endnote body.');
    expect(xml!).toContain('<w:separator/>');
    // Separator note (id=-1) must precede the normal note (id=1)
    expect(xml!.indexOf('w:id="-1"')).toBeLessThan(xml!.indexOf('w:id="1"'));
  });

  test('does not write the part when the document has no endnotes', async () => {
    const zip = new JSZip();
    await serializeEndnotesToZip(makeDoc(), zip, 6);
    expect(zip.file('word/endnotes.xml')).toBeNull();
  });

  test('registers the endnotes content-type override + document relationship', async () => {
    const zip = new JSZip();
    seedPackage(zip);
    const doc = makeDoc({ endnotes: [NORMAL_ENDNOTE], endnoteSeparators: [SEPARATOR_ENDNOTE] });
    await serializeEndnotesToZip(doc, zip, 6);

    const ct = await entryText(zip, '[Content_Types].xml');
    expect(ct!).toContain('PartName="/word/endnotes.xml"');
    expect(ct!).toContain('wordprocessingml.endnotes+xml');
    const rels = await entryText(zip, 'word/_rels/document.xml.rels');
    expect(rels!).toContain('Target="endnotes.xml"');
    expect(rels!).toContain('/relationships/endnotes');
  });
});

describe('serializeFootnotesToZip', () => {
  test('writes word/footnotes.xml with separators ahead of normal notes', async () => {
    const zip = new JSZip();
    const doc = makeDoc({ footnotes: [NORMAL_FOOTNOTE], footnoteSeparators: [SEPARATOR_FOOTNOTE] });
    await serializeFootnotesToZip(doc, zip, 6);

    const xml = await entryText(zip, 'word/footnotes.xml');
    expect(xml).not.toBeNull();
    expect(xml!).toContain('A footnote.');
    expect(xml!).toContain('<w:separator/>');
    expect(xml!.indexOf('w:id="-1"')).toBeLessThan(xml!.indexOf('w:id="1"'));
  });

  test('does not write the part when the document has no footnotes', async () => {
    const zip = new JSZip();
    await serializeFootnotesToZip(makeDoc(), zip, 6);
    expect(zip.file('word/footnotes.xml')).toBeNull();
  });

  test('registers footnotes part once — idempotent on a doc that already declares it', async () => {
    const zip = new JSZip();
    seedPackage(zip);
    const doc = makeDoc({ footnotes: [NORMAL_FOOTNOTE], footnoteSeparators: [SEPARATOR_FOOTNOTE] });

    await serializeFootnotesToZip(doc, zip, 6); // first author
    await serializeFootnotesToZip(doc, zip, 6); // re-save (round-trip)

    const ct = await entryText(zip, '[Content_Types].xml');
    expect(ct!).toContain('PartName="/word/footnotes.xml"');
    // Override appears exactly once (no duplicate on re-save).
    expect(ct!.split('PartName="/word/footnotes.xml"').length - 1).toBe(1);
    const rels = await entryText(zip, 'word/_rels/document.xml.rels');
    expect(rels!).toContain('Target="footnotes.xml"');
    expect(rels!.split('Target="footnotes.xml"').length - 1).toBe(1);
  });
});

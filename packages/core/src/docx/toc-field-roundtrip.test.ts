import { describe, test, expect } from 'bun:test';
import JSZip from 'jszip';
import { parseDocx } from './parser';
import { repackDocx } from './rezip';

/**
 * Table of Contents (and any multi-paragraph complex field) round-trip.
 *
 * A real Word TOC is a complex field whose `w:fldChar begin`/`instrText`/
 * `separate` live in the FIRST entry paragraph and whose `w:fldChar end` lives
 * in a LATER paragraph. The per-paragraph field state machine in
 * `paragraphParser/content.ts` only finalizes a field when it sees the matching
 * `end` in the same paragraph, so a cross-paragraph field used to be swallowed
 * (begin + instruction silently dropped). The fix flushes an unterminated
 * field's consumed runs back into position. This guards that behavior — and the
 * ordering of the begin runs vs. the first entry's hyperlink in the begin
 * paragraph.
 *
 * Fixture provenance: synthesized here (no binary).
 */

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships';

const CONTENT_TYPES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Types xmlns="${NS_CT}">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
  `</Types>`;

const PACKAGE_RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="${NS_PR}">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`;

const DOCUMENT_RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="${NS_PR}">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
  `</Relationships>`;

const STYLES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:styles xmlns:w="${NS_W}"><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style></w:styles>`;

// begin/instrText/separate + the first entry's hyperlink all share paragraph 1;
// the matching `end` is in paragraph 3.
const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>` +
  `<w:p>` +
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> TOC \\h \\o "1-3" </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:hyperlink w:anchor="_Toc1"><w:r><w:t xml:space="preserve">Heading One</w:t></w:r></w:hyperlink>` +
  `</w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">Heading One body</w:t></w:r></w:p>` +
  `<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>` +
  `</w:body></w:document>`;

async function buildTocDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES_XML);
  zip.file('_rels/.rels', PACKAGE_RELS_XML);
  zip.file('word/document.xml', DOCUMENT_XML);
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS_XML);
  zip.file('word/styles.xml', STYLES_XML);
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }));
}

async function bodyXmlOf(buf: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(buf);
  return zip.file('word/document.xml')!.async('text');
}

describe('TOC / cross-paragraph complex field round-trip', () => {
  test('the TOC field begin + instruction + end survive repack', async () => {
    const doc = await parseDocx(await buildTocDocx(), { preloadFonts: false });
    const out = new Uint8Array(await repackDocx(doc));
    const xml = await bodyXmlOf(out);

    expect((xml.match(/w:fldCharType="begin"/g) ?? []).length).toBe(1);
    expect((xml.match(/w:fldCharType="end"/g) ?? []).length).toBe(1);
    // The instruction text round-trips (this is what regressed: it used to vanish).
    expect(xml).toMatch(/<w:instrText[^>]*>\s*TOC/);
    // The entry content is preserved.
    expect(xml).toContain('Heading One');
  });

  test('begin runs keep their order before the first entry hyperlink', async () => {
    const doc = await parseDocx(await buildTocDocx(), { preloadFonts: false });
    const out = new Uint8Array(await repackDocx(doc));
    const xml = await bodyXmlOf(out);

    // The field begin must come BEFORE the first entry's hyperlink, not after it
    // (the flush splices runs back at the field's start index).
    expect(xml.indexOf('w:fldCharType="begin"')).toBeGreaterThan(-1);
    expect(xml.indexOf('w:fldCharType="begin"')).toBeLessThan(xml.indexOf('<w:hyperlink'));
  });
});

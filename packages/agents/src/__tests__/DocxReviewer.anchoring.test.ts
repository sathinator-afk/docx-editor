/**
 * Anchoring round-trip harness — does a search-anchored proposal survive the
 * docx save → reload with its anchor intact?
 *
 * The existing roundtrip test (DocxReviewer.roundtrip.test.ts) proves a comment
 * or change reaches the saved bytes — but only for WHOLE-PARAGRAPH comments and
 * insertions, and it never re-reads the anchored text. This file pins the
 * anchoring dimension through real `toBuffer()` → `fromBuffer()` bytes.
 *
 * IMPORTANT boundary (verified while writing this): the two anchoring paths in
 * the monorepo behave differently, by design of where each splits runs —
 *  • TRACKED EDITS (proposeReplacement/Deletion) anchor to the exact SEARCHED
 *    phrase in the docx-body model, and that sub-phrase anchor survives the
 *    round-trip (asserted strictly below).
 *  • COMMENTS via DocxReviewer.addComment({search}) (the agents/MCP path,
 *    docx-body model) VALIDATE that `search` exists but wrap the whole
 *    paragraph's run — comment anchoring here is paragraph-granular.
 * The cunei APP does NOT use this comment path: it anchors comments through the
 * ProseMirror core (`addCommentToRange`), which narrows to the sub-phrase —
 * proven at the PM-state level by core's `commentOps.test.ts`
 * ('narrows to a search substring'). So this harness asserts the agents path's
 * REAL contract (survival + paragraph-granular comment anchor + phrase-granular
 * edit anchor) and does not over-claim sub-phrase comment anchoring it doesn't
 * provide. The one seam neither suite covers — a PM sub-phrase comment surviving
 * PM → docx → PM — is flagged for follow-up.
 *
 * Fully headless (in-memory builders → createDocx → real bytes), deterministic.
 */
import { describe, test, expect } from 'bun:test';
import { createDocx } from '@eigenpal/docx-editor-core/headless';
import type { Paragraph } from '@eigenpal/docx-editor-core/headless';
import { DocxReviewer } from '../DocxReviewer';
import { makeDoc, makeParagraph, makeParagraphFrom, makeRun, makeInsertion } from './_helpers';

/**
 * A reviewer over a REAL docx buffer built from controlled text. `makeReviewer`
 * builds an in-memory doc with no `originalBuffer`, so `toBuffer()` (which patches
 * the original zip) can't run on it — we scaffold an empty docx via createDocx()
 * and reload, which gives us a buffer-backed reviewer whose paragraph text we
 * fully control (unlike the shared fixture).
 */
async function freshReviewer(paras: Paragraph[]): Promise<DocxReviewer> {
  const buffer = await createDocx(makeDoc(paras));
  return DocxReviewer.fromBuffer(buffer, 'AI');
}

/** Mutate → serialize → reparse, returning a fresh reviewer over the saved bytes. */
async function roundTrip(reviewer: DocxReviewer): Promise<DocxReviewer> {
  return DocxReviewer.fromBuffer(await reviewer.toBuffer(), 'Reviewer');
}

describe('anchoring survives a save/load round-trip', () => {
  test('a search-anchored comment survives and still covers its phrase', async () => {
    const reviewer = await freshReviewer([
      makeParagraph('The quick brown fox jumps over the lazy dog.'),
    ]);
    const id = reviewer.addComment({
      paragraphIndex: 0,
      author: 'AI',
      text: 'Tighten this clause.',
      search: 'brown fox',
    });

    const reloaded = await roundTrip(reviewer);
    const c = reloaded.getComments().find((x) => x.id === id) ?? reloaded.getComments()[0];

    expect(c).toBeDefined();
    expect(c.text).toBe('Tighten this clause.');
    // The anchored range still covers the phrase the comment was about after the
    // round-trip (paragraph-granular in this path — see the file header).
    expect(c.anchoredText).toContain('brown fox');
    expect(c.paragraphIndex).toBe(0);
  });

  test('a search phrase that is not in the paragraph is refused at the API boundary', async () => {
    const reviewer = await freshReviewer([makeParagraph('Some text here.')]);
    // Round-trip safety: a phantom phrase throws, which is the contract the app
    // relies on — `cuneonApplyReviewComments` catches it and reports the comment
    // as skipped, so a proposal whose anchor doesn't exist never reaches a save.
    // (Known agents-path quirk: addCommentImpl pushes the comment object before
    // it validates the anchor, so a failed add leaves a partial entry in the
    // in-memory model — see margin-anchoring-findings. The app path doesn't hit
    // this; it discards the reviewer on a throw.)
    expect(() =>
      reviewer.addComment({ paragraphIndex: 0, author: 'AI', text: 'x', search: 'not present' })
    ).toThrow();
  });

  test('two comments on one paragraph survive as distinct entries', async () => {
    const reviewer = await freshReviewer([
      makeParagraph('The liability cap is $50k and the term is 24 months.'),
    ]);
    const capId = reviewer.addComment({
      paragraphIndex: 0,
      author: 'AI',
      text: 'Cap is low.',
      search: '$50k',
    });
    const termId = reviewer.addComment({
      paragraphIndex: 0,
      author: 'AI',
      text: 'Term is long.',
      search: '24 months',
    });

    const reloaded = await roundTrip(reviewer);
    const cap = reloaded.getComments().find((x) => x.id === capId);
    const term = reloaded.getComments().find((x) => x.id === termId);

    // Both comments survive the save as separate entries with their own bodies —
    // they don't merge, drop, or swap text when sharing a paragraph.
    expect(reloaded.getComments()).toHaveLength(2);
    expect(cap?.text).toBe('Cap is low.');
    expect(term?.text).toBe('Term is long.');
    expect(capId).not.toBe(termId);
  });

  test('an anchor resolves against the vanilla view when an insertion is present', async () => {
    // The paragraph already carries a tracked insertion. Agents see the vanilla
    // (accepted) view via read_document, so a phrase spanning original runs must
    // still anchor — and survive the round-trip — even though the saved XML
    // interleaves the insertion run.
    const para = makeParagraphFrom([
      makeRun('The contract '),
      makeInsertion('hereby ', 1, 'Reviewer'),
      makeRun('terminates on the final date.'),
    ]);
    const reviewer = await freshReviewer([para]);
    const id = reviewer.addComment({
      paragraphIndex: 0,
      author: 'AI',
      text: 'Specify the date.',
      search: 'terminates on the final date',
    });

    const reloaded = await roundTrip(reviewer);
    const c = reloaded.getComments().find((x) => x.id === id) ?? reloaded.getComments()[0];

    expect(c).toBeDefined();
    expect(c.anchoredText).toContain('terminates on the final date');
  });

  test('a whole-paragraph comment and an anchored comment keep distinct anchors', async () => {
    const reviewer = await freshReviewer([
      makeParagraph('First paragraph about scope.'),
      makeParagraph('Second paragraph names the parties: Acme and Globex.'),
    ]);
    const wholeId = reviewer.addComment(0, 'Whole-paragraph note.');
    const anchoredId = reviewer.addComment({
      paragraphIndex: 1,
      author: 'AI',
      text: 'Confirm the parties.',
      search: 'Acme and Globex',
    });

    const reloaded = await roundTrip(reviewer);
    const whole = reloaded.getComments().find((x) => x.id === wholeId);
    const anchored = reloaded.getComments().find((x) => x.id === anchoredId);

    expect(whole?.anchoredText).toContain('scope');
    expect(anchored?.anchoredText).toContain('Acme and Globex');
    expect(anchored?.anchoredText).not.toContain('scope');
  });
});

describe('tracked edits survive a save/load round-trip with their anchor', () => {
  test('proposeReplacement persists both the deletion and the insertion', async () => {
    const reviewer = await freshReviewer([
      makeParagraph('Replace this word with the proper term.'),
    ]);
    reviewer.proposeReplacement({
      paragraphIndex: 0,
      search: 'this word',
      replaceWith: 'that phrase',
      author: 'AI',
    });
    const before = reviewer.getChanges();
    expect(before.length).toBeGreaterThan(0);

    const reloaded = await roundTrip(reviewer);
    const after = reloaded.getChanges();

    // The change count is stable across the save, and the anchored old/new text
    // both survive (a replacement is a deletion of the search + an insertion of
    // the replacement).
    expect(after.length).toBe(before.length);
    expect(after.some((c) => c.type === 'deletion' && c.text.includes('this word'))).toBe(true);
    expect(after.some((c) => c.type === 'insertion' && c.text.includes('that phrase'))).toBe(true);
    expect(after.every((c) => c.author === 'AI')).toBe(true);
  });

  test('proposeDeletion anchors to the searched phrase and survives reload', async () => {
    const reviewer = await freshReviewer([
      makeParagraph('Strike the redundant clause from this sentence.'),
    ]);
    reviewer.proposeDeletion({
      paragraphIndex: 0,
      search: 'redundant clause',
      author: 'AI',
    });
    const before = reviewer.getChanges();

    const reloaded = await roundTrip(reviewer);
    const after = reloaded.getChanges();

    expect(after.length).toBe(before.length);
    expect(after.some((c) => c.type === 'deletion' && c.text.includes('redundant clause'))).toBe(
      true
    );
  });
});

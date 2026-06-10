/**
 * Section Break Commands
 * @packageDocumentation
 * @public
 */

import type { Command } from 'prosemirror-state';
import { TextSelection } from 'prosemirror-state';

type InsertableSectionBreak = 'nextPage' | 'continuous';

/**
 * Insert a section break at the current cursor position.
 *
 * In OOXML a section break is the `sectPr` carried by the *last* paragraph of a
 * section. So we split the current paragraph at the cursor, mark the first half
 * with `sectionBreakType` (it becomes the section end), and leave the cursor in
 * the second half (the first paragraph of the new section). Content after the
 * cursor therefore flows into the new section — onto a new page for `nextPage`,
 * or in place for `continuous`.
 */
function insertSectionBreakAtCursor(breakType: InsertableSectionBreak): Command {
  return (state, dispatch) => {
    const { schema } = state;
    const paragraphType = schema.nodes.paragraph;
    if (!paragraphType) return false;

    if (dispatch) {
      const { $from } = state.selection;
      const tr = state.tr;
      let cursorPos: number;

      if ($from.parent.isTextblock) {
        // Position of the paragraph node the cursor sits in. Unaffected by the
        // split below, since the split happens at a later position.
        const paraPos = $from.before();
        tr.split($from.pos);
        const firstPara = tr.doc.nodeAt(paraPos);
        if (firstPara) {
          tr.setNodeMarkup(paraPos, undefined, {
            ...firstPara.attrs,
            sectionBreakType: breakType,
          });
        }
        // The split position maps to the start of the second paragraph's
        // content (the first paragraph of the new section) — put the cursor there.
        cursorPos = tr.mapping.map($from.pos);
      } else {
        // Not in a textblock — insert a section-ending empty paragraph here.
        const pos = $from.pos;
        tr.insert(pos, paragraphType.create({ sectionBreakType: breakType }));
        cursorPos = tr.mapping.map($from.pos);
      }

      tr.setSelection(TextSelection.create(tr.doc, cursorPos));
      dispatch(tr.scrollIntoView());
    }

    return true;
  };
}

/**
 * Insert a "next page" section break at the cursor — starts a new section on a
 * new page.
 */
export const insertSectionBreakNextPage: Command = insertSectionBreakAtCursor('nextPage');

/**
 * Insert a "continuous" section break at the cursor — starts a new section on
 * the same page.
 */
export const insertSectionBreakContinuous: Command = insertSectionBreakAtCursor('continuous');

import { describe, test, expect } from 'bun:test';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';
import { schema } from '../schema';
import { insertSectionBreakNextPage, insertSectionBreakContinuous } from './sectionBreak';

/**
 * Build an editor with a single paragraph of `text` and the cursor placed
 * `cursorOffset` characters into it (paragraph content starts at doc pos 1).
 */
function setup(text: string, cursorOffset: number) {
  const doc = schema.node('doc', { defaultTabStopTwips: null, watermark: null }, [
    schema.node('paragraph', {}, text ? [schema.text(text)] : []),
  ]);
  let state = EditorState.create({ doc });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1 + cursorOffset)));
  const dispatch = (tr: Transaction) => {
    state = state.apply(tr);
  };
  return {
    dispatch,
    get state() {
      return state;
    },
  };
}

describe('insertSectionBreak commands', () => {
  test('mid-text split: first paragraph ends the section, cursor lands in the second', () => {
    const ed = setup('HelloWorld', 5);
    insertSectionBreakNextPage(ed.state, ed.dispatch);

    const { doc, selection } = ed.state;
    expect(doc.childCount).toBe(2);
    expect(doc.child(0).attrs.sectionBreakType).toBe('nextPage');
    expect(doc.child(0).textContent).toBe('Hello');
    // The new section's first paragraph does not itself carry a break.
    expect(doc.child(1).attrs.sectionBreakType).toBeNull();
    expect(doc.child(1).textContent).toBe('World');
    // Cursor is inside the second paragraph (the start of the new section).
    expect(selection.$from.parent.textContent).toBe('World');
  });

  test('continuous variant marks the section end as continuous', () => {
    const ed = setup('HelloWorld', 5);
    insertSectionBreakContinuous(ed.state, ed.dispatch);

    const { doc } = ed.state;
    expect(doc.child(0).attrs.sectionBreakType).toBe('continuous');
    expect(doc.child(1).attrs.sectionBreakType).toBeNull();
  });

  test('end-of-paragraph: current paragraph ends the section, cursor in a new empty paragraph', () => {
    const ed = setup('Hello', 5);
    insertSectionBreakNextPage(ed.state, ed.dispatch);

    const { doc, selection } = ed.state;
    expect(doc.childCount).toBe(2);
    expect(doc.child(0).attrs.sectionBreakType).toBe('nextPage');
    expect(doc.child(0).textContent).toBe('Hello');
    expect(doc.child(1).textContent).toBe('');
    expect(selection.$from.parent.textContent).toBe('');
  });

  test('start-of-paragraph: an empty section-ending paragraph precedes the content', () => {
    const ed = setup('Hello', 0);
    insertSectionBreakNextPage(ed.state, ed.dispatch);

    const { doc, selection } = ed.state;
    expect(doc.childCount).toBe(2);
    expect(doc.child(0).attrs.sectionBreakType).toBe('nextPage');
    expect(doc.child(0).textContent).toBe('');
    expect(doc.child(1).textContent).toBe('Hello');
    expect(selection.$from.parent.textContent).toBe('Hello');
  });
});

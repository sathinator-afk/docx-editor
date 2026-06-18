/**
 * ProseMirror Plugins
 *
 * Selection tracker plugin for the DOCX editor.
 * Keymap plugins are now provided by the extension system.
 * @packageDocumentation
 * @public
 */

export {
  createSelectionTrackerPlugin,
  extractSelectionContext,
  getSelectionContext,
  selectionTrackerKey,
} from './selectionTracker';

export type { SelectionContext, SelectionChangeCallback } from './selectionTracker';

export {
  suggestionModeKey,
  createSuggestionModePlugin,
  toggleSuggestionMode,
  setSuggestionMode,
  isSuggestionModeActive,
  applyTrackedReplaceKeepingMarks,
  applyTrackedParagraphInsert,
} from './suggestionMode';
export {
  documentStylesKey,
  createDocumentStylesPlugin,
  getDocumentStyleResolver,
} from './documentStyles';
export { makeRevisionInfo } from './revisionIds';
export type { RevisionInfo } from '../../types/content/trackedChange';

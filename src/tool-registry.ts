import type { ToolDefinition } from './server.js';
import type { ToolName } from './config.js';

export async function getAllTools(): Promise<Map<ToolName, ToolDefinition>> {
  const tools = new Map<ToolName, ToolDefinition>();

  // word_get_document_info
  const infoMod = await import('./tools/word-get-document-info.js');
  tools.set('word_get_document_info', {
    name: 'word_get_document_info',
    description: 'Returns lightweight metadata for a DOCX file: page count, word count, comment counts, track changes status, and section count.',
    inputSchema: infoMod.WORD_GET_DOCUMENT_INFO_SCHEMA,
    handler: infoMod.wordGetDocumentInfo,
  });

  // word_get_content
  const contentMod = await import('./tools/word-get-content.js');
  tools.set('word_get_content', {
    name: 'word_get_content',
    description: 'Reads a DOCX file and returns structured text with paragraph structure and comments. Supports maxChars truncation.',
    inputSchema: contentMod.WORD_GET_CONTENT_SCHEMA,
    handler: contentMod.wordGetContent,
  });

  // word_get_sections
  const sectionsMod = await import('./tools/word-get-sections.js');
  tools.set('word_get_sections', {
    name: 'word_get_sections',
    description: 'Returns document structure as a list of sections (headings) with word count per section.',
    inputSchema: sectionsMod.WORD_GET_SECTIONS_SCHEMA,
    handler: sectionsMod.wordGetSections,
  });

  // word_get_section_content
  const sectionContentMod = await import('./tools/word-get-section-content.js');
  tools.set('word_get_section_content', {
    name: 'word_get_section_content',
    description: 'Returns full text content of a single section, limited by maxChars. Allows reading specific parts of large documents.',
    inputSchema: sectionContentMod.WORD_GET_SECTION_CONTENT_SCHEMA,
    handler: sectionContentMod.wordGetSectionContent,
  });

  // word_find_text
  const findMod = await import('./tools/word-find-text.js');
  tools.set('word_find_text', {
    name: 'word_find_text',
    description: 'Searches for exact text within a DOCX file, returning each match with surrounding context. Handles OOXML run fragmentation.',
    inputSchema: findMod.WORD_FIND_TEXT_SCHEMA,
    handler: findMod.wordFindText,
  });

  // word_search_replace
  const replaceMod = await import('./tools/word-search-replace.js');
  tools.set('word_search_replace', {
    name: 'word_search_replace',
    description: 'Finds and replaces text in a DOCX file, preserving formatting. Supports single replacement, by index, or replace-all.',
    inputSchema: replaceMod.WORD_SEARCH_REPLACE_SCHEMA,
    handler: replaceMod.wordSearchReplace,
  });

  // word_list_comments
  const listCommentsMod = await import('./tools/word-list-comments.js');
  tools.set('word_list_comments', {
    name: 'word_list_comments',
    description: 'Lists all comments and reply threads, including the exact text each comment refers to. Groups replies under parents.',
    inputSchema: listCommentsMod.WORD_LIST_COMMENTS_SCHEMA,
    handler: listCommentsMod.wordListComments,
  });

  // word_create_comment
  const createMod = await import('./tools/word-create-comment.js');
  tools.set('word_create_comment', {
    name: 'word_create_comment',
    description: 'Creates a new root comment anchored to a specific text fragment. Initializes comment threading if needed.',
    inputSchema: createMod.WORD_CREATE_COMMENT_SCHEMA,
    handler: createMod.wordCreateComment,
  });

  // word_reply_to_comment
  const replyMod = await import('./tools/word-reply-to-comment.js');
  tools.set('word_reply_to_comment', {
    name: 'word_reply_to_comment',
    description: 'Adds a reply to an existing comment, maintaining thread hierarchy via commentsExtended.xml.',
    inputSchema: replyMod.WORD_REPLY_TO_COMMENT_SCHEMA,
    handler: replyMod.wordReplyToComment,
  });

  // word_resolve_comment
  const resolveMod = await import('./tools/word-resolve-comment.js');
  tools.set('word_resolve_comment', {
    name: 'word_resolve_comment',
    description: 'Marks a comment as resolved in Word. Sets w15:done="1" in commentsExtended.xml.',
    inputSchema: resolveMod.WORD_RESOLVE_COMMENT_SCHEMA,
    handler: resolveMod.wordResolveComment,
  });

  // word_get_page_setup
  const pageSetupMod = await import('./tools/word-get-page-setup.js');
  tools.set('word_get_page_setup', {
    name: 'word_get_page_setup',
    description: 'Returns page layout: page size, orientation, margins, and per-section layout info.',
    inputSchema: pageSetupMod.WORD_GET_PAGE_SETUP_SCHEMA,
    handler: pageSetupMod.wordGetPageSetup,
  });

  // word_set_page_size
  const pageSizeMod = await import('./tools/word-set-page-size.js');
  tools.set('word_set_page_size', {
    name: 'word_set_page_size',
    description: 'Sets page size. Supports presets (letter, legal, folio, a4, executive) or custom dimensions.',
    inputSchema: pageSizeMod.WORD_SET_PAGE_SIZE_SCHEMA,
    handler: pageSizeMod.wordSetPageSize,
  });

  // word_set_orientation
  const orientMod = await import('./tools/word-set-orientation.js');
  tools.set('word_set_orientation', {
    name: 'word_set_orientation',
    description: 'Sets page orientation (portrait/landscape). Supports mixed-orientation documents.',
    inputSchema: orientMod.WORD_SET_ORIENTATION_SCHEMA,
    handler: orientMod.wordSetOrientation,
  });

  // word_set_margins
  const marginsMod = await import('./tools/word-set-margins.js');
  tools.set('word_set_margins', {
    name: 'word_set_margins',
    description: 'Sets page margins. Supports twips, points, inches, and centimeters.',
    inputSchema: marginsMod.WORD_SET_MARGINS_SCHEMA,
    handler: marginsMod.wordSetMargins,
  });

  // word_compare_versions
  const compareMod = await import('./tools/word-compare-versions.js');
  tools.set('word_compare_versions', {
    name: 'word_compare_versions',
    description: 'Compares two DOCX files and returns inserted, deleted, and changed text organized by section.',
    inputSchema: compareMod.WORD_COMPARE_VERSIONS_SCHEMA,
    handler: compareMod.wordCompareVersions,
  });

  // word_get_styles
  const stylesMod = await import('./tools/word-get-styles.js');
  tools.set('word_get_styles', {
    name: 'word_get_styles',
    description: 'Returns available styles, usage counts, and detects structural issues (e.g., heading-like paragraphs using Normal style).',
    inputSchema: stylesMod.WORD_GET_STYLES_SCHEMA,
    handler: stylesMod.wordGetStyles,
  });

  // word_apply_style
  const applyMod = await import('./tools/word-apply-style.js');
  tools.set('word_apply_style', {
    name: 'word_apply_style',
    description: 'Applies a paragraph style to one or more paragraphs. Modifies w:pStyle and removes redundant manual formatting.',
    inputSchema: applyMod.WORD_APPLY_STYLE_SCHEMA,
    handler: applyMod.wordApplyStyle,
  });

  // word_get_footnotes
  const footnotesMod = await import('./tools/word-get-footnotes.js');
  tools.set('word_get_footnotes', {
    name: 'word_get_footnotes',
    description: 'Returns all footnotes in the document with their text content. Essential for legal and academic documents.',
    inputSchema: footnotesMod.WORD_GET_FOOTNOTES_SCHEMA,
    handler: footnotesMod.wordGetFootnotes,
  });

  // word_insert_tracked_change
  const trackedChangeMod = await import('./tools/word-insert-tracked-change.js');
  tools.set('word_insert_tracked_change', {
    name: 'word_insert_tracked_change',
    description: 'Inserts a tracked change (deletion + optional insertion) into a DOCX file. Produces w:del/w:ins pairs visible in Word\'s Track Changes pane.',
    inputSchema: trackedChangeMod.WORD_INSERT_TRACKED_CHANGE_SCHEMA,
    handler: trackedChangeMod.wordInsertTrackedChange,
  });

  return tools;
}

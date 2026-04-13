import { validateDocxPath } from '../validation.js';

export const WORD_GET_SECTION_CONTENT_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    section: { type: 'string', description: 'Section title as returned by word_get_sections' },
    sectionIndex: { type: 'number', description: '1-based section index (alternative to title)' },
    maxChars: { type: 'number', description: 'Maximum characters. Default: 10000' },
  },
  required: ['filePath'],
} as const;

export async function wordGetSectionContent(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, section, sectionIndex, maxChars = 10000 } = args as { filePath: string; section?: string; sectionIndex?: number; maxChars?: number };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const { DocxDocument } = await import('../docx.js');
  const doc = await DocxDocument.load(filePath);
  const sections = await doc.getSections();

  let targetSection = section;
  if (!targetSection && sectionIndex) {
    const found = sections.find((s: { sectionIndex: number }) => s.sectionIndex === sectionIndex);
    if (!found) return { content: [{ type: 'text', text: `Section index ${sectionIndex} not found` }], isError: true };
    targetSection = found.title;
  }
  if (!targetSection) return { content: [{ type: 'text', text: 'Either section or sectionIndex is required' }], isError: true };

  const content = await doc.getSectionContent(targetSection, maxChars);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        ...content,
        _suggestions: {
          word_find_text: { tool: 'word_find_text', description: 'Search within this section' },
          word_list_comments: { tool: 'word_list_comments', description: 'View comments for this section' },
          word_search_replace: { tool: 'word_search_replace', description: 'Edit text in this section' },
        },
      }, null, 2),
    }],
  };
}

import { validateDocxPath } from '../validation.js';

export const WORD_GET_SECTIONS_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
  },
  required: ['filePath'],
} as const;

export async function wordGetSections(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath } = args as { filePath: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const { DocxDocument } = await import('../docx.js');
  const doc = await DocxDocument.load(filePath);
  const sections = await doc.getSections();

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        sections,
        _suggestions: {
          word_get_section_content: { tool: 'word_get_section_content', description: 'Read a specific section' },
          word_list_comments: { tool: 'word_list_comments', description: 'View comments by section' },
        },
      }, null, 2),
    }],
  };
}

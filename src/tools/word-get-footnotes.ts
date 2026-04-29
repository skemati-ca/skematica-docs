import { validateDocxPath } from '../validation.js';

export const WORD_GET_FOOTNOTES_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
  },
  required: ['filePath'],
} as const;

export async function wordGetFootnotes(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath } = args as { filePath: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const { DocxDocument } = await import('../docx.js');
  const doc = await DocxDocument.load(filePath);
  const footnotes = await doc.getFootnotes();

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        count: footnotes.length,
        footnotes,
        _suggestions: footnotes.length > 0
          ? {
              word_get_content: { tool: 'word_get_content', description: 'Read document body to see where footnotes are referenced' },
              word_find_text: { tool: 'word_find_text', description: 'Search body text to locate footnote anchors' },
            }
          : {
              word_get_content: { tool: 'word_get_content', description: 'Read document content' },
            },
      }, null, 2),
    }],
  };
}

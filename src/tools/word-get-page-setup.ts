import { validateDocxPath } from '../validation.js';

export const WORD_GET_PAGE_SETUP_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
  },
  required: ['filePath'],
} as const;

export async function wordGetPageSetup(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath } = args as { filePath: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const { DocxDocument } = await import('../docx.js');
  const doc = await DocxDocument.load(filePath);
  const pageInfo = await doc.getPageSetup();

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        ...pageInfo,
        margins: {
          ...pageInfo.margins,
          topIn: (pageInfo.margins.top / 1440).toFixed(2),
          bottomIn: (pageInfo.margins.bottom / 1440).toFixed(2),
          leftIn: (pageInfo.margins.left / 1440).toFixed(2),
          rightIn: (pageInfo.margins.right / 1440).toFixed(2),
        },
        _suggestions: {
          word_set_page_size: { tool: 'word_set_page_size', description: 'Change page size (letter, legal, A4, etc.)' },
          word_set_orientation: { tool: 'word_set_orientation', description: 'Change orientation (portrait/landscape)' },
          word_set_margins: { tool: 'word_set_margins', description: 'Adjust margins' },
        },
      }, null, 2),
    }],
  };
}

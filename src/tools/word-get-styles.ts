import { validateDocxPath } from '../validation.js';

export const WORD_GET_STYLES_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
  },
  required: ['filePath'],
} as const;

export async function wordGetStyles(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath } = args as { filePath: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const { DocxDocument } = await import('../docx.js');
  const doc = await DocxDocument.load(filePath);
  const stylesInfo = await doc.getStyles();

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        ...stylesInfo,
        _suggestions: {
          word_apply_style: { tool: 'word_apply_style', description: 'Fix structural issues by applying correct styles' },
          word_get_sections: { tool: 'word_get_sections', description: 'View current document structure' },
        },
      }, null, 2),
    }],
  };
}

import { validateDocxPath } from '../validation.js';

export const WORD_GET_DOCUMENT_INFO_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
  },
  required: ['filePath'],
} as const;

export async function wordGetDocumentInfo(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath } = args as { filePath: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const { DocxDocument } = await import('../docx.js');
  let doc;
  try {
    doc = await DocxDocument.load(filePath);
  } catch (e) {
    return { content: [{ type: 'text', text: (e as Error).message }], isError: true };
  }
  const meta = await doc.getMetadata();

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        ...meta,
        _suggestions: {
          word_get_sections: { tool: 'word_get_sections', description: 'View document structure' },
          word_list_comments: { tool: 'word_list_comments', description: `Review ${meta.unresolvedComments} unresolved comments` },
        },
      }, null, 2),
    }],
  };
}

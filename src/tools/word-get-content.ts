import { validateDocxPath } from '../validation.js';

export const WORD_GET_CONTENT_SCHEMA = {
  type: 'object',
  properties: {
    filePath: {
      type: 'string',
      description: 'Absolute path to the .docx file',
    },
    maxChars: {
      type: 'number',
      description: 'Maximum characters to return. Default: 50000. If exceeded, response is truncated with navigation guidance.',
    },
  },
  required: ['filePath'],
} as const;

export async function wordGetContent(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, maxChars = 50000 } = args as { filePath: string; maxChars?: number };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const { DocxDocument } = await import('../docx.js');
  const doc = await DocxDocument.load(filePath);
  const [text, paragraphs, comments, sections, listInfo] = await Promise.all([
    doc.getText(),
    doc.getParagraphs(),
    doc.getComments(),
    doc.getSections(),
    doc.getParagraphListInfo(),
  ]);

  const truncated = text.length > maxChars;
  const content = truncated ? text.substring(0, maxChars) : text;

  const suggestions: Record<string, Record<string, string>> = {};
  if (truncated) {
    suggestions.word_get_sections = { tool: 'word_get_sections', description: 'Navigate by sections instead of reading everything' };
  }
  suggestions.word_find_text = { tool: 'word_find_text', description: 'Search for specific text' };
  suggestions.word_list_comments = { tool: 'word_list_comments', description: 'Review pending comments' };

  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          text: content,
          paragraphs: paragraphs.map((p: string, i: number) => {
            const numPr = listInfo[i] ?? null;
            const entry: Record<string, unknown> = { index: i, text: p };
            if (numPr) {
              entry.isListItem = true;
              entry.listLevel = numPr.ilvl;
            }
            return entry;
          }),
          comments: (comments as unknown as Record<string, unknown>[]).map(flattenComment),
          truncated,
          totalChars: text.length,
          returnedChars: content.length,
          sectionCount: sections.length,
          _suggestions: suggestions,
        }, null, 2),
      },
    ],
  };
}

function flattenComment(c: Record<string, unknown>): Record<string, unknown> {
  const replies = (c.replies as Record<string, unknown>[] || []);
  return {
    id: c.id,
    author: c.author,
    text: c.text,
    date: c.date,
    parentId: c.parentId,
    isResolved: c.isResolved,
    commentedText: c.commentedText,
    replies: replies.map((r) => flattenComment(r)),
  };
}

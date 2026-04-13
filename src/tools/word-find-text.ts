import { validateDocxPath } from '../validation.js';

export const WORD_FIND_TEXT_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    searchText: { type: 'string', description: 'Text to search for (exact match)' },
    contextChars: { type: 'number', description: 'Characters before/after each match. Default: 100' },
    section: { type: 'string', description: 'Limit search to a specific section' },
  },
  required: ['filePath', 'searchText'],
} as const;

export async function wordFindText(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, searchText, contextChars = 100, section } = args as { filePath: string; searchText: string; contextChars?: number; section?: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const { DocxDocument } = await import('../docx.js');
  const doc = await DocxDocument.load(filePath);

  let text: string;
  if (section) {
    const sectionContent = await doc.getSectionContent(section);
    text = sectionContent.content;
  } else {
    text = await doc.getText();
  }

  const matches: Array<{ matchIndex: number; matchText: string; contextBefore: string; contextAfter: string }> = [];
  let pos = 0;
  let index = 0;

  while (true) {
    const found = text.indexOf(searchText, pos);
    if (found === -1) break;

    const start = Math.max(0, found - contextChars);
    const end = Math.min(text.length, found + searchText.length + contextChars);

    matches.push({
      matchIndex: index,
      matchText: searchText,
      contextBefore: text.substring(start, found),
      contextAfter: text.substring(found + searchText.length, end),
    });

    pos = found + searchText.length;
    index++;
  }

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        totalMatches: matches.length,
        matches,
        _suggestions: {
          word_search_replace: { tool: 'word_search_replace', description: matches.length > 1 ? 'Replace all or specific matches' : 'Replace this match' },
          word_find_text: { tool: 'word_find_text', description: 'Search for different text' },
        },
      }, null, 2),
    }],
  };
}

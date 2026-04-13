import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { validateDocxPath } from '../validation.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const WORD_RESOLVE_COMMENT_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    commentId: { type: 'number', description: 'ID of the comment to resolve' },
  },
  required: ['filePath', 'commentId'],
} as const;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', trimValues: true });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', format: true, indentBy: '  ', suppressEmptyNode: false });

export async function wordResolveComment(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, commentId } = args as { filePath: string; commentId: number };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const content = readFileSync(filePath);
  const zip = await JSZip.loadAsync(content);

  const extFile = zip.file('word/commentsExtended.xml');
  if (!extFile) return { content: [{ type: 'text', text: `Comment ${commentId} not found or no threading info.` }], isError: true };

  const commentsExtended = parser.parse(await extFile.async('text')) as Record<string, unknown>;
  const commentsEx = commentsExtended['w15:commentsEx'] as Record<string, unknown> | undefined;
  if (!commentsEx) return { content: [{ type: 'text', text: `Comment ${commentId} not found.` }], isError: true };

  const commentExList = ensureArray(commentsEx['w15:commentEx']);
  let found = false;

  for (const ex of commentExList) {
    const en = ex as Record<string, unknown>;
    if (en['@_w15:paraId']) {
      // Check if this is the comment or a reply to it
      const parentId = en['@_w15:parentId'];
      if (parentId === String(commentId)) {
        en['@_w15:done'] = '1';
        found = true;
      }
    }
  }

  if (!found) return { content: [{ type: 'text', text: `Comment ${commentId} not found in threading info.` }], isError: true };

  zip.file('word/commentsExtended.xml', builder.build(commentsExtended));
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const { writeFileSync } = await import('node:fs');
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        resolvedId: String(commentId),
        _suggestions: {
          word_list_comments: { tool: 'word_list_comments', description: 'View remaining unresolved comments' },
        },
      }, null, 2),
    }],
  };
}

function ensureArray(val: unknown): unknown[] {
  if (!val) return [];
  if (Array.isArray(val)) return val;
  return [val];
}

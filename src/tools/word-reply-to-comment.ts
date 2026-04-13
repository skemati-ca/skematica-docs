import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { validateDocxPath } from '../validation.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const WORD_REPLY_TO_COMMENT_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    commentId: { type: 'number', description: 'ID of the parent comment' },
    replyText: { type: 'string', description: 'Reply text' },
    author: { type: 'string', description: 'Reply author. Default: "Asistente IA"' },
  },
  required: ['filePath', 'commentId', 'replyText'],
} as const;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', trimValues: true });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', format: true, indentBy: '  ', suppressEmptyNode: false });

export async function wordReplyToComment(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, commentId, replyText, author = 'Asistente IA' } = args as { filePath: string; commentId: number; replyText: string; author?: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const content = readFileSync(filePath);
  const zip = await JSZip.loadAsync(content);

  // Load comments.xml
  const commentsFile = zip.file('word/comments.xml');
  if (!commentsFile) return { content: [{ type: 'text', text: 'No comments in this document.' }], isError: true };

  const commentsXml = parser.parse(await commentsFile.async('text')) as Record<string, unknown>;
  const comments = ensureArray(commentsXml?.['w:comments']?.['w:comment']);

  // Verify parent comment exists
  const parentExists = comments.some((c) => (c as Record<string, unknown>)['@_w:id'] === String(commentId));
  if (!parentExists) return { content: [{ type: 'text', text: `Comment ${commentId} not found.` }], isError: true };

  // Generate new comment ID
  let maxId = 0;
  for (const c of comments) {
    const id = Number((c as Record<string, unknown>)['@_w:id']);
    if (id > maxId) maxId = id;
  }
  const replyId = String(maxId + 1);

  // Add reply to comments.xml
  const newComment = {
    '@_w:id': replyId,
    '@_w:author': author,
    '@_w:initials': author.charAt(0),
    '@_w:date': new Date().toISOString(),
    'w:p': [{ 'w:r': [{ 'w:t': { '@_xml:space': 'preserve', '#text': replyText } }] }],
  };
  comments.push(newComment);

  // Load or create commentsExtended.xml
  const extFile = zip.file('word/commentsExtended.xml');
  let commentsExtended: Record<string, unknown>;

  if (extFile) {
    commentsExtended = parser.parse(await extFile.async('text')) as Record<string, unknown>;
  } else {
    commentsExtended = { 'w15:commentsEx': { '@_xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml' } };

    // Add relationship
    const relsFile = zip.file('word/_rels/document.xml.rels');
    if (relsFile) {
      const relsStr = await relsFile.async('text');
      const rels = parser.parse(relsStr) as Record<string, unknown>;
      const relsObj = rels['Relationships'] as Record<string, unknown> | undefined;
      if (relsObj) {
        const existing = ensureArray(relsObj['Relationship']);
        existing.push({
          '@_Id': 'rId_commentsExtended',
          '@_Type': 'http://schemas.microsoft.com/office/word/2012/wordml/commentsEx',
          '@_Target': 'commentsExtended.xml',
        });
        relsObj['Relationship'] = existing;
      }
      zip.file('word/_rels/document.xml.rels', builder.build(rels));
    }
  }

  // Add reply entry to commentsExtended.xml
  const paraId = generateParaId();
  const commentEx = {
    '@_w15:paraId': paraId,
    '@_w15:parentId': String(commentId),
    '@_w15:done': '0',
  };

  const commentsEx = commentsExtended['w15:commentsEx'] as Record<string, unknown>;
  if (!commentsEx) return { content: [{ type: 'text', text: 'Failed to initialize commentsExtended.xml' }], isError: true };

  const existingEx = ensureArray(commentsEx['w15:commentEx']);
  existingEx.push(commentEx);
  commentsEx['w15:commentEx'] = existingEx;

  // Save
  zip.file('word/comments.xml', builder.build(commentsXml));
  zip.file('word/commentsExtended.xml', builder.build(commentsExtended));

  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const { writeFileSync } = await import('node:fs');
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        id: replyId,
        parentId: String(commentId),
        author,
        text: replyText,
        _suggestions: {
          word_list_comments: { tool: 'word_list_comments', description: 'View full comment thread' },
          word_resolve_comment: { tool: 'word_resolve_comment', description: 'Mark this thread as resolved' },
        },
      }, null, 2),
    }],
  };
}

function generateParaId(): string {
  return Math.random().toString(16).substring(2, 10).toUpperCase();
}

function ensureArray(val: unknown): unknown[] {
  if (!val) return [];
  if (Array.isArray(val)) return val;
  return [val];
}

import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { validateDocxPath } from '../validation.js';
import { findAllTextInNode } from '../xml-utils.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const WORD_CREATE_COMMENT_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    searchText: { type: 'string', description: 'Text to anchor the comment to' },
    commentText: { type: 'string', description: 'Comment text' },
    author: { type: 'string', description: 'Comment author. Default: "Asistente IA"' },
  },
  required: ['filePath', 'searchText', 'commentText'],
} as const;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', trimValues: true });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', format: true, indentBy: '  ', suppressEmptyNode: false });

export async function wordCreateComment(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, searchText, commentText, author = 'Asistente IA' } = args as { filePath: string; searchText: string; commentText: string; author?: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const content = readFileSync(filePath);
  const zip = await JSZip.loadAsync(content);

  // Load document XML
  const docFile = zip.file('word/document.xml');
  if (!docFile) return { content: [{ type: 'text', text: 'Invalid .docx' }], isError: true };
  const docXml = parser.parse(await docFile.async('text')) as Record<string, unknown>;

  // Find the text to anchor to
  const fullText = extractFullText(docXml);
  const anchorPos = fullText.indexOf(searchText);
  if (anchorPos === -1) return { content: [{ type: 'text', text: `Text "${searchText}" not found in document.` }], isError: true };

  // Generate comment ID
  const commentsXml = await loadOrInitComments(zip);
  const commentId = getNextCommentId(commentsXml);

  // Add comment to comments.xml
  addCommentToXml(commentsXml, commentId, author, commentText);

  // Add comment anchors to document.xml
  addCommentAnchors(docXml, searchText, commentId);

  // Save
  await saveXmlPart(zip, 'word/document.xml', docXml, builder);
  await saveXmlPart(zip, 'word/comments.xml', commentsXml, builder);

  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const { writeFileSync } = await import('node:fs');
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        id: commentId,
        author,
        text: commentText,
        anchoredTo: searchText,
        _suggestions: {
          word_list_comments: { tool: 'word_list_comments', description: 'View all comments' },
          word_reply_to_comment: { tool: 'word_reply_to_comment', description: 'Reply to this comment' },
        },
      }, null, 2),
    }],
  };
}

function extractFullText(docXml: Record<string, unknown>): string {
  const body = docXml?.['w:document']?.['w:body'];
  if (!body) return '';
  const paragraphs = ensureArray(body?.['w:p']);
  return paragraphs.map((p) => findAllTextInNode(p as Record<string, unknown>).join('')).join('\n');
}

async function loadOrInitComments(zip: JSZip): Promise<Record<string, unknown>> {
  const commentsFile = zip.file('word/comments.xml');
  if (commentsFile) {
    const str = await commentsFile.async('text');
    return parser.parse(str) as Record<string, unknown>;
  }

  // Create new comments.xml
  const commentsXml = { 'w:comments': { '@_xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' } };

  // Add relationship
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (relsFile) {
    const relsStr = await relsFile.async('text');
    const rels = parser.parse(relsStr) as Record<string, unknown>;
    const relsObj = rels['Relationships'] as Record<string, unknown> | undefined;
    if (relsObj) {
      const existing = ensureArray(relsObj['Relationship']);
      existing.push({
        '@_Id': 'rId_comments',
        '@_Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
        '@_Target': 'comments.xml',
      });
      relsObj['Relationship'] = existing;
    }
    zip.file('word/_rels/document.xml.rels', builder.build(rels));
  }

  return commentsXml;
}

function getNextCommentId(commentsXml: Record<string, unknown>): string {
  const comments = ensureArray(commentsXml?.['w:comments']?.['w:comment']);
  let maxId = 0;
  for (const c of comments) {
    const id = Number((c as Record<string, unknown>)['@_w:id']);
    if (id > maxId) maxId = id;
  }
  return String(maxId + 1);
}

function addCommentToXml(commentsXml: Record<string, unknown>, id: string, author: string, text: string): void {
  const newComment = {
    '@_w:id': id,
    '@_w:author': author,
    '@_w:initials': author.charAt(0),
    '@_w:date': new Date().toISOString(),
    'w:p': [{ 'w:r': [{ 'w:t': { '@_xml:space': 'preserve', '#text': text } }] }],
  };

  const comments = commentsXml['w:comments'] as Record<string, unknown>;
  if (!comments) return;
  const existing = ensureArray(comments['w:comment']);
  existing.push(newComment);
  comments['w:comment'] = existing;
}

function addCommentAnchors(docXml: Record<string, unknown>, searchText: string, commentId: string): void {
  const body = docXml?.['w:document']?.['w:body'];
  if (!body) return;

  const paragraphs = ensureArray(body['w:p']);
  let globalOffset = 0;

  for (let i = 0; i < paragraphs.length; i++) {
    const pText = findAllTextInNode(paragraphs[i] as Record<string, unknown>).join('');
    const localPos = searchText.indexOf(pText);

    if (localPos !== -1 || (globalOffset <= searchText.length && globalOffset + pText.length >= searchText.length)) {
      // This paragraph contains the anchor text - add markers
      const pn = paragraphs[i] as Record<string, unknown>;

      // Add commentRangeEnd and commentReference at the end of paragraph
      const existingR = ensureArray(pn['w:r']);
      pn['w:r'] = [
        ...existingR,
        { 'w:commentRangeEnd': { '@_w:id': commentId } },
        { 'w:r': [{ 'w:commentReference': { '@_w:id': commentId } }] },
      ];
      return;
    }
    globalOffset += pText.length + 1;
  }
}

async function saveXmlPart(zip: JSZip, path: string, xml: Record<string, unknown>, b: XMLBuilder): Promise<void> {
  zip.file(path, b.build(xml));
}

function ensureArray(val: unknown): unknown[] {
  if (!val) return [];
  if (Array.isArray(val)) return val;
  return [val];
}

import { readFileSync, writeFileSync } from 'node:fs';
import JSZip from 'jszip';
import { XMLBuilder, XMLParser } from 'fast-xml-parser';
import { validateDocxPath } from '../validation.js';

export const WORD_ACCEPT_CHANGES_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    author: { type: 'string', description: 'Accept only changes by this author' },
    dateFrom: { type: 'string', description: 'Accept only changes on or after this ISO 8601 date' },
    dateTo: { type: 'string', description: 'Accept only changes on or before this ISO 8601 date' },
    changeId: { type: 'string', description: 'Accept only the change with this w:id' },
    section: { type: 'string', description: 'Reserved for section-scoped acceptance' },
  },
  required: ['filePath'],
} as const;

const orderedParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  trimValues: false,
});

const orderedBuilder = new XMLBuilder({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  format: false,
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
});

type ONode = Record<string, unknown>;

interface AcceptFilters {
  author?: string;
  dateFrom?: string;
  dateTo?: string;
  changeId?: string;
}

interface TransformResult {
  nodes: ONode[];
  acceptedCount: number;
}

const CHANGE_TAGS = new Set(['w:ins', 'w:del', 'w:pPrChange', 'w:rPrChange', 'w:numberingChange']);

export async function wordAcceptChanges(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, author, dateFrom, dateTo, changeId } = args as {
    filePath: string;
    author?: string;
    dateFrom?: string;
    dateTo?: string;
    changeId?: string;
  };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const zip = await JSZip.loadAsync(readFileSync(filePath));
  const docFile = zip.file('word/document.xml');
  if (!docFile) return { content: [{ type: 'text', text: 'Invalid .docx: missing word/document.xml' }], isError: true };

  const docOrdered = orderedParser.parse(await docFile.async('text')) as ONode[];
  const result = acceptChangesInNodes(docOrdered, { author, dateFrom, dateTo, changeId });

  zip.file('word/document.xml', orderedBuilder.build(result.nodes));
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        acceptedCount: result.acceptedCount,
        _suggestions: {
          word_get_document_info: { tool: 'word_get_document_info', description: 'Verify the document no longer reports tracked changes' },
          word_get_content: { tool: 'word_get_content', description: 'Review accepted document text' },
        },
      }, null, 2),
    }],
  };
}

function acceptChangesInNodes(nodes: ONode[], filters: AcceptFilters): TransformResult {
  const output: ONode[] = [];
  let acceptedCount = 0;

  for (const node of nodes) {
    const changeTag = getChangeTag(node);
    if (changeTag && matchesFilters(node, filters)) {
      acceptedCount++;
      if (changeTag === 'w:ins') {
        output.push(...((node[changeTag] as ONode[] | undefined) ?? []));
      }
      continue;
    }

    const transformed = transformNodeChildren(node, filters);
    output.push(transformed.node);
    acceptedCount += transformed.acceptedCount;
  }

  return { nodes: output, acceptedCount };
}

function transformNodeChildren(node: ONode, filters: AcceptFilters): { node: ONode; acceptedCount: number } {
  let acceptedCount = 0;
  const transformed: ONode = {};

  for (const [key, value] of Object.entries(node)) {
    if (Array.isArray(value)) {
      const result = acceptChangesInNodes(value as ONode[], filters);
      transformed[key] = result.nodes;
      acceptedCount += result.acceptedCount;
    } else {
      transformed[key] = value;
    }
  }

  return { node: transformed, acceptedCount };
}

function getChangeTag(node: ONode): string | null {
  return Object.keys(node).find((key) => CHANGE_TAGS.has(key)) ?? null;
}

function matchesFilters(node: ONode, filters: AcceptFilters): boolean {
  const attrs = node[':@'] as Record<string, string> | undefined;
  const id = attrs?.['@_w:id'];
  const author = attrs?.['@_w:author'];
  const date = attrs?.['@_w:date'];

  if (filters.changeId && id !== filters.changeId) return false;
  if (filters.author && author !== filters.author) return false;
  if (filters.dateFrom && (!date || Date.parse(date) < Date.parse(filters.dateFrom))) return false;
  if (filters.dateTo && (!date || Date.parse(date) > Date.parse(filters.dateTo))) return false;
  return true;
}

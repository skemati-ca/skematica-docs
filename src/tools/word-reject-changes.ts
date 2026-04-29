import { readFileSync, writeFileSync } from 'node:fs';
import JSZip from 'jszip';
import { XMLBuilder, XMLParser } from 'fast-xml-parser';
import { validateDocxPath } from '../validation.js';

export const WORD_REJECT_CHANGES_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    author: { type: 'string', description: 'Reject only changes by this author' },
    dateFrom: { type: 'string', description: 'Reject only changes on or after this ISO 8601 date' },
    dateTo: { type: 'string', description: 'Reject only changes on or before this ISO 8601 date' },
    changeId: { type: 'string', description: 'Reject only the change with this w:id' },
    section: { type: 'string', description: 'Reserved for section-scoped rejection' },
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

interface RejectFilters {
  author?: string;
  dateFrom?: string;
  dateTo?: string;
  changeId?: string;
}

interface TransformResult {
  nodes: ONode[];
  rejectedCount: number;
}

const CHANGE_TAGS = new Set(['w:ins', 'w:del']);

export async function wordRejectChanges(args: Record<string, unknown>): Promise<Record<string, unknown>> {
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
  const result = rejectChangesInNodes(docOrdered, { author, dateFrom, dateTo, changeId });

  zip.file('word/document.xml', orderedBuilder.build(result.nodes));
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        rejectedCount: result.rejectedCount,
        _suggestions: {
          word_get_document_info: { tool: 'word_get_document_info', description: 'Verify the document no longer reports tracked changes' },
          word_get_content: { tool: 'word_get_content', description: 'Review rejected document text' },
        },
      }, null, 2),
    }],
  };
}

function rejectChangesInNodes(nodes: ONode[], filters: RejectFilters): TransformResult {
  const output: ONode[] = [];
  let rejectedCount = 0;

  for (const node of nodes) {
    const changeTag = getChangeTag(node);
    if (changeTag && matchesFilters(node, filters)) {
      rejectedCount++;
      if (changeTag === 'w:del') {
        output.push(...restoreDeletedText((node[changeTag] as ONode[] | undefined) ?? []));
      }
      continue;
    }

    const transformed = transformNodeChildren(node, filters);
    output.push(transformed.node);
    rejectedCount += transformed.rejectedCount;
  }

  return { nodes: output, rejectedCount };
}

function transformNodeChildren(node: ONode, filters: RejectFilters): { node: ONode; rejectedCount: number } {
  let rejectedCount = 0;
  const transformed: ONode = {};

  for (const [key, value] of Object.entries(node)) {
    if (Array.isArray(value)) {
      const result = rejectChangesInNodes(value as ONode[], filters);
      transformed[key] = result.nodes;
      rejectedCount += result.rejectedCount;
    } else {
      transformed[key] = value;
    }
  }

  return { node: transformed, rejectedCount };
}

function restoreDeletedText(nodes: ONode[]): ONode[] {
  return nodes.map((node) => restoreDeletedTextInNode(node));
}

function restoreDeletedTextInNode(node: ONode): ONode {
  const restored: ONode = {};
  for (const [key, value] of Object.entries(node)) {
    if (key === 'w:delText') {
      restored['w:t'] = value;
    } else if (Array.isArray(value)) {
      restored[key] = value.map((child) => restoreDeletedTextInNode(child as ONode));
    } else {
      restored[key] = value;
    }
  }
  return restored;
}

function getChangeTag(node: ONode): string | null {
  return Object.keys(node).find((key) => CHANGE_TAGS.has(key)) ?? null;
}

function matchesFilters(node: ONode, filters: RejectFilters): boolean {
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

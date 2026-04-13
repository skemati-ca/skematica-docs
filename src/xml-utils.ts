import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import type { PageInfo, StyleInfo, CommentEntry } from './docx.js';

function asRecord(val: unknown): Record<string, unknown> | undefined {
  if (!val || typeof val !== 'object') return undefined;
  return val as Record<string, unknown>;
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  trimValues: true,
  updateTag: () => true,
});

const builder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  format: true,
  indentBy: '  ',
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
  preserveOrder: false,
});

export function parseXml(xmlStr: string): Record<string, unknown> {
  return parser.parse(xmlStr) as Record<string, unknown>;
}

export function buildXml(xml: Record<string, unknown>): string {
  return builder.build(xml);
}

export function extractParagraphs(docXml: Record<string, unknown>): string[] {
  const body = asRecord(asRecord(docXml?.['w:document'])?.['w:body']);
  if (!body) return [];
  const paragraphs = ensureArray(body?.['w:p']);
  return paragraphs.map((p) => extractParagraphText(asRecord(p) || {}));
}

export function extractParagraphStyles(docXml: Record<string, unknown>): string[] {
  const body = asRecord(asRecord(docXml?.['w:document'])?.['w:body']);
  if (!body) return [];
  const paragraphs = ensureArray(body?.['w:p']);
  return paragraphs.map((p) => {
    const pn = asRecord(p);
    const pPr = asRecord(pn?.['w:pPr']);
    const pStyle = asRecord(pPr?.['w:pStyle']);
    return (pStyle?.['@_w:val'] as string) || '';
  });
}

export function extractParagraphText(p: Record<string, unknown>): string {
  return findAllTextInNode(p).join('');
}

/**
 * Recursively finds all <w:t> text nodes within any OOXML node,
 * regardless of nesting depth (w:ins, w:del, w:proofErr, w:smartTag, etc.).
 * This handles OOXML run fragmentation correctly.
 */
export function findAllTextInNode(node: Record<string, unknown> | undefined): string[] {
  if (!node) return [];
  const texts: string[] = [];

  for (const [key, val] of Object.entries(node)) {
    if (key === 'w:t') {
      const tArr = ensureArray(val);
      for (const t of tArr) {
        const tn = asRecord(t);
        if (tn && '#text' in tn) {
          texts.push(String(tn['#text']));
        }
      }
    } else if (typeof val === 'object' && val !== null) {
      if (Array.isArray(val)) {
        for (const item of val) {
          if (typeof item === 'object' && item !== null) {
            texts.push(...findAllTextInNode(item as Record<string, unknown>));
          }
        }
      } else {
        texts.push(...findAllTextInNode(val as Record<string, unknown>));
      }
    }
  }

  return texts;
}

export function extractCommentRange(docXml: Record<string, unknown>, commentId: string): string {
  const body = asRecord(asRecord(docXml?.['w:document'])?.['w:body']);
  if (!body) return '';

  const children = flattenBodyChildren(body);
  let startIdx = -1;
  let endIdx = -1;

  for (let i = 0; i < children.length; i++) {
    const child = children[i] as Record<string, unknown>;
    const keys = Object.keys(child);
    if (keys.length === 0) continue;
    const tag = keys[0];
    const attrs = asRecord(child[tag]);

    if (tag === 'w:commentRangeStart' && attrs?.['@_w:id'] === commentId) {
      startIdx = i;
    }
    if (tag === 'w:commentRangeEnd' && attrs?.['@_w:id'] === commentId) {
      endIdx = i;
    }
  }

  if (startIdx === -1 || endIdx === -1 || endIdx <= startIdx) return '';

  // Collect text between start and end
  let text = '';
  for (let i = startIdx + 1; i < endIdx; i++) {
    const child = children[i] as Record<string, unknown>;
    const tag = Object.keys(child)[0];
    if (tag === 'w:p') {
      text += (text ? '\n' : '') + extractParagraphText(child);
    }
    if (tag === 'w:r') {
      text += extractRunText(child);
    }
  }

  return text.trim();
}

function flattenBodyChildren(body: Record<string, unknown>): Record<string, unknown>[] {
  const result: Record<string, unknown>[] = [];

  for (const [key, val] of Object.entries(body)) {
    if (key === 'w:p' || key === 'w:tbl' || key === 'w:sectPr' || key === 'w:commentRangeStart' || key === 'w:commentRangeEnd' || key === 'w:r') {
      const arr = ensureArray(val);
      for (const item of arr) {
        result.push({ [key]: item });
      }
    }
  }

  return result;
}

function extractRunText(r: Record<string, unknown>): string {
  const rn = asRecord(r);
  const t = rn?.['w:t'];
  if (!t) return '';
  return ensureArray(t)
    .map((tn) => {
      const tnn = asRecord(tn);
      return String(tnn?.['#text'] ?? '');
    })
    .join('');
}

export function parseComments(commentsXml: Record<string, unknown>): Omit<CommentEntry, 'parentId' | 'replies' | 'isResolved' | 'commentedText'>[] {
  const comments = commentsXml?.['w:comments']?.['w:comment'];
  if (!comments) return [];

  return ensureArray(comments).map((c) => {
    const cn = c as Record<string, unknown>;
    return {
      id: String(cn['@_w:id'] ?? ''),
      author: String(cn['@_w:author'] ?? ''),
      date: String(cn['@_w:date'] ?? ''),
      text: extractCommentText(cn),
    };
  });
}

function extractCommentText(c: Record<string, unknown>): string {
  const paragraphs = ensureArray(c?.['w:p']);
  return paragraphs
    .map((p) => extractParagraphText(p as Record<string, unknown>))
    .join('\n')
    .trim();
}

export function parseCommentsExtended(extXml: Record<string, unknown>): Record<string, { parentId: string | null; done: string }> {
  const commentsEx = extXml?.['w15:commentsEx']?.['w15:commentEx'];
  if (!commentsEx) return {};

  const result: Record<string, { parentId: string | null; done: string }> = {};
  for (const ex of ensureArray(commentsEx)) {
    const en = ex as Record<string, unknown>;
    const paraId = String(en['@_w15:paraId'] ?? '');
    const parentId = String(en['@_w15:parentId'] ?? '') || null;
    const done = String(en['@_w15:done'] ?? '0');
    if (paraId) {
      result[paraId] = { parentId, done };
    }
  }

  return result;
}

export function extractPageSetup(docXml: Record<string, unknown>): PageInfo {
  const body = docXml?.['w:document']?.['w:body'];
  if (!body) return defaultPageInfo();

  const sectPr = body['w:sectPr'] as Record<string, unknown> | undefined;
  if (!sectPr) return defaultPageInfo();

  const pgSz = sectPr['w:pgSz'] as Record<string, unknown> | undefined;
  const pgMar = sectPr['w:pgMar'] as Record<string, unknown> | undefined;
  const orient = pgSz ? String(pgSz['@_w:orient'] ?? '') : '';

  const width = pgSz ? Number(pgSz['@_w:w']) || 12240 : 12240;
  const height = pgSz ? Number(pgSz['@_w:h']) || 15840 : 15840;
  const orientation: 'portrait' | 'landscape' =
    orient === 'landscape' ? 'landscape' : width > height ? 'landscape' : 'portrait';

  return {
    pageSize: { width, height, label: lookupPageSizeLabel(width, height) },
    orientation,
    margins: {
      top: pgMar ? Number(pgMar['@_w:top']) || 1440 : 1440,
      bottom: pgMar ? Number(pgMar['@_w:bottom']) || 1440 : 1440,
      left: pgMar ? Number(pgMar['@_w:left']) || 1440 : 1440,
      right: pgMar ? Number(pgMar['@_w:right']) || 1440 : 1440,
    },
  };
}

function defaultPageInfo(): PageInfo {
  return {
    pageSize: { width: 12240, height: 15840, label: 'Letter' },
    orientation: 'portrait',
    margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
  };
}

function lookupPageSizeLabel(width: number, height: number): string {
  const presets: Record<string, { w: number; h: number }> = {
    Letter: { w: 12240, h: 15840 },
    Legal: { w: 12240, h: 20160 },
    Folio: { w: 12240, h: 19440 },
    A4: { w: 11906, h: 16838 },
    Executive: { w: 10440, h: 14400 },
  };
  for (const [label, dims] of Object.entries(presets)) {
    if (dims.w === width && dims.h === height) return label;
  }
  return 'Custom';
}

export function extractStyles(stylesXml: Record<string, unknown>): StyleInfo[] {
  const styles = stylesXml?.['w:styles']?.['w:style'];
  if (!styles) return [];

  return ensureArray(styles).map((s) => {
    const sn = s as Record<string, unknown>;
    const nameNode = sn?.['w:name'] as Record<string, unknown> | undefined;
    return {
      name: String(nameNode?.['@_w:val'] ?? sn['@_w:styleId'] ?? 'Unknown'),
      type: (sn['@_w:type'] as string) === 'paragraph' ? 'paragraph' : 'character',
      usageCount: 0,
    };
  });
}

function ensureArray(val: unknown): unknown[] {
  if (!val) return [];
  if (Array.isArray(val)) return val;
  return [val];
}

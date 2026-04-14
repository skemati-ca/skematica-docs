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

  return extractTextFromFlatChildren(body, commentId);
}

/**
 * Walks through a paragraph node and its children, collecting text between
 * commentRangeStart and commentRangeEnd markers.
 */
function extractTextFromFlatChildren(body: Record<string, unknown>, targetId: string): string {
  const texts: string[] = [];
  let isActive = false;
  
  // Process each key in the body object
  for (const [key, value] of Object.entries(body)) {
    if (key === 'w:p') {
      const paragraphs = ensureArray(value);
      for (const p of paragraphs) {
        const pText = extractTextFromParagraphWithComments(p as Record<string, unknown>, targetId);
        if (pText.found) {
          return pText.text;
        }
      }
    }
    // Handle comment range markers at body level
    if (key === 'w:commentRangeStart') {
      const arr = ensureArray(value);
      for (const m of arr) {
        const attrs = asRecord(m);
        if (attrs?.['@_w:id'] === targetId) {
          isActive = true;
        }
      }
    }
    if (key === 'w:commentRangeEnd') {
      const arr = ensureArray(value);
      for (const m of arr) {
        const attrs = asRecord(m);
        if (attrs?.['@_w:id'] === targetId) {
          isActive = false;
        }
      }
    }
    // Handle runs at body level
    if (key === 'w:r' && isActive) {
      const arr = ensureArray(value);
      for (const r of arr) {
        const rn = asRecord(r);
        const t = rn?.['w:t'];
        if (t) {
          const tArr = ensureArray(t);
          for (const tn of tArr) {
            const tnn = asRecord(tn);
            if (tnn && '#text' in tnn) {
              texts.push(String(tnn['#text']));
            }
          }
        }
      }
    }
  }
  
  return texts.join('').trim();
}

/**
 * Extracts text for a specific comment ID from a paragraph node.
 * Scans the paragraph's children for commentRangeStart/End markers.
 */
function extractTextFromParagraphWithComments(
  p: Record<string, unknown>,
  targetId: string
): { found: boolean; text: string } {
  if (!p) return { found: false, text: '' };
  
  let isActive = false;
  const texts: string[] = [];
  
  for (const [key, value] of Object.entries(p)) {
    if (key === 'w:commentRangeStart') {
      const arr = ensureArray(value);
      for (const m of arr) {
        const attrs = asRecord(m);
        if (attrs?.['@_w:id'] === targetId) {
          isActive = true;
        }
      }
    }
    if (key === 'w:commentRangeEnd') {
      const arr = ensureArray(value);
      for (const m of arr) {
        const attrs = asRecord(m);
        if (attrs?.['@_w:id'] === targetId) {
          isActive = false;
        }
      }
    }
    // Collect text from all text nodes (runs, inserts, deletions, etc.)
    if (key === 'w:r' || key === 'w:ins' || key === 'w:del') {
      const items = ensureArray(value);
      for (const item of items) {
        const inText = findAllTextInNode(item as Record<string, unknown>);
        if (isActive) {
          texts.push(...inText);
        }
      }
    }
  }
  
  if (texts.length > 0) {
    return { found: true, text: texts.join('').trim() };
  }
  
  return { found: false, text: '' };
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

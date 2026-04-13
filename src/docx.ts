import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import JSZip from 'jszip';

export interface DocxMetadata {
  pageCount: number;
  wordCount: number;
  characterCount: number;
  hasTrackChanges: boolean;
  commentCount: number;
  unresolvedComments: number;
  sectionCount: number;
  lastModifiedBy: string;
  format: string;
}

export interface CommentEntry {
  id: string;
  author: string;
  text: string;
  date: string;
  parentId: string | null;
  isResolved: boolean;
  replies: CommentEntry[];
  commentedText: string;
}

export interface SectionInfo {
  title: string;
  sectionIndex: number;
  wordCount: number;
}

export interface PageInfo {
  pageSize: { width: number; height: number; label: string };
  orientation: 'portrait' | 'landscape';
  margins: { top: number; bottom: number; left: number; right: number };
}

export interface TextMatch {
  matchIndex: number;
  matchText: string;
  contextBefore: string;
  contextAfter: string;
}

export interface StyleInfo {
  name: string;
  type: 'paragraph' | 'character';
  usageCount: number;
}

const PAGE_SIZE_PRESETS: Record<string, { width: number; height: number; label: string }> = {
  letter: { width: 12240, height: 15840, label: 'Letter' },
  legal: { width: 12240, height: 20160, label: 'Legal' },
  folio: { width: 12240, height: 19440, label: 'Folio' },
  a4: { width: 11906, height: 16838, label: 'A4' },
  executive: { width: 10440, height: 14400, label: 'Executive' },
};

export { PAGE_SIZE_PRESETS };

export class DocxDocument {
  private zip: JSZip;
  private content: Buffer;
  private textCache: string | null = null;
  private paragraphsCache: string[] | null = null;

  private constructor(content: Buffer, zip: JSZip) {
    this.content = content;
    this.zip = zip;
  }

  static async load(filePath: string): Promise<DocxDocument> {
    if (!existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`);
    }
    const content = readFileSync(filePath);
    const zip = await JSZip.loadAsync(content);
    if (!zip.file('word/document.xml')) {
      throw new Error('Invalid .docx: missing word/document.xml');
    }
    return new DocxDocument(content, zip);
  }

  async save(filePath: string): Promise<void> {
    const output = await this.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    writeFileSync(filePath, output);
  }

  async getText(): Promise<string> {
    if (this.textCache) return this.textCache;
    const docXml = await this.getDocumentXml();
    this.textCache = extractTextFromXml(docXml);
    return this.textCache;
  }

  async getParagraphs(): Promise<string[]> {
    if (this.paragraphsCache) return this.paragraphsCache;
    const docXml = await this.getDocumentXml();
    const { extractParagraphs } = await import('./xml-utils.js');
    this.paragraphsCache = extractParagraphs(docXml);
    return this.paragraphsCache;
  }

  async getMetadata(): Promise<DocxMetadata> {
    const text = await this.getText();
    await this.getParagraphs();
    const comments = await this.getComments();

    const hasTrackChanges = await this.hasTrackChanges();
    const resolved = countResolvedComments(comments);

    return {
      pageCount: estimatePageCount(text),
      wordCount: countWords(text),
      characterCount: text.length,
      hasTrackChanges,
      commentCount: countAllComments(comments),
      unresolvedComments: countAllComments(comments) - resolved,
      sectionCount: (await this.getSections()).length,
      lastModifiedBy: await this.getLastModifiedBy(),
      format: 'docx',
    };
  }

  async getSections(): Promise<SectionInfo[]> {
    const paragraphs = await this.getParagraphs();
    const styles = await this.getParagraphStyles();
    const sections: SectionInfo[] = [];
    let current: SectionInfo | null = null;

    for (let i = 0; i < paragraphs.length; i++) {
      const style = styles[i] || '';
      const isHeading = /^heading\s*\d*$/i.test(style);

      if (isHeading && paragraphs[i].trim()) {
        if (current) sections.push(current);
        current = { title: paragraphs[i].trim(), sectionIndex: sections.length + 1, wordCount: 0 };
      }
      if (current) {
        current.wordCount += countWords(paragraphs[i]);
      }
    }
    if (current) sections.push(current);

    if (sections.length === 0) {
      sections.push({
        title: 'Document Body',
        sectionIndex: 1,
        wordCount: countWords(paragraphs.join('\n')),
      });
    }

    return sections;
  }

  async getSectionContent(section: string, maxChars = 10000): Promise<{
    section: string;
    sectionIndex: number;
    wordCount: number;
    content: string;
    commentCount: number;
  }> {
    const sections = await this.getSections();
    const target = sections.find((s) => s.title === section);
    if (!target) throw new Error(`Section not found: ${section}`);

    const paragraphs = await this.getParagraphs();
    const styles = await this.getParagraphStyles();

    let collecting = false;
    let content = '';

    for (let i = 0; i < paragraphs.length; i++) {
      const isHeading = /^heading\s*\d*$/i.test(styles[i] || '');
      if (isHeading && paragraphs[i].trim() === section) {
        collecting = true;
        continue;
      }
      if (isHeading && paragraphs[i].trim() && collecting) break;
      if (collecting) {
        const newText = content ? `${content}\n${paragraphs[i]}` : paragraphs[i];
        if (newText.length > maxChars) {
          content += paragraphs[i].substring(0, maxChars - content.length);
          break;
        }
        content = newText;
      }
    }

    return {
      section: target.title,
      sectionIndex: target.sectionIndex,
      wordCount: countWords(content),
      content,
      commentCount: 0,
    };
  }

  async findText(searchText: string, contextChars = 100): Promise<TextMatch[]> {
    const text = await this.getText();
    const matches: TextMatch[] = [];
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

    return matches;
  }

  async getComments(): Promise<CommentEntry[]> {
    const docXml = await this.getDocumentXml();
    const commentsXml = await this.getCommentsXml();
    const extXml = await this.getCommentsExtendedXml();

    if (!commentsXml) return [];

    const { parseComments, parseCommentsExtended } = await import('./xml-utils.js');
    const comments = parseComments(commentsXml);
    const extended = extXml ? parseCommentsExtended(extXml) : {};

    // Build hierarchy
    const map = new Map<string, CommentEntry>();
    for (const c of comments) {
      const ext = extended[c.id];
      map.set(c.id, {
        ...c,
        parentId: ext?.parentId || null,
        isResolved: ext?.done === '1',
        replies: [],
        commentedText: await this.getCommentedText(c.id, docXml),
      });
    }

    const roots: CommentEntry[] = [];
    for (const entry of map.values()) {
      if (entry.parentId) {
        const parent = map.get(entry.parentId);
        if (parent) {
          parent.replies.push(entry);
        } else {
          roots.push(entry);
        }
      } else {
        roots.push(entry);
      }
    }

    return roots;
  }

  private async getCommentedText(commentId: string, docXml: Record<string, unknown>): Promise<string> {
    const { extractCommentRange } = await import('./xml-utils.js');
    return extractCommentRange(docXml, commentId);
  }

  async getPageSetup(): Promise<PageInfo> {
    const docXml = await this.getDocumentXml();
    const { extractPageSetup } = await import('./xml-utils.js');
    return extractPageSetup(docXml);
  }

  async getStyles(): Promise<{ availableStyles: StyleInfo[]; usedStyles: Record<string, number>; structureIssues: Array<{ paragraphIndex: number; expectedStyle: string; currentStyle: string }> }> {
    const stylesXml = await this.getStylesXml();
    const paragraphs = await this.getParagraphs();
    const styles = await this.getParagraphStyles();

    const available: StyleInfo[] = [];
    if (stylesXml) {
      const { extractStyles } = await import('./xml-utils.js');
      available.push(...extractStyles(stylesXml));
    }

    const used: Record<string, number> = {};
    for (const s of styles) {
      const name = s || 'Normal';
      used[name] = (used[name] || 0) + 1;
    }

    const issues: Array<{ paragraphIndex: number; expectedStyle: string; currentStyle: string }> = [];
    for (let i = 0; i < paragraphs.length; i++) {
      const text = paragraphs[i].trim();
      const isAllCaps = text === text.toUpperCase() && text.length > 5 && /[A-Z]/.test(text);
      const isNumbered = /^\d+[.)]\s/.test(text);
      const currentStyle = styles[i] || 'Normal';

      if ((isAllCaps || isNumbered) && currentStyle === 'Normal' && text.length < 100) {
        issues.push({
          paragraphIndex: i,
          expectedStyle: 'Heading',
          currentStyle: 'Normal',
        });
      }
    }

    return { availableStyles: available, usedStyles: used, structureIssues: issues };
  }

  async hasTrackChanges(): Promise<boolean> {
    const docXml = await this.getDocumentXml();
    const xmlStr = JSON.stringify(docXml);
    return xmlStr.includes('w:ins') || xmlStr.includes('w:del');
  }

  private async getLastModifiedBy(): Promise<string> {
    const coreXml = await this.getXmlPart('docProps/core.xml');
    if (!coreXml) return 'Unknown';
    const coreStr = JSON.stringify(coreXml);
    const match = coreStr.match(/dc:creator[^>]*>([^<]+)/);
    return match ? match[1].trim() : 'Unknown';
  }

  // Internal XML accessors
  private async getDocumentXml(): Promise<Record<string, unknown>> {
    return this.getXmlPart('word/document.xml') as Promise<Record<string, unknown>>;
  }

  private async getCommentsXml(): Promise<Record<string, unknown> | null> {
    return this.getXmlPart('word/comments.xml') as Promise<Record<string, unknown> | null>;
  }

  private async getCommentsExtendedXml(): Promise<Record<string, unknown> | null> {
    return this.getXmlPart('word/commentsExtended.xml') as Promise<Record<string, unknown> | null>;
  }

  private async getStylesXml(): Promise<Record<string, unknown> | null> {
    return this.getXmlPart('word/styles.xml') as Promise<Record<string, unknown> | null>;
  }

  private async getParagraphStyles(): Promise<string[]> {
    const docXml = await this.getDocumentXml();
    const { extractParagraphStyles } = await import('./xml-utils.js');
    return extractParagraphStyles(docXml);
  }

  private async getXmlPart(path: string): Promise<Record<string, unknown> | null> {
    const file = this.zip.file(path);
    if (!file) return null;
    const str = await file.async('text');
    const { parseXml } = await import('./xml-utils.js');
    return parseXml(str);
  }

  private async setXmlPart(path: string, xml: Record<string, unknown>): Promise<void> {
    const { buildXml } = await import('./xml-utils.js');
    const str = buildXml(xml);
    this.zip.file(path, str);
  }

  getZip(): JSZip {
    return this.zip;
  }
}

function extractTextFromXml(docXml: Record<string, unknown>): string {
  const body = docXml?.['w:document']?.['w:body'];
  if (!body) return '';

  const paragraphs = ensureArray((body as Record<string, unknown>)?.['w:p']);
  return paragraphs.map((p) => extractParagraphText(p as Record<string, unknown>)).join('\n');
}

function extractParagraphText(p: Record<string, unknown>): string {
  const runs = ensureArray(p?.['w:r']);
  return runs.map((r) => {
    const rn = r as Record<string, unknown>;
    const t = rn['w:t'];
    if (!t) return '';
    const texts = ensureArray(t);
    return texts.map((tn) => {
      const tnn = tn as Record<string, unknown>;
      return tnn['#text'] ?? '';
    }).join('');
  }).join('');
}

function ensureArray(val: unknown): unknown[] {
  if (!val) return [];
  if (Array.isArray(val)) return val;
  return [val];
}

function countWords(text: string): number {
  return text.split(/\s+/).filter((w) => w.length > 0).length;
}

function estimatePageCount(text: string): number {
  const words = countWords(text);
  return Math.max(1, Math.ceil(words / 300));
}

function countAllComments(comments: CommentEntry[]): number {
  let count = 0;
  for (const c of comments) {
    count += 1 + c.replies.length;
  }
  return count;
}

function countResolvedComments(comments: CommentEntry[]): number {
  let count = 0;
  for (const c of comments) {
    if (c.isResolved) count++;
    for (const r of c.replies) {
      if (r.isResolved) count++;
    }
  }
  return count;
}

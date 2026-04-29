import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { validateDocxPath } from '../validation.js';
import { findAllTextInNode, collectBodyParagraphs } from '../xml-utils.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const WORD_SEARCH_REPLACE_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    searchText: { type: 'string', description: 'Text to find' },
    replacementText: { type: 'string', description: 'Text to replace it with' },
    matchIndex: { type: 'number', description: 'Replace only this match (0-based). Omit for replace-all.' },
    section: { type: 'string', description: 'Replace only within a specific section' },
    author: { type: 'string', description: 'Author for revision entry. Default: "Asistente IA"' },
  },
  required: ['filePath', 'searchText', 'replacementText'],
} as const;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', trimValues: true });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', format: true, indentBy: '  ', suppressEmptyNode: false });

export async function wordSearchReplace(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, searchText, replacementText, matchIndex } = args as { filePath: string; searchText: string; replacementText: string; matchIndex?: number };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const content = readFileSync(filePath);
  const zip = await JSZip.loadAsync(content);
  const docFile = zip.file('word/document.xml');
  if (!docFile) return { content: [{ type: 'text', text: 'Invalid .docx: missing word/document.xml' }], isError: true };

  const docXmlStr = await docFile.async('text');
  const docXml = parser.parse(docXmlStr) as Record<string, unknown>;

  const fullText = extractFullText(docXml);
  const matches: number[] = [];
  let pos = 0;
  while (true) {
    const found = fullText.indexOf(searchText, pos);
    if (found === -1) break;
    matches.push(found);
    pos = found + searchText.length;
  }

  if (matches.length === 0) {
    return { content: [{ type: 'text', text: `Text "${searchText}" not found in document.` }], isError: true };
  }

  const targets = matchIndex !== undefined ? [matches[matchIndex]] : matches;
  if (targets.some((t) => t === undefined)) {
    return { content: [{ type: 'text', text: `matchIndex ${matchIndex} out of range. Found ${matches.length} matches.` }], isError: true };
  }

  replaceInXml(docXml, searchText, replacementText, targets as number[]);

  const newDocXmlStr = builder.build(docXml);
  zip.file('word/document.xml', newDocXmlStr);
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const { writeFileSync } = await import('node:fs');
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        replacedCount: targets.length,
        replacedLocations: targets.map((idx: number) => ({ matchIndex: matches.indexOf(idx), position: idx })),
        _suggestions: {
          word_find_text: { tool: 'word_find_text', description: 'Verify no remaining instances' },
          word_get_content: { tool: 'word_get_content', description: 'Review the result' },
        },
      }, null, 2),
    }],
  };
}

function extractFullText(docXml: Record<string, unknown>): string {
  const body = docXml?.['w:document']?.['w:body'] as Record<string, unknown> | undefined;
  if (!body) return '';
  return collectBodyParagraphs(body).map((p) => findAllTextInNode(p).join('')).join('\n');
}

function replaceInXml(docXml: Record<string, unknown>, searchText: string, replacementText: string, matchPositions: number[]): void {
  const body = docXml?.['w:document']?.['w:body'] as Record<string, unknown> | undefined;
  if (!body) return;

  const paragraphs = collectBodyParagraphs(body);
  let globalOffset = 0;

  for (const p of paragraphs) {
    const pn = p as Record<string, unknown>;
    const pText = findAllTextInNode(pn).join('');
    const pStart = globalOffset;
    const pEnd = globalOffset + pText.length;

    for (const matchPos of matchPositions) {
      if (matchPos >= pStart && matchPos < pEnd) {
        const localPos = matchPos - pStart;
        replaceInParagraph(pn, searchText, replacementText, localPos);
      }
    }

    globalOffset += pText.length + 1;
  }
}

function replaceInParagraph(p: Record<string, unknown>, searchText: string, replacementText: string, localPos: number): void {
  const runs = ensureArray(p?.['w:r']) as Record<string, unknown>[];
  let offset = 0;

  for (const r of runs) {
    const rn = r as Record<string, unknown>;
    const textNodes = ensureArray(rn?.['w:t']) as Record<string, unknown>[];

    for (const tn of textNodes) {
      const tnn = tn as Record<string, unknown>;
      const text = String(tnn['#text'] ?? '');

      if (localPos >= offset && localPos < offset + text.length) {
        const inNodePos = localPos - offset;
        const currentText = String(tnn['#text'] ?? '');
        if (currentText.substring(inNodePos, inNodePos + searchText.length) === searchText) {
          tnn['#text'] = currentText.substring(0, inNodePos) + replacementText + currentText.substring(inNodePos + searchText.length);
          return;
        }
      }
      offset += text.length;
    }
  }
}

function ensureArray(val: unknown): unknown[] {
  if (!val) return [];
  if (Array.isArray(val)) return val;
  return [val];
}

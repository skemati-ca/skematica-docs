import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { validateDocxPath } from '../validation.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const WORD_APPLY_STYLE_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    paragraphIndex: { type: 'number', description: '0-based paragraph index' },
    paragraphRange: { type: 'object', description: 'Range of paragraphs { start, end }', properties: { start: { type: 'number' }, end: { type: 'number' } } },
    style: { type: 'string', description: 'Style name (e.g., "Heading1", "Normal")' },
  },
  required: ['filePath', 'style'],
} as const;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', trimValues: true });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', format: true, indentBy: '  ', suppressEmptyNode: false });

export async function wordApplyStyle(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, paragraphIndex, paragraphRange, style } = args as { filePath: string; paragraphIndex?: number; paragraphRange?: { start: number; end: number }; style: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  if (paragraphIndex === undefined && !paragraphRange) {
    return { content: [{ type: 'text', text: 'Either paragraphIndex or paragraphRange is required.' }], isError: true };
  }

  const content = readFileSync(filePath);
  const zip = await JSZip.loadAsync(content);
  const docFile = zip.file('word/document.xml');
  if (!docFile) return { content: [{ type: 'text', text: 'Invalid .docx' }], isError: true };

  const docXml = parser.parse(await docFile.async('text')) as Record<string, unknown>;
  const body = docXml?.['w:document']?.['w:body'] as Record<string, unknown> | undefined;
  if (!body) return { content: [{ type: 'text', text: 'No document body found.' }], isError: true };

  const paragraphs = ensureArray(body['w:p']) as Record<string, unknown>[];
  const targetIndex = paragraphIndex ?? 0;
  const indices: number[] = paragraphRange
    ? Array.from({ length: paragraphRange.end - paragraphRange.start + 1 }, (_, i) => paragraphRange.start + i)
    : [targetIndex];

  let appliedCount = 0;
  const appliedTo: number[] = [];

  for (const idx of indices) {
    if (idx < 0 || idx >= paragraphs.length) continue;
    const p = paragraphs[idx];

    let pPr = p['w:pPr'] as Record<string, unknown> | undefined;
    if (!pPr) {
      pPr = {};
      p['w:pPr'] = pPr;
    }

    pPr['w:pStyle'] = { '@_w:val': style };
    appliedCount++;
    appliedTo.push(idx);
  }

  zip.file('word/document.xml', builder.build(docXml));
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const { writeFileSync } = await import('node:fs');
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        appliedCount,
        appliedTo,
        style,
        _suggestions: {
          word_get_styles: { tool: 'word_get_styles', description: 'Verify structure' },
          word_get_sections: { tool: 'word_get_sections', description: 'View new document structure' },
          word_apply_style: { tool: 'word_apply_style', description: 'Apply to more paragraphs' },
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

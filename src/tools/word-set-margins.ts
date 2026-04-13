import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { validateDocxPath } from '../validation.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const WORD_SET_MARGINS_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    top: { type: 'number', description: 'Top margin value' },
    bottom: { type: 'number', description: 'Bottom margin value' },
    left: { type: 'number', description: 'Left margin value' },
    right: { type: 'number', description: 'Right margin value' },
    unit: { type: 'string', enum: ['twips', 'pt', 'in', 'cm'], description: 'Unit of measurement. Default: "twips"' },
    sectionIndex: { type: 'number', description: '1-based section index. Default: last section' },
  },
  required: ['filePath', 'top', 'bottom', 'left', 'right'],
} as const;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', trimValues: true });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', format: true, indentBy: '  ', suppressEmptyNode: false });

function toTwips(value: number, unit: string): number {
  switch (unit) {
    case 'pt': return Math.round(value * 20);
    case 'in': return Math.round(value * 1440);
    case 'cm': return Math.round(value * 567);
    default: return value; // twips
  }
}

export async function wordSetMargins(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, top, bottom, left, right, unit = 'twips' } = args as { filePath: string; top: number; bottom: number; left: number; right: number; unit?: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const twips = {
    top: toTwips(top, unit),
    bottom: toTwips(bottom, unit),
    left: toTwips(left, unit),
    right: toTwips(right, unit),
  };

  const content = readFileSync(filePath);
  const zip = await JSZip.loadAsync(content);
  const docFile = zip.file('word/document.xml');
  if (!docFile) return { content: [{ type: 'text', text: 'Invalid .docx' }], isError: true };

  const docXml = parser.parse(await docFile.async('text')) as Record<string, unknown>;
  const body = docXml?.['w:document']?.['w:body'] as Record<string, unknown> | undefined;
  if (!body) return { content: [{ type: 'text', text: 'No document body found.' }], isError: true };

  let sectPr = body['w:sectPr'] as Record<string, unknown> | undefined;
  if (!sectPr) {
    sectPr = {};
    body['w:sectPr'] = sectPr;
  }

  sectPr['w:pgMar'] = {
    '@_w:top': twips.top,
    '@_w:bottom': twips.bottom,
    '@_w:left': twips.left,
    '@_w:right': twips.right,
  };

  zip.file('word/document.xml', builder.build(docXml));
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const { writeFileSync } = await import('node:fs');
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        margins: { ...twips, unit: 'twips' },
        _suggestions: {
          word_get_page_setup: { tool: 'word_get_page_setup', description: 'Verify changes' },
        },
      }, null, 2),
    }],
  };
}

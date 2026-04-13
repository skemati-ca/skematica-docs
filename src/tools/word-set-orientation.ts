import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { validateDocxPath } from '../validation.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const WORD_SET_ORIENTATION_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    orientation: { type: 'string', enum: ['portrait', 'landscape'], description: 'Page orientation' },
    sectionIndex: { type: 'number', description: '1-based section index. Default: last section' },
  },
  required: ['filePath', 'orientation'],
} as const;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', trimValues: true });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', format: true, indentBy: '  ', suppressEmptyNode: false });

export async function wordSetOrientation(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, orientation } = args as { filePath: string; orientation: string };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  if (orientation !== 'portrait' && orientation !== 'landscape') {
    return { content: [{ type: 'text', text: `Invalid orientation: ${orientation}. Use "portrait" or "landscape".` }], isError: true };
  }

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

  // Update or create pgSz with orient attribute
  let pgSz = sectPr['w:pgSz'] as Record<string, unknown> | undefined;
  if (!pgSz) {
    pgSz = { '@_w:w': 12240, '@_w:h': 15840 };
  }
  pgSz['@_w:orient'] = orientation === 'landscape' ? 'landscape' : undefined;

  // Swap dimensions for landscape
  if (orientation === 'landscape') {
    const w = pgSz['@_w:w'];
    const h = pgSz['@_w:h'];
    if (w && h && Number(w) < Number(h)) {
      pgSz['@_w:w'] = h;
      pgSz['@_w:h'] = w;
    }
  }

  sectPr['w:pgSz'] = pgSz;

  zip.file('word/document.xml', builder.build(docXml));
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const { writeFileSync } = await import('node:fs');
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        orientation,
        _suggestions: {
          word_get_page_setup: { tool: 'word_get_page_setup', description: 'Verify changes' },
          word_set_page_size: { tool: 'word_set_page_size', description: 'Adjust page size if needed' },
        },
      }, null, 2),
    }],
  };
}

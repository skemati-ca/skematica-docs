import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { validateDocxPath } from '../validation.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const WORD_SET_PAGE_SIZE_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    pageSize: { type: 'string', description: 'Preset: "letter", "legal", "folio", "a4", "executive", or JSON string of {width, height} in twips' },
    sectionIndex: { type: 'number', description: '1-based section index. Default: last section' },
  },
  required: ['filePath', 'pageSize'],
} as const;

const PAGE_SIZE_PRESETS: Record<string, { width: number; height: number }> = {
  letter: { width: 12240, height: 15840 },
  legal: { width: 12240, height: 20160 },
  folio: { width: 12240, height: 19440 },
  a4: { width: 11906, height: 16838 },
  executive: { width: 10440, height: 14400 },
};

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', trimValues: true });
const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', format: true, indentBy: '  ', suppressEmptyNode: false });

export async function wordSetPageSize(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, pageSize } = args as { filePath: string; pageSize: string; sectionIndex?: number };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  let width: number;
  let height: number;
  let label: string;

  if (typeof pageSize === 'string' && PAGE_SIZE_PRESETS[pageSize]) {
    const preset = PAGE_SIZE_PRESETS[pageSize];
    width = preset.width;
    height = preset.height;
    label = pageSize.charAt(0).toUpperCase() + pageSize.slice(1);
  } else {
    try {
      const dims = typeof pageSize === 'string' ? JSON.parse(pageSize) : pageSize;
      width = dims.width;
      height = dims.height;
      label = 'Custom';
    } catch {
      return { content: [{ type: 'text', text: `Invalid pageSize: ${pageSize}. Use preset name or {width, height} in twips.` }], isError: true };
    }
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

  // Update or create pgSz
  sectPr['w:pgSz'] = { '@_w:w': width, '@_w:h': height };

  zip.file('word/document.xml', builder.build(docXml));
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const { writeFileSync } = await import('node:fs');
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        pageSize: { width, height, label },
        _suggestions: {
          word_get_page_setup: { tool: 'word_get_page_setup', description: 'Verify changes' },
          word_set_orientation: { tool: 'word_set_orientation', description: 'Change orientation if needed' },
        },
      }, null, 2),
    }],
  };
}

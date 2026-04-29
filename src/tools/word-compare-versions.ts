import { validateDocxPath } from '../validation.js';
import { findAllTextInNode, collectBodyParagraphs } from '../xml-utils.js';

export const WORD_COMPARE_VERSIONS_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Path to the newer version' },
    compareWith: { type: 'string', description: 'Path to the older version' },
    maxChanges: { type: 'number', description: 'Maximum changes to return. Default: 50' },
    section: { type: 'string', description: 'Limit comparison to a specific section' },
  },
  required: ['filePath', 'compareWith'],
} as const;

export async function wordCompareVersions(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const { filePath, compareWith, maxChanges = 50 } = args as { filePath: string; compareWith: string; maxChanges?: number };

  const err1 = validateDocxPath(filePath);
  if (err1) return { content: [{ type: 'text', text: err1 }], isError: true };
  const err2 = validateDocxPath(compareWith);
  if (err2) return { content: [{ type: 'text', text: err2 }], isError: true };

  const { compareDocuments } = await import('@usejunior/docx-core');
  const { readFileSync } = await import('node:fs');
  const JSZip = (await import('jszip')).default;

  const original = readFileSync(compareWith);
  const revised = readFileSync(filePath);

  const result = await compareDocuments(original, revised, { engine: 'atomizer', ignoreFormatting: true });

  // Extract changes from the compared document
  const changedZip = await JSZip.loadAsync(result.document);
  const changedDocStr = await changedZip.file('word/document.xml')?.async('text');

  const changes: Array<{ type: string; text: string; section?: string }> = [];

  if (changedDocStr) {
    const { XMLParser } = await import('fast-xml-parser');
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', textNodeName: '#text', trimValues: true });
    const changedDoc = parser.parse(changedDocStr) as Record<string, unknown>;
    const body = changedDoc?.['w:document']?.['w:body'];

    if (body) {
      const paragraphs = collectBodyParagraphs(body as Record<string, unknown>);
      for (const p of paragraphs) {
        const pn = p as Record<string, unknown>;
        const runs = ensureArray(pn?.['w:r']);
        for (const r of runs) {
          const rn = r as Record<string, unknown>;
          for (const tag of ['w:ins', 'w:del']) {
            const tagged = rn[tag];
            if (tagged) {
              const taggedArr = ensureArray(tagged);
              for (const t of taggedArr) {
                const tn = t as Record<string, unknown>;
                const innerR = tn['w:r'];
                if (innerR) {
                  const text = extractRunText(innerR as Record<string, unknown>);
                  if (text) changes.push({ type: tag === 'w:ins' ? 'inserted' : 'deleted', text });
                }
              }
            }
          }
        }
      }
    }
  }

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        totalChanges: changes.length,
        returnedChanges: Math.min(changes.length, maxChanges),
        changes: changes.slice(0, maxChanges),
        stats: result.stats,
        _suggestions: {
          word_compare_versions: { tool: 'word_compare_versions', description: maxChanges < changes.length ? `Increase maxChanges (showing ${maxChanges} of ${changes.length})` : 'View full comparison' },
          word_find_text: { tool: 'word_find_text', description: 'Search for specific changes' },
          word_get_section_content: { tool: 'word_get_section_content', description: 'Read sections with many changes' },
        },
      }, null, 2),
    }],
  };
}

function extractRunText(r: Record<string, unknown>): string {
  return findAllTextInNode(r).join('');
}

function ensureArray(val: unknown): unknown[] {
  if (!val) return [];
  if (Array.isArray(val)) return val;
  return [val];
}

import { readFileSync, writeFileSync } from 'node:fs';
import JSZip from 'jszip';
import { validateDocxPath } from '../validation.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export const WORD_INSERT_TRACKED_CHANGE_SCHEMA = {
  type: 'object',
  properties: {
    filePath: { type: 'string', description: 'Absolute path to the .docx file' },
    searchText: { type: 'string', description: 'Existing text to mark as deleted' },
    replacementText: { type: 'string', description: 'New text to insert. Use empty string for a pure deletion.' },
    author: { type: 'string', description: 'Author name for the tracked change. Default: "Asistente IA"' },
    matchIndex: { type: 'number', description: 'Which occurrence to target (0-based). Default: 0 (first).' },
  },
  required: ['filePath', 'searchText', 'replacementText'],
} as const;

// Dedicated ordered parser/builder — isolated from the rest of the codebase.
// preserveOrder:true is required so we can insert w:del/w:ins as siblings of w:r
// at the paragraph level (they are never nested inside w:r).
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

function getRunText(runNode: ONode): string {
  const runChildren = runNode['w:r'] as ONode[] | undefined;
  if (!runChildren) return '';
  return runChildren
    .filter(c => 'w:t' in c)
    .flatMap(c => (c['w:t'] as ONode[]).map(t => String(t['#text'] ?? '')))
    .join('');
}

function getRprFromRun(runNode: ONode): ONode | null {
  const children = runNode['w:r'] as ONode[] | undefined;
  return children?.find(c => 'w:rPr' in c) ?? null;
}

function cloneNode(node: ONode): ONode {
  return JSON.parse(JSON.stringify(node));
}

function makeRun(text: string, rPrNode: ONode | null): ONode {
  const children: ONode[] = [];
  if (rPrNode) children.push(cloneNode(rPrNode));
  if (text) {
    const tNode: ONode = { 'w:t': [{ '#text': text }] };
    if (text !== text.trim()) {
      (tNode as ONode)[':@'] = { '@_xml:space': 'preserve' };
    }
    children.push(tNode);
  }
  return { 'w:r': children };
}

function makeDelRun(text: string, rPrNode: ONode | null): ONode {
  const children: ONode[] = [];
  if (rPrNode) children.push(cloneNode(rPrNode));
  if (text) children.push({ 'w:delText': [{ '#text': text }] });
  return { 'w:r': children };
}

function makeChangeElement(tag: 'w:del' | 'w:ins', id: number, author: string, date: string, innerRuns: ONode | ONode[]): ONode {
  return {
    [tag]: Array.isArray(innerRuns) ? innerRuns : [innerRuns],
    ':@': { '@_w:id': String(id), '@_w:author': author, '@_w:date': date },
  };
}

// Collect all paragraph children-arrays in document order (body + table cells).
function collectParaArrays(bodyChildren: ONode[]): ONode[][] {
  const result: ONode[][] = [];
  for (const child of bodyChildren) {
    if ('w:p' in child) {
      result.push(child['w:p'] as ONode[]);
    } else if ('w:tbl' in child) {
      for (const tblChild of child['w:tbl'] as ONode[]) {
        if ('w:tr' in tblChild) {
          for (const trChild of tblChild['w:tr'] as ONode[]) {
            if ('w:tc' in trChild) {
              for (const tcChild of trChild['w:tc'] as ONode[]) {
                if ('w:p' in tcChild) result.push(tcChild['w:p'] as ONode[]);
              }
            }
          }
        }
      }
    }
  }
  return result;
}

// Find the next available w:id by scanning the entire document tree.
function nextId(nodes: ONode[]): number {
  let max = 0;
  function walk(arr: ONode[]): void {
    for (const node of arr) {
      for (const [key, val] of Object.entries(node)) {
        if (key === ':@') {
          const id = Number((val as Record<string, string>)['@_w:id']);
          if (!isNaN(id)) max = Math.max(max, id);
        } else if (Array.isArray(val)) {
          walk(val as ONode[]);
        }
      }
    }
  }
  walk(nodes);
  return max + 1;
}

interface RunRangeSegment {
  run: { nodeIdx: number; node: ONode; text: string; start: number };
  text: string;
  localStart: number;
  localEnd: number;
}

function collectRunsCoveringRange(
  runs: Array<{ nodeIdx: number; node: ONode; text: string; start: number }>,
  start: number,
  end: number,
): RunRangeSegment[] {
  const result: RunRangeSegment[] = [];
  for (const run of runs) {
    const runEnd = run.start + run.text.length;
    if (runEnd <= start || run.start >= end) continue;

    const localStart = Math.max(0, start - run.start);
    const localEnd = Math.min(run.text.length, end - run.start);
    result.push({
      run,
      text: run.text.substring(localStart, localEnd),
      localStart,
      localEnd,
    });
  }
  return result;
}

function sameRunProperties(segments: RunRangeSegment[]): boolean {
  if (segments.length <= 1) return true;
  const first = JSON.stringify(getRprFromRun(segments[0].run.node));
  return segments.every((segment) => JSON.stringify(getRprFromRun(segment.run.node)) === first);
}

function splitRunAtTextOffset(run: { text: string }, offset: number): [string, string] {
  return [run.text.substring(0, offset), run.text.substring(offset)];
}

// Apply a tracked change to the target paragraph at the given match index.
// Returns true on success, false if the text is not found.
function applyChange(
  paraChildren: ONode[],
  searchText: string,
  replacementText: string,
  author: string,
  baseId: number,
  localMatchIndex: number,
): boolean {
  // Collect run nodes with their cumulative offset in the paragraph text.
  const runs: Array<{ nodeIdx: number; node: ONode; text: string; start: number }> = [];
  let offset = 0;
  for (let i = 0; i < paraChildren.length; i++) {
    const child = paraChildren[i];
    if ('w:r' in child) {
      const text = getRunText(child);
      runs.push({ nodeIdx: i, node: child, text, start: offset });
      offset += text.length;
    }
  }

  const paraText = runs.map(r => r.text).join('');

  // Find all occurrences of searchText in the concatenated paragraph text.
  const positions: number[] = [];
  let pos = 0;
  while (true) {
    const found = paraText.indexOf(searchText, pos);
    if (found === -1) break;
    positions.push(found);
    pos = found + searchText.length;
  }
  if (localMatchIndex >= positions.length) return false;

  const matchPos = positions[localMatchIndex];
  const matchEnd = matchPos + searchText.length;

  const segments = collectRunsCoveringRange(runs, matchPos, matchEnd);
  if (segments.length === 0) return false;

  const firstSegment = segments[0];
  const lastSegment = segments[segments.length - 1];
  const [beforeText] = splitRunAtTextOffset(firstSegment.run, firstSegment.localStart);
  const [, afterText] = splitRunAtTextOffset(lastSegment.run, lastSegment.localEnd);
  const firstRPr = getRprFromRun(firstSegment.run.node);
  const lastRPr = getRprFromRun(lastSegment.run.node);
  const date = new Date().toISOString();

  const newNodes: ONode[] = [];
  if (beforeText) newNodes.push(makeRun(beforeText, firstRPr));

  const delRuns = sameRunProperties(segments)
    ? [makeDelRun(searchText, firstRPr)]
    : segments.map((segment) => makeDelRun(segment.text, getRprFromRun(segment.run.node)));
  newNodes.push(makeChangeElement('w:del', baseId, author, date, delRuns));

  if (replacementText) {
    newNodes.push(makeChangeElement('w:ins', baseId + 1, author, date, makeRun(replacementText, firstRPr)));
  }
  if (afterText) newNodes.push(makeRun(afterText, lastRPr));

  const deleteCount = lastSegment.run.nodeIdx - firstSegment.run.nodeIdx + 1;
  paraChildren.splice(firstSegment.run.nodeIdx, deleteCount, ...newNodes);
  return true;
}

export async function wordInsertTrackedChange(args: Record<string, unknown>): Promise<Record<string, unknown>> {
  const {
    filePath,
    searchText,
    replacementText,
    author = 'Asistente IA',
    matchIndex = 0,
  } = args as { filePath: string; searchText: string; replacementText: string; author?: string; matchIndex?: number };

  const err = validateDocxPath(filePath);
  if (err) return { content: [{ type: 'text', text: err }], isError: true };

  const content = readFileSync(filePath);
  const zip = await JSZip.loadAsync(content);
  const docFile = zip.file('word/document.xml');
  if (!docFile) return { content: [{ type: 'text', text: 'Invalid .docx: missing word/document.xml' }], isError: true };

  const docXmlStr = await docFile.async('text');
  const docOrdered = orderedParser.parse(docXmlStr) as ONode[];

  const docEl = docOrdered.find(n => 'w:document' in n);
  if (!docEl) return { content: [{ type: 'text', text: 'Malformed document: missing w:document' }], isError: true };

  const bodyEl = (docEl['w:document'] as ONode[]).find(n => 'w:body' in n);
  if (!bodyEl) return { content: [{ type: 'text', text: 'Malformed document: missing w:body' }], isError: true };

  const bodyChildren = bodyEl['w:body'] as ONode[];
  const paraArrays = collectParaArrays(bodyChildren);

  const baseId = nextId(docOrdered);
  let globalCount = 0;
  let success = false;

  for (const paraChildren of paraArrays) {
    const paraText = paraChildren
      .filter(c => 'w:r' in c)
      .map(c => getRunText(c))
      .join('');

    let localCount = 0;
    let pos = 0;
    while (paraText.indexOf(searchText, pos) !== -1) {
      pos = paraText.indexOf(searchText, pos) + searchText.length;
      localCount++;
    }

    if (globalCount + localCount > matchIndex) {
      const localMatchIndex = matchIndex - globalCount;
      success = applyChange(paraChildren, searchText, replacementText, author, baseId, localMatchIndex);
      break;
    }
    globalCount += localCount;
  }

  if (!success) {
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          error: `"${searchText}" not found or spans multiple runs. For cross-run edits use word_search_replace.`,
        }),
      }],
      isError: true,
    };
  }

  const newDocXmlStr = orderedBuilder.build(docOrdered);
  zip.file('word/document.xml', newDocXmlStr);
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  writeFileSync(filePath, output);

  return {
    content: [{
      type: 'text',
      text: JSON.stringify({
        deletedText: searchText,
        insertedText: replacementText || '(deleted)',
        author,
        _suggestions: {
          word_find_text: { tool: 'word_find_text', description: 'Verify the text is no longer present as plain text' },
          word_get_content: { tool: 'word_get_content', description: 'Review the full document' },
          word_create_comment: { tool: 'word_create_comment', description: 'Add a comment explaining the change' },
        },
      }, null, 2),
    }],
  };
}

import { createHash } from 'node:crypto';

const PPR_ORDER = new Map([
  ['w:pStyle', 10],
  ['w:numPr', 20],
  ['w:spacing', 30],
  ['w:ind', 40],
  ['w:jc', 50],
  ['w:rPr', 999],
]);

const MAX_SIGNED_OOXML_ID = 0x7fffffff;

/**
 * Applies lightweight, write-safe OOXML invariants for Word document XML.
 */
export function normalizeDocumentXml(xml: string): string {
  return normalizeRsids(
    normalizeHighIds(
      reorderParagraphProperties(
        encodeSmartQuotes(addXmlSpacePreserve(xml))
      )
    )
  );
}

function addXmlSpacePreserve(xml: string): string {
  return xml.replace(/<w:t\b([^>]*)>([\s\S]*?)<\/w:t>/g, (match, attrs: string, text: string) => {
    if (/\bxml:space\s*=/.test(attrs)) return match;
    if (!hasBoundaryWhitespace(text)) return match;
    return `<w:t${attrs} xml:space="preserve">${text}</w:t>`;
  });
}

function hasBoundaryWhitespace(text: string): boolean {
  return /^\s/.test(text) || /\s$/.test(text);
}

function encodeSmartQuotes(xml: string): string {
  return xml
    .replaceAll('\u2018', '&#x2018;')
    .replaceAll('\u2019', '&#x2019;')
    .replaceAll('\u201C', '&#x201C;')
    .replaceAll('\u201D', '&#x201D;');
}

function reorderParagraphProperties(xml: string): string {
  return xml.replace(/<w:pPr\b([^>]*)>([\s\S]*?)<\/w:pPr>/g, (_match, attrs: string, content: string) => {
    const children = splitTopLevelElements(content);
    const indexed = children.map((raw, index) => ({
      raw,
      index,
      rank: PPR_ORDER.get(getElementName(raw) ?? '') ?? 900,
    }));
    indexed.sort((a, b) => a.rank - b.rank || a.index - b.index);
    return `<w:pPr${attrs}>${indexed.map((child) => child.raw).join('')}</w:pPr>`;
  });
}

function splitTopLevelElements(content: string): string[] {
  const result: string[] = [];
  let depth = 0;
  let start = -1;
  const tagPattern = /<[^>]+>/g;

  for (const match of content.matchAll(tagPattern)) {
    const tag = match[0];
    const index = match.index ?? 0;
    const isClosing = /^<\//.test(tag);
    const isSelfClosing = /\/>$/.test(tag);
    const isDeclaration = /^<\?/.test(tag) || /^<!--/.test(tag);

    if (isDeclaration) continue;

    if (!isClosing && depth === 0) {
      start = index;
    }

    if (!isClosing && !isSelfClosing) {
      depth++;
    } else if (isClosing) {
      depth--;
      if (depth === 0 && start !== -1) {
        result.push(content.slice(start, index + tag.length));
        start = -1;
      }
    }

    if (!isClosing && isSelfClosing && depth === 0) {
      result.push(tag);
      start = -1;
    }
  }

  return result.length > 0 ? result : [content];
}

function getElementName(raw: string): string | null {
  return raw.match(/^<([A-Za-z0-9:_-]+)/)?.[1] ?? null;
}

function normalizeHighIds(xml: string): string {
  return xml.replace(/\b((?:w14:paraId|w15:durableId))="([0-9A-Fa-f]+)"/g, (match, name: string, value: string) => {
    const numeric = Number.parseInt(value, 16);
    if (Number.isNaN(numeric) || numeric < MAX_SIGNED_OOXML_ID) return match;
    return `${name}="${makeEightDigitHex(`${name}:${value}`)}"`;
  });
}

function normalizeRsids(xml: string): string {
  return xml.replace(/\b(w:rsid[A-Za-z0-9]*)="([^"]*)"/g, (match, name: string, value: string) => {
    if (/^[0-9A-Fa-f]{8}$/.test(value)) return match;
    return `${name}="${makeEightDigitHex(`${name}:${value}`)}"`;
  });
}

function makeEightDigitHex(seed: string): string {
  const digest = createHash('sha1').update(seed).digest();
  const value = digest.readUInt32BE(0) & 0x7ffffffe;
  return value.toString(16).toUpperCase().padStart(8, '0');
}

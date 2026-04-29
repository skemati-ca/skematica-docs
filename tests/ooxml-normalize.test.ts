import { describe, expect, it } from 'vitest';
import { join } from 'node:path';
import { DocxDocument } from '../src/docx.js';
import { normalizeDocumentXml } from '../src/ooxml-normalize.js';

const fixturesDir = join(process.cwd(), 'tests', 'fixtures');

describe('normalizeDocumentXml', () => {
  it('adds xml:space="preserve" to w:t nodes with leading or trailing whitespace', () => {
    const xml = [
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      '<w:body><w:p><w:r><w:t> leading </w:t></w:r><w:r><w:t>inner</w:t></w:r></w:p></w:body>',
      '</w:document>',
    ].join('');

    const normalized = normalizeDocumentXml(xml);

    expect(normalized).toContain('<w:t xml:space="preserve"> leading </w:t>');
    expect(normalized).toContain('<w:t>inner</w:t>');
  });

  it('preserves smart quotes as numeric entities through serialization', () => {
    const xml = [
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      '<w:body><w:p><w:r><w:t>‘single’ “double”</w:t></w:r></w:p></w:body>',
      '</w:document>',
    ].join('');

    const normalized = normalizeDocumentXml(xml);

    expect(normalized).toContain('&#x2018;single&#x2019; &#x201C;double&#x201D;');
  });

  it('orders known w:pPr children with w:rPr last after paragraph properties', () => {
    const xml = [
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      '<w:body><w:p><w:pPr>',
      '<w:rPr><w:b/></w:rPr><w:jc w:val="center"/><w:pStyle w:val="Heading1"/>',
      '<w:ind w:left="720"/><w:spacing w:after="120"/><w:numPr><w:numId w:val="1"/></w:numPr>',
      '</w:pPr><w:r><w:t>Title</w:t></w:r></w:p></w:body></w:document>',
    ].join('');

    const normalized = normalizeDocumentXml(xml);
    const pPr = normalized.slice(normalized.indexOf('<w:pPr>'), normalized.indexOf('</w:pPr>'));

    expect(pPr.indexOf('<w:pStyle')).toBeLessThan(pPr.indexOf('<w:numPr'));
    expect(pPr.indexOf('<w:numPr')).toBeLessThan(pPr.indexOf('<w:spacing'));
    expect(pPr.indexOf('<w:spacing')).toBeLessThan(pPr.indexOf('<w:ind'));
    expect(pPr.indexOf('<w:ind')).toBeLessThan(pPr.indexOf('<w:jc'));
    expect(pPr.indexOf('<w:jc')).toBeLessThan(pPr.indexOf('<w:rPr'));
  });

  it('regenerates paraId and durableId values at or above 0x7FFFFFFF', () => {
    const xml = [
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ',
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">',
      '<w:body><w:p w14:paraId="7FFFFFFF" w15:durableId="80000000"><w:r><w:t>Text</w:t></w:r></w:p></w:body>',
      '</w:document>',
    ].join('');

    const normalized = normalizeDocumentXml(xml);

    expect(normalized).not.toContain('w14:paraId="7FFFFFFF"');
    expect(normalized).not.toContain('w15:durableId="80000000"');
    expect(normalized).toMatch(/w14:paraId="[0-9A-F]{8}"/);
    expect(normalized).toMatch(/w15:durableId="[0-9A-F]{8}"/);
  });

  it('coerces invalid RSID values to 8-digit hex', () => {
    const xml = [
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      '<w:body><w:p w:rsidR="bad" w:rsidRPr="123" w:rsidP="123456789"><w:r><w:t>Text</w:t></w:r></w:p></w:body>',
      '</w:document>',
    ].join('');

    const normalized = normalizeDocumentXml(xml);

    expect(normalized).not.toContain('w:rsidR="bad"');
    expect(normalized).not.toContain('w:rsidRPr="123"');
    expect(normalized).not.toContain('w:rsidP="123456789"');
    expect(normalized).toMatch(/w:rsidR="[0-9A-F]{8}"/);
    expect(normalized).toMatch(/w:rsidRPr="[0-9A-F]{8}"/);
    expect(normalized).toMatch(/w:rsidP="[0-9A-F]{8}"/);
  });

  it('is applied when DocxDocument writes word/document.xml parts', async () => {
    const doc = await DocxDocument.load(join(fixturesDir, 'simple.docx'));
    const writableDoc = doc as unknown as {
      setXmlPart(path: string, xml: Record<string, unknown>): Promise<void>;
      getZip(): { file(path: string): { async(type: 'text'): Promise<string> } | null };
    };

    await writableDoc.setXmlPart('word/document.xml', {
      'w:document': {
        '@_xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w:body': {
          'w:p': {
            '@_w:rsidR': 'invalid',
            'w:pPr': {
              'w:rPr': { 'w:b': '' },
              'w:pStyle': { '@_w:val': 'Heading1' },
            },
            'w:r': {
              'w:t': { '#text': ' “quoted” ' },
            },
          },
        },
      },
    });

    const documentXml = await writableDoc.getZip().file('word/document.xml')?.async('text');

    expect(documentXml).toContain('<w:t xml:space="preserve"> &#x201C;quoted&#x201D; </w:t>');
    expect(documentXml).toMatch(/w:rsidR="[0-9A-F]{8}"/);
    expect(documentXml?.indexOf('<w:pStyle')).toBeLessThan(documentXml?.indexOf('<w:rPr') ?? 0);
  });
});

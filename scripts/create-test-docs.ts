import JSZip from 'jszip';
import { writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';

function buildXml(tag: string, attrs: Record<string, string> | null, children: (string | { xml: string })[]): string {
  const attrStr = attrs
    ? Object.entries(attrs)
        .map(([k, v]) => ` ${k}="${v}"`)
        .join('')
    : '';

  const childStr = children
    .map((c) => (typeof c === 'string' ? c : c.xml))
    .join('');

  return `<${tag}${attrStr}>${childStr}</${tag}>`;
}

function buildRun(text: string): string {
  return buildXml('w:r', null, [buildXml('w:t', { 'xml:space': 'preserve' }, [text])]);
}

function buildParagraph(text: string, style?: string): string {
  const children: (string | { xml: string })[] = [];
  if (style) {
    children.push({
      xml: buildXml('w:pPr', null, [buildXml('w:pStyle', { 'w:val': style }, [])]),
    });
  }
  children.push(buildRun(text));
  return buildXml('w:p', null, children);
}

async function createDocx(
  filePath: string,
  paragraphs: Array<{ text: string; style?: string }>,
  options?: { comments?: Array<{ id: string; author: string; text: string; anchoredText: string }> }
): Promise<void> {
  const zip = new JSZip();

  // Content types
  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
  ];
  if (options?.comments) {
    overrides.push(
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
    );
  }

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${overrides.join('\n  ')}
</Types>`
  );

  // _rels/.rels
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  // Build document.xml
  const bodyChildren: string[] = [];
  for (const p of paragraphs) {
    bodyChildren.push(buildParagraph(p.text, p.style));
  }

  bodyChildren.push(
    buildXml('w:sectPr', null, [
      buildXml('w:pgSz', { 'w:w': '12240', 'w:h': '15840' }, []),
      buildXml('w:pgMar', { 'w:top': '1440', 'w:right': '1440', 'w:bottom': '1440', 'w:left': '1440' }, []),
    ])
  );

  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyChildren.join('\n    ')}
  </w:body>
</w:document>`
  );

  // word/styles.xml
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
  </w:style>
</w:styles>`
  );

  // word/comments.xml
  if (options?.comments) {
    const commentXml = options.comments
      .map(
        (c) =>
          `<w:comment w:id="${c.id}" w:author="${c.author}" w:initials="${c.author.charAt(0)}" w:date="${new Date().toISOString()}">
      <w:p>
        <w:r><w:t xml:space="preserve">${c.text}</w:t></w:r>
      </w:p>
    </w:comment>`
      )
      .join('\n  ');

    zip.file(
      'word/comments.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${commentXml}
</w:comments>`
    );
  }

  // word/_rels/document.xml.rels
  const rels = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
  ];
  if (options?.comments) {
    rels.push('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>');
  }

  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${rels.join('\n  ')}
</Relationships>`
  );

  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  mkdirSync(join(filePath, '..'), { recursive: true });
  writeFileSync(filePath, buf);
  console.log(`Created: ${filePath} (${(buf.length / 1024).toFixed(1)} KB)`);
}

const fixturesDir = join(process.cwd(), 'tests', 'fixtures');

// 1. Simple - no comments, no headings
createDocx(join(fixturesDir, 'simple.docx'), [
  { text: 'Contrato de Prestación de Servicios' },
  { text: 'Entre los suscritos a saber:' },
  { text: 'CLÁUSULA PRIMERA - OBJETO' },
  { text: 'El contratista se obliga a desarrollar software para el cliente.' },
  { text: 'CLÁUSULA SEGUNDA - VALOR' },
  { text: 'El valor del contrato es de $50.000.000 COP.' },
  { text: 'CLÁUSULA TERCERA - VIGENCIA' },
  { text: 'El contrato tendrá una vigencia de 12 meses a partir de la firma.' },
]);

// 2. With comments
createDocx(
  join(fixturesDir, 'with-comments.docx'),
  [
    { text: 'Contrato de Desarrollo de Software' },
    { text: 'CLÁUSULA PRIMERA - OBJETO' },
    { text: 'El contratista se obliga a desarrollar una plataforma web.' },
    { text: 'CLÁUSULA SEGUNDA - VALOR' },
    { text: 'El valor del contrato es de $2.400.000.000 COP para Q1 2026.' },
    { text: 'CLÁUSULA TERCERA - PENALIDADES' },
    { text: 'El incumplimiento generará una penalidad del 10% mensual.' },
  ],
  {
    comments: [
      { id: '1', author: 'María López', text: '¿Está actualizado este monto?', anchoredText: '2.400.000.000' },
      { id: '2', author: 'Carlos Ruiz', text: 'Revisar penalidad con el comité jurídico.', anchoredText: 'penalidad del 10%' },
    ],
  }
);

// 3. Structured with headings
createDocx(join(fixturesDir, 'structured.docx'), [
  { text: '1. Objeto del Contrato', style: 'Heading1' },
  { text: 'El contratista se obliga a prestar servicios de consultoría.' },
  { text: '2. Obligaciones del Contratista', style: 'Heading1' },
  { text: 'Entregar informes mensuales de avance.' },
  { text: 'Mantener confidencialidad sobre la información del cliente.' },
  { text: '3. Valor y Forma de Pago', style: 'Heading1' },
  { text: 'El valor total es de $100.000.000 COP pagaderos en 4 cuotas.' },
  { text: '4. Penalidades', style: 'Heading1' },
  { text: 'El retraso injustificado generará una multa del 5% mensual.' },
  { text: '5. Vigencia', style: 'Heading1' },
  { text: 'El contrato tendrá vigencia de 6 meses prorrogables.' },
]);

// 4. Large document for truncation testing
const largeParagraphs = [{ text: 'Documento de Prueba - Documento Grande', style: 'Heading1' }];
for (let i = 1; i <= 50; i++) {
  largeParagraphs.push({
    text: `Sección ${i} - Este es el párrafo número ${i} del documento de prueba para evaluar el funcionamiento del servidor MCP de skematica-docs en la extracción de contenido de archivos DOCX con muchos párrafos y texto extenso para pruebas de truncamiento.`,
  });
}
createDocx(join(fixturesDir, 'large.docx'), largeParagraphs);

// 5. Multi-section with Heading2
createDocx(join(fixturesDir, 'multi-section.docx'), [
  { text: 'Informe Legal', style: 'Heading1' },
  { text: 'Resumen ejecutivo del caso.' },
  { text: 'Antecedentes', style: 'Heading2' },
  { text: 'El caso comenzó en el año 2024 con una disputa contractual.' },
  { text: 'Análisis', style: 'Heading2' },
  { text: 'Se encontró que el contrato Q1 2026 tiene inconsistencias.' },
  { text: 'Conclusiones', style: 'Heading2' },
  { text: 'Se recomienda renegociar los términos del contrato.' },
]);

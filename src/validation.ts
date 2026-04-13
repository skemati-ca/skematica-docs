import { extname } from 'node:path';

const SUPPORTED_EXTENSIONS = ['.docx'];

const CONVERSION_INSTRUCTIONS = `Unsupported format: {ext}. Please convert to .docx first.

How to convert:
- In Word: File → Save As → Word Document (.docx)
- In Google Docs: Upload → File → Download → Microsoft Word (.docx)
- In LibreOffice/OpenOffice: Open → Save As → DOCX
- Safe online converters: CloudConvert (cloudconvert.com/doc-to-docx), Zamzar (zamzar.com/convert/doc-to-docx)`;

export function validateDocxPath(filePath: string): string | null {
  const ext = extname(filePath);

  if (ext === '') {
    return `No file extension detected. Supported formats: ${SUPPORTED_EXTENSIONS.join(', ')}.`;
  }

  if (ext === '.doc') {
    return CONVERSION_INSTRUCTIONS.replace('{ext}', ext);
  }

  if (!SUPPORTED_EXTENSIONS.includes(ext)) {
    return `Unsupported format: ${ext}. Supported formats: ${SUPPORTED_EXTENSIONS.join(', ')}.`;
  }

  return null;
}

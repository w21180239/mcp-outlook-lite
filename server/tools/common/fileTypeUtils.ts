import { Buffer } from 'buffer';
import { debug, warn } from '../../utils/logger.js';
import { parseExcelContent } from './excelParser.js';
import { parseOfficeDocument } from './documentParser.js';
export { parseExcelContent } from './excelParser.js';
export { parseOfficeDocument } from './documentParser.js';

// --- Detection utilities ---

export function formatFileSize(bytes: number) {
  if (bytes === null || bytes === undefined || isNaN(bytes) || typeof bytes !== 'number') {
    return 'Unknown size';
  }
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  if (bytes === 0) return '0 Bytes';
  const absBytes = Math.abs(bytes);
  const i = Math.floor(Math.log(absBytes) / Math.log(1024));
  const sizeIndex = Math.min(i, sizes.length - 1);
  const value = Math.round(absBytes / Math.pow(1024, sizeIndex) * 100) / 100;
  return value + ' ' + sizes[sizeIndex];
}

export function isTextContent(contentType: string | null, filename: string | null, contentBytes: string | null = null) {
  debug(`Debug: isTextContent check - contentType: "${contentType}", filename: "${filename}"`);

  const textTypes = [
    'text/',
    'application/json',
    'application/xml',
    'application/javascript',
    'application/typescript',
    'application/x-python',
    'application/x-sh',
    'application/sql'
  ];

  const textExtensions = [
    '.txt', '.md', '.csv', '.log', '.ini', '.cfg', '.conf',
    '.html', '.htm', '.xml', '.json', '.js', '.ts', '.py',
    '.sh', '.bash', '.sql', '.css', '.scss', '.less',
    '.yaml', '.yml', '.toml', '.properties', '.env'
  ];

  // Check content type first
  if (contentType) {
    const lowerContentType = contentType.toLowerCase();
    if (textTypes.some(type => lowerContentType.startsWith(type))) {
      debug(`Debug: Detected as text by contentType: ${contentType}`);
      return true;
    }
  }

  // Check file extension
  if (filename) {
    const ext = filename.toLowerCase().substring(filename.lastIndexOf('.'));
    if (textExtensions.includes(ext)) {
      debug(`Debug: Detected as text by extension: ${ext}`);
      return true;
    }
  }

  // If contentType is null/empty, try to detect from content
  if ((!contentType || contentType.trim() === '') && contentBytes) {
    try {
      const sampleContent = Buffer.from(contentBytes, 'base64').toString('utf8', 0, 200);
      if (sampleContent.includes('<!DOCTYPE html>') ||
          sampleContent.includes('<html>') ||
          sampleContent.includes('<?xml') ||
          sampleContent.startsWith('{') ||
          sampleContent.startsWith('[')) {
        debug(`Debug: Detected as text by content analysis`);
        return true;
      }
    } catch (error) {
      debug(`Debug: Content analysis failed: ${error.message}`);
    }
  }

  debug(`Debug: Detected as binary`);
  return false;
}

export function isExcelFile(contentType: string | null, filename: string | null) {
  const excelMimeTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
    'application/vnd.ms-excel', // .xls
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template', // .xltx
    'application/vnd.ms-excel.sheet.macroEnabled.12', // .xlsm
    'application/vnd.ms-excel.template.macroEnabled.12', // .xltm
    'application/vnd.ms-excel.addin.macroEnabled.12', // .xlam
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12' // .xlsb
  ];

  const excelExtensions = ['.xlsx', '.xls', '.xlsm', '.xltx', '.xltm', '.xlam', '.xlsb'];

  // Check content type
  if (contentType && excelMimeTypes.includes(contentType.toLowerCase())) {
    return true;
  }

  // Check file extension
  if (filename) {
    const ext = filename.toLowerCase().substring(filename.lastIndexOf('.'));
    return excelExtensions.includes(ext);
  }

  return false;
}

export function isOfficeDocument(contentType: string | null, filename: string | null) {
  const officeMimeTypes = [
    // PDF
    'application/pdf',
    // Word documents
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
    'application/msword', // .doc
    'application/vnd.openxmlformats-officedocument.wordprocessingml.template', // .dotx
    'application/vnd.ms-word.document.macroEnabled.12', // .docm
    'application/vnd.ms-word.template.macroEnabled.12', // .dotm
    // PowerPoint documents
    'application/vnd.openxmlformats-officedocument.presentationml.presentation', // .pptx
    'application/vnd.ms-powerpoint', // .ppt
    'application/vnd.openxmlformats-officedocument.presentationml.template', // .potx
    'application/vnd.openxmlformats-officedocument.presentationml.slideshow', // .ppsx
    'application/vnd.ms-powerpoint.addin.macroEnabled.12', // .ppam
    'application/vnd.ms-powerpoint.presentation.macroEnabled.12', // .pptm
    'application/vnd.ms-powerpoint.template.macroEnabled.12', // .potm
    'application/vnd.ms-powerpoint.slideshow.macroEnabled.12', // .ppsm
    // OpenDocument formats
    'application/vnd.oasis.opendocument.text', // .odt
    'application/vnd.oasis.opendocument.presentation', // .odp
    'application/vnd.oasis.opendocument.spreadsheet', // .ods
    // RTF
    'application/rtf',
    'text/rtf'
  ];

  const officeExtensions = [
    '.pdf',
    '.doc', '.docx', '.docm', '.dotx', '.dotm',
    '.ppt', '.pptx', '.pptm', '.potx', '.potm', '.ppsx', '.ppsm', '.ppam',
    '.odt', '.odp', '.ods',
    '.rtf'
  ];

  // Check content type
  if (contentType && officeMimeTypes.includes(contentType.toLowerCase())) {
    return true;
  }

  // Check file extension
  if (filename) {
    const ext = filename.toLowerCase().substring(filename.lastIndexOf('.'));
    return officeExtensions.includes(ext);
  }

  return false;
}

// --- Content decoder ---

export async function decodeContent(contentBytes: string, contentType: string | null, filename: string | null, maxTextSize = 1024 * 1024) {
  try {
    const buffer = Buffer.from(contentBytes, 'base64');
    const decodedSize = buffer.length;

    debug(`Debug: decodeContent - size: ${decodedSize}, contentType: "${contentType}", filename: "${filename}"`);

    if (isTextContent(contentType, filename, contentBytes)) {
      if (decodedSize <= maxTextSize) {
        return {
          type: 'text',
          content: buffer.toString('utf8'),
          size: decodedSize,
          sizeFormatted: formatFileSize(decodedSize),
          encoding: 'utf8'
        };
      } else {
        return {
          type: 'text',
          content: `[Text file too large to display: ${formatFileSize(decodedSize)}]`,
          contentBytes,
          size: decodedSize,
          sizeFormatted: formatFileSize(decodedSize),
          encoding: 'base64_preserved',
          note: 'File exceeds display limit, use contentBytes for full content'
        };
      }
    } else if (isExcelFile(contentType, filename)) {
      debug(`Debug: Detected Excel file, attempting to parse`);
      const excelData = parseExcelContent(contentBytes, filename);
      if ((excelData as any).type === 'excel_error') {
        warn(`Excel parse failed for ${filename}: ${(excelData as any).error}`);
        return {
          type: 'error',
          content: `[Excel parse failed: ${(excelData as any).error}]`,
          contentBytes,
          encoding: 'base64_fallback',
          error: (excelData as any).error
        };
      }
      return {
        type: 'excel',
        content: excelData,
        size: decodedSize,
        sizeFormatted: formatFileSize(decodedSize),
        encoding: 'parsed',
        contentBytes,
        note: 'Excel file parsed and data extracted. Use contentBytes for raw file access.'
      };
    } else if (isOfficeDocument(contentType, filename)) {
      debug(`Debug: Detected office document, attempting to parse`);
      const officeData = await parseOfficeDocument(contentBytes, filename);
      if ((officeData as any).type === 'office_error') {
        warn(`Office document parse failed for ${filename}: ${(officeData as any).error}`);
        return {
          type: 'error',
          content: `[Office parse failed: ${(officeData as any).error}]`,
          contentBytes,
          encoding: 'base64_fallback',
          error: (officeData as any).error
        };
      }
      return {
        type: 'office',
        content: officeData,
        size: decodedSize,
        sizeFormatted: formatFileSize(decodedSize),
        encoding: 'parsed',
        contentBytes,
        note: 'Office document parsed and text extracted. Use contentBytes for raw file access.'
      };
    } else {
      return {
        type: 'binary',
        content: `[Binary file: ${contentType || 'unknown type'}, ${formatFileSize(decodedSize)}]`,
        contentBytes,
        size: decodedSize,
        sizeFormatted: formatFileSize(decodedSize),
        encoding: 'base64',
        note: 'Binary file preserved as Base64, decode with Buffer.from(contentBytes, "base64") if needed'
      };
    }
  } catch (error) {
    return {
      type: 'error',
      content: `[Failed to decode content: ${error.message}]`,
      contentBytes,
      encoding: 'base64_fallback',
      error: error.message
    };
  }
}

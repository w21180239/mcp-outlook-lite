import { Buffer } from 'buffer';
import * as XLSX from 'xlsx';
import officeParser from 'officeparser';

// --- Detection utilities ---

export function formatFileSize(bytes) {
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

export function isTextContent(contentType, filename, contentBytes = null) {
  console.error(`Debug: isTextContent check - contentType: "${contentType}", filename: "${filename}"`);

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
      console.error(`Debug: Detected as text by contentType: ${contentType}`);
      return true;
    }
  }

  // Check file extension
  if (filename) {
    const ext = filename.toLowerCase().substring(filename.lastIndexOf('.'));
    if (textExtensions.includes(ext)) {
      console.error(`Debug: Detected as text by extension: ${ext}`);
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
        console.error(`Debug: Detected as text by content analysis`);
        return true;
      }
    } catch (error) {
      console.error(`Debug: Content analysis failed: ${error.message}`);
    }
  }

  console.error(`Debug: Detected as binary`);
  return false;
}

export function isExcelFile(contentType, filename) {
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

export function isOfficeDocument(contentType, filename) {
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

// --- Parsing utilities ---

export function parseExcelContent(contentBytes, filename, maxSheets = 10, maxRowsPerSheet = 1000) {
  try {
    console.error(`Debug: Parsing Excel file: ${filename}`);

    // Decode Base64 to buffer
    const buffer = Buffer.from(contentBytes, 'base64');

    // Parse Excel file
    const workbook = XLSX.read(buffer, { type: 'buffer' });

    const result = {
      type: 'excel',
      filename: filename,
      sheets: [],
      summary: {
        totalSheets: workbook.SheetNames.length,
        sheetNames: workbook.SheetNames
      }
    };

    // Process up to maxSheets sheets
    const sheetsToProcess = workbook.SheetNames.slice(0, maxSheets);

    for (const sheetName of sheetsToProcess) {
      console.error(`Debug: Processing sheet: ${sheetName}`);

      const worksheet = workbook.Sheets[sheetName];

      // Get sheet range
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      const totalRows = range.e.r - range.s.r + 1;
      const totalCols = range.e.c - range.s.c + 1;

      // Limit rows to prevent overwhelming output
      const rowsToProcess = Math.min(totalRows, maxRowsPerSheet);

      // Convert to JSON with limited rows
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1, // Use array format instead of object
        range: rowsToProcess < totalRows ? `${worksheet['!ref'].split(':')[0]}:${XLSX.utils.encode_cell({r: range.s.r + rowsToProcess - 1, c: range.e.c})}` : undefined
      });

      const sheetInfo = {
        name: sheetName,
        dimensions: {
          rows: totalRows,
          columns: totalCols,
          range: worksheet['!ref'] || 'A1:A1'
        },
        data: jsonData,
        truncated: rowsToProcess < totalRows,
        displayedRows: jsonData.length,
        note: rowsToProcess < totalRows ? `Sheet truncated to ${maxRowsPerSheet} rows (total: ${totalRows})` : undefined
      };

      result.sheets.push(sheetInfo);
    }

    if (workbook.SheetNames.length > maxSheets) {
      result.summary.note = `Only first ${maxSheets} sheets displayed (total: ${workbook.SheetNames.length})`;
    }

    console.error(`Debug: Successfully parsed Excel file with ${result.sheets.length} sheets`);
    return result;

  } catch (error) {
    console.error(`Debug: Excel parsing failed: ${error.message}`);
    return {
      type: 'excel_error',
      error: `Failed to parse Excel file: ${error.message}`,
      note: 'File may be corrupted or in an unsupported Excel format'
    };
  }
}

export function parseOfficeDocument(contentBytes, filename, maxTextLength = 50000) {
  try {
    console.error(`Debug: Parsing office document: ${filename}`);

    // Decode Base64 to buffer
    const buffer = Buffer.from(contentBytes, 'base64');

    // Parse office document using officeParser
    // NOTE: officeparser uses non-standard (data, err) callback signature
    return new Promise((resolve) => {
      officeParser.parseOffice(buffer, (data, err) => {
        if (err) {
          console.error(`Debug: Office parsing failed: ${err}`);
          resolve({
            type: 'office_error',
            error: `Failed to parse office document: ${err}`,
            note: 'File may be corrupted, password-protected, or in an unsupported format'
          });
          return;
        }

        // Extract and process the text content
        const extractedText = data || '';
        const textLength = extractedText.length;
        const truncated = textLength > maxTextLength;
        const displayText = truncated ? extractedText.substring(0, maxTextLength) + '...' : extractedText;

        const result = {
          type: 'office_document',
          filename: filename,
          content: {
            text: displayText,
            extractedLength: textLength,
            truncated: truncated,
            truncatedLength: truncated ? maxTextLength : undefined,
            note: truncated ? `Text truncated to ${maxTextLength} characters (total: ${textLength})` : undefined
          },
          metadata: {
            originalSize: buffer.length,
            textLength: textLength,
            hasContent: textLength > 0
          }
        };

        console.error(`Debug: Successfully parsed office document with ${textLength} characters of text`);
        resolve(result);
      });
    });

  } catch (error) {
    console.error(`Debug: Office parsing failed: ${error.message}`);
    return Promise.resolve({
      type: 'office_error',
      error: `Failed to parse office document: ${error.message}`,
      note: 'File may be corrupted, password-protected, or in an unsupported format'
    });
  }
}

// --- Content decoder ---

export async function decodeContent(contentBytes, contentType, filename, maxTextSize = 1024 * 1024) {
  try {
    const buffer = Buffer.from(contentBytes, 'base64');
    const decodedSize = buffer.length;

    console.error(`Debug: decodeContent - size: ${decodedSize}, contentType: "${contentType}", filename: "${filename}"`);

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
      console.error(`Debug: Detected Excel file, attempting to parse`);
      const excelData = parseExcelContent(contentBytes, filename);
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
      console.error(`Debug: Detected office document, attempting to parse`);
      const officeData = await parseOfficeDocument(contentBytes, filename);
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

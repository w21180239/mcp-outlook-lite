import { Buffer } from 'buffer';
import officeParser from 'officeparser';

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

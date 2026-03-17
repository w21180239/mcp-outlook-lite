import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock officeparser before importing
vi.mock('officeparser', () => ({
  default: {
    parseOffice: vi.fn(),
  },
}));

import { parseOfficeDocument } from '../../../tools/common/documentParser.js';
import officeParser from 'officeparser';

describe('parseOfficeDocument', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  it('should extract text from document', async () => {
    officeParser.parseOffice.mockImplementation((buffer, callback) => {
      // Non-standard (data, err) signature
      callback('Hello world document text', undefined);
    });

    const result = await parseOfficeDocument('dGVzdA==', 'test.docx');

    expect(result.type).toBe('office_document');
    expect(result.filename).toBe('test.docx');
    expect(result.content.text).toBe('Hello world document text');
    expect(result.content.truncated).toBe(false);
    expect(result.metadata.hasContent).toBe(true);
    expect(result.metadata.textLength).toBe(25);
  });

  it('should truncate long text', async () => {
    const longText = 'x'.repeat(100);
    officeParser.parseOffice.mockImplementation((buffer, callback) => {
      callback(longText, undefined);
    });

    const result = await parseOfficeDocument('dGVzdA==', 'long.docx', 50);

    expect(result.type).toBe('office_document');
    expect(result.content.truncated).toBe(true);
    expect(result.content.text).toHaveLength(53); // 50 + '...'
    expect(result.content.truncatedLength).toBe(50);
    expect(result.content.extractedLength).toBe(100);
  });

  it('should handle parsing errors (non-standard data,err callback)', async () => {
    officeParser.parseOffice.mockImplementation((buffer, callback) => {
      // Non-standard: (data, err) — error is second argument
      callback(undefined, 'Parsing failed: corrupt file');
    });

    const result = await parseOfficeDocument('dGVzdA==', 'corrupt.docx');

    expect(result.type).toBe('office_error');
    expect(result.error).toContain('Failed to parse office document');
    expect(result.note).toContain('corrupted');
  });

  it('should handle empty document', async () => {
    officeParser.parseOffice.mockImplementation((buffer, callback) => {
      callback('', undefined);
    });

    const result = await parseOfficeDocument('dGVzdA==', 'empty.docx');

    expect(result.type).toBe('office_document');
    expect(result.content.text).toBe('');
    expect(result.metadata.hasContent).toBe(false);
    expect(result.metadata.textLength).toBe(0);
  });

  it('should return error for unsupported types/exceptions', async () => {
    // When parseOffice throws synchronously inside the Promise constructor,
    // it becomes an unhandled rejection. Test the outer catch by passing
    // invalid base64 that causes Buffer.from to succeed but parseOffice to
    // call back with an error.
    officeParser.parseOffice.mockImplementation((buffer, callback) => {
      callback(undefined, 'Unsupported format');
    });

    const result = await parseOfficeDocument('dGVzdA==', 'bad.xyz');

    expect(result.type).toBe('office_error');
    expect(result.error).toContain('Unsupported format');
  });
});

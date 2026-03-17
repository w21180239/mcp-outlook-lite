import { describe, it, expect, vi } from 'vitest';
import { Buffer } from 'buffer';

// Mock the parser modules before importing fileTypeUtils
vi.mock('../../../tools/common/excelParser.js', () => ({
  parseExcelContent: vi.fn(),
}));

vi.mock('../../../tools/common/documentParser.js', () => ({
  parseOfficeDocument: vi.fn(),
}));

const { parseExcelContent } = await import('../../../tools/common/excelParser.js');
const { parseOfficeDocument } = await import('../../../tools/common/documentParser.js');
const { decodeContent } = await import('../../../tools/common/fileTypeUtils.js');

describe('decodeContent error branches', () => {
  it('should return error when parseExcelContent returns excel_error', async () => {
    parseExcelContent.mockReturnValue({
      type: 'excel_error',
      error: 'Corrupted workbook',
    });

    const base64 = Buffer.from('fake-excel-data').toString('base64');
    const result = await decodeContent(
      base64,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'broken.xlsx'
    );

    expect(result.type).toBe('error');
    expect(result.content).toContain('Excel parse failed');
    expect(result.error).toBe('Corrupted workbook');
    expect(result.encoding).toBe('base64_fallback');
    expect(result.contentBytes).toBe(base64);
  });

  it('should return error when parseOfficeDocument returns office_error', async () => {
    parseOfficeDocument.mockResolvedValue({
      type: 'office_error',
      error: 'Unsupported format',
    });

    const base64 = Buffer.from('fake-pdf-data').toString('base64');
    const result = await decodeContent(base64, 'application/pdf', 'broken.pdf');

    expect(result.type).toBe('error');
    expect(result.content).toContain('Office parse failed');
    expect(result.error).toBe('Unsupported format');
    expect(result.encoding).toBe('base64_fallback');
    expect(result.contentBytes).toBe(base64);
  });
});

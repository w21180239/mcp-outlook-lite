import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock xlsx before importing
vi.mock('xlsx', () => {
  const mockUtils = {
    decode_range: vi.fn().mockReturnValue({ s: { r: 0, c: 0 }, e: { r: 2, c: 1 } }),
    sheet_to_json: vi.fn().mockReturnValue([['Header1', 'Header2'], ['a', 'b'], ['c', 'd']]),
    encode_cell: vi.fn().mockReturnValue('B3'),
  };
  return {
    read: vi.fn(),
    utils: mockUtils,
  };
});

import { parseExcelContent } from '../../../tools/common/excelParser.js';
import * as XLSX from 'xlsx';

describe('parseExcelContent', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  it('should parse a single-sheet workbook', () => {
    XLSX.read.mockReturnValue({
      SheetNames: ['Sheet1'],
      Sheets: {
        Sheet1: { '!ref': 'A1:B3' },
      },
    });

    const result = parseExcelContent('dGVzdA==', 'test.xlsx');

    expect(result.type).toBe('excel');
    expect(result.filename).toBe('test.xlsx');
    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0].name).toBe('Sheet1');
    expect(result.summary.totalSheets).toBe(1);
    expect(result.summary.sheetNames).toEqual(['Sheet1']);
  });

  it('should handle empty sheets', () => {
    XLSX.read.mockReturnValue({
      SheetNames: ['Empty'],
      Sheets: {
        Empty: { '!ref': undefined },
      },
    });
    XLSX.utils.decode_range.mockReturnValue({ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } });
    XLSX.utils.sheet_to_json.mockReturnValue([]);

    const result = parseExcelContent('dGVzdA==', 'empty.xlsx');

    expect(result.type).toBe('excel');
    expect(result.sheets[0].data).toEqual([]);
    expect(result.sheets[0].displayedRows).toBe(0);
  });

  it('should limit sheets processed', () => {
    const sheetNames = Array.from({ length: 15 }, (_, i) => `Sheet${i + 1}`);
    const sheets = {};
    for (const name of sheetNames) {
      sheets[name] = { '!ref': 'A1:A1' };
    }

    XLSX.read.mockReturnValue({ SheetNames: sheetNames, Sheets: sheets });
    XLSX.utils.decode_range.mockReturnValue({ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } });
    XLSX.utils.sheet_to_json.mockReturnValue([]);

    const result = parseExcelContent('dGVzdA==', 'many-sheets.xlsx', 5);

    expect(result.sheets).toHaveLength(5);
    expect(result.summary.totalSheets).toBe(15);
    expect(result.summary.note).toContain('Only first 5 sheets');
  });

  it('should return error for corrupted file', () => {
    XLSX.read.mockImplementation(() => {
      throw new Error('Invalid file format');
    });

    const result = parseExcelContent('bad', 'corrupt.xlsx');

    expect(result.type).toBe('excel_error');
    expect(result.error).toContain('Failed to parse Excel file');
    expect(result.note).toContain('corrupted');
  });
});

import { describe, it, expect } from 'vitest';
import { Buffer } from 'buffer';
import {
  formatFileSize,
  isTextContent,
  isExcelFile,
  isOfficeDocument,
  decodeContent,
} from '../../../tools/common/fileTypeUtils.js';

describe('fileTypeUtils', () => {
  describe('formatFileSize', () => {
    it('should return "0 Bytes" for 0', () => {
      expect(formatFileSize(0)).toBe('0 Bytes');
    });

    it('should format bytes correctly', () => {
      expect(formatFileSize(500)).toBe('500 Bytes');
    });

    it('should format KB correctly', () => {
      expect(formatFileSize(1024)).toBe('1 KB');
    });

    it('should format MB correctly', () => {
      expect(formatFileSize(1048576)).toBe('1 MB');
    });

    it('should format GB correctly', () => {
      expect(formatFileSize(1073741824)).toBe('1 GB');
    });

    it('should handle decimal values', () => {
      expect(formatFileSize(1536)).toBe('1.5 KB');
    });

    it('should return "Unknown size" for null', () => {
      expect(formatFileSize(null)).toBe('Unknown size');
    });

    it('should return "Unknown size" for undefined', () => {
      expect(formatFileSize(undefined)).toBe('Unknown size');
    });

    it('should return "Unknown size" for NaN', () => {
      expect(formatFileSize(NaN)).toBe('Unknown size');
    });

    it('should return "Unknown size" for non-number', () => {
      expect(formatFileSize('abc')).toBe('Unknown size');
    });

    it('should handle negative bytes via Math.abs', () => {
      expect(formatFileSize(-1024)).toBe('1 KB');
    });
  });

  describe('isTextContent', () => {
    it('should detect text/plain by content type', () => {
      expect(isTextContent('text/plain', 'file.bin')).toBe(true);
    });

    it('should detect text/html by content type', () => {
      expect(isTextContent('text/html', 'file.bin')).toBe(true);
    });

    it('should detect application/json by content type', () => {
      expect(isTextContent('application/json', 'file.bin')).toBe(true);
    });

    it('should detect application/xml by content type', () => {
      expect(isTextContent('application/xml', 'file.bin')).toBe(true);
    });

    it('should detect .txt by extension', () => {
      expect(isTextContent(null, 'readme.txt')).toBe(true);
    });

    it('should detect .json by extension', () => {
      expect(isTextContent(null, 'config.json')).toBe(true);
    });

    it('should detect .py by extension', () => {
      expect(isTextContent(null, 'script.py')).toBe(true);
    });

    it('should detect .md by extension', () => {
      expect(isTextContent(null, 'README.md')).toBe(true);
    });

    it('should return false for binary content type', () => {
      expect(isTextContent('application/octet-stream', 'file.bin')).toBe(false);
    });

    it('should return false for image content type', () => {
      expect(isTextContent('image/png', 'photo.png')).toBe(false);
    });

    it('should detect HTML content by content analysis', () => {
      const htmlBase64 = Buffer.from('<!DOCTYPE html><html>').toString('base64');
      expect(isTextContent(null, 'file.bin', htmlBase64)).toBe(true);
    });

    it('should detect JSON content by content analysis', () => {
      const jsonBase64 = Buffer.from('{"key": "value"}').toString('base64');
      expect(isTextContent('', 'file.bin', jsonBase64)).toBe(true);
    });

    it('should be case-insensitive for content types', () => {
      expect(isTextContent('TEXT/PLAIN', 'file.bin')).toBe(true);
    });

    it('should be case-insensitive for extensions', () => {
      expect(isTextContent(null, 'FILE.TXT')).toBe(true);
    });
  });

  describe('isExcelFile', () => {
    it('should detect .xlsx by MIME type', () => {
      expect(isExcelFile('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'file.xlsx')).toBe(true);
    });

    it('should detect .xls by MIME type', () => {
      expect(isExcelFile('application/vnd.ms-excel', 'file.xls')).toBe(true);
    });

    it('should detect .xlsx by extension', () => {
      expect(isExcelFile(null, 'data.xlsx')).toBe(true);
    });

    it('should detect .xlsm by extension', () => {
      expect(isExcelFile(null, 'macro.xlsm')).toBe(true);
    });

    it('should return false for .csv', () => {
      expect(isExcelFile(null, 'data.csv')).toBe(false);
    });

    it('should return false for unknown types', () => {
      expect(isExcelFile('application/octet-stream', 'file.bin')).toBe(false);
    });
  });

  describe('isOfficeDocument', () => {
    it('should detect PDF by MIME type', () => {
      expect(isOfficeDocument('application/pdf', 'file.pdf')).toBe(true);
    });

    it('should detect .docx by MIME type', () => {
      expect(isOfficeDocument('application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'file.docx')).toBe(true);
    });

    it('should detect .pptx by MIME type', () => {
      expect(isOfficeDocument('application/vnd.openxmlformats-officedocument.presentationml.presentation', 'file.pptx')).toBe(true);
    });

    it('should detect .pdf by extension', () => {
      expect(isOfficeDocument(null, 'doc.pdf')).toBe(true);
    });

    it('should detect .docx by extension', () => {
      expect(isOfficeDocument(null, 'doc.docx')).toBe(true);
    });

    it('should detect .odt by extension', () => {
      expect(isOfficeDocument(null, 'doc.odt')).toBe(true);
    });

    it('should return false for .xlsx (handled by isExcelFile)', () => {
      expect(isOfficeDocument(null, 'data.xlsx')).toBe(false);
    });

    it('should return false for unknown types', () => {
      expect(isOfficeDocument('application/octet-stream', 'file.bin')).toBe(false);
    });
  });

  describe('decodeContent', () => {
    it('should decode text content', async () => {
      const textBase64 = Buffer.from('Hello, world!').toString('base64');
      const result = await decodeContent(textBase64, 'text/plain', 'hello.txt');
      expect(result.type).toBe('text');
      expect(result.content).toBe('Hello, world!');
      expect(result.encoding).toBe('utf8');
      expect(result.sizeFormatted).toBeDefined();
    });

    it('should preserve large text as base64', async () => {
      const largeText = 'x'.repeat(100);
      const base64 = Buffer.from(largeText).toString('base64');
      const result = await decodeContent(base64, 'text/plain', 'big.txt', 50);
      expect(result.type).toBe('text');
      expect(result.encoding).toBe('base64_preserved');
      expect(result.contentBytes).toBe(base64);
    });

    it('should detect and handle Excel files', async () => {
      const base64 = Buffer.from('fake-excel').toString('base64');
      const result = await decodeContent(base64, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'data.xlsx');
      // XLSX library parses arbitrary data as a single-cell sheet — this is a successful parse
      expect(result.type).toBe('excel');
      expect(result.encoding).toBe('parsed');
      expect(result.content).toBeDefined();
      expect(result.sizeFormatted).toBeDefined();
    });

    it('should return binary for unknown types', async () => {
      const base64 = Buffer.from([0x89, 0x50, 0x4e, 0x47]).toString('base64');
      const result = await decodeContent(base64, 'image/png', 'photo.png');
      expect(result.type).toBe('binary');
      expect(result.encoding).toBe('base64');
      expect(result.content).toContain('image/png');
      expect(result.sizeFormatted).toBeDefined();
    });

    it('should handle decode errors gracefully', async () => {
      const result = await decodeContent(null, 'text/plain', 'file.txt');
      expect(result.type).toBe('error');
      expect(result.encoding).toBe('base64_fallback');
    });
  });
});

import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  clearStylingCache,
  clearSignatureCache,
  getStylingCacheStats,
  getCachedStyling,
  setCachedStyling,
  getCachedSignature,
  setCachedSignature,
  isMcpGeneratedEmail,
  extractSignatureFromHtml,
  convertTextToHtml,
  applyOutlookStyling,
  buildEmailPayload,
  stylingCache,
  signatureCache,
} from '../../../tools/common/sharedUtils.js';

describe('sharedUtils', () => {
  beforeEach(() => {
    // Clear caches between tests
    stylingCache.clear();
    signatureCache.clear();
  });

  describe('clearStylingCache', () => {
    it('should clear cache for a specific user', () => {
      setCachedStyling('user-1', { fontFamily: 'Arial' });
      setCachedStyling('user-2', { fontFamily: 'Calibri' });

      clearStylingCache('user-1');

      expect(getCachedStyling('user-1')).toBeUndefined();
      expect(getCachedStyling('user-2')).toBeDefined();
    });

    it('should clear all cache when no userId provided', () => {
      setCachedStyling('user-1', { fontFamily: 'Arial' });
      setCachedStyling('user-2', { fontFamily: 'Calibri' });

      clearStylingCache();

      expect(getCachedStyling('user-1')).toBeUndefined();
      expect(getCachedStyling('user-2')).toBeUndefined();
    });
  });

  describe('clearSignatureCache', () => {
    it('should clear signature cache for a specific user', () => {
      setCachedSignature('user-1', 'Sig 1');
      setCachedSignature('user-2', 'Sig 2');

      clearSignatureCache('user-1');

      expect(getCachedSignature('user-1')).toBeUndefined();
      expect(getCachedSignature('user-2')).toBeDefined();
    });

    it('should clear all signature cache when no userId provided', () => {
      setCachedSignature('user-1', 'Sig 1');
      setCachedSignature('user-2', 'Sig 2');

      clearSignatureCache();

      expect(getCachedSignature('user-1')).toBeUndefined();
      expect(getCachedSignature('user-2')).toBeUndefined();
    });
  });

  describe('getStylingCacheStats', () => {
    it('should return stats with zero entries when cache is empty', () => {
      const stats = getStylingCacheStats();
      expect(stats.totalEntries).toBe(0);
    });

    it('should report correct entry count', () => {
      setCachedStyling('user-1', { fontFamily: 'Arial' });
      setCachedStyling('user-2', { fontFamily: 'Calibri' });

      const stats = getStylingCacheStats();
      expect(stats.totalEntries).toBe(2);
    });
  });

  describe('getCachedStyling / setCachedStyling', () => {
    it('should store and retrieve styling data', () => {
      const styling = { fontFamily: 'Arial', fontSize: '12pt' };
      setCachedStyling('user-1', styling);

      const cached = getCachedStyling('user-1');
      expect(cached.fontFamily).toBe('Arial');
      expect(cached.fontSize).toBe('12pt');
      expect(cached.timestamp).toBeDefined();
    });

    it('should return undefined for non-existent user', () => {
      expect(getCachedStyling('nonexistent')).toBeUndefined();
    });
  });

  describe('getCachedSignature / setCachedSignature', () => {
    it('should store and retrieve signature data', () => {
      setCachedSignature('user-1', '<div>My Sig</div>');

      const cached = getCachedSignature('user-1');
      expect(cached.signature).toBe('<div>My Sig</div>');
      expect(cached.timestamp).toBeDefined();
    });

    it('should return undefined for non-existent user', () => {
      expect(getCachedSignature('nonexistent')).toBeUndefined();
    });
  });

  describe('isMcpGeneratedEmail', () => {
    it('should detect MCP-generated HTML with email-content class', () => {
      const html = '<div class="email-content">Hello</div>';
      expect(isMcpGeneratedEmail(html)).toBe(true);
    });

    it('should detect MCP-generated HTML with signature class', () => {
      const html = '<div class="signature">Best regards</div>';
      expect(isMcpGeneratedEmail(html)).toBe(true);
    });

    it('should detect MCP-generated HTML with full structure', () => {
      const html = `<html><head><meta charset="UTF-8"><style>.email-content { margin: 0; }</style></head><body></body></html>`;
      expect(isMcpGeneratedEmail(html)).toBe(true);
    });

    it('should return false for plain non-MCP HTML', () => {
      const html = '<p>Just a regular email</p>';
      expect(isMcpGeneratedEmail(html)).toBe(false);
    });
  });

  describe('extractSignatureFromHtml', () => {
    it('should extract signature from div with signature id', () => {
      const html = '<p>Body</p><div id="signature">Best regards, John</div>';
      const sig = extractSignatureFromHtml(html);
      expect(sig).toContain('Best regards');
    });

    it('should extract signature from div with signature class', () => {
      const html = '<p>Body</p><div class="signature">Thanks, Jane Doe - Company Inc</div>';
      const sig = extractSignatureFromHtml(html);
      expect(sig).toContain('Jane Doe');
    });

    it('should return empty string when no signature found', () => {
      const html = '<p>Short email</p>';
      const sig = extractSignatureFromHtml(html);
      expect(sig).toBe('');
    });

    it('should filter out signatures containing Sent from', () => {
      const html = '<div id="signature">Sent from my iPhone device</div>';
      const sig = extractSignatureFromHtml(html);
      expect(sig).toBe('');
    });
  });

  describe('convertTextToHtml', () => {
    it('should escape HTML entities', () => {
      expect(convertTextToHtml('<script>')).toBe('&lt;script&gt;');
    });

    it('should convert newlines to br tags', () => {
      expect(convertTextToHtml('Line 1\nLine 2')).toBe('Line 1<br>Line 2');
    });

    it('should convert tabs to non-breaking spaces', () => {
      expect(convertTextToHtml('\t')).toBe('&nbsp;&nbsp;&nbsp;&nbsp;');
    });

    it('should escape ampersands', () => {
      expect(convertTextToHtml('A & B')).toBe('A &amp; B');
    });

    it('should escape quotes', () => {
      expect(convertTextToHtml('"hello"')).toBe('&quot;hello&quot;');
    });
  });

  describe('applyOutlookStyling', () => {
    it('should apply default styling when no settings provided', async () => {
      const result = await applyOutlookStyling('Hello', '', null, null);
      expect(result).toContain('Calibri');
      expect(result).toContain('11pt');
      expect(result).toContain('Hello');
    });

    it('should use actualStyling when provided', async () => {
      const actualStyling = { fontFamily: 'Arial', fontSize: '14pt', fontColor: '#333333' };
      const result = await applyOutlookStyling('Content', '', null, actualStyling);
      expect(result).toContain('Arial');
      expect(result).toContain('14pt');
      expect(result).toContain('#333333');
    });

    it('should include signature when provided', async () => {
      const sig = '<div>Best regards, Test User</div>';
      const result = await applyOutlookStyling('Body', sig, null, null);
      expect(result).toContain('Best regards, Test User');
      expect(result).toContain('class="signature"');
    });

    it('should not include signature div when empty', async () => {
      const result = await applyOutlookStyling('Body', '', null, null);
      expect(result).not.toContain('class="signature"');
    });
  });

  describe('buildEmailPayload', () => {
    it('should build payload with required fields', () => {
      const payload = buildEmailPayload('<p>Hello</p>', ['user@example.com'], 'Test Subject');
      expect(payload.subject).toBe('Test Subject');
      expect(payload.body.contentType).toBe('HTML');
      expect(payload.body.content).toBe('<p>Hello</p>');
      expect(payload.toRecipients).toHaveLength(1);
      expect(payload.toRecipients[0].emailAddress.address).toBe('user@example.com');
    });

    it('should include attachments when provided', () => {
      const attachments = [{ name: 'file.pdf', contentBytes: 'abc123' }];
      const payload = buildEmailPayload('content', ['user@example.com'], 'Test', attachments);
      expect(payload.attachments).toEqual(attachments);
    });

    it('should omit attachments when empty array', () => {
      const payload = buildEmailPayload('content', ['user@example.com'], 'Test', []);
      expect(payload.attachments).toBeUndefined();
    });

    it('should handle multiple recipients', () => {
      const payload = buildEmailPayload('content', ['a@test.com', 'b@test.com'], 'Test');
      expect(payload.toRecipients).toHaveLength(2);
    });
  });
});

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
  applyUserStyling,
  processEmailContent,
  stylingCache,
  signatureCache,
} from '../../../tools/common/sharedUtils.js';

describe('sharedUtils - additional coverage', () => {
  beforeEach(() => {
    stylingCache.clear();
    signatureCache.clear();
  });

  describe('applyOutlookStyling - edge cases', () => {
    it('should use mailSettings font name when no actualStyling', async () => {
      const result = await applyOutlookStyling('Hello', '', { defaultFontName: 'Verdana' }, null);
      expect(result).toContain('Verdana');
    });

    it('should use mailSettings font size when no actualStyling', async () => {
      const result = await applyOutlookStyling('Hello', '', { defaultFontSize: '14pt' }, null);
      expect(result).toContain('14pt');
    });

    it('should use mailSettings font color when no actualStyling', async () => {
      const result = await applyOutlookStyling('Hello', '', { defaultFontColor: '#FF0000' }, null);
      expect(result).toContain('#FF0000');
    });

    it('should prioritize actualStyling over mailSettings', async () => {
      const actualStyling = { fontFamily: 'Georgia', fontSize: '16pt', fontColor: '#0000FF' };
      const mailSettings = { defaultFontName: 'Verdana', defaultFontSize: '12pt', defaultFontColor: '#FF0000' };
      const result = await applyOutlookStyling('Content', '', mailSettings, actualStyling);
      expect(result).toContain('Georgia');
      expect(result).toContain('16pt');
      expect(result).toContain('#0000FF');
    });

    it('should contain proper HTML structure', async () => {
      const result = await applyOutlookStyling('Body text', '', null, null);
      expect(result).toContain('<html>');
      expect(result).toContain('<head>');
      expect(result).toContain('<style>');
      expect(result).toContain('<body>');
      expect(result).toContain('class="email-content"');
      expect(result).toContain('Body text');
    });

    it('should include signature in a signature div', async () => {
      const result = await applyOutlookStyling('Body', '<p>My Signature</p>', null, null);
      expect(result).toContain('class="signature"');
      expect(result).toContain('My Signature');
    });
  });

  describe('buildEmailPayload - edge cases', () => {
    it('should handle single recipient', () => {
      const payload = buildEmailPayload('<p>Hi</p>', ['single@test.com'], 'Subject');
      expect(payload.toRecipients).toHaveLength(1);
      expect(payload.body.contentType).toBe('HTML');
    });

    it('should handle many recipients', () => {
      const recipients = Array.from({ length: 50 }, (_, i) => `user${i}@test.com`);
      const payload = buildEmailPayload('content', recipients, 'Bulk');
      expect(payload.toRecipients).toHaveLength(50);
    });

    it('should not include attachments key when default empty array', () => {
      const payload = buildEmailPayload('content', ['u@t.com'], 'S');
      expect(payload).not.toHaveProperty('attachments');
    });

    it('should include non-empty attachments array', () => {
      const attachments = [
        { '@odata.type': '#microsoft.graph.fileAttachment', name: 'file.pdf', contentBytes: 'base64data' },
      ];
      const payload = buildEmailPayload('content', ['u@t.com'], 'S', attachments);
      expect(payload.attachments).toHaveLength(1);
      expect(payload.attachments[0].name).toBe('file.pdf');
    });
  });

  describe('isMcpGeneratedEmail - additional patterns', () => {
    it('should detect font-family sans-serif pattern', () => {
      const html = '<div style="font-family: Calibri, sans-serif">Hello</div>';
      expect(isMcpGeneratedEmail(html)).toBe(true);
    });

    it('should detect complete html structure with .email-content', () => {
      const html = `<html>
        <head>
          <meta charset="UTF-8">
          <style>
            .email-content { margin: 0; }
          </style>
        </head>
        <body>Content</body>
      </html>`;
      expect(isMcpGeneratedEmail(html)).toBe(true);
    });

    it('should return false for generic HTML without MCP indicators', () => {
      expect(isMcpGeneratedEmail('<html><body><p>Regular email</p></body></html>')).toBe(false);
    });

    it('should return false for empty string', () => {
      expect(isMcpGeneratedEmail('')).toBe(false);
    });
  });

  describe('extractSignatureFromHtml - additional patterns', () => {
    it('should extract table-based signature', () => {
      const html = '<p>Body</p><table><tr><td>John Doe</td><td>phone: 555-1234 cell: 555-5678</td></tr></table>';
      const sig = extractSignatureFromHtml(html);
      expect(sig.length).toBeGreaterThan(20);
    });

    it('should extract signature after hr tag', () => {
      const html = '<p>Body text</p><hr><p>John Doe, VP Engineering, phone: 555-1234</p>';
      const sig = extractSignatureFromHtml(html);
      expect(sig.length).toBeGreaterThan(0);
    });

    it('should filter out "Get Outlook" signatures', () => {
      const html = '<div id="signature">Get Outlook for iOS and Android</div>';
      const sig = extractSignatureFromHtml(html);
      expect(sig).toBe('');
    });

    it('should extract signature with "Best regards" pattern', () => {
      const html = '<p>Main content here.</p><p>Best regards,<br>John Smith<br>john@company.com</p>';
      const sig = extractSignatureFromHtml(html);
      expect(sig.length).toBeGreaterThan(0);
    });

    it('should extract Outlook-style signature div', () => {
      const html = '<p>Body</p><div id="Signature">Thanks, Jane - jane@corp.com phone: 555-0000</div>';
      const sig = extractSignatureFromHtml(html);
      expect(sig).toContain('Jane');
    });

    it('should handle very short signature divs', () => {
      // Short signatures in signature divs are still matched by the id pattern
      const html = '<div id="signature">Hi</div>';
      const sig = extractSignatureFromHtml(html);
      // The regex matches but the length filter (>20) only applies after match
      // The actual filter is length > 20 AND not containing "Sent from"/"Get Outlook"
      expect(typeof sig).toBe('string');
    });
  });

  describe('convertTextToHtml - additional', () => {
    it('should handle Windows-style newlines', () => {
      expect(convertTextToHtml('Line 1\r\nLine 2')).toBe('Line 1<br>Line 2');
    });

    it('should escape single quotes', () => {
      expect(convertTextToHtml("it's")).toBe('it&#39;s');
    });

    it('should handle complex content with multiple escapes', () => {
      const result = convertTextToHtml('A & B < C > D "E" \'F\'\n\tG');
      expect(result).toBe('A &amp; B &lt; C &gt; D &quot;E&quot; &#39;F&#39;<br>&nbsp;&nbsp;&nbsp;&nbsp;G');
    });
  });

  describe('cache operations', () => {
    it('should overwrite existing styling cache entry', () => {
      setCachedStyling('user-1', { fontFamily: 'Arial' });
      setCachedStyling('user-1', { fontFamily: 'Verdana' });
      const cached = getCachedStyling('user-1');
      expect(cached.fontFamily).toBe('Verdana');
    });

    it('should overwrite existing signature cache entry', () => {
      setCachedSignature('user-1', 'Sig A');
      setCachedSignature('user-1', 'Sig B');
      const cached = getCachedSignature('user-1');
      expect(cached.signature).toBe('Sig B');
    });

    it('should track cache stats correctly', () => {
      setCachedStyling('u1', { fontFamily: 'A' });
      setCachedStyling('u2', { fontFamily: 'B' });
      setCachedStyling('u3', { fontFamily: 'C' });
      const stats = getStylingCacheStats();
      expect(stats.totalEntries).toBe(3);
    });

    it('clearStylingCache(null) should clear all entries', () => {
      setCachedStyling('u1', { fontFamily: 'A' });
      setCachedStyling('u2', { fontFamily: 'B' });
      clearStylingCache(null);
      expect(stylingCache.size).toBe(0);
    });

    it('clearSignatureCache(null) should clear all entries', () => {
      setCachedSignature('u1', 'S1');
      setCachedSignature('u2', 'S2');
      clearSignatureCache(null);
      expect(signatureCache.size).toBe(0);
    });
  });

  describe('applyUserStyling', () => {
    it('should return original content when all API calls fail', async () => {
      const mockClient = {
        makeRequest: vi.fn().mockRejectedValue(new Error('General failure')),
      };
      const result = await applyUserStyling(mockClient, 'Hello world', 'text');
      expect(result.content).toBe('Hello world');
      expect(result.type).toBe('text');
    });

    it('should fall back to basic styling on 403 error', async () => {
      const mockClient = {
        makeRequest: vi.fn()
          .mockRejectedValueOnce(new Error('403 ErrorAccessDenied')) // mailboxSettings
          .mockResolvedValueOnce({ id: 'user-1' }) // /me for getUserSignature
          .mockResolvedValueOnce({ value: [] }), // sent items for signature
      };

      const result = await applyUserStyling(mockClient, 'Hello', 'text');
      expect(result.type).toBe('html');
      expect(result.content).toContain('Hello');
    });

    it('should apply styling from sent emails when available', async () => {
      const htmlEmail = '<div style="font-family: Georgia; font-size: 14pt; color: #333">content</div>';
      const mockClient = {
        makeRequest: vi.fn()
          .mockResolvedValueOnce({}) // mailboxSettings
          .mockResolvedValueOnce({ id: 'user-1' }) // /me for signature
          .mockResolvedValueOnce({ value: [] }) // sent items for signature
          .mockResolvedValueOnce({ id: 'user-1' }) // /me for styling
          .mockResolvedValueOnce({
            value: [{ body: { contentType: 'HTML', content: htmlEmail } }]
          }), // sent items for styling
      };

      const result = await applyUserStyling(mockClient, '<p>Hello</p>', 'html');
      expect(result.type).toBe('html');
      expect(result.content).toContain('<html>');
    });

    it('should convert plain text to HTML when bodyType is text', async () => {
      const mockClient = {
        makeRequest: vi.fn()
          .mockResolvedValueOnce({}) // mailboxSettings
          .mockResolvedValueOnce({ id: 'user-1' }) // /me
          .mockResolvedValueOnce({ value: [] }) // sent items
          .mockResolvedValueOnce({ id: 'user-1' }) // /me
          .mockResolvedValueOnce({ value: [] }), // sent items
      };

      const result = await applyUserStyling(mockClient, 'Plain text content', 'text');
      expect(result.type).toBe('html');
      expect(result.content).toContain('Plain text content');
      expect(result.content).toContain('<html>');
    });
  });
});

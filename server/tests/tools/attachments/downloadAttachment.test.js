import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { downloadAttachmentTool } from '../../../tools/attachments/downloadAttachment.js';

describe('downloadAttachmentTool', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  it('should return validation error when messageId is missing', async () => {
    const result = await downloadAttachmentTool(authManager, { attachmentId: 'att-1' });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('messageId');
  });

  it('should return validation error when attachmentId is missing', async () => {
    const result = await downloadAttachmentTool(authManager, { messageId: 'msg-1' });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('attachmentId');
  });

  it('should return attachment metadata without content when includeContent is false', async () => {
    graphApiClient.makeRequest.mockResolvedValue({
      id: 'att-1',
      name: 'report.pdf',
      contentType: 'application/pdf',
      size: 1024,
      isInline: false,
      lastModifiedDateTime: '2024-01-01T00:00:00Z',
      '@odata.type': '#microsoft.graph.fileAttachment',
    });

    const result = await downloadAttachmentTool(authManager, {
      messageId: 'msg-1',
      attachmentId: 'att-1',
      includeContent: false,
    });

    expect(result.isError).toBeUndefined();
    const data = JSON.parse(result.content[0].text);
    expect(data.contentIncluded).toBe(false);
    expect(data.name).toBe('report.pdf');
    expect(data.content).toBeUndefined();
  });

  it('should download and decode file attachment content', async () => {
    const textContent = 'Hello, world!';
    const base64Content = Buffer.from(textContent).toString('base64');

    // First call: metadata
    graphApiClient.makeRequest
      .mockResolvedValueOnce({
        id: 'att-1',
        name: 'hello.txt',
        contentType: 'text/plain',
        size: 13,
        isInline: false,
        lastModifiedDateTime: '2024-01-01T00:00:00Z',
        '@odata.type': '#microsoft.graph.fileAttachment',
      })
      // Second call: full attachment with contentBytes
      .mockResolvedValueOnce({
        id: 'att-1',
        name: 'hello.txt',
        contentType: 'text/plain',
        size: 13,
        contentBytes: base64Content,
        '@odata.type': '#microsoft.graph.fileAttachment',
      });

    const result = await downloadAttachmentTool(authManager, {
      messageId: 'msg-1',
      attachmentId: 'att-1',
      includeContent: true,
      decodeContent: true,
    });

    expect(result.isError).toBeUndefined();
    const data = JSON.parse(result.content[0].text);
    expect(data.contentIncluded).toBe(true);
    expect(data.content).toBe(textContent);
    expect(data.encoding).toBe('utf8');
    expect(graphApiClient.makeRequest).toHaveBeenCalledTimes(2);
  });

  it('should return raw base64 when decodeContent is false', async () => {
    const base64Content = Buffer.from('raw data').toString('base64');

    graphApiClient.makeRequest
      .mockResolvedValueOnce({
        id: 'att-1',
        name: 'file.bin',
        contentType: 'application/octet-stream',
        size: 8,
        isInline: false,
        lastModifiedDateTime: '2024-01-01T00:00:00Z',
        '@odata.type': '#microsoft.graph.fileAttachment',
      })
      .mockResolvedValueOnce({
        id: 'att-1',
        name: 'file.bin',
        contentBytes: base64Content,
        '@odata.type': '#microsoft.graph.fileAttachment',
      });

    const result = await downloadAttachmentTool(authManager, {
      messageId: 'msg-1',
      attachmentId: 'att-1',
      includeContent: true,
      decodeContent: false,
    });

    const data = JSON.parse(result.content[0].text);
    expect(data.contentIncluded).toBe(true);
    expect(data.contentBytes).toBe(base64Content);
    expect(data.encoding).toBe('base64');
    expect(data.note).toContain('Raw Base64');
  });

  it('should handle item attachment with embedded content', async () => {
    graphApiClient.makeRequest
      .mockResolvedValueOnce({
        id: 'att-1',
        name: 'Embedded Message',
        contentType: 'message/rfc822',
        size: 500,
        isInline: false,
        lastModifiedDateTime: '2024-01-01T00:00:00Z',
        '@odata.type': '#microsoft.graph.itemAttachment',
      })
      .mockResolvedValueOnce({
        id: 'att-1',
        name: 'Embedded Message',
        item: { subject: 'Test', body: { content: 'Hello' } },
        '@odata.type': '#microsoft.graph.itemAttachment',
      });

    const result = await downloadAttachmentTool(authManager, {
      messageId: 'msg-1',
      attachmentId: 'att-1',
      includeContent: true,
    });

    const data = JSON.parse(result.content[0].text);
    expect(data.contentIncluded).toBe(true);
    expect(data.itemContent.subject).toBe('Test');
    expect(data.encoding).toBe('json');
  });

  it('should handle reference attachment with sourceUrl', async () => {
    graphApiClient.makeRequest
      .mockResolvedValueOnce({
        id: 'att-1',
        name: 'Shared Document.docx',
        contentType: null,
        size: 0,
        isInline: false,
        lastModifiedDateTime: '2024-01-01T00:00:00Z',
        '@odata.type': '#microsoft.graph.referenceAttachment',
      })
      .mockResolvedValueOnce({
        id: 'att-1',
        sourceUrl: 'https://sharepoint.com/file.docx',
        providerType: 'oneDriveBusiness',
        thumbnailUrl: null,
        previewUrl: null,
        permission: 'view',
        isFolder: false,
        '@odata.type': '#microsoft.graph.referenceAttachment',
      });

    const result = await downloadAttachmentTool(authManager, {
      messageId: 'msg-1',
      attachmentId: 'att-1',
      includeContent: true,
    });

    const data = JSON.parse(result.content[0].text);
    expect(data.contentIncluded).toBe(false);
    expect(data.sourceUrl).toBe('https://sharepoint.com/file.docx');
    expect(data.providerType).toBe('oneDriveBusiness');
  });

  it('should return error when makeRequest throws', async () => {
    authManager.ensureAuthenticated.mockResolvedValue(undefined);
    graphApiClient.makeRequest.mockRejectedValue(new Error('Network failure'));

    const result = await downloadAttachmentTool(authManager, {
      messageId: 'msg-1',
      attachmentId: 'att-1',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Network failure');
  });
});

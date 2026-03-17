import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { listAttachmentsTool } from '../../../tools/attachments/listAttachments.js';
import { addAttachmentTool } from '../../../tools/attachments/addAttachment.js';
import { scanAttachmentsTool } from '../../../tools/attachments/scanAttachments.js';

describe('Attachment Tools', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  describe('listAttachmentsTool', () => {
    it('should list attachments for a message', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [
          {
            id: 'att-1',
            name: 'file.pdf',
            contentType: 'application/pdf',
            size: 1024,
            isInline: false,
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
          },
        ],
      });

      const result = await listAttachmentsTool(authManager, { messageId: 'msg-1' });

      expect(result.content[0].type).toBe('text');
      const data = JSON.parse(result.content[0].text);
      expect(data.totalAttachments).toBe(1);
      expect(data.attachments[0].name).toBe('file.pdf');
    });

    it('should return error when messageId is missing', async () => {
      const result = await listAttachmentsTool(authManager, {});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('messageId');
    });
  });

  describe('addAttachmentTool', () => {
    it('should add attachment to message', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'att-new' });

      const result = await addAttachmentTool(authManager, {
        messageId: 'msg-1',
        name: 'test.txt',
        contentType: 'text/plain',
        contentBytes: 'dGVzdA==',
      });

      expect(result.content[0].text).toContain('added successfully');
      expect(result.content[0].text).toContain('test.txt');
    });

    it('should return error when name is missing', async () => {
      const result = await addAttachmentTool(authManager, {
        messageId: 'msg-1',
        contentType: 'text/plain',
        contentBytes: 'dGVzdA==',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('name');
    });

    it('should return error when contentBytes is missing', async () => {
      const result = await addAttachmentTool(authManager, {
        messageId: 'msg-1',
        name: 'test.txt',
        contentType: 'text/plain',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('contentBytes');
    });
  });

  describe('scanAttachmentsTool', () => {
    it('should scan attachments and report suspicious files', async () => {
      graphApiClient.makeRequest
        .mockResolvedValueOnce({
          value: [
            {
              id: 'msg-1',
              subject: 'Test',
              from: { emailAddress: { address: 'a@b.com' } },
              receivedDateTime: '2024-01-01T00:00:00Z',
              hasAttachments: true,
            },
          ],
        })
        .mockResolvedValueOnce({
          value: [
            {
              id: 'att-1',
              name: 'malware.exe',
              contentType: 'application/x-msdownload',
              size: 5000,
              isInline: false,
            },
          ],
        });

      const result = await scanAttachmentsTool(authManager, {});

      expect(result.content[0].type).toBe('text');
      const data = JSON.parse(result.content[0].text);
      expect(data.summary.suspiciousAttachments).toBe(1);
      expect(data.suspiciousAttachments[0].attachment.name).toBe('malware.exe');
    });
  });
});

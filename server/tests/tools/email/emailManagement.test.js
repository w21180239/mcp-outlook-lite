import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import {
  deleteEmailTool,
  moveEmailTool,
  markAsReadTool,
  flagEmailTool,
  categorizeEmailTool,
  archiveEmailTool,
} from '../../../tools/email/emailManagement.js';

describe('Email Management Tools', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  describe('deleteEmailTool', () => {
    it('should permanently delete an email', async () => {
      graphApiClient.makeRequest.mockResolvedValue({});

      const result = await deleteEmailTool(authManager, {
        messageId: 'msg-1',
        permanentDelete: true,
      });

      expect(result.content[0].type).toBe('text');
      expect(result.content[0].text).toContain('permanently deleted');
      expect(graphApiClient.makeRequest).toHaveBeenCalledWith(
        '/me/messages/msg-1', {}, 'DELETE'
      );
    });

    it('should soft delete (move to Deleted Items)', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [{ id: 'deleted-folder-id' }],
      });
      graphApiClient.postWithRetry.mockResolvedValue({});

      const result = await deleteEmailTool(authManager, { messageId: 'msg-1' });

      expect(result.content[0].text).toContain('Deleted Items');
    });

    it('should return error when messageId is missing', async () => {
      const result = await deleteEmailTool(authManager, {});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('messageId');
    });
  });

  describe('moveEmailTool', () => {
    it('should move email to destination folder', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'new-msg-id' });

      const result = await moveEmailTool(authManager, {
        messageId: 'msg-1',
        destinationFolderId: 'folder-1',
      });

      expect(result.content[0].text).toContain('moved successfully');
    });

    it('should return error when destinationFolderId is missing', async () => {
      const result = await moveEmailTool(authManager, { messageId: 'msg-1' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('destinationFolderId');
    });
  });

  describe('markAsReadTool', () => {
    it('should mark email as read', async () => {
      graphApiClient.makeRequest.mockResolvedValue({});

      const result = await markAsReadTool(authManager, { messageId: 'msg-1' });

      expect(result.content[0].text).toContain('marked as read');
    });

    it('should return error when messageId is missing', async () => {
      const result = await markAsReadTool(authManager, {});

      expect(result.isError).toBe(true);
    });
  });

  describe('flagEmailTool', () => {
    it('should flag email', async () => {
      graphApiClient.makeRequest.mockResolvedValue({});

      const result = await flagEmailTool(authManager, {
        messageId: 'msg-1',
        flagStatus: 'flagged',
      });

      expect(result.content[0].text).toContain('flagged');
    });

    it('should return error for invalid flag status', async () => {
      const result = await flagEmailTool(authManager, {
        messageId: 'msg-1',
        flagStatus: 'invalid',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('flagStatus');
    });
  });

  describe('categorizeEmailTool', () => {
    it('should categorize email', async () => {
      graphApiClient.makeRequest.mockResolvedValue({});

      const result = await categorizeEmailTool(authManager, {
        messageId: 'msg-1',
        categories: ['Red', 'Blue'],
      });

      expect(result.content[0].text).toContain('categories updated');
    });

    it('should return error when categories is not an array', async () => {
      const result = await categorizeEmailTool(authManager, {
        messageId: 'msg-1',
        categories: 'not-array',
      });

      expect(result.isError).toBe(true);
    });
  });

  describe('archiveEmailTool', () => {
    it('should archive email', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [{ id: 'archive-folder-id' }],
      });
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'new-msg-id' });

      const result = await archiveEmailTool(authManager, { messageId: 'msg-1' });

      expect(result.content[0].text).toContain('archived successfully');
    });

    it('should return error when messageId is missing', async () => {
      const result = await archiveEmailTool(authManager, {});

      expect(result.isError).toBe(true);
    });
  });
});

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { replyToEmailTool, replyAllTool } from '../../../tools/email/replyEmail.js';

// Mock sharedUtils to avoid actual Graph API calls for styling
vi.mock('../../../tools/common/sharedUtils.js', () => ({
  applyUserStyling: vi.fn().mockResolvedValue({ content: 'styled reply', type: 'html' }),
}));

describe('replyToEmailTool', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  it('should return validation error when messageId is missing', async () => {
    const result = await replyToEmailTool(authManager, {
      body: 'Reply text',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('messageId');
  });

  it('should return validation error when both body and comment are missing', async () => {
    const result = await replyToEmailTool(authManager, {
      messageId: 'msg-1',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('body/comment');
  });

  it('should reply to email successfully with body', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'reply-1' });

    const result = await replyToEmailTool(authManager, {
      messageId: 'msg-1',
      body: 'Thanks for the update',
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Reply draft created successfully');
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages/msg-1/createReply',
      expect.objectContaining({
        message: expect.objectContaining({
          body: expect.objectContaining({
            contentType: 'HTML',
            content: 'styled reply',
          }),
        }),
      })
    );
  });

  it('should reply to email successfully with comment', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'reply-2' });

    const result = await replyToEmailTool(authManager, {
      messageId: 'msg-1',
      comment: 'Quick comment',
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Reply draft created successfully');
  });

  it('should not apply styling when preserveUserStyling is false', async () => {
    const { applyUserStyling } = await import('../../../tools/common/sharedUtils.js');
    applyUserStyling.mockClear();

    graphApiClient.postWithRetry.mockResolvedValue({ id: 'reply-3' });

    const result = await replyToEmailTool(authManager, {
      messageId: 'msg-1',
      body: 'Plain reply',
      preserveUserStyling: false,
    });

    expect(applyUserStyling).not.toHaveBeenCalled();
    expect(result.isError).toBeUndefined();
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages/msg-1/createReply',
      expect.objectContaining({
        message: expect.objectContaining({
          body: expect.objectContaining({
            contentType: 'Text',
            content: 'Plain reply',
          }),
        }),
      })
    );
  });

  it('should handle API errors gracefully', async () => {
    graphApiClient.postWithRetry.mockRejectedValue(new Error('Server error'));

    const result = await replyToEmailTool(authManager, {
      messageId: 'msg-1',
      body: 'Reply',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Failed to reply to email');
  });
});

describe('replyAllTool', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  it('should return validation error when messageId is missing', async () => {
    const result = await replyAllTool(authManager, {
      body: 'Reply all text',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('messageId');
  });

  it('should return validation error when both body and comment are missing', async () => {
    const result = await replyAllTool(authManager, {
      messageId: 'msg-1',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('body/comment');
  });

  it('should reply all successfully', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'reply-all-1' });

    const result = await replyAllTool(authManager, {
      messageId: 'msg-1',
      body: 'Reply to all',
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Reply-all draft created successfully');
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages/msg-1/createReplyAll',
      expect.any(Object)
    );
  });

  it('should handle API errors gracefully', async () => {
    graphApiClient.postWithRetry.mockRejectedValue(new Error('Timeout'));

    const result = await replyAllTool(authManager, {
      messageId: 'msg-1',
      body: 'Reply',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Failed to reply all to email');
  });
});

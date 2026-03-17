import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { createDraftTool } from '../../../tools/email/createDraft.js';

// Mock sharedUtils to avoid actual Graph API calls for styling
vi.mock('../../../tools/common/sharedUtils.js', () => ({
  applyUserStyling: vi.fn().mockResolvedValue({ content: 'styled body', type: 'html' }),
}));

describe('createDraftTool', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  it('should return validation error when to is missing and no replyToMessageId', async () => {
    const result = await createDraftTool(authManager, {
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('to');
  });

  it('should return validation error when to is empty array and no replyToMessageId', async () => {
    const result = await createDraftTool(authManager, {
      to: [],
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('to');
  });

  it('should return validation error when subject is missing and no replyToMessageId', async () => {
    const result = await createDraftTool(authManager, {
      to: ['user@example.com'],
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('subject');
  });

  it('should create a new draft successfully', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'draft-123' });

    const result = await createDraftTool(authManager, {
      to: ['user@example.com'],
      subject: 'Test Draft',
      body: 'Hello World',
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Draft created successfully');
    expect(result.content[0].text).toContain('draft-123');
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages',
      expect.objectContaining({
        subject: 'Test Draft',
        importance: 'normal',
      })
    );
  });

  it('should create a draft with cc and bcc recipients', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'draft-456' });

    const result = await createDraftTool(authManager, {
      to: ['user@example.com'],
      cc: ['cc@example.com'],
      bcc: ['bcc@example.com'],
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Draft created successfully');
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages',
      expect.objectContaining({
        ccRecipients: [{ emailAddress: { address: 'cc@example.com' } }],
        bccRecipients: [{ emailAddress: { address: 'bcc@example.com' } }],
      })
    );
  });

  it('should create a reply draft when replyToMessageId is provided', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'reply-draft-789' });

    const result = await createDraftTool(authManager, {
      replyToMessageId: 'original-msg-1',
      body: 'Reply content',
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Draft created successfully');
    expect(result.content[0].text).toContain('reply-draft-789');
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages/original-msg-1/createReply',
      expect.objectContaining({
        message: expect.objectContaining({
          body: expect.objectContaining({
            content: 'styled body',
          }),
        }),
      })
    );
  });

  it('should skip validation for to/subject when replyToMessageId is provided', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'reply-draft' });

    const result = await createDraftTool(authManager, {
      replyToMessageId: 'msg-1',
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Draft created successfully');
  });

  it('should set importance on new drafts', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'draft-high' });

    await createDraftTool(authManager, {
      to: ['user@example.com'],
      subject: 'Urgent',
      body: 'Important',
      importance: 'high',
    });

    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages',
      expect.objectContaining({
        importance: 'high',
      })
    );
  });

  it('should handle API errors gracefully', async () => {
    graphApiClient.postWithRetry.mockRejectedValue(new Error('Network error'));

    const result = await createDraftTool(authManager, {
      to: ['user@example.com'],
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Failed to create draft');
  });

  it('should not apply styling when preserveUserStyling is false', async () => {
    const { applyUserStyling } = await import('../../../tools/common/sharedUtils.js');
    applyUserStyling.mockClear();

    graphApiClient.postWithRetry.mockResolvedValue({ id: 'draft-no-style' });

    await createDraftTool(authManager, {
      to: ['user@example.com'],
      subject: 'Test',
      body: 'Hello',
      preserveUserStyling: false,
    });

    expect(applyUserStyling).not.toHaveBeenCalled();
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages',
      expect.objectContaining({
        body: expect.objectContaining({
          contentType: 'Text',
          content: 'Hello',
        }),
      })
    );
  });
});

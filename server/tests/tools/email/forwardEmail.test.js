import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { forwardEmailTool } from '../../../tools/email/forwardEmail.js';

// Mock sharedUtils to avoid actual Graph API calls for styling
vi.mock('../../../tools/common/sharedUtils.js', () => ({
  applyUserStyling: vi.fn().mockResolvedValue({ content: 'styled forward', type: 'html' }),
}));

describe('forwardEmailTool', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  it('should return validation error when messageId is missing', async () => {
    const result = await forwardEmailTool(authManager, {
      to: ['user@example.com'],
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('messageId');
  });

  it('should return validation error when to is missing', async () => {
    const result = await forwardEmailTool(authManager, {
      messageId: 'msg-1',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('to');
  });

  it('should return validation error when to is empty array', async () => {
    const result = await forwardEmailTool(authManager, {
      messageId: 'msg-1',
      to: [],
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('to');
  });

  it('should forward email successfully', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'fwd-1' });

    const result = await forwardEmailTool(authManager, {
      messageId: 'msg-1',
      to: ['recipient@example.com'],
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Email forwarded successfully');
    expect(result.content[0].text).toContain('recipient@example.com');
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages/msg-1/forward',
      expect.objectContaining({
        toRecipients: [{ emailAddress: { address: 'recipient@example.com' } }],
      })
    );
  });

  it('should forward email with body text', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'fwd-2' });

    const result = await forwardEmailTool(authManager, {
      messageId: 'msg-1',
      to: ['recipient@example.com'],
      body: 'FYI - see below',
    });

    expect(result.isError).toBeUndefined();
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages/msg-1/forward',
      expect.objectContaining({
        comment: expect.any(String),
      })
    );
  });

  it('should forward to multiple recipients', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({ id: 'fwd-3' });

    const result = await forwardEmailTool(authManager, {
      messageId: 'msg-1',
      to: ['a@example.com', 'b@example.com'],
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('a@example.com, b@example.com');
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages/msg-1/forward',
      expect.objectContaining({
        toRecipients: [
          { emailAddress: { address: 'a@example.com' } },
          { emailAddress: { address: 'b@example.com' } },
        ],
      })
    );
  });

  it('should not apply styling when preserveUserStyling is false', async () => {
    const { applyUserStyling } = await import('../../../tools/common/sharedUtils.js');
    applyUserStyling.mockClear();

    graphApiClient.postWithRetry.mockResolvedValue({ id: 'fwd-4' });

    const result = await forwardEmailTool(authManager, {
      messageId: 'msg-1',
      to: ['user@example.com'],
      body: 'Plain forward',
      preserveUserStyling: false,
    });

    expect(applyUserStyling).not.toHaveBeenCalled();
    expect(result.isError).toBeUndefined();
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/messages/msg-1/forward',
      expect.objectContaining({
        comment: 'Plain forward',
      })
    );
  });

  it('should handle API errors gracefully', async () => {
    graphApiClient.postWithRetry.mockRejectedValue(new Error('Forbidden'));

    const result = await forwardEmailTool(authManager, {
      messageId: 'msg-1',
      to: ['user@example.com'],
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Failed to forward email');
  });

  it('should strip HTML tags from styled content in comment field', async () => {
    const { applyUserStyling } = await import('../../../tools/common/sharedUtils.js');
    applyUserStyling.mockResolvedValue({ content: '<p>styled <b>forward</b></p>', type: 'html' });

    graphApiClient.postWithRetry.mockResolvedValue({ id: 'fwd-5' });

    await forwardEmailTool(authManager, {
      messageId: 'msg-1',
      to: ['user@example.com'],
      body: 'FYI',
    });

    // The forward API uses comment field, and HTML tags should be stripped
    const callArgs = graphApiClient.postWithRetry.mock.calls[0][1];
    expect(callArgs.comment).not.toContain('<p>');
    expect(callArgs.comment).not.toContain('<b>');
  });
});

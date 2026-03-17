import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { sendEmailTool } from '../../../tools/email/sendEmail.js';

// Mock sharedUtils to avoid actual Graph API calls for styling
vi.mock('../../../tools/common/sharedUtils.js', () => ({
  applyUserStyling: vi.fn().mockResolvedValue({ content: 'styled body', type: 'html' }),
  clearStylingCache: vi.fn(),
}));

describe('sendEmailTool', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
    graphApiClient.makeRequest.mockResolvedValue({ id: 'user-1' });
  });

  it('should return validation error when to field is missing', async () => {
    const result = await sendEmailTool(authManager, {
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('to');
  });

  it('should return validation error when to is an empty array', async () => {
    const result = await sendEmailTool(authManager, {
      to: [],
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('to');
  });

  it('should return validation error for invalid email format in to', async () => {
    const result = await sendEmailTool(authManager, {
      to: ['not-an-email'],
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('to');
  });

  it('should return validation error for invalid cc email', async () => {
    const result = await sendEmailTool(authManager, {
      to: ['valid@example.com'],
      cc: ['bad-email'],
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('cc');
  });

  it('should return validation error for invalid bcc email', async () => {
    const result = await sendEmailTool(authManager, {
      to: ['valid@example.com'],
      bcc: ['bad-email'],
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('bcc');
  });

  it('should send email successfully with valid recipients', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({});

    const result = await sendEmailTool(authManager, {
      to: ['user@example.com'],
      subject: 'Test Subject',
      body: 'Hello World',
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Email sent successfully');
    expect(graphApiClient.postWithRetry).toHaveBeenCalledWith(
      '/me/sendMail',
      expect.objectContaining({
        message: expect.objectContaining({
          subject: 'Test Subject',
        }),
      })
    );
  });

  it('should send email with valid cc and bcc', async () => {
    graphApiClient.postWithRetry.mockResolvedValue({});

    const result = await sendEmailTool(authManager, {
      to: ['user@example.com'],
      cc: ['cc@example.com'],
      bcc: ['bcc@example.com'],
      subject: 'Test',
      body: 'Hello',
    });

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Email sent successfully');
  });
});

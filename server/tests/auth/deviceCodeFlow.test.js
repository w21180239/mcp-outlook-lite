import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Mock config before importing module under test
vi.mock('../../auth/config.js', () => ({
  authConfig: {
    oauth: {
      deviceCodeUrl: (tenantId) => `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/devicecode`,
      tokenUrl: (tenantId) => `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      scope: 'Mail.Read Mail.Send',
    },
  },
}));

vi.mock('../../utils/mcpErrorResponse.js', () => ({
  createAuthError: (message, retryable) => {
    const err = new Error(`Authentication failed: ${message}`);
    err.isError = true;
    err.retryable = retryable;
    return err;
  },
}));

const { isHeadlessEnvironment, authenticateWithDeviceCode } = await import('../../auth/deviceCodeFlow.js');

describe('isHeadlessEnvironment', () => {
  const originalEnv = { ...process.env };
  const originalPlatform = process.platform;

  afterEach(() => {
    // Restore env
    process.env = { ...originalEnv };
    Object.defineProperty(process, 'platform', { value: originalPlatform });
  });

  it('returns true when MCP_OUTLOOK_DEVICE_CODE is "1"', () => {
    process.env.MCP_OUTLOOK_DEVICE_CODE = '1';
    expect(isHeadlessEnvironment()).toBe(true);
  });

  it('returns true when MCP_OUTLOOK_DEVICE_CODE is "true"', () => {
    process.env.MCP_OUTLOOK_DEVICE_CODE = 'true';
    expect(isHeadlessEnvironment()).toBe(true);
  });

  it('returns false when MCP_OUTLOOK_DEVICE_CODE is "0"', () => {
    process.env.MCP_OUTLOOK_DEVICE_CODE = '0';
    delete process.env.SSH_CLIENT;
    delete process.env.SSH_TTY;
    delete process.env.SSH_CONNECTION;
    delete process.env.container;
    delete process.env.DOCKER_CONTAINER;
    expect(isHeadlessEnvironment()).toBe(false);
  });

  it('returns true when SSH_CLIENT is set', () => {
    delete process.env.MCP_OUTLOOK_DEVICE_CODE;
    process.env.SSH_CLIENT = '192.168.1.1 12345 22';
    expect(isHeadlessEnvironment()).toBe(true);
  });

  it('returns true when SSH_TTY is set', () => {
    delete process.env.MCP_OUTLOOK_DEVICE_CODE;
    delete process.env.SSH_CLIENT;
    process.env.SSH_TTY = '/dev/pts/0';
    expect(isHeadlessEnvironment()).toBe(true);
  });

  it('returns true when SSH_CONNECTION is set', () => {
    delete process.env.MCP_OUTLOOK_DEVICE_CODE;
    delete process.env.SSH_CLIENT;
    delete process.env.SSH_TTY;
    process.env.SSH_CONNECTION = '192.168.1.1 12345 192.168.1.2 22';
    expect(isHeadlessEnvironment()).toBe(true);
  });

  it('returns true on Linux without DISPLAY or WAYLAND_DISPLAY', () => {
    delete process.env.MCP_OUTLOOK_DEVICE_CODE;
    delete process.env.SSH_CLIENT;
    delete process.env.SSH_TTY;
    delete process.env.SSH_CONNECTION;
    delete process.env.DISPLAY;
    delete process.env.WAYLAND_DISPLAY;
    delete process.env.container;
    delete process.env.DOCKER_CONTAINER;
    Object.defineProperty(process, 'platform', { value: 'linux' });
    expect(isHeadlessEnvironment()).toBe(true);
  });

  it('returns false on Linux with DISPLAY set', () => {
    delete process.env.MCP_OUTLOOK_DEVICE_CODE;
    delete process.env.SSH_CLIENT;
    delete process.env.SSH_TTY;
    delete process.env.SSH_CONNECTION;
    delete process.env.container;
    delete process.env.DOCKER_CONTAINER;
    process.env.DISPLAY = ':0';
    Object.defineProperty(process, 'platform', { value: 'linux' });
    expect(isHeadlessEnvironment()).toBe(false);
  });

  it('returns true when container env var is set (Docker)', () => {
    delete process.env.MCP_OUTLOOK_DEVICE_CODE;
    delete process.env.SSH_CLIENT;
    delete process.env.SSH_TTY;
    delete process.env.SSH_CONNECTION;
    process.env.container = 'docker';
    expect(isHeadlessEnvironment()).toBe(true);
  });

  it('returns true when DOCKER_CONTAINER is set', () => {
    delete process.env.MCP_OUTLOOK_DEVICE_CODE;
    delete process.env.SSH_CLIENT;
    delete process.env.SSH_TTY;
    delete process.env.SSH_CONNECTION;
    delete process.env.container;
    process.env.DOCKER_CONTAINER = '1';
    expect(isHeadlessEnvironment()).toBe(true);
  });

  it('returns false in a normal desktop environment', () => {
    delete process.env.MCP_OUTLOOK_DEVICE_CODE;
    delete process.env.SSH_CLIENT;
    delete process.env.SSH_TTY;
    delete process.env.SSH_CONNECTION;
    delete process.env.container;
    delete process.env.DOCKER_CONTAINER;
    Object.defineProperty(process, 'platform', { value: 'darwin' });
    expect(isHeadlessEnvironment()).toBe(false);
  });
});

describe('authenticateWithDeviceCode', () => {
  let originalFetch;

  beforeEach(() => {
    originalFetch = global.fetch;
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    global.fetch = originalFetch;
    vi.restoreAllMocks();
  });

  it('throws when device code request fails', async () => {
    global.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 400,
      text: async () => 'Bad Request',
    });

    await expect(authenticateWithDeviceCode('client-id', 'tenant-id'))
      .rejects.toThrow('Authentication failed');
  });

  it('completes successfully when token is returned on first poll', async () => {
    const deviceCodeResponse = {
      device_code: 'device-123',
      user_code: 'ABCD-EFGH',
      verification_uri: 'https://microsoft.com/devicelogin',
      expires_in: 900,
      interval: 0, // Use 0 to avoid real delays
      message: 'Go to https://microsoft.com/devicelogin and enter code ABCD-EFGH',
    };

    const tokenResponse = {
      access_token: 'test-access-token',
      refresh_token: 'test-refresh-token',
      expires_in: 3600,
      token_type: 'Bearer',
    };

    let callCount = 0;
    global.fetch = vi.fn().mockImplementation(async (url) => {
      callCount++;
      if (callCount === 1) {
        // Device code request
        return { ok: true, json: async () => deviceCodeResponse };
      }
      // Token poll - success immediately
      return { ok: true, json: async () => tokenResponse };
    });

    const result = await authenticateWithDeviceCode('client-id', 'tenant-id');

    expect(result.access_token).toBe('test-access-token');
    expect(result.refresh_token).toBe('test-refresh-token');
    expect(result.token_type).toBe('Bearer');
  });

  it('retries when authorization_pending is returned', async () => {
    const deviceCodeResponse = {
      device_code: 'device-123',
      user_code: 'ABCD-EFGH',
      verification_uri: 'https://microsoft.com/devicelogin',
      expires_in: 900,
      interval: 0,
      message: 'Enter code',
    };

    const tokenResponse = {
      access_token: 'final-token',
      expires_in: 3600,
      token_type: 'Bearer',
    };

    let callCount = 0;
    global.fetch = vi.fn().mockImplementation(async () => {
      callCount++;
      if (callCount === 1) {
        return { ok: true, json: async () => deviceCodeResponse };
      }
      if (callCount === 2) {
        // First poll: authorization_pending
        return { ok: false, json: async () => ({ error: 'authorization_pending' }) };
      }
      // Second poll: success
      return { ok: true, json: async () => tokenResponse };
    });

    const result = await authenticateWithDeviceCode('client-id', 'tenant-id');
    expect(result.access_token).toBe('final-token');
    expect(callCount).toBe(3);
  });

  it('throws when user declines authorization', async () => {
    const deviceCodeResponse = {
      device_code: 'device-123',
      user_code: 'ABCD-EFGH',
      verification_uri: 'https://microsoft.com/devicelogin',
      expires_in: 900,
      interval: 0,
      message: 'Enter code',
    };

    let callCount = 0;
    global.fetch = vi.fn().mockImplementation(async () => {
      callCount++;
      if (callCount === 1) {
        return { ok: true, json: async () => deviceCodeResponse };
      }
      return { ok: false, json: async () => ({ error: 'authorization_declined' }) };
    });

    await expect(authenticateWithDeviceCode('client-id', 'tenant-id'))
      .rejects.toThrow('Authentication failed');
  });
});

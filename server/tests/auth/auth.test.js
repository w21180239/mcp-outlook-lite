import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock all dependencies before importing the module under test
vi.mock('@microsoft/microsoft-graph-client', () => ({
  Client: { init: vi.fn(() => ({ api: vi.fn() })) },
}));

vi.mock('../../auth/tokenManager.js', () => ({
  TokenManager: vi.fn().mockImplementation(() => ({
    getPKCEVerifier: vi.fn().mockResolvedValue('test-verifier'),
    getRefreshToken: vi.fn().mockResolvedValue('test-refresh-token'),
    storeTokens: vi.fn().mockResolvedValue(undefined),
    clearTokens: vi.fn().mockResolvedValue(undefined),
    getAccessToken: vi.fn().mockResolvedValue('test-token'),
    isAuthenticated: vi.fn().mockResolvedValue(false),
    getTokenMetadata: vi.fn().mockResolvedValue(null),
  })),
}));

vi.mock('../../auth/config.js', () => ({
  authConfig: {
    oauth: {
      tokenUrl: (tenantId) => `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      scope: 'openid profile offline_access',
    },
  },
}));

vi.mock('../../graph/graphClient.js', () => ({
  GraphApiClient: vi.fn().mockImplementation(() => ({
    initialize: vi.fn().mockResolvedValue(undefined),
  })),
}));

vi.mock('../../auth/browserLauncher.js', () => ({
  openBrowser: vi.fn(),
}));

vi.mock('../../auth/templates.js', () => ({
  getSuccessPage: vi.fn(() => ''),
  getErrorPage: vi.fn(() => ''),
  getFailurePage: vi.fn(() => ''),
}));

const { OutlookAuthManager } = await import('../../auth/auth.js');

describe('OutlookAuthManager - OAuth error sanitization', () => {
  let manager;

  beforeEach(() => {
    vi.clearAllMocks();
    manager = new OutlookAuthManager('test-client-id', 'test-tenant-id');
    manager.lastUsedPort = 3000;
  });

  describe('exchangeCodeForToken', () => {
    it('should NOT leak raw error body from Microsoft OAuth', async () => {
      const sensitiveBody = 'invalid_grant: AADSTS70000 sensitive_detail';
      global.fetch = vi.fn().mockResolvedValue({
        ok: false,
        status: 400,
        text: async () => sensitiveBody,
      });

      try {
        await manager.exchangeCodeForToken('test-code');
        expect.fail('Should have thrown');
      } catch (error) {
        const errorText = JSON.stringify(error);
        expect(errorText).not.toContain('sensitive_detail');
        expect(errorText).not.toContain('AADSTS70000');
        expect(error.content[0].text).toContain('Token exchange failed');
        expect(error.content[0].text).toContain('400');
      }
    });
  });

  describe('refreshAccessToken', () => {
    it('should NOT leak raw error body from Microsoft OAuth', async () => {
      const sensitiveBody = 'error=invalid_token&error_description=sensitive_info';
      global.fetch = vi.fn().mockResolvedValue({
        ok: false,
        status: 401,
        text: async () => sensitiveBody,
      });

      try {
        await manager.refreshAccessToken();
        expect.fail('Should have thrown');
      } catch (error) {
        const errorText = JSON.stringify(error);
        expect(errorText).not.toContain('sensitive_info');
        expect(error.content[0].text).toContain('Token refresh failed');
      }
    });
  });
});

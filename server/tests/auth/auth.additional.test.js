import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock all dependencies before importing the module under test
const mockTokenManager = {
  getPKCEVerifier: vi.fn().mockResolvedValue('test-verifier'),
  getRefreshToken: vi.fn().mockResolvedValue('test-refresh-token'),
  storeTokens: vi.fn().mockResolvedValue(undefined),
  clearTokens: vi.fn().mockResolvedValue(undefined),
  getAccessToken: vi.fn().mockResolvedValue('test-token'),
  isAuthenticated: vi.fn().mockResolvedValue(false),
  getTokenMetadata: vi.fn().mockResolvedValue(null),
  generateCodeVerifier: vi.fn().mockReturnValue('verifier'),
  generateCodeChallenge: vi.fn().mockReturnValue('challenge'),
  storePKCEVerifier: vi.fn().mockResolvedValue(undefined),
};

vi.mock('../../auth/tokenManager.js', () => ({
  TokenManager: vi.fn().mockImplementation(() => mockTokenManager),
}));

const mockGraphClientApi = vi.fn().mockReturnValue({
  get: vi.fn().mockResolvedValue({ id: 'uid', displayName: 'User', mail: 'user@test.com' }),
});

vi.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: vi.fn(() => ({ api: mockGraphClientApi })),
  },
}));

vi.mock('../../auth/config.js', () => ({
  authConfig: {
    oauth: {
      authorizeUrl: (t) => `https://login.microsoftonline.com/${t}/oauth2/v2.0/authorize`,
      tokenUrl: (t) => `https://login.microsoftonline.com/${t}/oauth2/v2.0/token`,
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

vi.mock('../../auth/deviceCodeFlow.js', () => ({
  isHeadlessEnvironment: vi.fn().mockReturnValue(false),
  authenticateWithDeviceCode: vi.fn(),
}));

vi.mock('../../auth/templates.js', () => ({
  getSuccessPage: vi.fn(() => ''),
  getErrorPage: vi.fn(() => ''),
  getFailurePage: vi.fn(() => ''),
}));

const { OutlookAuthManager } = await import('../../auth/auth.js');

describe('OutlookAuthManager - additional coverage', () => {
  let manager;

  beforeEach(() => {
    vi.clearAllMocks();
    manager = new OutlookAuthManager('test-client-id', 'test-tenant-id');
    manager.lastUsedPort = 3000;
    // Reset mock defaults
    mockTokenManager.isAuthenticated.mockResolvedValue(false);
    mockTokenManager.getTokenMetadata.mockResolvedValue(null);
    mockTokenManager.getAccessToken.mockResolvedValue('test-token');
    mockGraphClientApi.mockReturnValue({
      get: vi.fn().mockResolvedValue({ id: 'uid', displayName: 'User', mail: 'user@test.com' }),
    });
  });

  describe('authenticate', () => {
    it('should return success when token is already valid', async () => {
      mockTokenManager.isAuthenticated.mockResolvedValue(true);
      const result = await manager.authenticate();
      expect(result.success).toBe(true);
      expect(result.user.displayName).toBe('User');
    });

    it('should attempt silent refresh when metadata exists', async () => {
      mockTokenManager.isAuthenticated.mockResolvedValue(false);
      mockTokenManager.getTokenMetadata.mockResolvedValue({
        accessTokenExpiry: Date.now() - 1000,
        refreshTokenExpiry: Date.now() + 86400000,
        lastRefresh: Date.now() - 3600000,
      });
      mockTokenManager.getRefreshToken.mockResolvedValue('refresh-token');

      // Mock fetch for token refresh
      global.fetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ access_token: 'new-at', refresh_token: 'new-rt', expires_in: 3600 }),
      });

      const result = await manager.authenticate();
      expect(result.success).toBe(true);
    });

    it('should return MCP error as-is when error has isError flag', async () => {
      mockTokenManager.isAuthenticated.mockRejectedValue({ isError: true, content: [{ type: 'text', text: 'auth err' }] });

      const result = await manager.authenticate();
      expect(result.success).toBe(false);
      expect(result.error.isError).toBe(true);
    });

    it('should wrap non-MCP errors in auth error format', async () => {
      mockTokenManager.isAuthenticated.mockRejectedValue(new Error('network fail'));

      const result = await manager.authenticate();
      expect(result.success).toBe(false);
      expect(result.error.isError).toBe(true);
      expect(result.error.content[0].text).toContain('network fail');
    });
  });

  describe('ensureAuthenticated', () => {
    it('should call authenticate when not authenticated', async () => {
      manager.isAuthenticated = false;
      manager.graphClient = null;

      mockTokenManager.isAuthenticated.mockResolvedValue(true);

      const result = await manager.ensureAuthenticated();
      expect(result).toBeDefined();
    });

    it('should refresh token when access token needs refresh (MCP error)', async () => {
      manager.isAuthenticated = true;
      manager.graphClient = { api: mockGraphClientApi };

      const mcpError = {
        isError: true,
        _errorDetails: { needsRefresh: true },
        content: [{ type: 'text', text: 'needs refresh' }],
      };
      mockTokenManager.getAccessToken.mockRejectedValue(mcpError);
      mockTokenManager.getRefreshToken.mockResolvedValue('rt');

      global.fetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ access_token: 'new-at', refresh_token: 'new-rt', expires_in: 3600 }),
      });

      const result = await manager.ensureAuthenticated();
      expect(result).toBeDefined();
    });

    it('should throw MCP error without needsRefresh', async () => {
      manager.isAuthenticated = true;
      manager.graphClient = { api: mockGraphClientApi };

      const mcpError = {
        isError: true,
        content: [{ type: 'text', text: 'some error' }],
      };
      mockTokenManager.getAccessToken.mockRejectedValue(mcpError);

      await expect(manager.ensureAuthenticated()).rejects.toEqual(mcpError);
    });

    it('should handle "needs refresh" message in standard Error', async () => {
      manager.isAuthenticated = true;
      manager.graphClient = { api: mockGraphClientApi };

      mockTokenManager.getAccessToken.mockRejectedValue(new Error('Token needs refresh'));
      mockTokenManager.getRefreshToken.mockResolvedValue('rt');

      global.fetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ access_token: 'new-at', refresh_token: 'new-rt', expires_in: 3600 }),
      });

      const result = await manager.ensureAuthenticated();
      expect(result).toBeDefined();
    });

    it('should convert unknown errors to tool errors', async () => {
      manager.isAuthenticated = true;
      manager.graphClient = { api: mockGraphClientApi };

      mockTokenManager.getAccessToken.mockRejectedValue(new Error('Something else'));

      await expect(manager.ensureAuthenticated()).rejects.toHaveProperty('isError', true);
    });
  });

  describe('validateAuthentication', () => {
    it('should return user info on success', async () => {
      manager.graphClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ id: '1', displayName: 'Test', mail: 'test@t.com' }),
        }),
      };

      const result = await manager.validateAuthentication();
      expect(result.success).toBe(true);
      expect(manager.isAuthenticated).toBe(true);
    });

    it('should use userPrincipalName when mail is absent', async () => {
      manager.graphClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({ id: '1', displayName: 'Test', userPrincipalName: 'upn@t.com' }),
        }),
      };

      const result = await manager.validateAuthentication();
      expect(result.user.mail).toBe('upn@t.com');
    });

    it('should throw MCP error on API failure', async () => {
      manager.graphClient = {
        api: vi.fn().mockReturnValue({
          get: vi.fn().mockRejectedValue(new Error('API fail')),
        }),
      };

      await expect(manager.validateAuthentication()).rejects.toHaveProperty('isError', true);
      expect(manager.isAuthenticated).toBe(false);
    });
  });

  describe('getGraphClient / getGraphApiClient', () => {
    it('should throw auth error when not authenticated (getGraphClient)', () => {
      manager.graphClient = null;
      expect(() => manager.getGraphClient()).toThrow();
    });

    it('should throw auth error when not authenticated (getGraphApiClient)', () => {
      manager.graphApiClient = null;
      expect(() => manager.getGraphApiClient()).toThrow();
    });

    it('should return client when authenticated (getGraphClient)', () => {
      const client = { api: vi.fn() };
      manager.graphClient = client;
      expect(manager.getGraphClient()).toBe(client);
    });
  });

  describe('logout', () => {
    it('should clear all state', async () => {
      manager.graphClient = {};
      manager.graphApiClient = {};
      manager.isAuthenticated = true;
      manager.authenticationRecord = {};

      await manager.logout();

      expect(mockTokenManager.clearTokens).toHaveBeenCalled();
      expect(manager.graphClient).toBeNull();
      expect(manager.graphApiClient).toBeNull();
      expect(manager.isAuthenticated).toBe(false);
      expect(manager.authenticationRecord).toBeNull();
    });
  });

  describe('refreshAccessToken', () => {
    it('should store new tokens on success', async () => {
      global.fetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ access_token: 'new-at', refresh_token: 'new-rt', expires_in: 3600 }),
      });

      const result = await manager.refreshAccessToken();
      expect(result).toBe(true);
      expect(mockTokenManager.storeTokens).toHaveBeenCalledWith('new-at', 'new-rt', 3600);
    });

    it('should keep old refresh token when new one is not provided', async () => {
      mockTokenManager.getRefreshToken.mockResolvedValue('old-rt');
      global.fetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({ access_token: 'new-at', expires_in: 3600 }),
      });

      await manager.refreshAccessToken();
      expect(mockTokenManager.storeTokens).toHaveBeenCalledWith('new-at', 'old-rt', 3600);
    });

    it('should clear tokens and rethrow MCP errors on failure', async () => {
      global.fetch = vi.fn().mockResolvedValue({
        ok: false,
        status: 400,
        text: async () => 'bad request',
      });

      await expect(manager.refreshAccessToken()).rejects.toHaveProperty('isError', true);
      expect(mockTokenManager.clearTokens).toHaveBeenCalled();
    });
  });
});

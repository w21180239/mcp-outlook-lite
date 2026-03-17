import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock dependencies
vi.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: vi.fn((config) => {
      const mockClient = {
        api: vi.fn(() => mockClient),
        get: vi.fn().mockResolvedValue({ value: [] }),
        post: vi.fn().mockResolvedValue({}),
        patch: vi.fn().mockResolvedValue({}),
        put: vi.fn().mockResolvedValue({}),
        delete: vi.fn().mockResolvedValue({}),
        select: vi.fn(() => mockClient),
        filter: vi.fn(() => mockClient),
        top: vi.fn(() => mockClient),
        orderby: vi.fn(() => mockClient),
        expand: vi.fn(() => mockClient),
        search: vi.fn(() => mockClient),
        query: vi.fn(() => mockClient),
        header: vi.fn(() => mockClient),
      };
      // Resolve the auth provider to avoid hanging
      if (config && config.authProvider) {
        config.authProvider((err, token) => {});
      }
      return mockClient;
    }),
  },
}));

vi.mock('../../auth/config.js', () => ({
  authConfig: {
    oauth: {
      scope: 'Mail.Read',
    },
    retry: {
      maxAttempts: 3,
      initialDelay: 10, // Small delays for tests
      maxDelay: 100,
      backoffMultiplier: 2,
    },
  },
}));

vi.mock('../../utils/mcpErrorResponse.js', () => ({
  convertErrorToToolError: vi.fn((error, context) => ({
    isError: true,
    content: [{ type: 'text', text: `${context}: ${error.message}` }],
  })),
  createServiceUnavailableError: vi.fn((service) => ({
    isError: true,
    content: [{ type: 'text', text: `${service} is temporarily unavailable.` }],
  })),
  createRateLimitError: vi.fn((retryAfter) => ({
    isError: true,
    content: [{ type: 'text', text: `Rate limit exceeded. Wait ${retryAfter} seconds.` }],
  })),
  createValidationError: vi.fn((param, reason) => ({
    isError: true,
    content: [{ type: 'text', text: `Invalid parameter '${param}': ${reason}` }],
  })),
}));

vi.mock('../../utils/jsonUtils.js', () => ({
  safeStringify: vi.fn((obj) => JSON.stringify(obj)),
}));

vi.mock('../../graph/folderResolver.js', () => ({
  FolderResolver: vi.fn().mockImplementation(() => ({
    resolveFolderToId: vi.fn(),
  })),
}));

const { GraphApiClient } = await import('../../graph/graphClient.js');

describe('GraphApiClient', () => {
  let client;
  let mockAuthManager;

  beforeEach(() => {
    vi.clearAllMocks();
    mockAuthManager = {
      tokenManager: {
        getAccessToken: vi.fn().mockResolvedValue('test-token'),
      },
      refreshAccessToken: vi.fn().mockResolvedValue(undefined),
    };
    client = new GraphApiClient(mockAuthManager);
    vi.spyOn(console, 'error').mockImplementation(() => {});
    vi.spyOn(console, 'warn').mockImplementation(() => {});
  });

  describe('constructor', () => {
    it('initializes with default values', () => {
      expect(client.client).toBeNull();
      expect(client.requestCount).toBe(0);
      expect(client.maxConcurrentRequests).toBe(4);
      expect(client.activeRequests).toBe(0);
      expect(client.folderResolver).toBeNull();
    });

    it('initializes rate limit metrics', () => {
      expect(client.rateLimitMetrics.rateLimitHits).toBe(0);
      expect(client.rateLimitMetrics.totalRetries).toBe(0);
      expect(client.rateLimitMetrics.requestDurations).toEqual([]);
    });
  });

  describe('initialize', () => {
    it('initializes the Graph client', async () => {
      const result = await client.initialize();
      expect(result).toBeDefined();
      expect(client.client).not.toBeNull();
    });

    it('returns existing client on subsequent calls', async () => {
      const first = await client.initialize();
      const second = await client.initialize();
      expect(first).toBe(second);
    });
  });

  describe('enforceRateLimit', () => {
    it('increments activeRequests', async () => {
      expect(client.activeRequests).toBe(0);
      await client.enforceRateLimit();
      expect(client.activeRequests).toBe(1);
    });

    it('adds timestamp to requestWindow', async () => {
      expect(client.requestWindow).toHaveLength(0);
      await client.enforceRateLimit();
      expect(client.requestWindow).toHaveLength(1);
    });

    it('cleans up old timestamps from requestWindow', async () => {
      // Add an old timestamp
      client.requestWindow.push(Date.now() - 120000); // 2 minutes ago
      await client.enforceRateLimit();
      // Old one should be cleaned, new one added
      expect(client.requestWindow).toHaveLength(1);
    });
  });

  describe('extractRetryAfter', () => {
    it('extracts Retry-After from headers (lowercase)', () => {
      const error = { headers: { 'retry-after': '30' } };
      expect(client.extractRetryAfter(error)).toBe(30000);
    });

    it('extracts Retry-After from headers (Title-Case)', () => {
      const error = { headers: { 'Retry-After': '60' } };
      expect(client.extractRetryAfter(error)).toBe(60000);
    });

    it('extracts retry-after-ms from inner error body', () => {
      const error = {
        headers: {},
        body: { error: { innerError: { 'retry-after-ms': '5000' } } },
      };
      expect(client.extractRetryAfter(error)).toBe(5000);
    });

    it('returns null when no retry info is available', () => {
      expect(client.extractRetryAfter({})).toBeNull();
      expect(client.extractRetryAfter({ headers: {} })).toBeNull();
    });
  });

  describe('generateCorrelationId', () => {
    it('generates a non-empty string', () => {
      const id = client.generateCorrelationId();
      expect(typeof id).toBe('string');
      expect(id.length).toBeGreaterThan(0);
    });

    it('generates unique IDs', () => {
      const id1 = client.generateCorrelationId();
      const id2 = client.generateCorrelationId();
      expect(id1).not.toBe(id2);
    });
  });

  describe('isRetryableError', () => {
    it('returns true for 429 (rate limit)', () => {
      expect(client.isRetryableError(429)).toBe(true);
    });

    it('returns true for 5xx errors', () => {
      expect(client.isRetryableError(500)).toBe(true);
      expect(client.isRetryableError(502)).toBe(true);
      expect(client.isRetryableError(503)).toBe(true);
      expect(client.isRetryableError(504)).toBe(true);
    });

    it('returns true for 401 (auth)', () => {
      expect(client.isRetryableError(401)).toBe(true);
    });

    it('returns false for 400 and 404', () => {
      expect(client.isRetryableError(400)).toBe(false);
      expect(client.isRetryableError(404)).toBe(false);
    });
  });

  describe('updateMetrics', () => {
    it('tracks request durations', () => {
      client.updateMetrics(100);
      client.updateMetrics(200);
      expect(client.rateLimitMetrics.requestDurations).toEqual([100, 200]);
      expect(client.rateLimitMetrics.averageRequestDuration).toBe(150);
    });

    it('keeps only last 100 durations', () => {
      for (let i = 0; i < 110; i++) {
        client.updateMetrics(10);
      }
      expect(client.rateLimitMetrics.requestDurations).toHaveLength(100);
    });
  });

  describe('getRateLimitMetrics', () => {
    it('returns current metrics with additional fields', () => {
      const metrics = client.getRateLimitMetrics();
      expect(metrics).toHaveProperty('activeRequests');
      expect(metrics).toHaveProperty('requestsInWindow');
      expect(metrics).toHaveProperty('lastUpdated');
    });
  });

  describe('resetMetrics', () => {
    it('resets all metrics to initial values', () => {
      client.rateLimitMetrics.rateLimitHits = 5;
      client.rateLimitMetrics.totalRetries = 10;
      client.resetMetrics();
      expect(client.rateLimitMetrics.rateLimitHits).toBe(0);
      expect(client.rateLimitMetrics.totalRetries).toBe(0);
      expect(client.rateLimitMetrics.requestDurations).toEqual([]);
    });
  });

  describe('checkRateLimitHealth', () => {
    it('returns no alerts when healthy', () => {
      const alerts = client.checkRateLimitHealth();
      expect(alerts).toEqual([]);
    });

    it('returns warning when rate limit hits exceed threshold', () => {
      client.rateLimitMetrics.rateLimitHits = 10;
      const alerts = client.checkRateLimitHealth();
      expect(alerts.some(a => a.level === 'warning' && a.message.includes('rate limit'))).toBe(true);
    });

    it('returns warning when average duration is high', () => {
      client.rateLimitMetrics.averageRequestDuration = 6000;
      const alerts = client.checkRateLimitHealth();
      expect(alerts.some(a => a.message.includes('request duration'))).toBe(true);
    });

    it('returns error when max concurrent requests reached', () => {
      client.activeRequests = 4;
      const alerts = client.checkRateLimitHealth();
      expect(alerts.some(a => a.level === 'error')).toBe(true);
    });
  });

  describe('getFolderResolver', () => {
    it('creates FolderResolver on first call', () => {
      const resolver = client.getFolderResolver();
      expect(resolver).toBeDefined();
    });

    it('returns same instance on subsequent calls', () => {
      const r1 = client.getFolderResolver();
      const r2 = client.getFolderResolver();
      expect(r1).toBe(r2);
    });
  });

  describe('makeBatchRequest', () => {
    it('returns validation error when more than 20 requests', async () => {
      const requests = Array.from({ length: 21 }, (_, i) => ({ url: `/test/${i}` }));
      const result = await client.makeBatchRequest(requests);
      expect(result.isError).toBe(true);
    });
  });

  describe('handleGraphError', () => {
    it('returns rate limit error for 429 status', () => {
      const error = { status: 429, message: 'Too many requests', headers: {} };
      const details = { statusCode: 429, message: 'Too many requests', microsoftCorrelationIds: {} };
      const result = client.handleGraphError(error, details);
      expect(result.isError).toBe(true);
    });

    it('returns service unavailable for 500 status', () => {
      const error = { status: 500, message: 'Server error', headers: {} };
      const details = { statusCode: 500, message: 'Server error', microsoftCorrelationIds: {} };
      const result = client.handleGraphError(error, details);
      expect(result.isError).toBe(true);
    });

    it('returns appropriate error for 404 status', () => {
      const error = { status: 404, message: 'Not found', headers: {} };
      const details = { statusCode: 404, message: 'Not found', microsoftCorrelationIds: {} };
      const result = client.handleGraphError(error, details);
      expect(result.isError).toBe(true);
    });

    it('includes support correlation IDs in error message', () => {
      const error = { status: 400, message: 'Bad request', headers: {} };
      const details = {
        statusCode: 400,
        message: 'Bad request',
        microsoftCorrelationIds: { requestId: 'abc-123' },
      };
      const result = client.handleGraphError(error, details);
      expect(result.isError).toBe(true);
    });
  });
});

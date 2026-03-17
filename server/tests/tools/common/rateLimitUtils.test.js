import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { getRateLimitMetricsTool, resetRateLimitMetricsTool } from '../../../tools/common/rateLimitUtils.js';

describe('getRateLimitMetricsTool', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  it('should return rate limit metrics when getRateLimitMetrics exists', async () => {
    const mockMetrics = {
      requestCount: 42,
      throttleCount: 3,
      lastResetTime: '2024-01-01T00:00:00Z',
      currentWindow: {
        remainingRequests: 100,
        resetTime: '2024-01-01T01:00:00Z',
      },
    };
    graphApiClient.getRateLimitMetrics = vi.fn().mockReturnValue(mockMetrics);

    const result = await getRateLimitMetricsTool(authManager, {});

    expect(result.isError).toBeUndefined();
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.requestCount).toBe(42);
    expect(parsed.throttleCount).toBe(3);
  });

  it('should return default metrics when getRateLimitMetrics does not exist', async () => {
    // graphApiClient from mock doesn't have getRateLimitMetrics
    const result = await getRateLimitMetricsTool(authManager, {});

    expect(result.isError).toBeUndefined();
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.requestCount).toBe(0);
    expect(parsed.throttleCount).toBe(0);
    expect(parsed.currentWindow.remainingRequests).toBe('unknown');
  });

  it('should handle authentication errors', async () => {
    authManager.ensureAuthenticated.mockRejectedValue(new Error('Auth failed'));

    const result = await getRateLimitMetricsTool(authManager, {});

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Failed to get rate limit metrics');
  });
});

describe('resetRateLimitMetricsTool', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  it('should reset metrics when resetRateLimitMetrics exists', async () => {
    graphApiClient.resetRateLimitMetrics = vi.fn();

    const result = await resetRateLimitMetricsTool(authManager, {});

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Rate limit metrics reset successfully');
    expect(graphApiClient.resetRateLimitMetrics).toHaveBeenCalled();
  });

  it('should succeed even when resetRateLimitMetrics does not exist', async () => {
    // graphApiClient from mock doesn't have resetRateLimitMetrics
    const result = await resetRateLimitMetricsTool(authManager, {});

    expect(result.isError).toBeUndefined();
    expect(result.content[0].text).toContain('Rate limit metrics reset successfully');
  });

  it('should handle authentication errors', async () => {
    authManager.ensureAuthenticated.mockRejectedValue(new Error('Not authenticated'));

    const result = await resetRateLimitMetricsTool(authManager, {});

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Failed to reset rate limit metrics');
  });
});

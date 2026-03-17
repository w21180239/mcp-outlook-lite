import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { HttpConnectionPool, createGraphConnectionPool } from '../../graph/httpConfig.js';

describe('HttpConnectionPool', () => {
  let pool;

  afterEach(() => {
    if (pool) {
      pool.destroy();
      pool = null;
    }
  });

  describe('constructor', () => {
    it('should create with default options', () => {
      pool = new HttpConnectionPool();
      expect(pool.options.maxSockets).toBe(10);
      expect(pool.options.maxFreeSockets).toBe(5);
      expect(pool.options.timeout).toBe(30000);
      expect(pool.options.keepAlive).toBe(true);
    });

    it('should accept custom options', () => {
      pool = new HttpConnectionPool({ maxSockets: 20, timeout: 60000 });
      expect(pool.options.maxSockets).toBe(20);
      expect(pool.options.timeout).toBe(60000);
    });

    it('should create http and https agents', () => {
      pool = new HttpConnectionPool();
      expect(pool.httpAgent).toBeDefined();
      expect(pool.httpsAgent).toBeDefined();
    });

    it('should initialize metrics', () => {
      pool = new HttpConnectionPool();
      expect(pool.metrics.socketsCreated).toBe(0);
      expect(pool.metrics.activeSockets).toBe(0);
      expect(pool.metrics.requestsMade).toBe(0);
    });

    it('should set up connection monitor', () => {
      pool = new HttpConnectionPool();
      expect(pool.connectionMonitor).toBeDefined();
    });

    it('should allow disabling keepAlive', () => {
      pool = new HttpConnectionPool({ keepAlive: false });
      expect(pool.options.keepAlive).toBe(false);
    });

    it('should allow disabling metrics', () => {
      pool = new HttpConnectionPool({ enableMetrics: false });
      expect(pool.options.enableMetrics).toBe(false);
    });
  });

  describe('getAgent', () => {
    it('should return https agent for https URLs', () => {
      pool = new HttpConnectionPool();
      const agent = pool.getAgent('https://graph.microsoft.com/v1.0/me');
      expect(agent).toBe(pool.httpsAgent);
    });

    it('should return http agent for http URLs', () => {
      pool = new HttpConnectionPool();
      const agent = pool.getAgent('http://localhost:3000/callback');
      expect(agent).toBe(pool.httpAgent);
    });
  });

  describe('getAxiosConfig', () => {
    it('should return config with httpsAgent for https base URL', () => {
      pool = new HttpConnectionPool();
      const config = pool.getAxiosConfig('https://graph.microsoft.com');
      expect(config.httpsAgent).toBe(pool.httpsAgent);
      expect(config.timeout).toBe(30000);
      expect(config.headers['Connection']).toBe('keep-alive');
      expect(config.headers['User-Agent']).toBe('mcp-outlook-lite/2.0.0');
    });

    it('should return config with httpAgent for http base URL', () => {
      pool = new HttpConnectionPool();
      const config = pool.getAxiosConfig('http://localhost:3000');
      expect(config.httpAgent).toBe(pool.httpAgent);
    });

    it('should use default base URL when none provided', () => {
      pool = new HttpConnectionPool();
      const config = pool.getAxiosConfig();
      expect(config.httpsAgent).toBe(pool.httpsAgent);
    });

    it('should have transformResponse that parses JSON', () => {
      pool = new HttpConnectionPool();
      const config = pool.getAxiosConfig();
      const transform = config.transformResponse[0];
      expect(transform('{"key":"value"}')).toEqual({ key: 'value' });
    });

    it('should have transformResponse that returns raw data on parse failure', () => {
      pool = new HttpConnectionPool();
      const config = pool.getAxiosConfig();
      const transform = config.transformResponse[0];
      expect(transform('not json')).toBe('not json');
    });

    it('should have validateStatus that accepts 4xx', () => {
      pool = new HttpConnectionPool();
      const config = pool.getAxiosConfig();
      expect(config.validateStatus(200)).toBe(true);
      expect(config.validateStatus(404)).toBe(true);
      expect(config.validateStatus(500)).toBe(false);
    });
  });

  describe('getFetchConfig', () => {
    it('should return config with agent and headers', () => {
      pool = new HttpConnectionPool();
      const config = pool.getFetchConfig('https://graph.microsoft.com/v1.0/me');
      expect(config.agent).toBe(pool.httpsAgent);
      expect(config.timeout).toBe(30000);
      expect(config.headers['Connection']).toBe('keep-alive');
    });
  });

  describe('trackRequest', () => {
    it('should increment requestsMade and record connection time', () => {
      pool = new HttpConnectionPool();
      const startTime = performance.now() - 100;
      pool.trackRequest(startTime);
      expect(pool.metrics.requestsMade).toBe(1);
      expect(pool.metrics.connectionTimes).toHaveLength(1);
    });

    it('should do nothing when metrics are disabled', () => {
      pool = new HttpConnectionPool({ enableMetrics: false });
      pool.trackRequest(performance.now());
      expect(pool.metrics.requestsMade).toBe(0);
    });
  });

  describe('getMetrics', () => {
    it('should return metrics with agent stats and config', () => {
      pool = new HttpConnectionPool();
      const metrics = pool.getMetrics();
      expect(metrics.httpAgent).toBeDefined();
      expect(metrics.httpsAgent).toBeDefined();
      expect(metrics.config.maxSockets).toBe(10);
      expect(metrics.config.keepAlive).toBe(true);
    });
  });

  describe('getAgentStats', () => {
    it('should return stats structure with zero totals for fresh agent', () => {
      pool = new HttpConnectionPool();
      const stats = pool.getAgentStats(pool.httpAgent);
      expect(stats.totalSockets).toBe(0);
      expect(stats.totalRequests).toBe(0);
    });
  });

  describe('updateMetrics', () => {
    it('should calculate pool utilization', () => {
      pool = new HttpConnectionPool();
      pool.metrics.activeSockets = 5;
      pool.updateMetrics();
      expect(pool.metrics.poolUtilization).toBe(25); // 5 / (10*2) * 100
    });

    it('should calculate average connection time', () => {
      pool = new HttpConnectionPool();
      pool.metrics.connectionTimes = [100, 200, 300];
      pool.updateMetrics();
      expect(pool.metrics.averageConnectionTime).toBe(200);
    });

    it('should trim connection times to last 100', () => {
      pool = new HttpConnectionPool();
      pool.metrics.connectionTimes = new Array(150).fill(10);
      pool.updateMetrics();
      expect(pool.metrics.connectionTimes).toHaveLength(100);
    });

    it('should skip when metrics disabled', () => {
      pool = new HttpConnectionPool({ enableMetrics: false });
      pool.metrics.activeSockets = 5;
      pool.updateMetrics();
      expect(pool.metrics.poolUtilization).toBe(0);
    });
  });

  describe('healthCheck', () => {
    it('should report healthy for fresh pool', async () => {
      pool = new HttpConnectionPool();
      const health = await pool.healthCheck();
      expect(health.healthy).toBe(true);
      // Fresh pool may have info-level issues (low reuse rate) but no errors
      expect(health.issues.filter(i => i.level === 'error')).toEqual([]);
    });

    it('should warn on high pool utilization', async () => {
      pool = new HttpConnectionPool();
      pool.metrics.poolUtilization = 95;
      const health = await pool.healthCheck();
      expect(health.issues.some(i => i.level === 'warning' && i.message.includes('utilization'))).toBe(true);
    });

    it('should error on high socket error rate', async () => {
      pool = new HttpConnectionPool();
      pool.metrics.socketErrors = 20;
      pool.metrics.socketsCreated = 100;
      const health = await pool.healthCheck();
      expect(health.healthy).toBe(false);
      expect(health.issues.some(i => i.level === 'error')).toBe(true);
    });

    it('should warn on high timeout rate', async () => {
      pool = new HttpConnectionPool();
      pool.metrics.socketTimeouts = 10;
      pool.metrics.socketsCreated = 100;
      const health = await pool.healthCheck();
      expect(health.issues.some(i => i.message.includes('timeout'))).toBe(true);
    });

    it('should note low connection reuse rate', async () => {
      pool = new HttpConnectionPool();
      pool.metrics.connectionsReused = 1;
      pool.metrics.requestsMade = 100;
      const health = await pool.healthCheck();
      expect(health.issues.some(i => i.level === 'info')).toBe(true);
    });
  });

  describe('resetMetrics', () => {
    it('should reset all metrics to initial values', () => {
      pool = new HttpConnectionPool();
      pool.metrics.requestsMade = 100;
      pool.metrics.socketErrors = 5;
      pool.resetMetrics();
      expect(pool.metrics.requestsMade).toBe(0);
      expect(pool.metrics.socketErrors).toBe(0);
    });
  });

  describe('destroy', () => {
    it('should clear interval and destroy agents', () => {
      pool = new HttpConnectionPool();
      const clearSpy = vi.spyOn(global, 'clearInterval');
      const httpDestroy = vi.spyOn(pool.httpAgent, 'destroy');
      const httpsDestroy = vi.spyOn(pool.httpsAgent, 'destroy');

      pool.destroy();

      expect(clearSpy).toHaveBeenCalled();
      expect(httpDestroy).toHaveBeenCalled();
      expect(httpsDestroy).toHaveBeenCalled();

      clearSpy.mockRestore();
      pool = null; // prevent afterEach from calling destroy again
    });
  });

  describe('cleanupStaleConnections', () => {
    it('should not throw on fresh pool', () => {
      pool = new HttpConnectionPool();
      expect(() => pool.cleanupStaleConnections()).not.toThrow();
    });
  });
});

describe('createGraphConnectionPool', () => {
  it('should create a pool with Graph API defaults', () => {
    const pool = createGraphConnectionPool();
    expect(pool).toBeInstanceOf(HttpConnectionPool);
    expect(pool.options.maxSockets).toBe(10);
    expect(pool.options.enableMetrics).toBe(true);
    pool.destroy();
  });

  it('should allow overriding defaults', () => {
    const pool = createGraphConnectionPool({ maxSockets: 25 });
    expect(pool.options.maxSockets).toBe(25);
    pool.destroy();
  });
});

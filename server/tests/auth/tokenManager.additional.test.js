import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ENCKEY_PATH = path.join(__dirname, '../../../.tokens/.enckey');

// Mock keytar as unavailable
vi.mock('keytar', () => {
  throw new Error('keytar not available');
});

// Mock node-persist with in-memory store
const memStore = {};
vi.mock('node-persist', () => ({
  default: {
    init: vi.fn(),
    setItem: vi.fn((key, val) => { memStore[key] = val; }),
    getItem: vi.fn((key) => memStore[key]),
    removeItem: vi.fn((key) => { delete memStore[key]; }),
  },
}));

describe('TokenManager - additional coverage', () => {
  let TokenManager;
  let tm;

  beforeEach(async () => {
    // Clean up key file
    try { fs.unlinkSync(ENCKEY_PATH); } catch {}
    // Clear in-memory store
    Object.keys(memStore).forEach(k => delete memStore[k]);

    vi.resetModules();
    const mod = await import('../../auth/tokenManager.js');
    TokenManager = mod.TokenManager;
    tm = new TokenManager('test-client-id');
  });

  afterEach(() => {
    try { fs.unlinkSync(ENCKEY_PATH); } catch {}
  });

  describe('encrypt / decrypt', () => {
    it('should round-trip encrypt and decrypt', async () => {
      await tm.initialize();
      const original = 'my-secret-token-12345';
      const encrypted = tm.encrypt(original);
      expect(encrypted).not.toBe(original);
      expect(encrypted).toContain(':'); // iv:ciphertext format
      const decrypted = tm.decrypt(encrypted);
      expect(decrypted).toBe(original);
    });

    it('should produce different ciphertexts for the same plaintext (random IV)', async () => {
      await tm.initialize();
      const text = 'same-text';
      const enc1 = tm.encrypt(text);
      const enc2 = tm.encrypt(text);
      expect(enc1).not.toBe(enc2);
      // But both decrypt to the same value
      expect(tm.decrypt(enc1)).toBe(text);
      expect(tm.decrypt(enc2)).toBe(text);
    });

    it('should handle empty string', async () => {
      await tm.initialize();
      const encrypted = tm.encrypt('');
      expect(tm.decrypt(encrypted)).toBe('');
    });

    it('should handle long strings', async () => {
      await tm.initialize();
      const long = 'a'.repeat(10000);
      expect(tm.decrypt(tm.encrypt(long))).toBe(long);
    });
  });

  describe('storeTokens', () => {
    it('should store tokens with fallback storage (keytar unavailable)', async () => {
      await tm.storeTokens('access-tok', 'refresh-tok', 7200);
      expect(memStore['fallback_access_token']).toBeDefined();
      expect(memStore['fallback_refresh_token']).toBeDefined();
      expect(memStore['token-metadata']).toBeDefined();
      expect(memStore['token-metadata'].accessTokenExpiry).toBeGreaterThan(Date.now());
    });

    it('should store metadata with correct expiry calculations', async () => {
      const before = Date.now();
      await tm.storeTokens('at', 'rt', 3600);
      const metadata = memStore['token-metadata'];
      expect(metadata.accessTokenExpiry).toBeGreaterThanOrEqual(before + 3600000);
      expect(metadata.refreshTokenExpiry).toBeGreaterThanOrEqual(before + 90 * 24 * 60 * 60 * 1000);
      expect(metadata.lastRefresh).toBeGreaterThanOrEqual(before);
    });

    it('should skip refresh token storage if empty', async () => {
      await tm.storeTokens('at', '', 3600);
      expect(memStore['fallback_refresh_token']).toBeUndefined();
    });

    it('should default expiresIn to 3600', async () => {
      const before = Date.now();
      await tm.storeTokens('at', 'rt');
      const metadata = memStore['token-metadata'];
      expect(metadata.accessTokenExpiry).toBeGreaterThanOrEqual(before + 3600000);
    });
  });

  describe('getAccessToken', () => {
    it('should return decrypted access token when valid', async () => {
      await tm.storeTokens('my-access-token', 'my-refresh-token', 7200);
      // Make the token not expired (the default 7200s is well within the 55-min threshold)
      const token = await tm.getAccessToken();
      expect(token).toBe('my-access-token');
    });

    it('should throw auth error when no metadata exists', async () => {
      await tm.initialize();
      // No tokens stored
      try {
        await tm.getAccessToken();
        expect.fail('Should have thrown');
      } catch (error) {
        expect(error.isError).toBe(true);
        expect(error.content[0].text).toContain('No token metadata');
      }
    });

    it('should throw needsRefresh error when token is about to expire', async () => {
      await tm.initialize();
      // Store metadata with token expiring very soon (within 55 min threshold)
      memStore['token-metadata'] = {
        accessTokenExpiry: Date.now() + 1000, // expires in 1 second
        refreshTokenExpiry: Date.now() + 86400000,
        lastRefresh: Date.now() - 3600000,
      };

      try {
        await tm.getAccessToken();
        expect.fail('Should have thrown');
      } catch (error) {
        expect(error.isError).toBe(true);
        expect(error._errorDetails.needsRefresh).toBe(true);
      }
    });

    it('should throw auth error when no token found in any store', async () => {
      await tm.initialize();
      memStore['token-metadata'] = {
        accessTokenExpiry: Date.now() + 7200000,
        refreshTokenExpiry: Date.now() + 86400000,
        lastRefresh: Date.now(),
      };
      // No actual token stored

      try {
        await tm.getAccessToken();
        expect.fail('Should have thrown');
      } catch (error) {
        expect(error.isError).toBe(true);
        expect(error.content[0].text).toContain('No access token');
      }
    });
  });

  describe('getRefreshToken', () => {
    it('should return decrypted refresh token when valid', async () => {
      await tm.storeTokens('at', 'my-refresh-token', 7200);
      const token = await tm.getRefreshToken();
      expect(token).toBe('my-refresh-token');
    });

    it('should throw when no metadata exists', async () => {
      await tm.initialize();
      try {
        await tm.getRefreshToken();
        expect.fail('Should have thrown');
      } catch (error) {
        expect(error.isError).toBe(true);
      }
    });

    it('should throw when refresh token is expired', async () => {
      await tm.initialize();
      memStore['token-metadata'] = {
        accessTokenExpiry: Date.now() + 7200000,
        refreshTokenExpiry: Date.now() - 1000, // expired
        lastRefresh: Date.now() - 86400000,
      };

      try {
        await tm.getRefreshToken();
        expect.fail('Should have thrown');
      } catch (error) {
        expect(error.isError).toBe(true);
        expect(error.content[0].text).toContain('expired');
      }
    });

    it('should throw when no refresh token found in any store', async () => {
      await tm.initialize();
      memStore['token-metadata'] = {
        accessTokenExpiry: Date.now() + 7200000,
        refreshTokenExpiry: Date.now() + 86400000,
        lastRefresh: Date.now(),
      };

      try {
        await tm.getRefreshToken();
        expect.fail('Should have thrown');
      } catch (error) {
        expect(error.isError).toBe(true);
        expect(error.content[0].text).toContain('No refresh token');
      }
    });
  });

  describe('clearTokens', () => {
    it('should remove all tokens and metadata', async () => {
      await tm.storeTokens('at', 'rt', 3600);
      expect(memStore['fallback_access_token']).toBeDefined();

      await tm.clearTokens();

      expect(memStore['fallback_access_token']).toBeUndefined();
      expect(memStore['fallback_refresh_token']).toBeUndefined();
      expect(memStore['token-metadata']).toBeUndefined();
    });
  });

  describe('PKCE verifier', () => {
    it('generateCodeVerifier should return a base64url string', () => {
      const verifier = tm.generateCodeVerifier();
      expect(typeof verifier).toBe('string');
      expect(verifier.length).toBeGreaterThan(0);
      // base64url doesn't contain + or /
      expect(verifier).not.toMatch(/[+/=]/);
    });

    it('generateCodeChallenge should return a sha256 base64url hash', () => {
      const verifier = 'test-verifier';
      const challenge = tm.generateCodeChallenge(verifier);
      expect(typeof challenge).toBe('string');
      // Verify it's the correct hash
      const expected = crypto.createHash('sha256').update(verifier).digest('base64url');
      expect(challenge).toBe(expected);
    });

    it('storePKCEVerifier and getPKCEVerifier should round-trip', async () => {
      await tm.storePKCEVerifier('my-verifier');
      expect(memStore['pkce_verifier']).toBe('my-verifier');

      const retrieved = await tm.getPKCEVerifier();
      expect(retrieved).toBe('my-verifier');
      // Should be removed after retrieval
      expect(memStore['pkce_verifier']).toBeUndefined();
    });

    it('getPKCEVerifier should throw when verifier not found', async () => {
      await tm.initialize();
      try {
        await tm.getPKCEVerifier();
        expect.fail('Should have thrown');
      } catch (error) {
        expect(error.isError).toBe(true);
        expect(error.content[0].text).toContain('PKCE verifier');
      }
    });
  });

  describe('isAuthenticated', () => {
    it('should return true when getAccessToken succeeds', async () => {
      await tm.storeTokens('at', 'rt', 7200);
      const result = await tm.isAuthenticated();
      expect(result).toBe(true);
    });

    it('should return false when getAccessToken fails', async () => {
      await tm.initialize();
      // No tokens stored
      const result = await tm.isAuthenticated();
      expect(result).toBe(false);
    });
  });

  describe('getTokenMetadata', () => {
    it('should return stored metadata', async () => {
      await tm.storeTokens('at', 'rt', 3600);
      const metadata = await tm.getTokenMetadata();
      expect(metadata).toBeDefined();
      expect(metadata.accessTokenExpiry).toBeDefined();
      expect(metadata.refreshTokenExpiry).toBeDefined();
      expect(metadata.lastRefresh).toBeDefined();
    });

    it('should return undefined when no metadata exists', async () => {
      await tm.initialize();
      const metadata = await tm.getTokenMetadata();
      expect(metadata).toBeUndefined();
    });
  });

  describe('initialize', () => {
    it('should not re-initialize if already initialized', async () => {
      const { default: storage } = await import('node-persist');
      await tm.initialize();
      const initCount = storage.init.mock.calls.length;
      await tm.initialize();
      // Should not have called init again
      expect(storage.init.mock.calls.length).toBe(initCount);
    });
  });
});

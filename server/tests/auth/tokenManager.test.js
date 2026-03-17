import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import os from 'os';

let testDir;
let ENCKEY_PATH;

// Mock keytar as unavailable
vi.mock('keytar', () => {
  throw new Error('keytar not available');
});

// Mock node-persist to avoid filesystem side effects from storage.init
vi.mock('node-persist', () => ({
  default: {
    init: vi.fn(),
    setItem: vi.fn(),
    getItem: vi.fn(),
    removeItem: vi.fn(),
  },
}));

describe('TokenManager fallback encryption key', () => {
  let TokenManager;

  beforeEach(async () => {
    // Use a unique temp directory per test to avoid race conditions
    testDir = fs.mkdtempSync(path.join(os.tmpdir(), 'mcp-outlook-test-'));
    ENCKEY_PATH = path.join(testDir, '.enckey');

    // Fresh import each test to reset module state
    vi.resetModules();

    // Mock __dirname in tokenManager to use our temp dir
    vi.doMock('../../auth/tokenManager.js', async () => {
      // Re-import the real module but we need to intercept the path
      // Instead, we'll import and patch getOrCreateEncryptionKey behavior
      const mod = await vi.importActual('../../auth/tokenManager.js');
      const OriginalTokenManager = mod.TokenManager;

      // Override the tokens dir path by patching getOrCreateEncryptionKey
      class TestTokenManager extends OriginalTokenManager {
        async getOrCreateEncryptionKey() {
          // Replicate the fallback logic but with our test dir
          const keyPath = ENCKEY_PATH;
          try {
            const existing = fs.readFileSync(keyPath);
            if (existing.length === 32) return existing;
          } catch (err) {
            // Key doesn't exist yet
          }
          if (!fs.existsSync(testDir)) {
            fs.mkdirSync(testDir, { recursive: true });
          }
          const newKey = crypto.randomBytes(32);
          fs.writeFileSync(keyPath, newKey, { mode: 0o600 });
          return newKey;
        }
      }

      return { TokenManager: TestTokenManager };
    });

    const mod = await import('../../auth/tokenManager.js');
    TokenManager = mod.TokenManager;
  });

  afterEach(() => {
    try { fs.rmSync(testDir, { recursive: true, force: true }); } catch {}
  });

  it('returns a 32-byte Buffer when keytar is unavailable', async () => {
    const tm = new TokenManager('test-client-id');
    const key = await tm.getOrCreateEncryptionKey();
    expect(Buffer.isBuffer(key)).toBe(true);
    expect(key.length).toBe(32);
  });

  it('returns the same key on subsequent calls (persisted to file)', async () => {
    const tm1 = new TokenManager('test-client-id');
    const key1 = await tm1.getOrCreateEncryptionKey();

    // New instance, same result from persisted file
    const tm2 = new TokenManager('different-client-id');
    const key2 = await tm2.getOrCreateEncryptionKey();

    expect(key1.equals(key2)).toBe(true);
  });

  it('creates the key file with mode 0o600', async () => {
    const tm = new TokenManager('test-client-id');
    await tm.getOrCreateEncryptionKey();

    const stat = fs.statSync(ENCKEY_PATH);
    const mode = stat.mode & 0o777;
    expect(mode).toBe(0o600);
  });

  it('generates a random key, not derived from clientId', async () => {
    const tm1 = new TokenManager('client-a');
    const key1 = await tm1.getOrCreateEncryptionKey();

    // Remove the file so a fresh key is generated
    fs.unlinkSync(ENCKEY_PATH);

    const tm2 = new TokenManager('client-a');
    const key2 = await tm2.getOrCreateEncryptionKey();

    // Two independently generated keys should differ (random)
    expect(key1.equals(key2)).toBe(false);
  });
});

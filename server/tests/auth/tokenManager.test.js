import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TOKENS_DIR = path.join(__dirname, '../../../.tokens');
const ENCKEY_PATH = path.join(TOKENS_DIR, '.enckey');

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
    // Clean up any persisted key file
    try { fs.unlinkSync(ENCKEY_PATH); } catch {}

    // Fresh import each test to reset module state
    vi.resetModules();
    const mod = await import('../../auth/tokenManager.js');
    TokenManager = mod.TokenManager;
  });

  afterEach(() => {
    try { fs.unlinkSync(ENCKEY_PATH); } catch {}
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

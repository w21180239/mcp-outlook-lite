// Import keytar as optional dependency - it may not be available in all environments
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let keytar: any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
let keytarImportPromise: Promise<any> | undefined;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const getKeytar = async (): Promise<any> => {
  if (keytar === undefined) {
    if (!keytarImportPromise) {
      keytarImportPromise = (async () => {
        try {
          const keytarModule = await import('keytar');
          return keytarModule.default;
        } catch (error: unknown) {
          const msg = error instanceof Error ? error.message : String(error);
          console.error('Keytar not available - using fallback token storage:', msg);
          return null;
        }
      })();
    }
    keytar = await keytarImportPromise;
  }
  return keytar;
};
import storage from 'node-persist';
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createAuthError, convertErrorToToolError } from '../utils/mcpErrorResponse.js';
import { debug, warn } from '../utils/logger.js';
import type { TokenMetadata } from '../types.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SERVICE_NAME = 'outlook-mcp';
const ENCRYPTION_KEY_ACCOUNT = 'encryption-key';
const ACCESS_TOKEN_ACCOUNT = 'access-token';
const REFRESH_TOKEN_ACCOUNT = 'refresh-token';
const TOKEN_METADATA_KEY = 'token-metadata';

export class TokenManager {
  clientId: string;
  storageInitialized: boolean;
  encryptionKey: Buffer | null;

  constructor(clientId: string) {
    this.clientId = clientId;
    this.storageInitialized = false;
    this.encryptionKey = null;
  }

  async initialize(): Promise<void> {
    if (this.storageInitialized) return;
    await storage.init({ dir: path.join(__dirname, '../../.tokens'), logging: false });
    this.encryptionKey = await this.getOrCreateEncryptionKey();
    this.storageInitialized = true;
  }

  async getOrCreateEncryptionKey(): Promise<Buffer> {
    try {
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        const existingKey = await keytarInstance.getPassword(SERVICE_NAME, ENCRYPTION_KEY_ACCOUNT);
        if (existingKey) return Buffer.from(existingKey, 'base64');
        const newKey = crypto.randomBytes(32);
        await keytarInstance.setPassword(SERVICE_NAME, ENCRYPTION_KEY_ACCOUNT, newKey.toString('base64'));
        return newKey;
      }
    } catch (_error) { warn('keytar unavailable for encryption key, using file fallback:', _error); }
    
    const tokensDir = path.join(__dirname, '../../.tokens');
    const keyPath = path.join(tokensDir, '.enckey');
    try {
      const existing = fs.readFileSync(keyPath);
      if (existing.length === 32) return existing;
    } catch (err: any) {
      if (err?.code !== 'ENOENT') {
        warn('Unexpected error reading encryption key file:', err?.message || err);
      }
    }

    if (!fs.existsSync(tokensDir)) fs.mkdirSync(tokensDir, { recursive: true });
    const newKey = crypto.randomBytes(32);
    fs.writeFileSync(keyPath, newKey, { mode: 0o600 });
    return newKey;
  }

  encrypt(text: string): string {
    const iv = crypto.randomBytes(16);
    const cipher = crypto.createCipheriv('aes-256-cbc', this.encryptionKey!, iv);
    let encrypted = cipher.update(text, 'utf8', 'hex');
    encrypted += cipher.final('hex');
    return iv.toString('hex') + ':' + encrypted;
  }

  decrypt(encryptedText: string): string {
    const parts = encryptedText.split(':');
    const iv = Buffer.from(parts.shift()!, 'hex');
    const encrypted = parts.join(':');
    const decipher = crypto.createDecipheriv('aes-256-cbc', this.encryptionKey!, iv);
    let decrypted = decipher.update(encrypted, 'hex', 'utf8');
    decrypted += decipher.final('utf8');
    return decrypted;
  }

  async storeTokens(accessToken: string, refreshToken: string, expiresIn: number = 3600): Promise<void> {
    await this.initialize();
    let usingFallback = false;
    try {
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        await keytarInstance.setPassword(SERVICE_NAME, ACCESS_TOKEN_ACCOUNT, this.encrypt(accessToken));
        if (refreshToken) await keytarInstance.setPassword(SERVICE_NAME, REFRESH_TOKEN_ACCOUNT, this.encrypt(refreshToken));
      } else {
        usingFallback = true;
        await storage.setItem('fallback_access_token', this.encrypt(accessToken));
        if (refreshToken) await storage.setItem('fallback_refresh_token', this.encrypt(refreshToken));
      }
    } catch (_error) {
      warn('keytar failed during token storage, falling back to file storage:', _error);
      usingFallback = true;
      await storage.setItem('fallback_access_token', this.encrypt(accessToken));
      if (refreshToken) await storage.setItem('fallback_refresh_token', this.encrypt(refreshToken));
    }
    const metadata: TokenMetadata = {
      accessTokenExpiry: Date.now() + (expiresIn * 1000),
      refreshTokenExpiry: Date.now() + (90 * 24 * 60 * 60 * 1000),
      lastRefresh: Date.now(),
    };
    await storage.setItem(TOKEN_METADATA_KEY, metadata);
    if (usingFallback) console.error('Tokens stored securely in encrypted file storage');
  }

  async getAccessToken(): Promise<string> {
    try {
      await this.initialize();
      const metadata = await storage.getItem(TOKEN_METADATA_KEY) as TokenMetadata | undefined;
      if (!metadata) throw createAuthError('No token metadata found', true);
      const refreshThreshold = 55 * 60 * 1000;
      const shouldRefresh = Date.now() > (metadata.accessTokenExpiry - refreshThreshold);
      if (shouldRefresh) {
        const error = createAuthError('Access token needs refresh', true) as Record<string, unknown>;
        error._errorDetails = { ...(error._errorDetails as Record<string, unknown> || {}), needsRefresh: true };
        throw error;
      }
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        try {
          const encryptedToken = await keytarInstance.getPassword(SERVICE_NAME, ACCESS_TOKEN_ACCOUNT);
          if (encryptedToken) return this.decrypt(encryptedToken);
        } catch (_e) { debug('keytar read failed for access token, trying fallback:', _e); }
      }
      const fallbackToken = await storage.getItem('fallback_access_token') as string | undefined;
      if (fallbackToken) return this.decrypt(fallbackToken);
      throw createAuthError('No access token found', true);
    } catch (error: unknown) {
      if ((error as Record<string, unknown>).isError) throw error;
      throw convertErrorToToolError(error, 'Failed to retrieve access token');
    }
  }

  async getRefreshToken(): Promise<string> {
    try {
      await this.initialize();
      const metadata = await storage.getItem(TOKEN_METADATA_KEY) as TokenMetadata | undefined;
      if (!metadata) throw createAuthError('No token metadata found', true);
      if (Date.now() > metadata.refreshTokenExpiry) throw createAuthError('Refresh token has expired', true);
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        try {
          const encryptedToken = await keytarInstance.getPassword(SERVICE_NAME, REFRESH_TOKEN_ACCOUNT);
          if (encryptedToken) return this.decrypt(encryptedToken);
        } catch (_e) { debug('keytar read failed for refresh token, trying fallback:', _e); }
      }
      const fallbackToken = await storage.getItem('fallback_refresh_token') as string | undefined;
      if (fallbackToken) return this.decrypt(fallbackToken);
      throw createAuthError('No refresh token found', true);
    } catch (error: unknown) {
      if ((error as Record<string, unknown>).isError) throw error;
      throw convertErrorToToolError(error, 'Failed to retrieve refresh token');
    }
  }

  async clearTokens(): Promise<void> {
    await this.initialize();
    const keytarInstance = await getKeytar();
    if (keytarInstance) {
      try {
        await keytarInstance.deletePassword(SERVICE_NAME, ACCESS_TOKEN_ACCOUNT);
        await keytarInstance.deletePassword(SERVICE_NAME, REFRESH_TOKEN_ACCOUNT);
      } catch (_e) { debug('keytar delete failed during token cleanup:', _e); }
    }
    await storage.removeItem('fallback_access_token');
    await storage.removeItem('fallback_refresh_token');
    await storage.removeItem(TOKEN_METADATA_KEY);
  }

  generateCodeVerifier(): string { return crypto.randomBytes(32).toString('base64url'); }
  generateCodeChallenge(verifier: string): string { return crypto.createHash('sha256').update(verifier).digest('base64url'); }

  async storePKCEVerifier(verifier: string): Promise<void> {
    await this.initialize();
    await storage.setItem('pkce_verifier', verifier);
  }

  async getPKCEVerifier(): Promise<string> {
    try {
      await this.initialize();
      const verifier = await storage.getItem('pkce_verifier') as string | undefined;
      await storage.removeItem('pkce_verifier');
      if (!verifier) throw createAuthError('PKCE verifier not found or expired', true);
      return verifier;
    } catch (error: unknown) {
      if ((error as Record<string, unknown>).isError) throw error;
      throw convertErrorToToolError(error, 'Failed to retrieve PKCE verifier');
    }
  }

  async isAuthenticated(): Promise<boolean> {
    try { await this.getAccessToken(); return true; } catch (_e) { return false; }
  }

  async getTokenMetadata(): Promise<TokenMetadata | undefined> {
    await this.initialize();
    return await storage.getItem(TOKEN_METADATA_KEY) as TokenMetadata | undefined;
  }
}

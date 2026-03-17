// OAuth 2.0 authentication is handled manually with PKCE flow
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenManager } from './tokenManager.js';
import { authConfig } from './config.js';
import { GraphApiClient } from '../graph/graphClient.js';
import { createAuthError, convertErrorToToolError } from '../utils/mcpErrorResponse.js';
import { openBrowser } from './browserLauncher.js';
import { isHeadlessEnvironment, authenticateWithDeviceCode as performDeviceCodeAuth } from './deviceCodeFlow.js';
import { getSuccessPage, getErrorPage, getFailurePage } from './templates.js';
import http from 'http';
import url from 'url';
import crypto from 'crypto';
import type { AddressInfo } from 'net';
import type { MCPResponse } from '../types.js';

interface TokenResponse {
  access_token: string;
  refresh_token: string;
  expires_in: number;
}

interface AuthResultSuccess {
  success: true;
  user: {
    id: string;
    displayName: string;
    mail: string;
  };
}

interface AuthResultFailure {
  success: false;
  error: MCPResponse;
}

type AuthResult = AuthResultSuccess | AuthResultFailure;

export class OutlookAuthManager {
  clientId: string;
  tenantId: string;
  tokenManager: TokenManager;
  graphClient: Client | null;
  graphApiClient: GraphApiClient | null;
  isAuthenticated: boolean;
  authenticationRecord: unknown;
  lastUsedPort: number | null;

  constructor(clientId: string, tenantId: string) {
    this.clientId = clientId;
    this.tenantId = tenantId;
    this.tokenManager = new TokenManager(clientId);
    this.graphClient = null;
    this.graphApiClient = null;
    this.isAuthenticated = false;
    this.authenticationRecord = null;
    this.lastUsedPort = null;
  }

  async authenticate(): Promise<AuthResult> {
    try {
      const isTokenValid = await this.tokenManager.isAuthenticated();

      if (isTokenValid) {
        await this.initializeGraphClient();
        return await this.validateAuthentication();
      }

      // Token expired — try refreshing before falling back to interactive login
      try {
        const metadata = await this.tokenManager.getTokenMetadata();
        if (metadata) {
          console.error('Access token expired, attempting silent refresh...');
          await this.refreshAccessToken();
          console.error('Silent refresh succeeded — no browser login needed.');
          return await this.validateAuthentication();
        }
      } catch (refreshError: unknown) {
        const msg = refreshError instanceof Error ? refreshError.message : String(refreshError);
        console.error('Silent refresh failed, falling back to interactive login:', msg);
      }

      // Use interactive authentication with PKCE for delegated access
      return await this.authenticateInteractive();
    } catch (error: unknown) {
      console.error('Authentication error:', error);
      this.isAuthenticated = false;
      const err = error as Record<string, unknown>;
      if (err.isError) {
        // Already an MCP error, return as-is
        return {
          success: false,
          error: err as unknown as MCPResponse,
        };
      }
      const msg = error instanceof Error ? error.message : String(error);
      return {
        success: false,
        error: createAuthError(msg, true),
      };
    }
  }

  async authenticateViaDeviceCode(): Promise<AuthResult> {
    const tokenResponse = await performDeviceCodeAuth(this.clientId, this.tenantId);

    await this.tokenManager.storeTokens(
      tokenResponse.access_token,
      tokenResponse.refresh_token ?? '',
      tokenResponse.expires_in
    );

    await this.initializeGraphClient();
    return await this.validateAuthentication();
  }

  async authenticateInteractive(): Promise<AuthResult> {
    // Use device code flow in headless environments
    if (isHeadlessEnvironment()) {
      return await this.authenticateViaDeviceCode();
    }

    const codeVerifier = this.tokenManager.generateCodeVerifier();
    const codeChallenge = this.tokenManager.generateCodeChallenge(codeVerifier);
    await this.tokenManager.storePKCEVerifier(codeVerifier);

    const authorizationCode = await this.getAuthorizationCode(codeChallenge);

    if (!authorizationCode) {
      return {
        success: false,
        error: createAuthError('Failed to get authorization code', true),
      };
    }

    const tokenResponse = await this.exchangeCodeForToken(authorizationCode as string);

    await this.tokenManager.storeTokens(
      tokenResponse.access_token,
      tokenResponse.refresh_token,
      tokenResponse.expires_in
    );

    await this.initializeGraphClient();
    return await this.validateAuthentication();
  }

  async getAuthorizationCode(codeChallenge: string): Promise<unknown> {
    return new Promise((resolve, reject) => {
      const state = crypto.randomBytes(16).toString('hex');
      const authUrl = new URL(authConfig.oauth.authorizeUrl(this.tenantId));

      authUrl.searchParams.append('client_id', this.clientId);
      authUrl.searchParams.append('response_type', 'code');
      authUrl.searchParams.append('scope', authConfig.oauth.scope);
      authUrl.searchParams.append('state', state);
      authUrl.searchParams.append('code_challenge', codeChallenge);
      authUrl.searchParams.append('code_challenge_method', 'S256');
      authUrl.searchParams.append('prompt', 'select_account');


      const server = http.createServer(async (req, res) => {
        const parsedUrl = url.parse(req.url!, true);

        if (parsedUrl.pathname === '/callback') {
          const code = parsedUrl.query.code;
          const returnedState = parsedUrl.query.state;

          if (returnedState !== state) {
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end(getErrorPage());
            server.close();
            reject(createAuthError('State mismatch - possible CSRF attack', false));
            return;
          }

          if (code) {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(getSuccessPage());
            server.close();
            resolve(code);
          } else {
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end(getFailurePage());
            server.close();
            reject(createAuthError('No authorization code received', true));
          }
        }
      });

      server.listen(0, () => {
        const addr = server.address() as AddressInfo;
        const port = addr.port;
        this.lastUsedPort = port;
        console.error(`\nListening for authentication callback on port ${port}...`);

        // Update redirect URI with actual port
        authUrl.searchParams.set('redirect_uri', `http://localhost:${port}/callback`);

        console.error(`\nOpening your browser for Microsoft account selection...`);
        console.error(`If the browser doesn't open automatically, please visit:`);
        console.error(authUrl.toString());

        // Attempt to open the browser automatically
        openBrowser(authUrl.toString());
      });

      const timeoutHandle = setTimeout(() => {
        server.close();
        reject(createAuthError('Authentication timeout - please try again', true));
      }, 5 * 60 * 1000); // 5 minute timeout

      // Clean up timeout when auth completes (success or failure)
      const originalResolve = resolve;
      const originalReject = reject;
      resolve = ((value: unknown) => { clearTimeout(timeoutHandle); originalResolve(value); }) as any;
      reject = ((reason: unknown) => { clearTimeout(timeoutHandle); originalReject(reason); }) as any;
    });
  }

  async exchangeCodeForToken(code: string): Promise<TokenResponse> {
    const codeVerifier = await this.tokenManager.getPKCEVerifier();

    const tokenUrl = authConfig.oauth.tokenUrl(this.tenantId);
    const params = new URLSearchParams({
      client_id: this.clientId,
      scope: authConfig.oauth.scope,
      code: code,
      redirect_uri: `http://localhost:${this.lastUsedPort}/callback`,
      grant_type: 'authorization_code',
      code_verifier: codeVerifier,
    });

    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params.toString(),
    });

    if (!response.ok) {
      const errorBody = await response.text();
      console.error('Token exchange failed (raw):', errorBody);
      throw createAuthError(`Token exchange failed (HTTP ${response.status}). Check server logs for details.`, true);
    }

    return await response.json() as TokenResponse;
  }

  async refreshAccessToken(): Promise<boolean> {
    try {
      const refreshToken = await this.tokenManager.getRefreshToken();
      const tokenUrl = authConfig.oauth.tokenUrl(this.tenantId);

      const params = new URLSearchParams({
        client_id: this.clientId,
        scope: authConfig.oauth.scope,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
      });

      const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params.toString(),
      });

      if (!response.ok) {
        const errorBody = await response.text();
        console.error('Token refresh failed (raw):', errorBody);
        throw createAuthError(`Token refresh failed (HTTP ${response.status}). Check server logs for details.`, true);
      }

      const tokenResponse = await response.json() as TokenResponse;

      await this.tokenManager.storeTokens(
        tokenResponse.access_token,
        tokenResponse.refresh_token || refreshToken,
        tokenResponse.expires_in
      );

      await this.initializeGraphClient();
      return true;
    } catch (error: unknown) {
      console.error('Token refresh failed:', error);
      await this.tokenManager.clearTokens();
      const err = error as Record<string, unknown>;
      if (err.isError) {
        throw error;
      }
      throw convertErrorToToolError(error, 'Token refresh failed');
    }
  }

  async initializeGraphClient(): Promise<void> {
    const authProvider = {
      getAccessToken: async (): Promise<string> => {
        try {
          return await this.tokenManager.getAccessToken();
        } catch (error: unknown) {
          const err = error as Error;
          if (err.message.includes('needs refresh')) {
            await this.refreshAccessToken();
            return await this.tokenManager.getAccessToken();
          }
          throw error;
        }
      },
    };

    this.graphClient = Client.init({
      authProvider: (done) => {
        authProvider.getAccessToken()
          .then(token => done(null, token))
          .catch(error => done(error, null));
      },
      defaultVersion: 'v1.0',
    });

    // Initialize the enhanced GraphApiClient
    this.graphApiClient = new GraphApiClient(this);
    await this.graphApiClient.initialize();
  }

  async validateAuthentication(): Promise<AuthResultSuccess> {
    try {
      const user = await this.graphClient!.api('/me').get();
      this.isAuthenticated = true;

      return {
        success: true,
        user: {
          id: user.id,
          displayName: user.displayName,
          mail: user.mail || user.userPrincipalName,
        },
      };
    } catch (error: unknown) {
      this.isAuthenticated = false;
      const err = error as Record<string, unknown>;
      if (err.isError) {
        throw error;
      }
      throw convertErrorToToolError(error, 'User validation failed');
    }
  }

  async ensureAuthenticated(): Promise<Client> {
    if (!this.isAuthenticated || !this.graphClient) {
      const result = await this.authenticate();
      if (!result.success) {
        const failResult = result as AuthResultFailure;
        const err = failResult.error as unknown as Record<string, unknown>;
        if (err.isError) {
          throw failResult.error;
        }
        throw createAuthError(`Authentication failed: ${failResult.error}`, true);
      }
    }

    try {
      await this.tokenManager.getAccessToken();
    } catch (error: unknown) {
      const err = error as Record<string, unknown>;
      if (err.isError) {
        if (err._errorDetails && (err._errorDetails as Record<string, unknown>).needsRefresh) {
          await this.refreshAccessToken();
        } else {
          throw error;
        }
      } else if (error instanceof Error && error.message.includes('needs refresh')) {
        await this.refreshAccessToken();
      } else {
        throw convertErrorToToolError(error, 'Token validation failed');
      }
    }

    return this.graphClient!;
  }

  getGraphClient(): Client {
    if (!this.graphClient) {
      throw createAuthError('Not authenticated. Call authenticate() first.', true);
    }
    return this.graphClient;
  }

  getGraphApiClient(): GraphApiClient {
    if (!this.graphApiClient) {
      throw createAuthError('Not authenticated. Call authenticate() first.', true);
    }
    return this.graphApiClient;
  }

  async logout(): Promise<void> {
    await this.tokenManager.clearTokens();
    this.graphClient = null;
    this.graphApiClient = null;
    this.isAuthenticated = false;
    this.authenticationRecord = null;
  }
}

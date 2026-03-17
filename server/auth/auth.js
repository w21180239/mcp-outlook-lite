// OAuth 2.0 authentication is handled manually with PKCE flow
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenManager } from './tokenManager.js';
import { authConfig } from './config.js';
import { GraphApiClient } from '../graph/graphClient.js';
import { createAuthError, convertErrorToToolError } from '../utils/mcpErrorResponse.js';
import { openBrowser } from './browserLauncher.js';
import { getSuccessPage, getErrorPage, getFailurePage } from './templates.js';
import http from 'http';
import url from 'url';
import crypto from 'crypto';

export class OutlookAuthManager {
  constructor(clientId, tenantId) {
    this.clientId = clientId;
    this.tenantId = tenantId;
    this.tokenManager = new TokenManager(clientId);
    this.graphClient = null;
    this.graphApiClient = null;
    this.isAuthenticated = false;
    this.authenticationRecord = null;
    this.lastUsedPort = null;
  }

  async authenticate() {
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
      } catch (refreshError) {
        console.error('Silent refresh failed, falling back to interactive login:', refreshError.message || refreshError);
      }

      // Use interactive authentication with PKCE for delegated access
      return await this.authenticateInteractive();
    } catch (error) {
      console.error('Authentication error:', error);
      this.isAuthenticated = false;
      if (error.isError) {
        // Already an MCP error, return as-is
        return {
          success: false,
          error: error,
        };
      }
      return {
        success: false,
        error: createAuthError(error.message, true),
      };
    }
  }

  async authenticateInteractive() {
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

    const tokenResponse = await this.exchangeCodeForToken(authorizationCode);

    await this.tokenManager.storeTokens(
      tokenResponse.access_token,
      tokenResponse.refresh_token,
      tokenResponse.expires_in
    );

    await this.initializeGraphClient();
    return await this.validateAuthentication();
  }

  async getAuthorizationCode(codeChallenge) {
    return new Promise((resolve, reject) => {
      const state = crypto.randomBytes(16).toString('hex');
      const authUrl = new URL(authConfig.oauth.authorizeUrl(this.tenantId));

      authUrl.searchParams.append('client_id', this.clientId);
      authUrl.searchParams.append('response_type', 'code');
      // redirect_uri will be set after server starts and we know the port
      authUrl.searchParams.append('scope', authConfig.oauth.scope);
      authUrl.searchParams.append('state', state);
      authUrl.searchParams.append('code_challenge', codeChallenge);
      authUrl.searchParams.append('code_challenge_method', 'S256');
      authUrl.searchParams.append('prompt', 'select_account');


      const server = http.createServer(async (req, res) => {
        const parsedUrl = url.parse(req.url, true);

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
        const port = server.address().port;
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

      setTimeout(() => {
        server.close();
        reject(createAuthError('Authentication timeout - please try again', true));
      }, 5 * 60 * 1000); // 5 minute timeout
    });
  }

  async exchangeCodeForToken(code) {
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
      const error = await response.text();
      throw createAuthError(`Token exchange failed: ${error}`, true);
    }

    return await response.json();
  }

  async refreshAccessToken() {
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
        const error = await response.text();
        throw createAuthError(`Token refresh failed: ${error}`, true);
      }

      const tokenResponse = await response.json();

      await this.tokenManager.storeTokens(
        tokenResponse.access_token,
        tokenResponse.refresh_token || refreshToken,
        tokenResponse.expires_in
      );

      await this.initializeGraphClient();
      return true;
    } catch (error) {
      console.error('Token refresh failed:', error);
      await this.tokenManager.clearTokens();
      if (error.isError) {
        // Already an MCP error, re-throw as-is
        throw error;
      }
      throw convertErrorToToolError(error, 'Token refresh failed');
    }
  }

  async initializeGraphClient() {
    const authProvider = {
      getAccessToken: async () => {
        try {
          return await this.tokenManager.getAccessToken();
        } catch (error) {
          if (error.message.includes('needs refresh')) {
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

  async validateAuthentication() {
    try {
      const user = await this.graphClient.api('/me').get();
      this.isAuthenticated = true;

      return {
        success: true,
        user: {
          id: user.id,
          displayName: user.displayName,
          mail: user.mail || user.userPrincipalName,
        },
      };
    } catch (error) {
      this.isAuthenticated = false;
      if (error.isError) {
        // Already an MCP error, re-throw as-is
        throw error;
      }
      throw convertErrorToToolError(error, 'User validation failed');
    }
  }

  async ensureAuthenticated() {
    if (!this.isAuthenticated || !this.graphClient) {
      const result = await this.authenticate();
      if (!result.success) {
        if (result.error.isError) {
          // Already an MCP error, re-throw as-is
          throw result.error;
        }
        throw createAuthError(`Authentication failed: ${result.error}`, true);
      }
    }

    try {
      await this.tokenManager.getAccessToken();
    } catch (error) {
      if (error.isError) {
        // Handle MCP errors from token manager
        if (error._errorDetails && error._errorDetails.needsRefresh) {
          await this.refreshAccessToken();
        } else {
          throw error;
        }
      } else if (error.message.includes('needs refresh')) {
        await this.refreshAccessToken();
      } else {
        throw convertErrorToToolError(error, 'Token validation failed');
      }
    }

    return this.graphClient;
  }

  getGraphClient() {
    if (!this.graphClient) {
      throw createAuthError('Not authenticated. Call authenticate() first.', true);
    }
    return this.graphClient;
  }

  getGraphApiClient() {
    if (!this.graphApiClient) {
      throw createAuthError('Not authenticated. Call authenticate() first.', true);
    }
    return this.graphApiClient;
  }

  async logout() {
    await this.tokenManager.clearTokens();
    this.graphClient = null;
    this.graphApiClient = null;
    this.isAuthenticated = false;
    this.authenticationRecord = null;
  }
}
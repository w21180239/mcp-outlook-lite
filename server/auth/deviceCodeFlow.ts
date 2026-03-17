import { authConfig } from './config.js';
import { createAuthError } from '../utils/mcpErrorResponse.js';

interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message: string;
}

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  token_type: string;
}

/**
 * Detect if running in a headless environment (no browser available).
 */
export function isHeadlessEnvironment(): boolean {
  // Explicit opt-in via env var
  if (process.env.MCP_OUTLOOK_DEVICE_CODE === '1' || process.env.MCP_OUTLOOK_DEVICE_CODE === 'true') {
    return true;
  }

  // SSH session
  if (process.env.SSH_CLIENT || process.env.SSH_TTY || process.env.SSH_CONNECTION) {
    return true;
  }

  // No display on Linux
  if (process.platform === 'linux' && !process.env.DISPLAY && !process.env.WAYLAND_DISPLAY) {
    return true;
  }

  // Docker / container
  if (process.env.container || process.env.DOCKER_CONTAINER) {
    return true;
  }

  return false;
}

/**
 * Request a device code from Azure AD.
 */
async function requestDeviceCode(clientId: string, tenantId: string): Promise<DeviceCodeResponse> {
  const deviceCodeUrl = authConfig.oauth.deviceCodeUrl(tenantId);

  const params = new URLSearchParams({
    client_id: clientId,
    scope: authConfig.oauth.scope,
  });

  const response = await fetch(deviceCodeUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });

  if (!response.ok) {
    const errorBody = await response.text();
    console.error('Device code request failed:', errorBody);
    throw createAuthError(`Device code request failed (HTTP ${response.status})`, true);
  }

  return await response.json() as DeviceCodeResponse;
}

/**
 * Poll the token endpoint until the user completes authentication.
 */
async function pollForToken(
  clientId: string,
  tenantId: string,
  deviceCode: string,
  interval: number,
  expiresIn: number
): Promise<TokenResponse> {
  const tokenUrl = authConfig.oauth.tokenUrl(tenantId);
  const deadline = Date.now() + expiresIn * 1000;

  const params = new URLSearchParams({
    client_id: clientId,
    grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
    device_code: deviceCode,
  });

  while (Date.now() < deadline) {
    await new Promise(resolve => setTimeout(resolve, interval * 1000));

    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });

    const body = await response.json() as Record<string, unknown>;

    if (response.ok) {
      return body as unknown as TokenResponse;
    }

    const error = body.error as string;
    if (error === 'authorization_pending') {
      // User hasn't completed auth yet, keep polling
      continue;
    } else if (error === 'slow_down') {
      // Increase polling interval by 5 seconds
      interval += 5;
      continue;
    } else if (error === 'authorization_declined') {
      throw createAuthError('User declined the authentication request', true);
    } else if (error === 'expired_token') {
      throw createAuthError('Device code expired. Please try again.', true);
    } else {
      throw createAuthError(`Device code auth failed: ${error}`, true);
    }
  }

  throw createAuthError('Device code authentication timed out', true);
}

/**
 * Perform device code flow authentication.
 * Prints instructions to stderr and polls for completion.
 */
export async function authenticateWithDeviceCode(
  clientId: string,
  tenantId: string
): Promise<TokenResponse> {
  console.error('\n=== Device Code Authentication ===');
  console.error('Browser-based login is not available in this environment.');
  console.error('Using device code flow instead.\n');

  const deviceCodeResponse = await requestDeviceCode(clientId, tenantId);

  // Print instructions to stderr (stdout is reserved for MCP protocol)
  console.error(deviceCodeResponse.message);
  console.error(`\nCode: ${deviceCodeResponse.user_code}`);
  console.error(`URL:  ${deviceCodeResponse.verification_uri}`);
  console.error('\nWaiting for authentication...\n');

  const tokenResponse = await pollForToken(
    clientId,
    tenantId,
    deviceCodeResponse.device_code,
    deviceCodeResponse.interval,
    deviceCodeResponse.expires_in
  );

  console.error('Device code authentication successful!\n');
  return tokenResponse;
}

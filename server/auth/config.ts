interface OAuthConfig {
  authorizeUrl: (tenantId: string) => string;
  tokenUrl: (tenantId: string) => string;
  deviceCodeUrl: (tenantId: string) => string;
  scope: string;
  redirectUri: string;
}

interface AuthConfigType {
  oauth: OAuthConfig;
  token: {
    accessTokenTTL: number;
    refreshThreshold: number;
    refreshTokenTTL: number;
  };
  retry: {
    maxAttempts: number;
    initialDelay: number;
    maxDelay: number;
    backoffMultiplier: number;
  };
  security: {
    usePKCE: boolean;
    encryptTokens: boolean;
    auditLogging: boolean;
  };
}

export const authConfig: AuthConfigType = {
  oauth: {
    authorizeUrl: (tenantId: string) =>
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
    tokenUrl: (tenantId: string) =>
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    deviceCodeUrl: (tenantId: string) =>
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/devicecode`,
    scope: [
      'Mail.Read',
      'Mail.ReadWrite',
      'Mail.Send',
      'Calendars.Read',
      'Calendars.ReadWrite',
      'Contacts.Read',
      'Contacts.ReadWrite',
      'Tasks.Read',
      'Tasks.ReadWrite',
      'User.Read',
      'MailboxSettings.ReadWrite',
      'Sites.Read.All',
      'Sites.ReadWrite.All',
      'Files.Read.All',
      'Files.ReadWrite.All',
      'offline_access',
    ].join(' '),
    redirectUri: process.env.MCP_OUTLOOK_REDIRECT_URI || 'http://localhost:0/callback',
  },

  token: {
    accessTokenTTL: 60 * 60 * 1000,
    refreshThreshold: 55 * 60 * 1000,
    refreshTokenTTL: 90 * 24 * 60 * 60 * 1000,
  },

  retry: {
    maxAttempts: 3,
    initialDelay: 1000,
    maxDelay: 30000,
    backoffMultiplier: 2,
  },

  security: {
    usePKCE: true,
    encryptTokens: true,
    auditLogging: true,
  },
};

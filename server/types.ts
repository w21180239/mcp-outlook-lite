import type { OutlookAuthManager } from './auth/auth.js';

export interface MCPResponse {
  content: Array<{ type: string; text: string }>;
  isError?: boolean;
  _errorDetails?: Record<string, unknown>;
}

export interface MCPErrorResponse extends MCPResponse {
  isError: true;
}

export interface AuthResult {
  success: boolean;
  user?: {
    id: string;
    displayName: string;
    mail: string;
  };
  error?: MCPErrorResponse;
}

export interface TokenMetadata {
  accessTokenExpiry: number;
  refreshTokenExpiry: number;
  lastRefresh: number;
}

export interface ToolHandler {
  (authManager: OutlookAuthManager, args: Record<string, unknown>): Promise<MCPResponse>;
}

export interface GraphApiResponse {
  value?: unknown[];
  [key: string]: unknown;
}

export interface ToolSchema {
  name: string;
  description: string;
  inputSchema: {
    type: string;
    properties?: Record<string, unknown>;
    required?: string[];
  };
}

export interface ToolSchemaMap {
  [key: string]: ToolSchema;
}

export interface DecodedContent {
  type: string;
  content: unknown;
  size?: number;
  sizeFormatted?: string;
  encoding?: string;
  contentBytes?: string;
  note?: string;
  error?: string;
}

export interface FileInfo {
  success: boolean;
  filePath?: string;
  filename?: string;
  originalFilename?: string;
  size?: number;
  sizeFormatted?: string;
  originalSize?: number;
  mimeType?: string | null;
  encoding?: string;
  createdAt?: string;
  workDirectory?: string;
  note?: string;
  error?: string;
}

export interface StyledBody {
  content: string;
  type: string;
}

export interface CacheStyling {
  fontFamily?: string | null;
  fontSize?: string | null;
  fontColor?: string | null;
  timestamp?: number;
}

export interface CacheSignature {
  signature: string;
  timestamp: number;
}

export interface PromptDefinition {
  name: string;
  description: string;
  arguments: Array<{
    name: string;
    description: string;
    required?: boolean;
  }>;
}

export interface PromptResult {
  messages: Array<{
    role: string;
    content: {
      type: string;
      text: string;
    };
  }>;
}

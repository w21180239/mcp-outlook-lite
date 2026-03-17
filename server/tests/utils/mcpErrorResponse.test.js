import { describe, it, expect } from 'vitest';
import {
  convertErrorToToolError,
  createToolError,
  createProtocolError,
  ErrorCodes,
  createValidationError,
  createAuthError,
  createRateLimitError,
  createServiceUnavailableError,
} from '../../utils/mcpErrorResponse.js';

describe('convertErrorToToolError', () => {
  it('should not include stack trace in error details', () => {
    const error = new Error('something broke');
    const result = convertErrorToToolError(error, 'test context');

    expect(result._errorDetails).toBeDefined();
    expect(result._errorDetails).not.toHaveProperty('stack');
  });

  it('should still include the error name in details', () => {
    const error = new TypeError('bad type');
    const result = convertErrorToToolError(error, 'test context');

    expect(result._errorDetails.name).toBe('TypeError');
  });

  it('returns MCP error objects as-is', () => {
    const mcpError = { isError: true, content: [{ type: 'text', text: 'already an error' }] };
    const result = convertErrorToToolError(mcpError);
    expect(result).toBe(mcpError);
  });

  it('handles non-Error objects gracefully', () => {
    const result = convertErrorToToolError('string error', 'context');
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('context');
  });

  it('handles null error', () => {
    const result = convertErrorToToolError(null, 'context');
    expect(result.isError).toBe(true);
  });

  it('preserves statusCode from error', () => {
    const error = new Error('Not found');
    error.statusCode = 404;
    const result = convertErrorToToolError(error);
    expect(result._errorDetails.statusCode).toBe(404);
  });

  it('preserves error code', () => {
    const error = new Error('Token expired');
    error.code = 'InvalidAuthenticationToken';
    const result = convertErrorToToolError(error);
    expect(result._errorDetails.code).toBe('InvalidAuthenticationToken');
  });
});

describe('createToolError', () => {
  it('creates error with message and isError flag', () => {
    const result = createToolError('Something went wrong');
    expect(result.isError).toBe(true);
    expect(result.content[0].type).toBe('text');
    expect(result.content[0].text).toBe('Something went wrong');
  });

  it('includes details when provided', () => {
    const result = createToolError('Error', { code: 'E001' });
    expect(result._errorDetails.code).toBe('E001');
  });

  it('omits _errorDetails when details is empty', () => {
    const result = createToolError('Error');
    expect(result).not.toHaveProperty('_errorDetails');
  });
});

describe('createProtocolError', () => {
  it('creates JSON-RPC 2.0 error', () => {
    const result = createProtocolError(-32600, 'Invalid Request');
    expect(result.jsonrpc).toBe('2.0');
    expect(result.error.code).toBe(-32600);
    expect(result.error.message).toBe('Invalid Request');
  });

  it('includes data when provided', () => {
    const result = createProtocolError(-32602, 'Invalid params', { param: 'foo' });
    expect(result.error.data.param).toBe('foo');
  });

  it('omits data when null', () => {
    const result = createProtocolError(-32700, 'Parse error');
    expect(result.error).not.toHaveProperty('data');
  });
});

describe('ErrorCodes', () => {
  it('has standard JSON-RPC error codes', () => {
    expect(ErrorCodes.PARSE_ERROR).toBe(-32700);
    expect(ErrorCodes.INVALID_REQUEST).toBe(-32600);
    expect(ErrorCodes.METHOD_NOT_FOUND).toBe(-32601);
    expect(ErrorCodes.INTERNAL_ERROR).toBe(-32603);
  });
});

describe('createValidationError', () => {
  it('creates validation error with param name and reason', () => {
    const result = createValidationError('email', 'must be a valid email address');
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('email');
    expect(result.content[0].text).toContain('must be a valid email address');
    expect(result._errorDetails.type).toBe('validation');
    expect(result._errorDetails.parameter).toBe('email');
  });
});

describe('createAuthError', () => {
  it('creates auth error with retryable suggestion', () => {
    const result = createAuthError('Token expired', true);
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Authentication failed');
    expect(result._errorDetails.retryable).toBe(true);
    expect(result._errorDetails.suggestion).toContain('re-authenticate');
  });

  it('creates non-retryable auth error', () => {
    const result = createAuthError('Bad credentials', false);
    expect(result._errorDetails.retryable).toBe(false);
    expect(result._errorDetails.suggestion).toContain('Check your credentials');
  });
});

describe('createRateLimitError', () => {
  it('creates rate limit error with retry time', () => {
    const result = createRateLimitError(30);
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('30 seconds');
    expect(result._errorDetails.retryAfter).toBe(30);
    expect(result._errorDetails.retryable).toBe(true);
  });
});

describe('createServiceUnavailableError', () => {
  it('creates service unavailable error with default service name', () => {
    const result = createServiceUnavailableError();
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Microsoft Graph API');
    expect(result._errorDetails.retryable).toBe(true);
  });

  it('uses custom service name', () => {
    const result = createServiceUnavailableError('SharePoint');
    expect(result.content[0].text).toContain('SharePoint');
  });
});

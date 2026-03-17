import { describe, it, expect } from 'vitest';
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';

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
});

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { debug } from '../../utils/logger.js';

describe('debug logger', () => {
  let consoleSpy;

  beforeEach(() => {
    consoleSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    consoleSpy.mockRestore();
    delete process.env.DEBUG;
  });

  it('should log when DEBUG is set', () => {
    process.env.DEBUG = '1';
    debug('test message');
    expect(consoleSpy).toHaveBeenCalledWith('test message');
  });

  it('should NOT log when DEBUG is unset', () => {
    delete process.env.DEBUG;
    debug('test message');
    expect(consoleSpy).not.toHaveBeenCalled();
  });

  it('should pass multiple arguments', () => {
    process.env.DEBUG = '1';
    debug('msg', 42, { key: 'val' });
    expect(consoleSpy).toHaveBeenCalledWith('msg', 42, { key: 'val' });
  });
});

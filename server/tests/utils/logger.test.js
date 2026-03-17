import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { debug, warn } from '../../utils/logger.js';

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

describe('warn logger', () => {
  let consoleSpy;

  beforeEach(() => {
    consoleSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    consoleSpy.mockRestore();
    delete process.env.DEBUG;
  });

  it('should always log regardless of DEBUG', () => {
    delete process.env.DEBUG;
    warn('warning message');
    expect(consoleSpy).toHaveBeenCalledWith('[WARN]', 'warning message');
  });

  it('should pass multiple arguments with [WARN] prefix', () => {
    warn('msg', 42);
    expect(consoleSpy).toHaveBeenCalledWith('[WARN]', 'msg', 42);
  });
});

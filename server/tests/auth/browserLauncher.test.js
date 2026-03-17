import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { execFile } from 'child_process';

vi.mock('child_process', () => ({
  execFile: vi.fn(),
}));

const { openBrowser } = await import('../../auth/browserLauncher.js');

describe('browserLauncher', () => {
  const originalPlatform = process.platform;

  beforeEach(() => {
    vi.clearAllMocks();
  });

  afterEach(() => {
    Object.defineProperty(process, 'platform', { value: originalPlatform });
  });

  it('should use "open" on darwin', () => {
    Object.defineProperty(process, 'platform', { value: 'darwin' });
    openBrowser('https://example.com');
    expect(execFile).toHaveBeenCalledWith('open', ['https://example.com'], expect.any(Function));
  });

  it('should use "cmd /c start" on win32', () => {
    Object.defineProperty(process, 'platform', { value: 'win32' });
    openBrowser('https://example.com');
    expect(execFile).toHaveBeenCalledWith('cmd', ['/c', 'start', '', 'https://example.com'], expect.any(Function));
  });

  it('should use "xdg-open" on linux (default)', () => {
    Object.defineProperty(process, 'platform', { value: 'linux' });
    openBrowser('https://example.com');
    expect(execFile).toHaveBeenCalledWith('xdg-open', ['https://example.com'], expect.any(Function));
  });

  it('should log warning when execFile errors', () => {
    Object.defineProperty(process, 'platform', { value: 'darwin' });
    const consoleSpy = vi.spyOn(console, 'error').mockImplementation(() => {});

    openBrowser('https://example.com');

    // Get the callback and invoke it with an error
    const callback = execFile.mock.calls[0][2];
    callback(new Error('spawn failed'));

    expect(consoleSpy).toHaveBeenCalledWith(expect.stringContaining('Could not open browser'));
    consoleSpy.mockRestore();
  });
});

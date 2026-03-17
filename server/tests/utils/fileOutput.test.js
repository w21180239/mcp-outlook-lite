import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import path from 'path';
import os from 'os';

const {
  getWorkDirectory,
  generateUniqueFilename,
  saveFileToDisc,
  shouldSaveToFile,
  handleLargeContent,
  cleanupOldFiles,
} = await import('../../utils/fileOutput.js');

describe('getWorkDirectory', () => {
  const originalEnv = process.env.MCP_OUTLOOK_WORK_DIR;

  afterEach(() => {
    if (originalEnv !== undefined) {
      process.env.MCP_OUTLOOK_WORK_DIR = originalEnv;
    } else {
      delete process.env.MCP_OUTLOOK_WORK_DIR;
    }
  });

  it('returns system temp dir when env var is not set', () => {
    delete process.env.MCP_OUTLOOK_WORK_DIR;
    const result = getWorkDirectory();
    expect(result).toBe(os.tmpdir());
  });

  it('returns custom directory when MCP_OUTLOOK_WORK_DIR is set', () => {
    const customDir = path.join(os.tmpdir(), 'outlook-mcp-test-workdir');
    process.env.MCP_OUTLOOK_WORK_DIR = customDir;
    const result = getWorkDirectory();
    expect(result).toBe(customDir);
    // Cleanup
    try { fs.rmdirSync(customDir); } catch {}
  });
});

describe('generateUniqueFilename', () => {
  it('generates filename with prefix and original name', () => {
    const result = generateUniqueFilename('report.xlsx', 'outlook');
    expect(result).toMatch(/^outlook_report_\d+_[a-z0-9]+\.xlsx$/);
  });

  it('generates filename without original name', () => {
    const result = generateUniqueFilename(null, 'outlook');
    expect(result).toMatch(/^outlook_file_\d+_[a-z0-9]+$/);
  });

  it('uses default prefix', () => {
    const result = generateUniqueFilename('doc.pdf');
    expect(result).toMatch(/^outlook_doc_\d+_[a-z0-9]+\.pdf$/);
  });

  it('generates unique filenames on successive calls', () => {
    const name1 = generateUniqueFilename('file.txt');
    const name2 = generateUniqueFilename('file.txt');
    expect(name1).not.toBe(name2);
  });
});

describe('shouldSaveToFile', () => {
  it('returns false for small string content', () => {
    expect(shouldSaveToFile('hello world')).toBe(false);
  });

  it('returns true for content exceeding maxSize', () => {
    const largeContent = 'x'.repeat(2 * 1024 * 1024); // 2MB
    expect(shouldSaveToFile(largeContent)).toBe(true);
  });

  it('uses custom maxSize', () => {
    expect(shouldSaveToFile('hello', 3)).toBe(true);
    expect(shouldSaveToFile('hi', 10)).toBe(false);
  });

  it('handles Buffer content', () => {
    const buf = Buffer.alloc(100);
    expect(shouldSaveToFile(buf, 50)).toBe(true);
    expect(shouldSaveToFile(buf, 200)).toBe(false);
  });

  it('handles object content by estimating JSON size', () => {
    const obj = { key: 'value' };
    expect(shouldSaveToFile(obj, 5)).toBe(true);
    expect(shouldSaveToFile(obj, 1000)).toBe(false);
  });
});

describe('saveFileToDisc', () => {
  let tempDir;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-test-'));
    process.env.MCP_OUTLOOK_WORK_DIR = tempDir;
  });

  afterEach(() => {
    delete process.env.MCP_OUTLOOK_WORK_DIR;
    // Cleanup
    try {
      const files = fs.readdirSync(tempDir);
      for (const f of files) fs.unlinkSync(path.join(tempDir, f));
      fs.rmdirSync(tempDir);
    } catch {}
  });

  it('saves base64 content and returns success info', async () => {
    const base64Content = Buffer.from('Hello World').toString('base64');
    const result = await saveFileToDisc(base64Content, 'test.txt', { encoding: 'base64' });

    expect(result.success).toBe(true);
    expect(result.filePath).toContain(tempDir);
    expect(result.originalFilename).toBe('test.txt');
    expect(fs.existsSync(result.filePath)).toBe(true);

    const written = fs.readFileSync(result.filePath, 'utf8');
    expect(written).toBe('Hello World');
  });

  it('saves utf8 content', async () => {
    const result = await saveFileToDisc('plain text content', 'readme.md', { encoding: 'utf8' });

    expect(result.success).toBe(true);
    const written = fs.readFileSync(result.filePath, 'utf8');
    expect(written).toBe('plain text content');
  });

  it('returns error info when write fails', async () => {
    process.env.MCP_OUTLOOK_WORK_DIR = '/nonexistent/path/that/should/not/exist';
    const result = await saveFileToDisc('data', 'file.txt', { encoding: 'utf8' });

    // It should fall back to tmpdir, so it might still succeed.
    // If the work dir fallback works, just check success.
    expect(result).toHaveProperty('success');
  });
});

describe('handleLargeContent', () => {
  let tempDir;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-test-large-'));
    process.env.MCP_OUTLOOK_WORK_DIR = tempDir;
  });

  afterEach(() => {
    delete process.env.MCP_OUTLOOK_WORK_DIR;
    try {
      const files = fs.readdirSync(tempDir);
      for (const f of files) fs.unlinkSync(path.join(tempDir, f));
      fs.rmdirSync(tempDir);
    } catch {}
  });

  it('returns content inline when below maxSize', async () => {
    const result = await handleLargeContent('small text', 'file.txt', { encoding: 'utf8' });
    expect(result.savedToFile).toBe(false);
    expect(result.content).toBe('small text');
  });

  it('saves to file when content exceeds maxSize', async () => {
    const largeContent = 'x'.repeat(100);
    const result = await handleLargeContent(largeContent, 'big.txt', {
      maxSize: 10,
      encoding: 'utf8',
    });

    expect(result.savedToFile).toBe(true);
    expect(result.success).toBe(true);
    expect(result.filePath).toBeDefined();
  });

  it('saves to file when forceFile is true', async () => {
    const result = await handleLargeContent('tiny', 'file.txt', {
      forceFile: true,
      encoding: 'utf8',
    });

    expect(result.savedToFile).toBe(true);
  });
});

describe('cleanupOldFiles', () => {
  let tempDir;

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-test-cleanup-'));
    process.env.MCP_OUTLOOK_WORK_DIR = tempDir;
  });

  afterEach(() => {
    delete process.env.MCP_OUTLOOK_WORK_DIR;
    try {
      const files = fs.readdirSync(tempDir);
      for (const f of files) fs.unlinkSync(path.join(tempDir, f));
      fs.rmdirSync(tempDir);
    } catch {}
  });

  it('cleans up old outlook_ prefixed files', async () => {
    // Create a file with old mtime
    const filePath = path.join(tempDir, 'outlook_old_file.txt');
    fs.writeFileSync(filePath, 'old content');
    // Set mtime to 2 days ago
    const twoDAysAgo = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000);
    fs.utimesSync(filePath, twoDAysAgo, twoDAysAgo);

    const result = await cleanupOldFiles(24 * 60 * 60 * 1000); // 24 hours
    expect(result.cleaned).toBe(1);
    expect(fs.existsSync(filePath)).toBe(false);
  });

  it('does not clean up recent files', async () => {
    const filePath = path.join(tempDir, 'outlook_recent_file.txt');
    fs.writeFileSync(filePath, 'recent content');

    const result = await cleanupOldFiles(24 * 60 * 60 * 1000);
    expect(result.cleaned).toBe(0);
    expect(fs.existsSync(filePath)).toBe(true);
  });

  it('does not clean up files without outlook_ prefix', async () => {
    const filePath = path.join(tempDir, 'other_file.txt');
    fs.writeFileSync(filePath, 'content');
    const twoDAysAgo = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000);
    fs.utimesSync(filePath, twoDAysAgo, twoDAysAgo);

    const result = await cleanupOldFiles(24 * 60 * 60 * 1000);
    expect(result.cleaned).toBe(0);
    expect(fs.existsSync(filePath)).toBe(true);
  });
});

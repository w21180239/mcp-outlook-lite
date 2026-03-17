import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../../utils/mcpErrorResponse.js', () => ({
  convertErrorToToolError: vi.fn((error, context) => ({
    isError: true,
    content: [{ type: 'text', text: `${context}: ${error.message}` }],
  })),
  createValidationError: vi.fn((param, reason) => ({
    isError: true,
    content: [{ type: 'text', text: `Invalid parameter '${param}': ${reason}` }],
  })),
}));

const { FolderResolver } = await import('../../graph/folderResolver.js');

describe('FolderResolver', () => {
  let resolver;
  let mockGraphClient;

  const mockFolders = [
    { id: 'id-inbox', displayName: 'Inbox', parentFolderId: null },
    { id: 'id-sent', displayName: 'Sent Items', parentFolderId: null },
    { id: 'id-drafts', displayName: 'Drafts', parentFolderId: null },
    { id: 'id-junk', displayName: 'Junk Email', parentFolderId: null },
  ];

  beforeEach(() => {
    mockGraphClient = {
      makeRequest: vi.fn().mockResolvedValue({ value: mockFolders }),
    };
    resolver = new FolderResolver(mockGraphClient);
    vi.spyOn(console, 'error').mockImplementation(() => {});
  });

  describe('constructor', () => {
    it('initializes with empty caches', () => {
      expect(resolver.foldersByName.size).toBe(0);
      expect(resolver.foldersById.size).toBe(0);
      expect(resolver.foldersList).toHaveLength(0);
      expect(resolver.lastCacheUpdate).toBe(0);
    });
  });

  describe('shouldRefreshCache', () => {
    it('returns true when cache is empty', () => {
      expect(resolver.shouldRefreshCache()).toBe(true);
    });

    it('returns true when cache has expired', () => {
      resolver.foldersList = [{ id: '1' }];
      resolver.lastCacheUpdate = Date.now() - 10 * 60 * 1000; // 10 minutes ago
      expect(resolver.shouldRefreshCache()).toBe(true);
    });

    it('returns false when cache is fresh', () => {
      resolver.foldersList = [{ id: '1' }];
      resolver.lastCacheUpdate = Date.now(); // just now
      expect(resolver.shouldRefreshCache()).toBe(false);
    });
  });

  describe('refreshFolderCache', () => {
    it('populates folder caches from API response', async () => {
      await resolver.refreshFolderCache();

      expect(resolver.foldersList).toHaveLength(4);
      expect(resolver.foldersByName.has('inbox')).toBe(true);
      expect(resolver.foldersByName.has('sent items')).toBe(true);
      expect(resolver.foldersById.has('id-inbox')).toBe(true);
    });

    it('clears old cache before refreshing', async () => {
      resolver.foldersByName.set('old-folder', { id: 'old' });
      await resolver.refreshFolderCache();
      expect(resolver.foldersByName.has('old-folder')).toBe(false);
    });

    it('updates lastCacheUpdate timestamp', async () => {
      expect(resolver.lastCacheUpdate).toBe(0);
      await resolver.refreshFolderCache();
      expect(resolver.lastCacheUpdate).toBeGreaterThan(0);
    });

    it('throws when API call fails', async () => {
      mockGraphClient.makeRequest.mockRejectedValue(new Error('API error'));
      await expect(resolver.refreshFolderCache()).rejects.toThrow('API error');
    });
  });

  describe('resolveFolderToId', () => {
    it('throws when folderNameOrId is empty', async () => {
      await expect(resolver.resolveFolderToId('')).rejects.toThrow('required');
      await expect(resolver.resolveFolderToId(null)).rejects.toThrow('required');
    });

    it('returns "inbox" for special inbox folder name', async () => {
      const result = await resolver.resolveFolderToId('inbox');
      expect(result).toBe('inbox');
    });

    it('returns "inbox" case-insensitively', async () => {
      const result = await resolver.resolveFolderToId('Inbox');
      expect(result).toBe('inbox');
    });

    it('returns the ID directly if it looks like a folder ID', async () => {
      const longId = 'AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAABs+U/dOA==';
      const result = await resolver.resolveFolderToId(longId);
      expect(result).toBe(longId);
    });

    it('resolves folder name to ID via cache', async () => {
      const result = await resolver.resolveFolderToId('Drafts');
      expect(result).toBe('id-drafts');
    });

    it('resolves folder name case-insensitively', async () => {
      const result = await resolver.resolveFolderToId('sent items');
      expect(result).toBe('id-sent');
    });

    it('throws when folder name is not found', async () => {
      await expect(resolver.resolveFolderToId('Nonexistent'))
        .rejects.toThrow("Folder 'Nonexistent' not found");
    });

    it('refreshes cache and retries when folder not found initially', async () => {
      // First call returns base folders, second call adds the new folder
      mockGraphClient.makeRequest
        .mockResolvedValueOnce({ value: mockFolders })
        .mockResolvedValueOnce({
          value: [...mockFolders, { id: 'id-new', displayName: 'NewFolder', parentFolderId: null }],
        });

      const result = await resolver.resolveFolderToId('NewFolder');
      expect(result).toBe('id-new');
      expect(mockGraphClient.makeRequest).toHaveBeenCalledTimes(2);
    });
  });

  describe('resolveFoldersToIds', () => {
    it('returns empty array for empty input', async () => {
      expect(await resolver.resolveFoldersToIds([])).toEqual([]);
      expect(await resolver.resolveFoldersToIds(null)).toEqual([]);
    });

    it('resolves multiple folder names', async () => {
      const result = await resolver.resolveFoldersToIds(['inbox', 'Drafts']);
      expect(result).toEqual(['inbox', 'id-drafts']);
    });

    it('throws when one folder cannot be resolved', async () => {
      await expect(resolver.resolveFoldersToIds(['inbox', 'Nonexistent']))
        .rejects.toThrow("Failed to resolve folder 'Nonexistent'");
    });
  });

  describe('getFolderInfo', () => {
    it('returns cached folder info', async () => {
      await resolver.refreshFolderCache();
      const result = await resolver.getFolderInfo('Drafts');
      expect(result.id).toBe('id-drafts');
      expect(result.displayName).toBe('Drafts');
    });

    it('makes API call when folder not in cache by ID', async () => {
      await resolver.refreshFolderCache();

      // 'inbox' resolves to ID 'inbox' (special case), which is not in foldersById cache
      // so it falls through to direct API call
      mockGraphClient.makeRequest.mockResolvedValueOnce({
        id: 'inbox-real-id',
        displayName: 'Inbox',
        parentFolderId: null,
        totalItemCount: 50,
        unreadItemCount: 5,
      });

      const result = await resolver.getFolderInfo('inbox');
      expect(result.id).toBe('inbox-real-id');
      expect(result.displayName).toBe('Inbox');
      expect(result.totalItemCount).toBe(50);
    });
  });

  describe('listAllFolders', () => {
    it('returns all folders sorted by displayName', async () => {
      const result = await resolver.listAllFolders();
      expect(result).toHaveLength(4);
      expect(result[0].displayName).toBe('Drafts');
      expect(result[1].displayName).toBe('Inbox');
    });

    it('refreshes cache if needed', async () => {
      await resolver.listAllFolders();
      expect(mockGraphClient.makeRequest).toHaveBeenCalled();
    });
  });
});

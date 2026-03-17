import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { listFoldersTool } from '../../../tools/folders/listFolders.js';
import { createFolderTool } from '../../../tools/folders/createFolder.js';
import { renameFolderTool } from '../../../tools/folders/renameFolder.js';
import { getFolderStatsTool } from '../../../tools/folders/getFolderStats.js';

describe('Folder Tools', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  describe('listFoldersTool', () => {
    it('should list mail folders', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [
          {
            id: 'folder-1',
            displayName: 'Inbox',
            parentFolderId: null,
            childFolderCount: 2,
            unreadItemCount: 5,
            totalItemCount: 100,
            isHidden: false,
          },
        ],
      });

      const result = await listFoldersTool(authManager, {});

      expect(result.content[0].type).toBe('text');
      const data = JSON.parse(result.content[0].text);
      expect(data.folders).toHaveLength(1);
      expect(data.folders[0].name).toBe('Inbox');
    });
  });

  describe('createFolderTool', () => {
    it('should create a folder', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'new-folder' });

      const result = await createFolderTool(authManager, {
        displayName: 'My Folder',
      });

      expect(result.content[0].text).toContain('created successfully');
      expect(result.content[0].text).toContain('My Folder');
    });

    it('should return error when displayName is missing', async () => {
      const result = await createFolderTool(authManager, {});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('displayName');
    });
  });

  describe('renameFolderTool', () => {
    it('should rename a folder', async () => {
      graphApiClient.makeRequest.mockResolvedValue({});

      const result = await renameFolderTool(authManager, {
        folderId: 'folder-1',
        newDisplayName: 'Renamed Folder',
      });

      expect(result.content[0].text).toContain('renamed');
      expect(result.content[0].text).toContain('Renamed Folder');
    });

    it('should return error when folderId is missing', async () => {
      const result = await renameFolderTool(authManager, {
        newDisplayName: 'Name',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('folderId');
    });

    it('should return error when newDisplayName is missing', async () => {
      const result = await renameFolderTool(authManager, {
        folderId: 'folder-1',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('newDisplayName');
    });
  });

  describe('getFolderStatsTool', () => {
    it('should get folder statistics', async () => {
      graphApiClient.makeRequest
        .mockResolvedValueOnce({
          id: 'folder-1',
          displayName: 'Inbox',
          totalItemCount: 50,
          unreadItemCount: 10,
          childFolderCount: 1,
          isHidden: false,
          parentFolderId: 'root',
        })
        .mockResolvedValueOnce({
          value: [
            {
              id: 'sub-1',
              displayName: 'Sub',
              totalItemCount: 5,
              unreadItemCount: 2,
            },
          ],
        });

      const result = await getFolderStatsTool(authManager, {
        folderId: 'folder-1',
      });

      expect(result.content[0].type).toBe('text');
      const data = JSON.parse(result.content[0].text);
      expect(data.totalItems).toBe(50);
      expect(data.unreadItems).toBe(10);
      expect(data.subfolders).toHaveLength(1);
    });

    it('should return error when folderId is missing', async () => {
      const result = await getFolderStatsTool(authManager, {});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('folderId');
    });
  });
});

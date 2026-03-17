import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { listSharePointFilesTool, getSharePointFileTool } from '../../../tools/sharepoint/getSharePointFile.js';

describe('SharePoint Tools', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  describe('listSharePointFilesTool', () => {
    it('should list files from user OneDrive by default', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [
          {
            id: 'file-1',
            name: 'document.docx',
            size: 2048,
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-02T00:00:00Z',
            webUrl: 'https://example.sharepoint.com/file',
            file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
          },
          {
            id: 'folder-1',
            name: 'Photos',
            size: 0,
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            webUrl: 'https://example.sharepoint.com/folder',
            folder: { childCount: 5 },
          },
        ],
      });

      const result = await listSharePointFilesTool(authManager, {});

      expect(result.content[0].type).toBe('text');
      const data = JSON.parse(result.content[0].text);
      expect(data.files).toHaveLength(2);
      expect(data.files[0].name).toBe('document.docx');
      expect(data.files[0].type).toBe('file');
      expect(data.files[1].type).toBe('folder');
      expect(data.count).toBe(2);
    });

    it('should use driveId when provided', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: [] });

      await listSharePointFilesTool(authManager, { driveId: 'drive-123' });

      expect(graphApiClient.makeRequest).toHaveBeenCalledWith(
        '/drives/drive-123/root/children',
        expect.any(Object)
      );
    });

    it('should handle API errors gracefully', async () => {
      graphApiClient.makeRequest.mockRejectedValue(new Error('Access denied'));

      const result = await listSharePointFilesTool(authManager, {});

      expect(result.isError).toBe(true);
    });
  });

  describe('getSharePointFileTool', () => {
    it('should return validation error when neither sharePointUrl nor fileId provided', async () => {
      const result = await getSharePointFileTool(authManager, {});
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('sharePointUrl or fileId');
    });

    it('should return error for non-SharePoint URL', async () => {
      const result = await getSharePointFileTool(authManager, {
        sharePointUrl: 'https://example.com/file.docx',
      });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('SharePoint');
    });

    it('should fetch file by fileId', async () => {
      // ensureAuthenticated returns the graph client
      authManager.ensureAuthenticated.mockResolvedValue({});

      graphApiClient.makeRequest.mockResolvedValue({
        id: 'item-1',
        name: 'budget.xlsx',
        size: 4096,
        createdDateTime: '2024-01-01T00:00:00Z',
        lastModifiedDateTime: '2024-02-01T00:00:00Z',
        webUrl: 'https://company.sharepoint.com/budget.xlsx',
        file: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      });

      const result = await getSharePointFileTool(authManager, { fileId: 'item-1' });

      expect(result.isError).toBeUndefined();
      const data = JSON.parse(result.content[0].text);
      // handleLargeContent wraps the response; the original data may be nested under content
      const fileData = data.success ? data : data.content;
      expect(fileData.success).toBe(true);
      expect(fileData.file.name).toBe('budget.xlsx');
      expect(graphApiClient.makeRequest).toHaveBeenCalledWith(
        '/drives/me/items/item-1',
        expect.any(Object)
      );
    });
  });
});

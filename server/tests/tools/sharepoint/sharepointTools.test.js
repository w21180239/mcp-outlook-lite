import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { listSharePointFilesTool } from '../../../tools/sharepoint/getSharePointFile.js';

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
});

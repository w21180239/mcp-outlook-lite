import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import {
  listSharePointFilesTool,
  getSharePointFileTool,
  resolveSharePointLinkTool,
} from '../../../tools/sharepoint/getSharePointFile.js';

describe('SharePoint Tools - additional coverage', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  describe('listSharePointFilesTool', () => {
    it('should use siteId and driveId when both provided', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: [] });
      await listSharePointFilesTool(authManager, { siteId: 'site-1', driveId: 'drive-1' });
      expect(graphApiClient.makeRequest).toHaveBeenCalledWith(
        '/sites/site-1/drives/drive-1/root/children',
        expect.any(Object)
      );
    });

    it('should use folderId when provided', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: [] });
      await listSharePointFilesTool(authManager, { driveId: 'drive-1', folderId: 'folder-1' });
      expect(graphApiClient.makeRequest).toHaveBeenCalledWith(
        '/drives/drive-1/items/folder-1/children',
        expect.any(Object)
      );
    });

    it('should pass limit and orderBy params', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: [] });
      await listSharePointFilesTool(authManager, { limit: 10, orderBy: 'lastModifiedDateTime desc' });
      expect(graphApiClient.makeRequest).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({ top: 10, orderby: 'lastModifiedDateTime desc' })
      );
    });

    it('should default to 50 items ordered by name', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: [] });
      await listSharePointFilesTool(authManager, {});
      expect(graphApiClient.makeRequest).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({ top: 50, orderby: 'name' })
      );
    });

    it('should handle empty value response', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: [] });
      const result = await listSharePointFilesTool(authManager, {});
      const data = JSON.parse(result.content[0].text);
      expect(data.count).toBe(0);
      expect(data.files).toEqual([]);
    });
  });

  describe('getSharePointFileTool - URL patterns', () => {
    it('should return error for direct OneDrive personal URL', async () => {
      const result = await getSharePointFileTool(authManager, {
        sharePointUrl: 'https://company-my.sharepoint.com/personal/user_company_com/Documents',
      });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('sharing link');
    });

    it('should return error for generic SharePoint URL', async () => {
      const result = await getSharePointFileTool(authManager, {
        sharePointUrl: 'https://company.sharepoint.com/SitePages/Home.aspx',
      });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Unsupported');
    });

    it('should attempt to resolve sharing link URLs', async () => {
      graphApiClient.makeRequest
        .mockRejectedValueOnce(new Error('Not found'))
        .mockRejectedValueOnce(new Error('Not found'))
        .mockRejectedValueOnce(new Error('Not found'));

      const result = await getSharePointFileTool(authManager, {
        sharePointUrl: 'https://company.sharepoint.com/:x:/r/sites/Team/Shared%20Documents/file.xlsx?d=abc123',
      });
      // Should have tried resolution strategies and then returned error
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to resolve');
    });

    it('should handle sharing link with e= parameter', async () => {
      graphApiClient.makeRequest
        .mockRejectedValueOnce(new Error('Not found'))
        .mockRejectedValueOnce(new Error('Not found'))
        .mockRejectedValueOnce(new Error('Not found'));

      const result = await getSharePointFileTool(authManager, {
        sharePointUrl: 'https://company.sharepoint.com/:w:/g/personal/user/doc.docx?e=token123',
      });
      expect(result.isError).toBe(true);
    });

    it('should use custom driveId when fetching by fileId', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        id: 'item-1', name: 'file.txt', size: 100,
        createdDateTime: '2024-01-01T00:00:00Z',
        lastModifiedDateTime: '2024-01-01T00:00:00Z',
        webUrl: 'https://example.com/file.txt',
      });

      await getSharePointFileTool(authManager, { fileId: 'item-1', driveId: 'custom-drive' });
      expect(graphApiClient.makeRequest).toHaveBeenCalledWith(
        '/drives/custom-drive/items/item-1',
        expect.any(Object)
      );
    });

    it('should handle team site URL', async () => {
      const result = await getSharePointFileTool(authManager, {
        sharePointUrl: 'https://company.sharepoint.com/sites/TeamSite/Docs',
      });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('team_site');
    });
  });

  describe('resolveSharePointLinkTool', () => {
    it('should return validation error when no URL provided', async () => {
      const result = await resolveSharePointLinkTool(authManager, {});
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('sharePointUrl');
    });

    it('should return error for non-SharePoint URL', async () => {
      const result = await resolveSharePointLinkTool(authManager, {
        sharePointUrl: 'https://google.com/file',
      });
      expect(result.isError).toBe(true);
    });

    it('should resolve a sharing link successfully', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        id: 'file-1',
        name: 'report.pdf',
        size: 5000,
        createdDateTime: '2024-01-01T00:00:00Z',
        lastModifiedDateTime: '2024-02-01T00:00:00Z',
        webUrl: 'https://company.sharepoint.com/report.pdf',
        file: { mimeType: 'application/pdf' },
        '@microsoft.graph.downloadUrl': 'https://download.url/file',
      });

      const result = await resolveSharePointLinkTool(authManager, {
        sharePointUrl: 'https://company.sharepoint.com/:b:/r/sites/Team/report.pdf?d=abc',
      });
      const data = JSON.parse(result.content[0].text);
      expect(data.success).toBe(true);
      expect(data.file.name).toBe('report.pdf');
      expect(data.file.sharing.resolved).toBe(true);
    });

    it('should include permissions when requested', async () => {
      graphApiClient.makeRequest
        .mockResolvedValueOnce({
          id: 'file-1', name: 'doc.docx', size: 1000,
          createdDateTime: '2024-01-01', lastModifiedDateTime: '2024-01-01',
          webUrl: 'https://example.com', parentReference: { driveId: 'drv-1' },
        })
        .mockResolvedValueOnce({ value: [{ id: 'perm-1', roles: ['read'] }] });

      const result = await resolveSharePointLinkTool(authManager, {
        sharePointUrl: 'https://company.sharepoint.com/:w:/r/sites/Team/doc.docx?d=abc',
        includePermissions: true,
      });
      const data = JSON.parse(result.content[0].text);
      expect(data.file.sharing.permissions).toHaveLength(1);
    });

    it('should handle permission fetch failure gracefully', async () => {
      graphApiClient.makeRequest
        .mockResolvedValueOnce({
          id: 'file-1', name: 'doc.docx', size: 1000,
          createdDateTime: '2024-01-01', lastModifiedDateTime: '2024-01-01',
          webUrl: 'https://example.com', parentReference: { driveId: 'drv-1' },
        })
        .mockRejectedValueOnce(new Error('Forbidden'));

      const result = await resolveSharePointLinkTool(authManager, {
        sharePointUrl: 'https://company.sharepoint.com/:w:/r/sites/Team/doc.docx?d=abc',
        includePermissions: true,
      });
      const data = JSON.parse(result.content[0].text);
      expect(data.file.sharing.permissionsError).toContain('Forbidden');
    });

    it('should handle missing driveId in permissions request', async () => {
      graphApiClient.makeRequest.mockResolvedValueOnce({
        id: 'file-1', name: 'doc.docx', size: 1000,
        createdDateTime: '2024-01-01', lastModifiedDateTime: '2024-01-01',
        webUrl: 'https://example.com',
        // No parentReference
      });

      const result = await resolveSharePointLinkTool(authManager, {
        sharePointUrl: 'https://company.sharepoint.com/:w:/r/sites/Team/doc.docx?d=abc',
        includePermissions: true,
      });
      const data = JSON.parse(result.content[0].text);
      expect(data.file.sharing.permissionsError).toContain('driveId');
    });
  });

  describe('getSharePointFileTool - content download', () => {
    it('should return non-error response when fetching by fileId', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        id: 'item-1', name: 'file.txt', size: 100,
        createdDateTime: '2024-01-01T00:00:00Z',
        lastModifiedDateTime: '2024-01-01T00:00:00Z',
        webUrl: 'https://example.com/file.txt',
        file: { mimeType: 'text/plain' },
      });

      const result = await getSharePointFileTool(authManager, { fileId: 'item-1' });
      expect(result.isError).toBeUndefined();
      expect(result.content).toBeDefined();
      expect(result.content[0].text).toContain('file.txt');
    });
  });
});

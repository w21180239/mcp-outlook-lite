import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { listRulesTool, createRuleTool, deleteRuleTool } from '../../../tools/rules/manageRules.js';

describe('Rules Tools', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  describe('listRulesTool', () => {
    it('should list inbox rules', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [
          {
            id: 'rule-1',
            displayName: 'Move spam',
            isEnabled: true,
            sequence: 1,
            conditions: { senderContains: ['spam@'] },
            actions: { moveToFolder: 'junk' },
          },
        ],
      });

      const result = await listRulesTool(authManager, {});

      expect(result.content[0].type).toBe('text');
      const data = JSON.parse(result.content[0].text);
      expect(data).toHaveLength(1);
      expect(data[0].displayName).toBe('Move spam');
    });

    it('should return message when no rules found', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: [] });

      const result = await listRulesTool(authManager, {});

      expect(result.content[0].text).toContain('No inbox rules');
    });
  });

  describe('createRuleTool', () => {
    it('should create a rule', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'rule-new' });

      const result = await createRuleTool(authManager, {
        displayName: 'Test Rule',
        senderContains: ['test@example.com'],
        moveToFolder: 'folder-1',
      });

      expect(result.content[0].text).toContain('created successfully');
      expect(result.content[0].text).toContain('Test Rule');
    });

    it('should return error when displayName is missing', async () => {
      const result = await createRuleTool(authManager, {
        senderContains: ['test@'],
        moveToFolder: 'folder-1',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('displayName');
    });

    it('should return error when senderContains is missing', async () => {
      const result = await createRuleTool(authManager, {
        displayName: 'Rule',
        moveToFolder: 'folder-1',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('senderContains');
    });

    it('should return error when moveToFolder is missing', async () => {
      const result = await createRuleTool(authManager, {
        displayName: 'Rule',
        senderContains: ['test@'],
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('moveToFolder');
    });
  });

  describe('deleteRuleTool', () => {
    it('should delete a rule', async () => {
      graphApiClient.deleteWithRetry.mockResolvedValue({});

      const result = await deleteRuleTool(authManager, { ruleId: 'rule-1' });

      expect(result.content[0].text).toContain('deleted successfully');
    });

    it('should return error when ruleId is missing', async () => {
      const result = await deleteRuleTool(authManager, {});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('ruleId');
    });
  });
});

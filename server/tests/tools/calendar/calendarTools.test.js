import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import { listEventsTool } from '../../../tools/calendar/listEvents.js';
import { createEventTool } from '../../../tools/calendar/createEvent.js';
import { getEventTool, updateEventTool, deleteEventTool } from '../../../tools/calendar/eventManagement.js';

// Mock applyUserStyling to avoid actual API calls
vi.mock('../../../tools/common/sharedUtils.js', () => ({
  applyUserStyling: vi.fn().mockResolvedValue({ content: 'styled body', type: 'html' }),
}));

describe('Calendar Tools', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  describe('listEventsTool', () => {
    it('should list events', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [
          {
            id: 'evt-1',
            subject: 'Meeting',
            start: { dateTime: '2024-01-01T10:00:00' },
            end: { dateTime: '2024-01-01T11:00:00' },
            location: { displayName: 'Room A' },
            attendees: [{ emailAddress: { address: 'a@b.com' } }],
            bodyPreview: 'Preview text',
          },
        ],
      });

      const result = await listEventsTool(authManager, { limit: 10 });

      expect(result.content[0].type).toBe('text');
      const data = JSON.parse(result.content[0].text);
      expect(data.events).toHaveLength(1);
      expect(data.events[0].subject).toBe('Meeting');
    });
  });

  describe('createEventTool', () => {
    it('should create an event', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'new-evt' });

      const result = await createEventTool(authManager, {
        subject: 'Test Event',
        start: { dateTime: '2024-01-01T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2024-01-01T11:00:00', timeZone: 'UTC' },
      });

      expect(result.content[0].text).toContain('created successfully');
      expect(result.content[0].text).toContain('new-evt');
    });
  });

  describe('getEventTool', () => {
    it('should get event details', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        id: 'evt-1',
        subject: 'Meeting',
        start: { dateTime: '2024-01-01T10:00:00' },
        end: { dateTime: '2024-01-01T11:00:00' },
        body: { contentType: 'Text', content: 'Body' },
        categories: [],
      });

      const result = await getEventTool(authManager, { eventId: 'evt-1' });

      expect(result.content[0].type).toBe('text');
      const data = JSON.parse(result.content[0].text);
      expect(data.subject).toBe('Meeting');
    });

    it('should return error when eventId is missing', async () => {
      const result = await getEventTool(authManager, {});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('eventId');
    });
  });

  describe('updateEventTool', () => {
    it('should update an event', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ id: 'evt-1' });

      const result = await updateEventTool(authManager, {
        eventId: 'evt-1',
        subject: 'Updated Meeting',
      });

      expect(result.content[0].text).toContain('updated successfully');
    });

    it('should return error when eventId is missing', async () => {
      const result = await updateEventTool(authManager, {});

      expect(result.isError).toBe(true);
    });
  });

  describe('deleteEventTool', () => {
    it('should delete an event', async () => {
      graphApiClient.makeRequest.mockResolvedValue({});

      const result = await deleteEventTool(authManager, { eventId: 'evt-1' });

      expect(result.content[0].text).toContain('deleted successfully');
    });

    it('should return error when eventId is missing', async () => {
      const result = await deleteEventTool(authManager, {});

      expect(result.isError).toBe(true);
    });
  });
});

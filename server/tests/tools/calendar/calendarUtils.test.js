import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createMockAuthManager } from '../../helpers/mockAuthManager.js';
import {
  createRecurringEventTool,
  findMeetingTimesTool,
  checkAvailabilityTool,
  scheduleOnlineMeetingTool,
  listCalendarsTool,
  getCalendarViewTool,
  getBusyTimesTool,
  buildRecurrencePatternTool,
  createRecurrenceHelperTool,
  checkCalendarPermissionsTool,
} from '../../../tools/calendar/calendarUtils.js';
import {
  getEventTool,
  updateEventTool,
  deleteEventTool,
  respondToInviteTool,
  validateEventDateTimesTool,
} from '../../../tools/calendar/eventManagement.js';
import { createEventTool } from '../../../tools/calendar/createEvent.js';

// Mock applyUserStyling to avoid actual API calls
vi.mock('../../../tools/common/sharedUtils.js', () => ({
  applyUserStyling: vi.fn().mockResolvedValue({ content: 'styled body', type: 'html' }),
}));

describe('Calendar Utils Tools', () => {
  let authManager;
  let graphApiClient;

  beforeEach(() => {
    authManager = createMockAuthManager();
    graphApiClient = authManager.getGraphApiClient();
  });

  // ─── createRecurringEventTool ───────────────────────────────────────

  describe('createRecurringEventTool', () => {
    const baseArgs = {
      subject: 'Weekly Standup',
      start: { dateTime: '2024-06-01T09:00:00', timeZone: 'UTC' },
      end: { dateTime: '2024-06-01T09:30:00', timeZone: 'UTC' },
      recurrencePattern: {
        pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday'] },
        range: { type: 'noEnd' },
      },
    };

    it('should create a recurring event with minimal args', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'recurring-1' });

      const result = await createRecurringEventTool(authManager, baseArgs);

      expect(result.content[0].text).toContain('Recurring Event');
      expect(result.content[0].text).toContain('recurring-1');
      expect(result.content[0].text).toContain('created successfully');
    });

    it('should include location when provided', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'recurring-2' });

      await createRecurringEventTool(authManager, {
        ...baseArgs,
        location: 'Room B',
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.location).toEqual({ displayName: 'Room B' });
    });

    it('should map attendees correctly', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'recurring-3' });

      await createRecurringEventTool(authManager, {
        ...baseArgs,
        attendees: ['alice@test.com', 'bob@test.com'],
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.attendees).toHaveLength(2);
      expect(postedEvent.attendees[0]).toEqual({
        emailAddress: { address: 'alice@test.com' },
        type: 'required',
      });
    });

    it('should set online meeting fields when isOnlineMeeting is true', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({
        id: 'recurring-4',
        onlineMeeting: { joinUrl: 'https://teams.link/join' },
      });

      const result = await createRecurringEventTool(authManager, {
        ...baseArgs,
        isOnlineMeeting: true,
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.isOnlineMeeting).toBe(true);
      expect(postedEvent.onlineMeetingProvider).toBe('teamsForBusiness');
      expect(result.content[0].text).toContain('Teams meeting');
      expect(result.content[0].text).toContain('Join URL');
    });

    it('should apply user styling when body is provided', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'recurring-5' });

      await createRecurringEventTool(authManager, {
        ...baseArgs,
        body: 'Some notes',
        preserveUserStyling: true,
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.body.content).toBe('styled body');
      expect(postedEvent.body.contentType).toBe('HTML');
    });

    it('should handle API errors gracefully', async () => {
      graphApiClient.postWithRetry.mockRejectedValue(new Error('Network failure'));

      const result = await createRecurringEventTool(authManager, baseArgs);

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to create recurring event');
    });
  });

  // ─── findMeetingTimesTool ───────────────────────────────────────────

  describe('findMeetingTimesTool', () => {
    it('should return validation error when attendees is empty', async () => {
      const result = await findMeetingTimesTool(authManager, { attendees: [] });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('attendees');
    });

    it('should return validation error when attendees is not provided', async () => {
      const result = await findMeetingTimesTool(authManager, {});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('attendees');
    });

    it('should find meeting times successfully', async () => {
      const mockSchedule = { value: [{ scheduleId: 'user@test.com', availabilityView: '0020' }] };
      graphApiClient.postWithRetry.mockResolvedValue(mockSchedule);

      const result = await findMeetingTimesTool(authManager, {
        attendees: ['user@test.com'],
        timeConstraint: {
          start: '2024-06-01T08:00:00',
          end: '2024-06-01T17:00:00',
        },
        maxCandidates: 5,
        meetingDuration: 30,
      });

      expect(result.isError).toBeUndefined();
      const data = JSON.parse(result.content[0].text);
      expect(data.value).toBeDefined();
    });

    it('should handle API errors', async () => {
      graphApiClient.postWithRetry.mockRejectedValue(new Error('Service unavailable'));

      const result = await findMeetingTimesTool(authManager, {
        attendees: ['user@test.com'],
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to find meeting times');
    });
  });

  // ─── checkAvailabilityTool ──────────────────────────────────────────

  describe('checkAvailabilityTool', () => {
    it('should return validation error when schedules is empty', async () => {
      const result = await checkAvailabilityTool(authManager, {
        schedules: [],
        startTime: '2024-06-01T08:00:00',
        endTime: '2024-06-01T17:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('schedules');
    });

    it('should return validation error when startTime is missing', async () => {
      const result = await checkAvailabilityTool(authManager, {
        schedules: ['user@test.com'],
        endTime: '2024-06-01T17:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('startTime');
    });

    it('should return validation error when endTime is missing', async () => {
      const result = await checkAvailabilityTool(authManager, {
        schedules: ['user@test.com'],
        startTime: '2024-06-01T08:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('endTime');
    });

    it('should check availability successfully', async () => {
      const mockResponse = {
        value: [{ scheduleId: 'user@test.com', availabilityView: '0020000' }],
      };
      graphApiClient.postWithRetry.mockResolvedValue(mockResponse);

      const result = await checkAvailabilityTool(authManager, {
        schedules: ['user@test.com'],
        startTime: '2024-06-01T08:00:00',
        endTime: '2024-06-01T17:00:00',
        availabilityViewInterval: 30,
      });

      expect(result.isError).toBeUndefined();
      const callArgs = graphApiClient.postWithRetry.mock.calls[0];
      expect(callArgs[0]).toBe('/me/calendar/getSchedule');
      expect(callArgs[1].availabilityViewInterval).toBe(30);
    });

    it('should handle API errors', async () => {
      graphApiClient.postWithRetry.mockRejectedValue(new Error('Forbidden'));

      const result = await checkAvailabilityTool(authManager, {
        schedules: ['user@test.com'],
        startTime: '2024-06-01T08:00:00',
        endTime: '2024-06-01T17:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to check availability');
    });
  });

  // ─── scheduleOnlineMeetingTool ──────────────────────────────────────

  describe('scheduleOnlineMeetingTool', () => {
    const meetingArgs = {
      subject: 'Sprint Planning',
      startTime: '2024-06-01T10:00:00',
      endTime: '2024-06-01T11:00:00',
    };

    it('should schedule an online meeting', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({
        id: 'meeting-1',
        onlineMeeting: { joinUrl: 'https://teams.link/abc' },
      });

      const result = await scheduleOnlineMeetingTool(authManager, meetingArgs);

      expect(result.content[0].text).toContain('scheduled successfully');
      expect(result.content[0].text).toContain('meeting-1');
      expect(result.content[0].text).toContain('Join URL');
    });

    it('should set isOnlineMeeting and provider on posted event', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'meeting-2' });

      await scheduleOnlineMeetingTool(authManager, meetingArgs);

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.isOnlineMeeting).toBe(true);
      expect(postedEvent.onlineMeetingProvider).toBe('teamsForBusiness');
    });

    it('should use custom meeting provider', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'meeting-3' });

      await scheduleOnlineMeetingTool(authManager, {
        ...meetingArgs,
        meetingProvider: 'skypeForConsumer',
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.onlineMeetingProvider).toBe('skypeForConsumer');
    });

    it('should include attendees when provided', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'meeting-4' });

      await scheduleOnlineMeetingTool(authManager, {
        ...meetingArgs,
        attendees: ['dev@test.com'],
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.attendees).toHaveLength(1);
      expect(postedEvent.attendees[0].emailAddress.address).toBe('dev@test.com');
    });

    it('should apply user styling to body', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'meeting-5' });

      await scheduleOnlineMeetingTool(authManager, {
        ...meetingArgs,
        body: 'Agenda here',
        preserveUserStyling: true,
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.body.content).toBe('styled body');
    });

    it('should not include join URL when not returned', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'meeting-6' });

      const result = await scheduleOnlineMeetingTool(authManager, meetingArgs);

      expect(result.content[0].text).not.toContain('Join URL');
    });

    it('should handle errors', async () => {
      graphApiClient.postWithRetry.mockRejectedValue(new Error('Quota exceeded'));

      const result = await scheduleOnlineMeetingTool(authManager, meetingArgs);

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to schedule online meeting');
    });
  });

  // ─── listCalendarsTool ──────────────────────────────────────────────

  describe('listCalendarsTool', () => {
    it('should list calendars', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [
          {
            id: 'cal-1',
            name: 'Calendar',
            color: 'auto',
            isDefaultCalendar: true,
            canShare: true,
            canViewPrivateItems: true,
            canEdit: true,
            owner: { name: 'User', address: 'user@test.com' },
          },
        ],
      });

      const result = await listCalendarsTool(authManager, {});

      const data = JSON.parse(result.content[0].text);
      expect(data.calendars).toHaveLength(1);
      expect(data.calendars[0].name).toBe('Calendar');
      expect(data.count).toBe(1);
    });

    it('should cap top at 1000', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: [] });

      await listCalendarsTool(authManager, { top: 5000 });

      const options = graphApiClient.makeRequest.mock.calls[0][1];
      expect(options.top).toBe(1000);
    });

    it('should return empty array when no calendars', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: null });

      const result = await listCalendarsTool(authManager, {});

      const data = JSON.parse(result.content[0].text);
      expect(data.calendars).toEqual([]);
      expect(data.count).toBe(0);
    });

    it('should handle errors', async () => {
      graphApiClient.makeRequest.mockRejectedValue(new Error('Unauthorized'));

      const result = await listCalendarsTool(authManager, {});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to list calendars');
    });
  });

  // ─── getCalendarViewTool ────────────────────────────────────────────

  describe('getCalendarViewTool', () => {
    it('should return validation error when startDateTime is missing', async () => {
      const result = await getCalendarViewTool(authManager, {
        endDateTime: '2024-06-02T00:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('startDateTime');
    });

    it('should return validation error when endDateTime is missing', async () => {
      const result = await getCalendarViewTool(authManager, {
        startDateTime: '2024-06-01T00:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('endDateTime');
    });

    it('should get calendar view for default calendar', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [
          {
            id: 'evt-1',
            subject: 'Lunch',
            start: { dateTime: '2024-06-01T12:00:00' },
            end: { dateTime: '2024-06-01T13:00:00' },
            location: { displayName: 'Cafe' },
            attendees: [],
            bodyPreview: 'Preview',
            organizer: { emailAddress: { address: 'boss@test.com' } },
            isAllDay: false,
            showAs: 'busy',
            importance: 'normal',
            sensitivity: 'normal',
            categories: ['lunch'],
            webLink: 'https://outlook.com/evt-1',
          },
        ],
      });

      const result = await getCalendarViewTool(authManager, {
        startDateTime: '2024-06-01T00:00:00',
        endDateTime: '2024-06-02T00:00:00',
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.events).toHaveLength(1);
      expect(data.events[0].subject).toBe('Lunch');
      expect(data.events[0].organizer).toBe('boss@test.com');

      // default endpoint, no calendarId
      expect(graphApiClient.makeRequest.mock.calls[0][0]).toBe('/me/calendarView');
    });

    it('should use specific calendar endpoint when calendarId is provided', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ value: [] });

      await getCalendarViewTool(authManager, {
        startDateTime: '2024-06-01T00:00:00',
        endDateTime: '2024-06-02T00:00:00',
        calendarId: 'cal-123',
      });

      expect(graphApiClient.makeRequest.mock.calls[0][0]).toBe(
        '/me/calendars/cal-123/calendarView'
      );
    });

    it('should handle missing location gracefully', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        value: [{ id: 'evt-2', subject: 'No location', location: null }],
      });

      const result = await getCalendarViewTool(authManager, {
        startDateTime: '2024-06-01T00:00:00',
        endDateTime: '2024-06-02T00:00:00',
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.events[0].location).toBe('No location');
    });

    it('should handle errors', async () => {
      graphApiClient.makeRequest.mockRejectedValue(new Error('Timeout'));

      const result = await getCalendarViewTool(authManager, {
        startDateTime: '2024-06-01T00:00:00',
        endDateTime: '2024-06-02T00:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to get calendar view');
    });
  });

  // ─── getBusyTimesTool ───────────────────────────────────────────────

  describe('getBusyTimesTool', () => {
    it('should return validation error when schedules is empty', async () => {
      const result = await getBusyTimesTool(authManager, {
        schedules: [],
        startTime: '2024-06-01T00:00:00',
        endTime: '2024-06-02T00:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('schedules');
    });

    it('should return validation error when startTime is missing', async () => {
      const result = await getBusyTimesTool(authManager, {
        schedules: ['user@test.com'],
        endTime: '2024-06-02T00:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('startTime');
    });

    it('should return validation error when endTime is missing', async () => {
      const result = await getBusyTimesTool(authManager, {
        schedules: ['user@test.com'],
        startTime: '2024-06-01T00:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('endTime');
    });

    it('should extract busy times from schedule data', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({
        value: [
          {
            busyViewData: ['0', '0', '2', '0'],
          },
        ],
      });

      const result = await getBusyTimesTool(authManager, {
        schedules: ['user@test.com'],
        startTime: '2024-06-01T08:00:00',
        endTime: '2024-06-01T12:00:00',
        availabilityViewInterval: 60,
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.busyTimes).toHaveLength(1);
      expect(data.busyTimes[0].user).toBe('user@test.com');
      expect(data.busyTimes[0].busyTimes).toHaveLength(1);
    });

    it('should handle no busy times', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({
        value: [
          {
            busyViewData: ['0', '0', '0'],
          },
        ],
      });

      const result = await getBusyTimesTool(authManager, {
        schedules: ['free@test.com'],
        startTime: '2024-06-01T08:00:00',
        endTime: '2024-06-01T11:00:00',
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.busyTimes[0].busyTimes).toHaveLength(0);
    });

    it('should handle errors', async () => {
      graphApiClient.postWithRetry.mockRejectedValue(new Error('Graph error'));

      const result = await getBusyTimesTool(authManager, {
        schedules: ['user@test.com'],
        startTime: '2024-06-01T08:00:00',
        endTime: '2024-06-01T12:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to get busy times');
    });
  });

  // ─── buildRecurrencePatternTool ─────────────────────────────────────

  describe('buildRecurrencePatternTool', () => {
    it('should build a daily pattern', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'daily',
        interval: 1,
        rangeType: 'noEnd',
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.recurrencePattern.pattern.type).toBe('daily');
      expect(data.recurrencePattern.pattern.interval).toBe(1);
      expect(data.description).toContain('Every day');
      expect(data.validation.valid).toBe(true);
    });

    it('should build a weekly pattern with days', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'weekly',
        interval: 2,
        daysOfWeek: ['monday', 'wednesday', 'friday'],
        rangeType: 'numbered',
        numberOfOccurrences: 20,
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.recurrencePattern.pattern.daysOfWeek).toEqual([
        'monday',
        'wednesday',
        'friday',
      ]);
      expect(data.recurrencePattern.range.numberOfOccurrences).toBe(20);
      expect(data.description).toContain('Every 2 weeks');
    });

    it('should build an absoluteMonthly pattern', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'absoluteMonthly',
        interval: 1,
        dayOfMonth: 15,
        rangeType: 'endDate',
        rangeEndDate: '2025-12-31',
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.recurrencePattern.pattern.dayOfMonth).toBe(15);
      expect(data.recurrencePattern.range.endDate).toBe('2025-12-31');
      expect(data.description).toContain('Monthly on day 15');
    });

    it('should build a relativeMonthly pattern', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'relativeMonthly',
        interval: 1,
        daysOfWeek: ['tuesday'],
        index: 'second',
        rangeType: 'noEnd',
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.recurrencePattern.pattern.index).toBe('second');
      expect(data.description).toContain('Monthly on the second tuesday');
    });

    it('should build an absoluteYearly pattern', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'absoluteYearly',
        dayOfMonth: 25,
        monthOfYear: 12,
        rangeType: 'noEnd',
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.recurrencePattern.pattern.dayOfMonth).toBe(25);
      expect(data.recurrencePattern.pattern.month).toBe(12);
      expect(data.description).toContain('Yearly on day 25 of month 12');
    });

    it('should build a relativeYearly pattern', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'relativeYearly',
        daysOfWeek: ['thursday'],
        index: 'fourth',
        monthOfYear: 11,
        rangeType: 'noEnd',
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.recurrencePattern.pattern.month).toBe(11);
      expect(data.recurrencePattern.pattern.index).toBe('fourth');
      expect(data.description).toContain('Yearly on the fourth thursday of month 11');
    });

    it('should reject invalid patternType', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'biweekly',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('patternType');
    });

    it('should reject invalid rangeType', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'daily',
        rangeType: 'forever',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('rangeType');
    });

    it('should reject invalid daysOfWeek', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'weekly',
        daysOfWeek: ['funday'],
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('daysOfWeek');
    });

    it('should reject invalid index for relativeMonthly', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'relativeMonthly',
        index: 'fifth',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('index');
    });

    it('should require rangeEndDate when rangeType is endDate', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'daily',
        rangeType: 'endDate',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('rangeEndDate');
    });

    it('should include startDate when provided', async () => {
      const result = await buildRecurrencePatternTool(authManager, {
        patternType: 'daily',
        rangeType: 'noEnd',
        rangeStartDate: '2024-06-01',
      });

      const data = JSON.parse(result.content[0].text);
      expect(data.recurrencePattern.range.startDate).toBe('2024-06-01');
    });
  });

  // ─── createRecurrenceHelperTool ─────────────────────────────────────

  describe('createRecurrenceHelperTool', () => {
    const baseHelperArgs = {
      eventTitle: 'Daily Standup',
      startDateTime: '2024-06-01T09:00:00',
      endDateTime: '2024-06-01T09:15:00',
    };

    it('should return validation error when startDateTime is missing', async () => {
      const result = await createRecurrenceHelperTool(authManager, {
        endDateTime: '2024-06-01T09:15:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('startDateTime');
    });

    it('should return validation error when endDateTime is missing', async () => {
      const result = await createRecurrenceHelperTool(authManager, {
        startDateTime: '2024-06-01T09:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('endDateTime');
    });

    it('should create a daily recurring event', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'helper-daily-1' });

      const result = await createRecurrenceHelperTool(authManager, {
        ...baseHelperArgs,
        recurrenceType: 'daily',
        endAfter: 'occurrences',
        occurrences: 5,
      });

      expect(result.content[0].text).toContain('created successfully');
      expect(result.content[0].text).toContain('helper-daily-1');
      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.recurrence.pattern.type).toBe('daily');
      expect(postedEvent.recurrence.range.numberOfOccurrences).toBe(5);
    });

    it('should create a weekly recurring event with weekDays', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'helper-weekly-1' });

      await createRecurrenceHelperTool(authManager, {
        ...baseHelperArgs,
        recurrenceType: 'weekly',
        weekDays: ['tuesday', 'thursday'],
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.recurrence.pattern.daysOfWeek).toEqual(['tuesday', 'thursday']);
    });

    it('should default weekly to monday when no weekDays', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'helper-weekly-2' });

      await createRecurrenceHelperTool(authManager, {
        ...baseHelperArgs,
        recurrenceType: 'weekly',
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.recurrence.pattern.daysOfWeek).toEqual(['monday']);
    });

    it('should create a monthly recurring event', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'helper-monthly-1' });

      await createRecurrenceHelperTool(authManager, {
        ...baseHelperArgs,
        recurrenceType: 'monthly',
        monthDay: 15,
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.recurrence.pattern.type).toBe('absoluteMonthly');
      expect(postedEvent.recurrence.pattern.dayOfMonth).toBe(15);
    });

    it('should create a yearly recurring event', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'helper-yearly-1' });

      await createRecurrenceHelperTool(authManager, {
        ...baseHelperArgs,
        recurrenceType: 'yearly',
        monthDay: 25,
        yearMonth: 12,
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.recurrence.pattern.type).toBe('absoluteYearly');
      expect(postedEvent.recurrence.pattern.month).toBe(12);
    });

    it('should return validation error for invalid recurrenceType', async () => {
      const result = await createRecurrenceHelperTool(authManager, {
        ...baseHelperArgs,
        recurrenceType: 'biweekly',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('recurrenceType');
    });

    it('should use endDate range when endAfter is date', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'helper-enddate-1' });

      await createRecurrenceHelperTool(authManager, {
        ...baseHelperArgs,
        recurrenceType: 'daily',
        endAfter: 'date',
        endDate: '2024-12-31',
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.recurrence.range.type).toBe('endDate');
      expect(postedEvent.recurrence.range.endDate).toBe('2024-12-31');
    });

    it('should return validation error when endAfter is date but endDate is missing', async () => {
      const result = await createRecurrenceHelperTool(authManager, {
        ...baseHelperArgs,
        recurrenceType: 'daily',
        endAfter: 'date',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('endDate');
    });

    it('should include location and attendees', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'helper-full-1' });

      await createRecurrenceHelperTool(authManager, {
        ...baseHelperArgs,
        location: 'Conf Room',
        attendees: ['a@test.com'],
      });

      const postedEvent = graphApiClient.postWithRetry.mock.calls[0][1];
      expect(postedEvent.location.displayName).toBe('Conf Room');
      expect(postedEvent.attendees[0].emailAddress.address).toBe('a@test.com');
    });

    it('should handle API errors', async () => {
      graphApiClient.postWithRetry.mockRejectedValue(new Error('API down'));

      const result = await createRecurrenceHelperTool(authManager, baseHelperArgs);

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to create recurring event');
    });
  });

  // ─── checkCalendarPermissionsTool ───────────────────────────────────

  describe('checkCalendarPermissionsTool', () => {
    it('should check default calendar permissions', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        id: 'cal-default',
        name: 'Calendar',
        canEdit: true,
        canShare: true,
        canViewPrivateItems: false,
        owner: { name: 'Me' },
        permissions: [],
      });

      const result = await checkCalendarPermissionsTool(authManager, {});

      const data = JSON.parse(result.content[0].text);
      expect(data.calendarId).toBe('cal-default');
      expect(data.canEdit).toBe(true);
      expect(graphApiClient.makeRequest.mock.calls[0][0]).toBe('/me/calendar');
    });

    it('should check specific calendar permissions', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        id: 'cal-shared',
        name: 'Shared Cal',
        canEdit: false,
        canShare: false,
        canViewPrivateItems: false,
        owner: { name: 'Other' },
      });

      const result = await checkCalendarPermissionsTool(authManager, {
        calendarId: 'cal-shared',
      });

      expect(graphApiClient.makeRequest.mock.calls[0][0]).toBe('/me/calendars/cal-shared');
      const data = JSON.parse(result.content[0].text);
      expect(data.canEdit).toBe(false);
    });

    it('should handle errors', async () => {
      graphApiClient.makeRequest.mockRejectedValue(new Error('Not found'));

      const result = await checkCalendarPermissionsTool(authManager, { calendarId: 'bad-id' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to check calendar permissions');
    });
  });

  // ─── eventManagement: respondToInviteTool ───────────────────────────

  describe('respondToInviteTool', () => {
    it('should return validation error when eventId is missing', async () => {
      const result = await respondToInviteTool(authManager, {});

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('eventId');
    });

    it('should return validation error for invalid response', async () => {
      const result = await respondToInviteTool(authManager, {
        eventId: 'evt-1',
        response: 'maybe',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('response');
    });

    it('should accept an invite', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({});

      const result = await respondToInviteTool(authManager, {
        eventId: 'evt-1',
        response: 'accept',
        comment: 'Looking forward to it',
      });

      expect(result.content[0].text).toContain('accept');
      expect(result.content[0].text).toContain('evt-1');
      const callArgs = graphApiClient.postWithRetry.mock.calls[0];
      expect(callArgs[0]).toBe('/me/events/evt-1/accept');
      expect(callArgs[1].comment).toBe('Looking forward to it');
    });

    it('should decline an invite', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({});

      const result = await respondToInviteTool(authManager, {
        eventId: 'evt-2',
        response: 'decline',
      });

      expect(result.content[0].text).toContain('decline');
      expect(graphApiClient.postWithRetry.mock.calls[0][0]).toBe('/me/events/evt-2/decline');
    });

    it('should tentatively accept an invite', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({});

      const result = await respondToInviteTool(authManager, {
        eventId: 'evt-3',
        response: 'tentativelyAccept',
      });

      expect(result.content[0].text).toContain('tentativelyAccept');
    });

    it('should handle errors', async () => {
      graphApiClient.postWithRetry.mockRejectedValue(new Error('Event not found'));

      const result = await respondToInviteTool(authManager, {
        eventId: 'evt-bad',
        response: 'accept',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to respond to invite');
    });
  });

  // ─── eventManagement: validateEventDateTimesTool ────────────────────

  describe('validateEventDateTimesTool', () => {
    it('should return validation error when startDateTime is missing', async () => {
      const result = await validateEventDateTimesTool(authManager, {
        endDateTime: '2024-06-01T11:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('startDateTime');
    });

    it('should return validation error when endDateTime is missing', async () => {
      const result = await validateEventDateTimesTool(authManager, {
        startDateTime: '2024-06-01T10:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('endDateTime');
    });

    it('should return error for invalid start date format', async () => {
      const result = await validateEventDateTimesTool(authManager, {
        startDateTime: 'not-a-date',
        endDateTime: '2024-06-01T11:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('startDateTime');
    });

    it('should return error for invalid end date format', async () => {
      const result = await validateEventDateTimesTool(authManager, {
        startDateTime: '2024-06-01T10:00:00',
        endDateTime: 'not-a-date',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('endDateTime');
    });

    it('should return error when start >= end', async () => {
      const result = await validateEventDateTimesTool(authManager, {
        startDateTime: '2024-06-01T12:00:00',
        endDateTime: '2024-06-01T10:00:00',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('dateRange');
    });

    it('should validate correct date times and calculate duration', async () => {
      const result = await validateEventDateTimesTool(authManager, {
        startDateTime: '2024-06-01T10:00:00Z',
        endDateTime: '2024-06-01T11:30:00Z',
        timeZone: 'Pacific/Auckland',
      });

      expect(result.isError).toBeUndefined();
      const data = JSON.parse(result.content[0].text);
      expect(data.valid).toBe(true);
      expect(data.duration.totalMinutes).toBe(90);
      expect(data.duration.hours).toBe(1);
      expect(data.duration.minutes).toBe(30);
      expect(data.duration.formatted).toBe('1h 30m');
      expect(data.startDateTime.timeZone).toBe('Pacific/Auckland');
    });
  });

  // ─── createEvent: additional coverage ───────────────────────────────

  describe('createEventTool (additional coverage)', () => {
    it('should create event with recurrence', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({ id: 'evt-recur-1' });

      const result = await createEventTool(authManager, {
        subject: 'Recurring Test',
        start: { dateTime: '2024-06-01T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2024-06-01T11:00:00', timeZone: 'UTC' },
        recurrence: {
          pattern: { type: 'daily', interval: 1 },
          range: { type: 'noEnd' },
        },
      });

      expect(result.content[0].text).toContain('(recurring)');
      expect(result.content[0].text).toContain('created successfully');
    });

    it('should reject recurrence missing pattern', async () => {
      const result = await createEventTool(authManager, {
        subject: 'Bad Recur',
        start: { dateTime: '2024-06-01T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2024-06-01T11:00:00', timeZone: 'UTC' },
        recurrence: { range: { type: 'noEnd' } },
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('recurrence');
    });

    it('should reject invalid recurrence pattern type', async () => {
      const result = await createEventTool(authManager, {
        subject: 'Bad Pattern',
        start: { dateTime: '2024-06-01T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2024-06-01T11:00:00', timeZone: 'UTC' },
        recurrence: {
          pattern: { type: 'biweekly', interval: 1 },
          range: { type: 'noEnd' },
        },
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('recurrence.pattern.type');
    });

    it('should reject invalid recurrence range type', async () => {
      const result = await createEventTool(authManager, {
        subject: 'Bad Range',
        start: { dateTime: '2024-06-01T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2024-06-01T11:00:00', timeZone: 'UTC' },
        recurrence: {
          pattern: { type: 'daily', interval: 1 },
          range: { type: 'infinite' },
        },
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('recurrence.range.type');
    });

    it('should create Teams meeting with join URL', async () => {
      graphApiClient.postWithRetry.mockResolvedValue({
        id: 'teams-evt-1',
        onlineMeeting: { joinUrl: 'https://teams.link/xyz' },
      });

      const result = await createEventTool(authManager, {
        subject: 'Teams Call',
        start: { dateTime: '2024-06-01T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2024-06-01T11:00:00', timeZone: 'UTC' },
        isOnlineMeeting: true,
      });

      expect(result.content[0].text).toContain('Teams meeting');
      expect(result.content[0].text).toContain('Join URL');
    });

    it('should handle API errors', async () => {
      graphApiClient.postWithRetry.mockRejectedValue(new Error('Server error'));

      const result = await createEventTool(authManager, {
        subject: 'Fail',
        start: { dateTime: '2024-06-01T10:00:00', timeZone: 'UTC' },
        end: { dateTime: '2024-06-01T11:00:00', timeZone: 'UTC' },
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to create event');
    });
  });

  // ─── eventManagement: updateEventTool (additional) ──────────────────

  describe('updateEventTool (additional coverage)', () => {
    it('should update event with body and apply styling', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ id: 'evt-upd-1' });

      const result = await updateEventTool(authManager, {
        eventId: 'evt-upd-1',
        body: 'New body content',
        preserveUserStyling: true,
      });

      expect(result.content[0].text).toContain('updated successfully');
      const callArgs = graphApiClient.makeRequest.mock.calls[0];
      expect(callArgs[2]).toBe('PATCH');
      expect(callArgs[1].body.body.content).toBe('styled body');
    });

    it('should update event with location, attendees, and online meeting', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ id: 'evt-upd-2' });

      await updateEventTool(authManager, {
        eventId: 'evt-upd-2',
        location: 'New Room',
        attendees: ['new@test.com'],
        isOnlineMeeting: true,
        onlineMeetingProvider: 'teamsForBusiness',
      });

      const updateBody = graphApiClient.makeRequest.mock.calls[0][1].body;
      expect(updateBody.location.displayName).toBe('New Room');
      expect(updateBody.attendees[0].emailAddress.address).toBe('new@test.com');
      expect(updateBody.isOnlineMeeting).toBe(true);
      expect(updateBody.onlineMeetingProvider).toBe('teamsForBusiness');
    });

    it('should update event with recurrence', async () => {
      graphApiClient.makeRequest.mockResolvedValue({ id: 'evt-upd-3' });
      const recurrence = { pattern: { type: 'daily', interval: 2 }, range: { type: 'noEnd' } };

      await updateEventTool(authManager, {
        eventId: 'evt-upd-3',
        recurrence,
      });

      const updateBody = graphApiClient.makeRequest.mock.calls[0][1].body;
      expect(updateBody.recurrence).toEqual(recurrence);
    });

    it('should handle API errors', async () => {
      graphApiClient.makeRequest.mockRejectedValue(new Error('Conflict'));

      const result = await updateEventTool(authManager, {
        eventId: 'evt-upd-fail',
        subject: 'Oops',
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to update event');
    });
  });

  // ─── eventManagement: deleteEventTool (additional) ──────────────────

  describe('deleteEventTool (additional coverage)', () => {
    it('should delete from specific calendar', async () => {
      graphApiClient.makeRequest.mockResolvedValue({});

      const result = await deleteEventTool(authManager, {
        eventId: 'evt-del-1',
        calendarId: 'cal-123',
      });

      expect(result.content[0].text).toContain('deleted successfully');
      expect(graphApiClient.makeRequest.mock.calls[0][0]).toBe(
        '/me/calendars/cal-123/events/evt-del-1'
      );
    });

    it('should handle API errors', async () => {
      graphApiClient.makeRequest.mockRejectedValue(new Error('Not found'));

      const result = await deleteEventTool(authManager, { eventId: 'evt-bad' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to delete event');
    });
  });

  // ─── eventManagement: getEventTool (additional) ─────────────────────

  describe('getEventTool (additional coverage)', () => {
    it('should get event from specific calendar', async () => {
      graphApiClient.makeRequest.mockResolvedValue({
        id: 'evt-cal-1',
        subject: 'Cal Event',
        body: { contentType: 'HTML', content: '<p>Hi</p>' },
        categories: ['work'],
      });

      const result = await getEventTool(authManager, {
        eventId: 'evt-cal-1',
        calendarId: 'cal-special',
      });

      expect(graphApiClient.makeRequest.mock.calls[0][0]).toBe(
        '/me/calendars/cal-special/events/evt-cal-1'
      );
      const data = JSON.parse(result.content[0].text);
      expect(data.id).toBe('evt-cal-1');
    });

    it('should handle API errors', async () => {
      graphApiClient.makeRequest.mockRejectedValue(new Error('Gone'));

      const result = await getEventTool(authManager, { eventId: 'evt-gone' });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Failed to get event');
    });
  });
});

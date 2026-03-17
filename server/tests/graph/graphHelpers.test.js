import { describe, it, expect } from 'vitest';
import { graphHelpers } from '../../graph/graphHelpers.js';

describe('graphHelpers', () => {
  // ── Email helpers ──────────────────────────────────────────────

  describe('email.buildMessageObject', () => {
    it('should build a basic message with a single recipient', () => {
      const msg = graphHelpers.email.buildMessageObject('user@test.com', 'Hi', 'body');
      expect(msg.subject).toBe('Hi');
      expect(msg.body).toEqual({ contentType: 'Text', content: 'body' });
      expect(msg.toRecipients).toEqual([{ emailAddress: { address: 'user@test.com' } }]);
    });

    it('should accept an array of recipients', () => {
      const msg = graphHelpers.email.buildMessageObject(['a@t.com', 'b@t.com'], 'S', 'B');
      expect(msg.toRecipients).toHaveLength(2);
    });

    it('should set HTML body type when option is html', () => {
      const msg = graphHelpers.email.buildMessageObject('u@t.com', 'S', '<b>hi</b>', { bodyType: 'html' });
      expect(msg.body.contentType).toBe('HTML');
    });

    it('should include cc recipients', () => {
      const msg = graphHelpers.email.buildMessageObject('u@t.com', 'S', 'B', { cc: ['cc@t.com'] });
      expect(msg.ccRecipients).toHaveLength(1);
      expect(msg.ccRecipients[0].emailAddress.address).toBe('cc@t.com');
    });

    it('should include bcc recipients', () => {
      const msg = graphHelpers.email.buildMessageObject('u@t.com', 'S', 'B', { bcc: ['bcc@t.com'] });
      expect(msg.bccRecipients).toHaveLength(1);
    });

    it('should include importance', () => {
      const msg = graphHelpers.email.buildMessageObject('u@t.com', 'S', 'B', { importance: 'high' });
      expect(msg.importance).toBe('high');
    });

    it('should include attachments', () => {
      const att = [{ name: 'f.txt' }];
      const msg = graphHelpers.email.buildMessageObject('u@t.com', 'S', 'B', { attachments: att });
      expect(msg.attachments).toEqual(att);
    });
  });

  describe('email.buildReplyObject', () => {
    it('should build a basic reply', () => {
      const reply = graphHelpers.email.buildReplyObject('thanks');
      expect(reply.comment).toBe('thanks');
      expect(reply.message).toBeUndefined();
    });

    it('should build a reply-all with cc', () => {
      const reply = graphHelpers.email.buildReplyObject('ok', { replyAll: true, cc: ['cc@t.com'] });
      expect(reply.message.ccRecipients).toHaveLength(1);
    });

    it('should build a reply-all without cc', () => {
      const reply = graphHelpers.email.buildReplyObject('ok', { replyAll: true });
      expect(reply.message).toBeDefined();
    });
  });

  describe('email.parseEmailAddress', () => {
    it('should return string emails as-is', () => {
      expect(graphHelpers.email.parseEmailAddress('test@example.com')).toBe('test@example.com');
    });

    it('should extract address from email object', () => {
      expect(graphHelpers.email.parseEmailAddress({ emailAddress: { address: 'a@b.com' } })).toBe('a@b.com');
    });

    it('should return unknown for missing address', () => {
      expect(graphHelpers.email.parseEmailAddress({})).toBe('unknown');
    });
  });

  describe('email.parseEmailName', () => {
    it('should return null for string input', () => {
      expect(graphHelpers.email.parseEmailName('test@example.com')).toBeNull();
    });

    it('should extract name from email object', () => {
      expect(graphHelpers.email.parseEmailName({ emailAddress: { name: 'John' } })).toBe('John');
    });

    it('should return null when name is missing', () => {
      expect(graphHelpers.email.parseEmailName({ emailAddress: {} })).toBeNull();
    });
  });

  // ── Calendar legacy helpers ────────────────────────────────────

  describe('_calendarLegacy.buildEventObject', () => {
    it('should build a basic event', () => {
      const evt = graphHelpers._calendarLegacy.buildEventObject('Meeting', '2024-01-01T10:00:00Z', '2024-01-01T11:00:00Z');
      expect(evt.subject).toBe('Meeting');
      expect(evt.start.dateTime).toBe('2024-01-01T10:00:00Z');
      expect(evt.end.timeZone).toBe('UTC');
    });

    it('should use dateTime/timeZone objects for start/end', () => {
      const evt = graphHelpers._calendarLegacy.buildEventObject(
        'Mtg',
        { dateTime: '2024-01-01T10:00:00', timeZone: 'PST' },
        { dateTime: '2024-01-01T11:00:00', timeZone: 'PST' }
      );
      expect(evt.start.timeZone).toBe('PST');
    });

    it('should include body, location, attendees, isAllDay, recurrence, online meeting', () => {
      const evt = graphHelpers._calendarLegacy.buildEventObject('Mtg', 's', 'e', {
        body: 'notes', bodyType: 'html', location: 'Room A',
        attendees: ['a@t.com'], isAllDay: true,
        recurrence: { pattern: {} }, isOnlineMeeting: true, onlineMeetingProvider: 'skypeForBusiness'
      });
      expect(evt.body.contentType).toBe('HTML');
      expect(evt.location.displayName).toBe('Room A');
      expect(evt.attendees[0].type).toBe('required');
      expect(evt.isAllDay).toBe(true);
      expect(evt.recurrence).toBeDefined();
      expect(evt.isOnlineMeeting).toBe(true);
      expect(evt.onlineMeetingProvider).toBe('skypeForBusiness');
    });

    it('should default onlineMeetingProvider to teamsForBusiness', () => {
      const evt = graphHelpers._calendarLegacy.buildEventObject('M', 's', 'e', { isOnlineMeeting: true });
      expect(evt.onlineMeetingProvider).toBe('teamsForBusiness');
    });
  });

  describe('_calendarLegacy.buildRecurrencePattern', () => {
    it('should build a weekly recurrence with endDate', () => {
      const rec = graphHelpers._calendarLegacy.buildRecurrencePattern(
        { type: 'weekly', interval: 2, daysOfWeek: ['monday'] },
        { type: 'endDate', startDate: '2024-01-01', endDate: '2024-06-01' }
      );
      expect(rec.pattern.type).toBe('weekly');
      expect(rec.pattern.interval).toBe(2);
      expect(rec.pattern.daysOfWeek).toEqual(['monday']);
      expect(rec.range.endDate).toBe('2024-06-01');
    });

    it('should build a monthly recurrence with dayOfMonth and numbered range', () => {
      const rec = graphHelpers._calendarLegacy.buildRecurrencePattern(
        { type: 'absoluteMonthly', dayOfMonth: 15 },
        { type: 'numbered', startDate: '2024-01-01', numberOfOccurrences: 10 }
      );
      expect(rec.pattern.dayOfMonth).toBe(15);
      expect(rec.range.numberOfOccurrences).toBe(10);
    });

    it('should default interval to 1', () => {
      const rec = graphHelpers._calendarLegacy.buildRecurrencePattern(
        { type: 'daily' },
        { type: 'noEnd', startDate: '2024-01-01' }
      );
      expect(rec.pattern.interval).toBe(1);
    });
  });

  describe('_calendarLegacy.parseDateTimeWithZone', () => {
    it('should return dateTime and timeZone', () => {
      const result = graphHelpers._calendarLegacy.parseDateTimeWithZone('2024-01-01T10:00:00', 'PST');
      expect(result.dateTime).toBe('2024-01-01T10:00:00');
      expect(result.timeZone).toBe('PST');
    });
  });

  // ── Timezone helpers ───────────────────────────────────────────

  describe('timezone.normalizeTimezone', () => {
    it('should return UTC for falsy input', () => {
      expect(graphHelpers.timezone.normalizeTimezone(null)).toBe('UTC');
      expect(graphHelpers.timezone.normalizeTimezone('')).toBe('UTC');
    });

    it('should map known abbreviations', () => {
      expect(graphHelpers.timezone.normalizeTimezone('EST')).toBe('Eastern Standard Time');
      expect(graphHelpers.timezone.normalizeTimezone('PST')).toBe('Pacific Standard Time');
    });

    it('should find partial matches', () => {
      expect(graphHelpers.timezone.normalizeTimezone('eastern')).toBe('Eastern Standard Time');
    });

    it('should return input if it contains spaces (assumed MS Graph format)', () => {
      expect(graphHelpers.timezone.normalizeTimezone('Custom Standard Time')).toBe('Custom Standard Time');
    });

    it('should return UTC for unknown single-word timezone', () => {
      expect(graphHelpers.timezone.normalizeTimezone('XYZABC')).toBe('UTC');
    });
  });

  describe('timezone.createDateTime', () => {
    it('should handle Date objects', () => {
      const d = new Date('2024-06-15T12:00:00Z');
      const result = graphHelpers.timezone.createDateTime(d);
      expect(result.dateTime).toBe(d.toISOString());
      expect(result.timeZone).toBe('UTC');
    });

    it('should handle ISO strings with T and Z', () => {
      const result = graphHelpers.timezone.createDateTime('2024-06-15T12:00:00Z');
      expect(result.dateTime).toBe('2024-06-15T12:00:00Z');
    });

    it('should append Z to ISO strings with T but no Z', () => {
      const result = graphHelpers.timezone.createDateTime('2024-06-15T12:00:00');
      expect(result.dateTime).toBe('2024-06-15T12:00:00Z');
    });

    it('should convert date strings without T', () => {
      const result = graphHelpers.timezone.createDateTime('2024-06-15');
      expect(result.dateTime).toContain('2024');
    });

    it('should return validation error for non-string, non-Date', () => {
      const result = graphHelpers.timezone.createDateTime(12345);
      expect(result.isError).toBe(true);
    });
  });

  describe('timezone.createDateTimeFromLocal', () => {
    it('should build a datetime string from components', () => {
      const result = graphHelpers.timezone.createDateTimeFromLocal(2024, 3, 15, 9, 30, 0, 'EST');
      expect(result.dateTime).toContain('2024-03-15T09:30:00');
      expect(result.timeZone).toBe('Eastern Standard Time');
    });

    it('should pad single-digit values', () => {
      const result = graphHelpers.timezone.createDateTimeFromLocal(2024, 1, 5, 8, 5, 3);
      expect(result.dateTime).toContain('01-05T08:05:03');
    });
  });

  describe('timezone.createAllDayDateTime', () => {
    it('should handle Date objects', () => {
      const result = graphHelpers.timezone.createAllDayDateTime(new Date('2024-06-15'));
      expect(result.dateTime).toContain('T00:00:00.0000000');
    });

    it('should handle date strings', () => {
      const result = graphHelpers.timezone.createAllDayDateTime('2024-06-15');
      expect(result.dateTime).toBe('2024-06-15T00:00:00.0000000');
    });

    it('should strip time portion from ISO strings', () => {
      const result = graphHelpers.timezone.createAllDayDateTime('2024-06-15T10:30:00Z');
      expect(result.dateTime).toBe('2024-06-15T00:00:00.0000000');
    });

    it('should return validation error for invalid input', () => {
      const result = graphHelpers.timezone.createAllDayDateTime(12345);
      expect(result.isError).toBe(true);
    });
  });

  describe('timezone.parseGraphDateTime', () => {
    it('should parse a valid graph datetime', () => {
      const result = graphHelpers.timezone.parseGraphDateTime({ dateTime: '2024-06-15T12:00:00Z' });
      expect(result).toBeInstanceOf(Date);
    });

    it('should return null for null/missing input', () => {
      expect(graphHelpers.timezone.parseGraphDateTime(null)).toBeNull();
      expect(graphHelpers.timezone.parseGraphDateTime({})).toBeNull();
    });
  });

  describe('timezone.now', () => {
    it('should return current time in graph format', () => {
      const result = graphHelpers.timezone.now();
      expect(result.dateTime).toBeDefined();
      expect(result.timeZone).toBe('UTC');
    });
  });

  describe('timezone.addDuration', () => {
    it('should add minutes to a datetime', () => {
      const start = { dateTime: '2024-06-15T12:00:00Z', timeZone: 'UTC' };
      const result = graphHelpers.timezone.addDuration(start, 30);
      const parsed = new Date(result.dateTime);
      expect(parsed.getUTCMinutes()).toBe(30);
    });

    it('should return null for invalid datetime', () => {
      expect(graphHelpers.timezone.addDuration({}, 30)).toBeNull();
    });
  });

  describe('timezone.dateRangesOverlap', () => {
    const make = (dt) => ({ dateTime: dt });

    it('should detect overlapping ranges', () => {
      expect(graphHelpers.timezone.dateRangesOverlap(
        make('2024-01-01T10:00:00Z'), make('2024-01-01T12:00:00Z'),
        make('2024-01-01T11:00:00Z'), make('2024-01-01T13:00:00Z')
      )).toBe(true);
    });

    it('should detect non-overlapping ranges', () => {
      expect(graphHelpers.timezone.dateRangesOverlap(
        make('2024-01-01T10:00:00Z'), make('2024-01-01T11:00:00Z'),
        make('2024-01-01T12:00:00Z'), make('2024-01-01T13:00:00Z')
      )).toBe(false);
    });

    it('should return false for invalid inputs', () => {
      expect(graphHelpers.timezone.dateRangesOverlap({}, {}, {}, {})).toBe(false);
    });
  });

  describe('timezone.validateDateTime', () => {
    it('should validate a correct datetime', () => {
      const result = graphHelpers.timezone.validateDateTime({ dateTime: '2024-06-15T12:00:00Z', timeZone: 'UTC' });
      expect(result.valid).toBe(true);
    });

    it('should reject non-object input', () => {
      expect(graphHelpers.timezone.validateDateTime(null).isError).toBe(true);
      expect(graphHelpers.timezone.validateDateTime('string').isError).toBe(true);
    });

    it('should reject missing dateTime property', () => {
      expect(graphHelpers.timezone.validateDateTime({ timeZone: 'UTC' }).isError).toBe(true);
    });

    it('should reject invalid dateTime value', () => {
      expect(graphHelpers.timezone.validateDateTime({ dateTime: 'not-a-date' }).isError).toBe(true);
    });
  });

  describe('timezone.createWorkingHours', () => {
    it('should return defaults', () => {
      const wh = graphHelpers.timezone.createWorkingHours();
      expect(wh.startTime).toBe('09:00:00');
      expect(wh.endTime).toBe('17:00:00');
      expect(wh.daysOfWeek).toHaveLength(5);
      expect(wh.timeZone).toBe('UTC');
    });

    it('should accept custom values', () => {
      const wh = graphHelpers.timezone.createWorkingHours('08:00:00', '16:00:00', ['monday'], 'EST');
      expect(wh.startTime).toBe('08:00:00');
      expect(wh.timeZone).toBe('Eastern Standard Time');
    });
  });

  // ── Enhanced calendar helpers ──────────────────────────────────

  describe('calendar.buildEventObject', () => {
    it('should build a basic event using timezone helpers', () => {
      const evt = graphHelpers.calendar.buildEventObject('Mtg', '2024-06-15T10:00:00Z', '2024-06-15T11:00:00Z');
      expect(evt.subject).toBe('Mtg');
      expect(evt.start.dateTime).toBeDefined();
      expect(evt.end.dateTime).toBeDefined();
    });

    it('should handle all-day events', () => {
      const evt = graphHelpers.calendar.buildEventObject('AllDay', '2024-06-15', '2024-06-16', { isAllDay: true });
      expect(evt.isAllDay).toBe(true);
      expect(evt.start.dateTime).toContain('T00:00:00.0000000');
    });

    it('should include body, location, attendees, recurrence, online meeting', () => {
      const evt = graphHelpers.calendar.buildEventObject('Mtg', '2024-06-15T10:00:00Z', '2024-06-15T11:00:00Z', {
        body: 'Notes', bodyType: 'html', location: 'Room B',
        attendees: ['x@t.com'], recurrence: {}, isOnlineMeeting: true
      });
      expect(evt.body.contentType).toBe('HTML');
      expect(evt.location.displayName).toBe('Room B');
      expect(evt.attendees).toHaveLength(1);
      expect(evt.isOnlineMeeting).toBe(true);
    });
  });

  describe('calendar.buildRecurrencePattern', () => {
    it('should delegate to the same logic as legacy', () => {
      const rec = graphHelpers.calendar.buildRecurrencePattern(
        { type: 'daily', interval: 1 },
        { type: 'noEnd', startDate: '2024-01-01' }
      );
      expect(rec.pattern.type).toBe('daily');
    });
  });

  describe('calendar.parseDateTimeWithZone', () => {
    it('should use timezone.createDateTime', () => {
      const result = graphHelpers.calendar.parseDateTimeWithZone('2024-06-15T10:00:00Z', 'PST');
      expect(result.timeZone).toBe('Pacific Standard Time');
    });
  });

  // ── Contact helpers ────────────────────────────────────────────

  describe('contact.buildContactObject', () => {
    it('should build basic contact with auto displayName', () => {
      const c = graphHelpers.contact.buildContactObject('John', 'Doe');
      expect(c.displayName).toBe('John Doe');
    });

    it('should use custom displayName', () => {
      const c = graphHelpers.contact.buildContactObject('J', 'D', { displayName: 'JD' });
      expect(c.displayName).toBe('JD');
    });

    it('should handle email addresses as strings', () => {
      const c = graphHelpers.contact.buildContactObject('J', 'D', { emailAddresses: ['j@t.com'] });
      expect(c.emailAddresses[0].address).toBe('j@t.com');
    });

    it('should handle email addresses as objects', () => {
      const c = graphHelpers.contact.buildContactObject('J', 'D', {
        emailAddresses: [{ address: 'j@t.com', name: 'J' }]
      });
      expect(c.emailAddresses[0].name).toBe('J');
    });

    it('should handle businessPhones as string', () => {
      const c = graphHelpers.contact.buildContactObject('J', 'D', { businessPhones: '555-1234' });
      expect(c.businessPhones).toEqual(['555-1234']);
    });

    it('should handle businessPhones as array', () => {
      const c = graphHelpers.contact.buildContactObject('J', 'D', { businessPhones: ['555-1234'] });
      expect(c.businessPhones).toEqual(['555-1234']);
    });

    it('should include all optional fields', () => {
      const c = graphHelpers.contact.buildContactObject('J', 'D', {
        mobilePhone: '555-5678', jobTitle: 'Dev', companyName: 'Acme',
        department: 'Eng', businessAddress: { street: '123 Main' }
      });
      expect(c.mobilePhone).toBe('555-5678');
      expect(c.jobTitle).toBe('Dev');
      expect(c.companyName).toBe('Acme');
      expect(c.department).toBe('Eng');
      expect(c.businessAddress.street).toBe('123 Main');
    });
  });

  // ── Task helpers ───────────────────────────────────────────────

  describe('task.buildTaskObject', () => {
    it('should build a basic task with default status', () => {
      const t = graphHelpers.task.buildTaskObject('Buy milk');
      expect(t.title).toBe('Buy milk');
      expect(t.status).toBe('notStarted');
    });

    it('should set custom status', () => {
      const t = graphHelpers.task.buildTaskObject('X', { status: 'completed' });
      expect(t.status).toBe('completed');
    });

    it('should include body, dueDateTime, startDateTime, importance, recurrence, categories', () => {
      const t = graphHelpers.task.buildTaskObject('T', {
        body: 'details', bodyType: 'html',
        dueDateTime: '2024-06-15', startDateTime: '2024-06-10',
        timeZone: 'EST', importance: 'high',
        recurrence: {}, categories: ['work']
      });
      expect(t.body.contentType).toBe('HTML');
      expect(t.dueDateTime.timeZone).toBe('EST');
      expect(t.startDateTime.dateTime).toBe('2024-06-10');
      expect(t.importance).toBe('high');
      expect(t.categories).toEqual(['work']);
    });
  });

  // ── General helpers ────────────────────────────────────────────

  describe('general.buildODataFilter', () => {
    it('should return null for empty/null filters', () => {
      expect(graphHelpers.general.buildODataFilter(null)).toBeNull();
      expect(graphHelpers.general.buildODataFilter({})).toBeNull();
    });

    it('should handle string values', () => {
      expect(graphHelpers.general.buildODataFilter({ name: 'John' })).toBe("name eq 'John'");
    });

    it('should handle boolean values', () => {
      expect(graphHelpers.general.buildODataFilter({ isRead: true })).toBe('isRead eq true');
    });

    it('should handle Date values', () => {
      const d = new Date('2024-06-15T00:00:00Z');
      const result = graphHelpers.general.buildODataFilter({ created: d });
      expect(result).toContain('created eq');
      expect(result).toContain('2024');
    });

    it('should skip null/undefined values', () => {
      expect(graphHelpers.general.buildODataFilter({ a: null, b: undefined, c: 'ok' })).toBe("c eq 'ok'");
    });

    it('should handle $gt operator', () => {
      const result = graphHelpers.general.buildODataFilter({ age: { $gt: 10 } });
      expect(result).toBe('age gt 10');
    });

    it('should handle $gte operator', () => {
      expect(graphHelpers.general.buildODataFilter({ age: { $gte: 10 } })).toBe('age ge 10');
    });

    it('should handle $lt operator', () => {
      expect(graphHelpers.general.buildODataFilter({ age: { $lt: 10 } })).toBe('age lt 10');
    });

    it('should handle $lte operator', () => {
      expect(graphHelpers.general.buildODataFilter({ age: { $lte: 10 } })).toBe('age le 10');
    });

    it('should handle $ne operator', () => {
      expect(graphHelpers.general.buildODataFilter({ status: { $ne: 'draft' } })).toBe("status ne 'draft'");
    });

    it('should handle $contains operator', () => {
      expect(graphHelpers.general.buildODataFilter({ subject: { $contains: 'hello' } })).toBe("contains(subject, 'hello')");
    });

    it('should handle $startswith operator', () => {
      expect(graphHelpers.general.buildODataFilter({ subject: { $startswith: 'Re' } })).toBe("startswith(subject, 'Re')");
    });

    it('should handle $gt with Date values', () => {
      const d = new Date('2024-01-01T00:00:00Z');
      const result = graphHelpers.general.buildODataFilter({ created: { $gt: d } });
      expect(result).toContain('created gt');
      expect(result).toContain('2024');
    });

    it('should combine multiple filters with and', () => {
      const result = graphHelpers.general.buildODataFilter({ isRead: true, subject: { $contains: 'hi' } });
      expect(result).toContain(' and ');
    });
  });

  describe('general.parseGraphError', () => {
    it('should parse Graph API error body', () => {
      const error = {
        body: { error: { code: 'ErrorNotFound', message: 'Item not found', innerError: {} } }
      };
      const result = graphHelpers.general.parseGraphError(error);
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Item not found');
    });

    it('should handle generic error with message', () => {
      const result = graphHelpers.general.parseGraphError({ message: 'Something failed' });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('Something failed');
    });

    it('should handle error without message', () => {
      const result = graphHelpers.general.parseGraphError({});
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toContain('unknown error');
    });
  });

  describe('general.formatFileSize', () => {
    it('should return "0 Bytes" for zero', () => {
      expect(graphHelpers.general.formatFileSize(0)).toBe('0 Bytes');
    });

    it('should format bytes', () => {
      expect(graphHelpers.general.formatFileSize(500)).toBe('500 Bytes');
    });

    it('should format KB', () => {
      expect(graphHelpers.general.formatFileSize(2048)).toBe('2 KB');
    });

    it('should format MB', () => {
      expect(graphHelpers.general.formatFileSize(1048576)).toBe('1 MB');
    });

    it('should format GB', () => {
      expect(graphHelpers.general.formatFileSize(1073741824)).toBe('1 GB');
    });

    it('should return "Unknown size" for null/undefined/NaN/non-number', () => {
      expect(graphHelpers.general.formatFileSize(null)).toBe('Unknown size');
      expect(graphHelpers.general.formatFileSize(undefined)).toBe('Unknown size');
      expect(graphHelpers.general.formatFileSize(NaN)).toBe('Unknown size');
      expect(graphHelpers.general.formatFileSize('abc')).toBe('Unknown size');
    });
  });
});

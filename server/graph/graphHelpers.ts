// Helper functions for common Graph API operations
import { convertErrorToToolError, createServiceUnavailableError, createRateLimitError, createValidationError } from '../utils/mcpErrorResponse.js';

export const graphHelpers = {
  // Email helpers
  email: {
    buildMessageObject(to: string | string[], subject: string, body: string, options: Record<string, any> = {}) {
      const message: Record<string, any> = {
        subject,
        body: {
          contentType: options.bodyType === 'html' ? 'HTML' : 'Text',
          content: body,
        },
        toRecipients: Array.isArray(to) ? 
          to.map((email: string) => ({ emailAddress: { address: email } })) :
          [{ emailAddress: { address: to } }],
      };

      if (options.cc) {
        message.ccRecipients = options.cc.map((email: string) => ({
          emailAddress: { address: email },
        }));
      }

      if (options.bcc) {
        message.bccRecipients = options.bcc.map((email: string) => ({
          emailAddress: { address: email },
        }));
      }

      if (options.importance) {
        message.importance = options.importance; // low, normal, high
      }

      if (options.attachments) {
        message.attachments = options.attachments;
      }

      return message;
    },

    buildReplyObject(body: string, options: Record<string, any> = {}) {
      const reply: Record<string, any> = {
        comment: body,
      };

      if (options.replyAll) {
        reply.message = {};
        if (options.cc) {
          reply.message.ccRecipients = options.cc.map((email: string) => ({
            emailAddress: { address: email },
          }));
        }
      }

      return reply;
    },

    parseEmailAddress(emailObject: Record<string, any> | string) {
      if (typeof emailObject === 'string') return emailObject;
      return emailObject?.emailAddress?.address || 'unknown';
    },

    parseEmailName(emailObject: Record<string, any> | string) {
      if (typeof emailObject === 'string') return null;
      return emailObject?.emailAddress?.name || null;
    },
  },

  // Calendar helpers
  _calendarLegacy: {
    buildEventObject(subject: string, start: Record<string, any> | string, end: Record<string, any> | string, options: Record<string, any> = {}) {
      const event: Record<string, any> = {
        subject,
        start: {
          dateTime: (start as any).dateTime || start,
          timeZone: (start as any).timeZone || 'UTC',
        },
        end: {
          dateTime: (end as any).dateTime || end,
          timeZone: (end as any).timeZone || 'UTC',
        },
      };

      if (options.body) {
        event.body = {
          contentType: options.bodyType === 'html' ? 'HTML' : 'Text',
          content: options.body,
        };
      }

      if (options.location) {
        event.location = {
          displayName: options.location,
        };
      }

      if (options.attendees) {
        event.attendees = options.attendees.map((email: string) => ({
          emailAddress: { address: email },
          type: 'required',
        }));
      }

      if (options.isAllDay) {
        event.isAllDay = true;
      }

      if (options.recurrence) {
        event.recurrence = options.recurrence;
      }

      if (options.isOnlineMeeting) {
        event.isOnlineMeeting = true;
        event.onlineMeetingProvider = options.onlineMeetingProvider || 'teamsForBusiness';
      }

      return event;
    },

    buildRecurrencePattern(pattern: Record<string, any>, range: Record<string, any>) {
      const recurrence: Record<string, any> = {
        pattern: {
          type: pattern.type, // daily, weekly, absoluteMonthly, relativeMonthly, absoluteYearly, relativeYearly
          interval: pattern.interval || 1,
        },
        range: {
          type: range.type, // endDate, noEnd, numbered
          startDate: range.startDate,
        },
      };

      if (pattern.daysOfWeek) {
        recurrence.pattern.daysOfWeek = pattern.daysOfWeek;
      }

      if (pattern.dayOfMonth) {
        recurrence.pattern.dayOfMonth = pattern.dayOfMonth;
      }

      if (range.type === 'endDate') {
        recurrence.range.endDate = range.endDate;
      } else if (range.type === 'numbered') {
        recurrence.range.numberOfOccurrences = range.numberOfOccurrences;
      }

      return recurrence;
    },

    parseDateTimeWithZone(dateTime: string, timeZone = 'UTC') {
      return {
        dateTime: dateTime,
        timeZone: timeZone,
      };
    },
  },

  // Timezone handling utilities
  timezone: {
    // Map common timezone names to Microsoft Graph timezone identifiers
    timezoneMap: {
      'UTC': 'UTC',
      'GMT': 'Greenwich Standard Time',
      'EST': 'Eastern Standard Time',
      'CST': 'Central Standard Time',
      'MST': 'Mountain Standard Time',
      'PST': 'Pacific Standard Time',
      'EDT': 'Eastern Daylight Time',
      'CDT': 'Central Daylight Time',
      'MDT': 'Mountain Daylight Time',
      'PDT': 'Pacific Daylight Time',
      'New York': 'Eastern Standard Time',
      'Chicago': 'Central Standard Time',
      'Denver': 'Mountain Standard Time',
      'Los Angeles': 'Pacific Standard Time',
      'London': 'GMT Standard Time',
      'Paris': 'W. Europe Standard Time',
      'Tokyo': 'Tokyo Standard Time',
      'Sydney': 'AUS Eastern Standard Time',
      'India': 'India Standard Time',
      'Beijing': 'China Standard Time'
    },

    // Detect and normalize timezone input
    normalizeTimezone(timezone: string) {
      if (!timezone) return 'UTC';
      
      // Check if it's already a valid Microsoft Graph timezone
      if ((this.timezoneMap as Record<string, string>)[timezone]) {
        return (this.timezoneMap as Record<string, string>)[timezone];
      }
      
      // Try to find a partial match
      const lowerTimezone = (timezone as string).toLowerCase();
      for (const [key, value] of Object.entries(this.timezoneMap) as [string, string][]) {
        if (key.toLowerCase().includes(lowerTimezone) || 
            value.toLowerCase().includes(lowerTimezone)) {
          return value;
        }
      }
      
      // If no match found, assume it's already a Microsoft Graph timezone or return UTC
      return timezone.includes(' ') ? timezone : 'UTC';
    },

    // Create a Microsoft Graph datetime object with timezone
    createDateTime(dateTime: Date | string, timeZone = 'UTC') {
      // Handle various input formats
      let normalizedDateTime;
      
      if (dateTime instanceof Date) {
        normalizedDateTime = dateTime.toISOString();
      } else if (typeof dateTime === 'string') {
        // Check if it's already in ISO format
        if (dateTime.includes('T') && dateTime.includes('Z')) {
          normalizedDateTime = dateTime;
        } else if (dateTime.includes('T')) {
          normalizedDateTime = dateTime + (dateTime.endsWith('Z') ? '' : 'Z');
        } else {
          // Assume it's a date string and convert
          normalizedDateTime = new Date(dateTime).toISOString();
        }
      } else {
        return createValidationError('dateTime', 'Expected Date object or ISO string');
      }

      return {
        dateTime: normalizedDateTime,
        timeZone: this.normalizeTimezone(timeZone)
      };
    },

    // Convert a local datetime to Microsoft Graph format
    createDateTimeFromLocal(year: number, month: number, day: number, hour = 0, minute = 0, second = 0, timeZone = 'UTC') {
      // Create date in the specified timezone (simplified approach)
      const dateStr = `${year}-${month.toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}T${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}:${second.toString().padStart(2, '0')}`;
      
      return this.createDateTime(dateStr, timeZone);
    },

    // Create an all-day event datetime
    createAllDayDateTime(date: Date | string, timeZone = 'UTC') {
      let dateStr;
      
      if (date instanceof Date) {
        dateStr = date.toISOString().split('T')[0];
      } else if (typeof date === 'string') {
        dateStr = date.split('T')[0];
      } else {
        return createValidationError('date', 'Invalid date format for all-day event');
      }
      
      return {
        dateTime: dateStr + 'T00:00:00.0000000',
        timeZone: this.normalizeTimezone(timeZone)
      };
    },

    // Parse Microsoft Graph datetime back to JavaScript Date
    parseGraphDateTime(graphDateTime: Record<string, any> | null) {
      if (!graphDateTime || !graphDateTime.dateTime) {
        return null;
      }
      
      return new Date(graphDateTime.dateTime);
    },

    // Get the current time in Microsoft Graph format
    now(timeZone = 'UTC') {
      return this.createDateTime(new Date(), timeZone);
    },

    // Add duration to a datetime
    addDuration(graphDateTime: Record<string, any>, durationMinutes: number) {
      const date = this.parseGraphDateTime(graphDateTime);
      if (!date) return null;
      
      date.setMinutes(date.getMinutes() + durationMinutes);
      
      return this.createDateTime(date, graphDateTime.timeZone);
    },

    // Check if two datetime ranges overlap
    dateRangesOverlap(start1: Record<string, any>, end1: Record<string, any>, start2: Record<string, any>, end2: Record<string, any>) {
      const s1 = this.parseGraphDateTime(start1);
      const e1 = this.parseGraphDateTime(end1);
      const s2 = this.parseGraphDateTime(start2);
      const e2 = this.parseGraphDateTime(end2);
      
      if (!s1 || !e1 || !s2 || !e2) return false;
      
      return s1 < e2 && s2 < e1;
    },

    // Validate a datetime object
    validateDateTime(dateTime: Record<string, any>) {
      if (!dateTime || typeof dateTime !== 'object') {
        return createValidationError('dateTime', 'DateTime must be an object');
      }
      
      if (!dateTime.dateTime) {
        return createValidationError('dateTime', 'dateTime property is required');
      }
      
      try {
        const date = new Date(dateTime.dateTime);
        if (isNaN(date.getTime())) {
          return createValidationError('dateTime', 'Invalid dateTime value');
        }
      } catch (error) {
        return createValidationError('dateTime', 'Invalid dateTime format');
      }
      
      if (dateTime.timeZone && !this.normalizeTimezone(dateTime.timeZone)) {
        return createValidationError('timeZone', 'Invalid timezone');
      }
      
      return { valid: true };
    },

    // Get working hours in Microsoft Graph format
    createWorkingHours(startTime = '09:00:00', endTime = '17:00:00', daysOfWeek: string[] = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'], timeZone = 'UTC') {
      return {
        daysOfWeek,
        startTime,
        endTime,
        timeZone: this.normalizeTimezone(timeZone)
      };
    }
  },

  // Enhanced calendar helpers with timezone support
  calendar: {
    buildEventObject(subject: string, start: Date | string, end: Date | string, options: Record<string, any> = {}) {
      const event: Record<string, any> = {
        subject,
        start: graphHelpers.timezone.createDateTime(start, options.startTimeZone),
        end: graphHelpers.timezone.createDateTime(end, options.endTimeZone || options.startTimeZone),
      };

      if (options.body) {
        event.body = {
          contentType: options.bodyType === 'html' ? 'HTML' : 'Text',
          content: options.body,
        };
      }

      if (options.location) {
        event.location = {
          displayName: options.location,
        };
      }

      if (options.attendees) {
        event.attendees = options.attendees.map((email: string) => ({
          emailAddress: { address: email },
          type: 'required',
        }));
      }

      if (options.isAllDay) {
        event.isAllDay = true;
        event.start = graphHelpers.timezone.createAllDayDateTime(start, options.startTimeZone);
        event.end = graphHelpers.timezone.createAllDayDateTime(end, options.endTimeZone || options.startTimeZone);
      }

      if (options.recurrence) {
        event.recurrence = options.recurrence;
      }

      if (options.isOnlineMeeting) {
        event.isOnlineMeeting = true;
        event.onlineMeetingProvider = options.onlineMeetingProvider || 'teamsForBusiness';
      }

      return event;
    },

    buildRecurrencePattern(pattern: Record<string, any>, range: Record<string, any>) {
      const recurrence: Record<string, any> = {
        pattern: {
          type: pattern.type, // daily, weekly, absoluteMonthly, relativeMonthly, absoluteYearly, relativeYearly
          interval: pattern.interval || 1,
        },
        range: {
          type: range.type, // endDate, noEnd, numbered
          startDate: range.startDate,
        },
      };

      if (pattern.daysOfWeek) {
        recurrence.pattern.daysOfWeek = pattern.daysOfWeek;
      }

      if (pattern.dayOfMonth) {
        recurrence.pattern.dayOfMonth = pattern.dayOfMonth;
      }

      if (range.type === 'endDate') {
        recurrence.range.endDate = range.endDate;
      } else if (range.type === 'numbered') {
        recurrence.range.numberOfOccurrences = range.numberOfOccurrences;
      }

      return recurrence;
    },

    parseDateTimeWithZone(dateTime: Date | string, timeZone = 'UTC') {
      return graphHelpers.timezone.createDateTime(dateTime, timeZone);
    },
  },

  // Contact helpers
  contact: {
    buildContactObject(givenName: string, surname: string, options: Record<string, any> = {}) {
      const contact: Record<string, any> = {
        givenName,
        surname,
      };

      if (options.displayName) {
        contact.displayName = options.displayName;
      } else {
        contact.displayName = `${givenName} ${surname}`;
      }

      if (options.emailAddresses) {
        contact.emailAddresses = options.emailAddresses.map((email: any) => ({
          address: email.address || email,
          name: email.name || contact.displayName,
        }));
      }

      if (options.businessPhones) {
        contact.businessPhones = Array.isArray(options.businessPhones) 
          ? options.businessPhones 
          : [options.businessPhones];
      }

      if (options.mobilePhone) {
        contact.mobilePhone = options.mobilePhone;
      }

      if (options.jobTitle) {
        contact.jobTitle = options.jobTitle;
      }

      if (options.companyName) {
        contact.companyName = options.companyName;
      }

      if (options.department) {
        contact.department = options.department;
      }

      if (options.businessAddress) {
        contact.businessAddress = options.businessAddress;
      }

      return contact;
    },
  },

  // Task helpers
  task: {
    buildTaskObject(title: string, options: Record<string, any> = {}) {
      const task: Record<string, any> = {
        title,
        status: options.status || 'notStarted', // notStarted, inProgress, completed, waitingOnOthers, deferred
      };

      if (options.body) {
        task.body = {
          contentType: options.bodyType === 'html' ? 'HTML' : 'Text',
          content: options.body,
        };
      }

      if (options.dueDateTime) {
        task.dueDateTime = {
          dateTime: options.dueDateTime,
          timeZone: options.timeZone || 'UTC',
        };
      }

      if (options.startDateTime) {
        task.startDateTime = {
          dateTime: options.startDateTime,
          timeZone: options.timeZone || 'UTC',
        };
      }

      if (options.importance) {
        task.importance = options.importance; // low, normal, high
      }

      if (options.recurrence) {
        task.recurrence = options.recurrence;
      }

      if (options.categories) {
        task.categories = options.categories;
      }

      return task;
    },
  },

  // General helpers
  general: {
    buildODataFilter(filters: Record<string, any>) {
      if (!filters || Object.keys(filters).length === 0) return null;

      const filterStrings = [];

      for (const [key, value] of Object.entries(filters)) {
        if (value === null || value === undefined) continue;

        if (typeof value === 'string') {
          filterStrings.push(`${key} eq '${value}'`);
        } else if (typeof value === 'boolean') {
          filterStrings.push(`${key} eq ${value}`);
        } else if (value instanceof Date) {
          filterStrings.push(`${key} eq ${value.toISOString()}`);
        } else if (typeof value === 'object') {
          // Handle complex filters like { $gt: date }
          for (const [operator, val] of Object.entries(value)) {
            switch (operator) {
              case '$gt':
                filterStrings.push(`${key} gt ${val instanceof Date ? val.toISOString() : val}`);
                break;
              case '$gte':
                filterStrings.push(`${key} ge ${val instanceof Date ? val.toISOString() : val}`);
                break;
              case '$lt':
                filterStrings.push(`${key} lt ${val instanceof Date ? val.toISOString() : val}`);
                break;
              case '$lte':
                filterStrings.push(`${key} le ${val instanceof Date ? val.toISOString() : val}`);
                break;
              case '$ne':
                filterStrings.push(`${key} ne '${val}'`);
                break;
              case '$contains':
                filterStrings.push(`contains(${key}, '${val}')`);
                break;
              case '$startswith':
                filterStrings.push(`startswith(${key}, '${val}')`);
                break;
            }
          }
        }
      }

      return filterStrings.join(' and ');
    },

    parseGraphError(error: Record<string, any>) {
      if (error.body?.error) {
        const graphError = {
          code: error.body.error.code,
          message: error.body.error.message,
          innerError: error.body.error.innerError,
        };
        
        // Return MCP error format instead of plain object
        const finalError = new Error(graphError.message || 'Graph API error') as any;
        finalError.code = graphError.code;
        finalError.innerError = graphError.innerError;
        
        return convertErrorToToolError(finalError, 'Graph API');
      }
      
      const message = error.message || 'An unknown error occurred';
      const finalError = new Error(message) as any;
      finalError.code = 'Unknown';
      
      return convertErrorToToolError(finalError, 'Graph API');
    },

    // Format file size for display
    formatFileSize(bytes: number) {
      // Handle null, undefined, or invalid byte values
      if (bytes === null || bytes === undefined || isNaN(bytes) || typeof bytes !== 'number') {
        return 'Unknown size';
      }
      
      const sizes = ['Bytes', 'KB', 'MB', 'GB'];
      if (bytes === 0) return '0 Bytes';
      
      // Ensure bytes is a positive number
      const absBytes = Math.abs(bytes);
      const i = Math.floor(Math.log(absBytes) / Math.log(1024));
      const sizeIndex = Math.min(i, sizes.length - 1); // Cap at largest unit
      
      const value = Math.round(absBytes / Math.pow(1024, sizeIndex) * 100) / 100;
      return value + ' ' + sizes[sizeIndex];
    },
  },
};
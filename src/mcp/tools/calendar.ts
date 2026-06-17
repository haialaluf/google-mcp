import { z } from 'zod';
import dayjs from 'dayjs';
import utc from 'dayjs/plugin/utc';
import timezone from 'dayjs/plugin/timezone';
import { calendar, EVENT_COLORS, type EventColor } from '../../lib/google.ts';
import type { ToolContext } from '../server.ts';

dayjs.extend(utc);
dayjs.extend(timezone);

// A datetime that already carries a 'Z' or '±HH:MM' offset, or an all-day date.
const HAS_ZONE = /([zZ]|[+-]\d{2}:?\d{2})$/;
const DATE_ONLY = /^\d{4}-\d{2}-\d{2}$/;

// Resolve the calendar's IANA timezone (e.g. 'Asia/Jerusalem'), defaulting to UTC.
async function resolveTimeZone(context: ToolContext, calendarId: string): Promise<string> {
  const { items } = await calendar.listCalendars({ accessToken: context.accessToken });
  const entry = calendarId === 'primary'
    ? items.find((c) => c.primary) ?? items[0]
    : items.find((c) => c.id === calendarId) ?? items.find((c) => c.primary);
  return entry?.timeZone ?? 'UTC';
}

// Interpret a naive datetime in the calendar's zone. Explicit offsets and
// all-day dates are left as-is — assume calendar time unless told otherwise.
function qualify(value: string, timeZone: string): string {
  const v = value.trim();
  if (DATE_ONLY.test(v) || HAS_ZONE.test(v)) return v;
  return dayjs.tz(v, timeZone).format(); // ISO 8601 with the zone's offset
}

export const calendarTools = {
  list_calendars: {
    product: 'calendar' as const,
    scopes: ['https://www.googleapis.com/auth/calendar.calendarlist.readonly'],
    description: 'List all calendars accessible to the user',
    parameters: z.object({}),
    execute: async (context: ToolContext) => {
      const result = await calendar.listCalendars({ accessToken: context.accessToken });
      return result.items.map((cal) => ({
        id: cal.id,
        name: cal.summary,
        description: cal.description,
        primary: cal.primary ?? false,
        accessRole: cal.accessRole,
        timeZone: cal.timeZone,
      }));
    },
  },

  check_availability: {
    product: 'calendar' as const,
    scopes: ['https://www.googleapis.com/auth/calendar.events'],
    description: 'Check calendar availability (busy/free times) without exposing event details',
    parameters: z.object({
      calendarId: z.string().default('primary').describe('Calendar ID to check'),
      timeMin: z.string().describe('Start of time range in ISO 8601 format'),
      timeMax: z.string().describe('End of time range in ISO 8601 format'),
    }),
    execute: async (context: ToolContext, params: {
      calendarId: string;
      timeMin: string;
      timeMax: string;
    }) => {
      const timeZone = await resolveTimeZone(context, params.calendarId);
      const result = await calendar.freeBusy({ accessToken: context.accessToken }, {
        timeMin: qualify(params.timeMin, timeZone),
        timeMax: qualify(params.timeMax, timeZone),
        calendarIds: [params.calendarId],
      });

      const calendarData = result.calendars[params.calendarId];
      if (!calendarData) {
        return { error: 'Calendar not found or not accessible' };
      }

      return { timeZone, busy: calendarData.busy };
    },
  },

  list_events: {
    product: 'calendar' as const,
    scopes: ['https://www.googleapis.com/auth/calendar.events'],
    description: 'List events from a calendar within a time range',
    parameters: z.object({
      calendarId: z.string().default('primary').describe('Calendar ID, defaults to primary calendar'),
      timeMin: z.string().optional().describe('Start time in ISO 8601 format (e.g., 2024-01-01T00:00:00Z)'),
      timeMax: z.string().optional().describe('End time in ISO 8601 format'),
      maxResults: z.number().optional().default(50).describe('Maximum number of events to return'),
      query: z.string().optional().describe('Free text search query'),
    }),
    execute: async (context: ToolContext, params: {
      calendarId: string;
      timeMin?: string;
      timeMax?: string;
      maxResults?: number;
      query?: string;
    }) => {
      const timeZone = await resolveTimeZone(context, params.calendarId);
      const result = await calendar.listEvents({ accessToken: context.accessToken }, params.calendarId, {
        timeMin: params.timeMin && qualify(params.timeMin, timeZone),
        timeMax: params.timeMax && qualify(params.timeMax, timeZone),
        maxResults: params.maxResults,
        q: params.query,
        timeZone,
      });
      const events = result.items.map((event) => ({
        id: event.id,
        summary: event.summary,
        description: event.description,
        start: event.start,
        end: event.end,
        location: event.location,
        status: event.status,
        link: event.htmlLink,
        attendees: event.attendees?.map((a) => ({
          email: a.email,
          status: a.responseStatus,
        })),
      }));
      return { timeZone, items: events };
    },
  },

  create_event: {
    product: 'calendar' as const,
    scopes: ['https://www.googleapis.com/auth/calendar.events'],
    description: 'Create a new calendar event. Attendees receive email invitations automatically.',
    parameters: z.object({
      calendarId: z.string().default('primary').describe('Calendar ID, defaults to primary calendar'),
      summary: z.string().describe('Event title'),
      description: z.string().optional().describe('Event description'),
      startDateTime: z.string().describe('Start time in ISO 8601 format'),
      endDateTime: z.string().describe('End time in ISO 8601 format'),
      timeZone: z.string().optional().describe('Time zone (e.g., America/New_York)'),
      location: z.string().optional().describe('Event location'),
      attendees: z.array(z.string()).optional().describe('List of attendee email addresses'),
      color: z.enum(['lavender', 'sage', 'grape', 'flamingo', 'banana', 'tangerine', 'peacock', 'graphite', 'blueberry', 'basil', 'tomato']).optional().describe('Event color'),
    }),
    execute: async (context: ToolContext, params: {
      calendarId: string;
      summary: string;
      description?: string;
      startDateTime: string;
      endDateTime: string;
      timeZone?: string;
      location?: string;
      attendees?: string[];
      color?: EventColor;
    }) => {
      // Naive datetimes are interpreted in the calendar's zone unless the
      // caller passed an explicit one.
      const timeZone = params.timeZone ?? await resolveTimeZone(context, params.calendarId);
      // Detect all-day events (YYYY-MM-DD format) vs timed events
      const isAllDay = (val: string) => /^\d{4}-\d{2}-\d{2}$/.test(val);
      const event = await calendar.createEvent(
        { accessToken: context.accessToken },
        params.calendarId,
        {
          summary: params.summary,
          description: params.description,
          start: isAllDay(params.startDateTime)
            ? { date: params.startDateTime }
            : { dateTime: params.startDateTime, timeZone },
          end: isAllDay(params.endDateTime)
            ? { date: params.endDateTime }
            : { dateTime: params.endDateTime, timeZone },
          location: params.location,
          attendees: params.attendees?.map((email) => ({ email })),
          guestsCanModify: false,
          guestsCanInviteOthers: false,
          guestsCanSeeOtherGuests: false,
          colorId: params.color ? EVENT_COLORS[params.color] : undefined,
        },
        'all'
      );
      return {
        id: event.id,
        summary: event.summary,
        start: event.start,
        end: event.end,
        link: event.htmlLink,
        timeZone,
      };
    },
  },

  update_event: {
    product: 'calendar' as const,
    scopes: ['https://www.googleapis.com/auth/calendar.events'],
    description: 'Update an existing calendar event',
    parameters: z.object({
      calendarId: z.string().default('primary').describe('Calendar ID'),
      eventId: z.string().describe('Event ID to update'),
      summary: z.string().optional().describe('New event title'),
      description: z.string().optional().describe('New event description'),
      startDateTime: z.string().optional().describe('New start time in ISO 8601 format'),
      endDateTime: z.string().optional().describe('New end time in ISO 8601 format'),
      timeZone: z.string().optional().describe('Time zone'),
      location: z.string().optional().describe('New event location'),
      color: z.enum(['lavender', 'sage', 'grape', 'flamingo', 'banana', 'tangerine', 'peacock', 'graphite', 'blueberry', 'basil', 'tomato']).optional().describe('Event color'),
    }),
    execute: async (context: ToolContext, params: {
      calendarId: string;
      eventId: string;
      summary?: string;
      description?: string;
      startDateTime?: string;
      endDateTime?: string;
      timeZone?: string;
      location?: string;
      color?: EventColor;
    }) => {
      const updateData: Record<string, unknown> = {};
      if (params.summary) updateData.summary = params.summary;
      if (params.description) updateData.description = params.description;
      if (params.location) updateData.location = params.location;
      if (params.color) updateData.colorId = EVENT_COLORS[params.color];
      // Detect all-day events (YYYY-MM-DD format) vs timed events
      const isAllDay = (val: string) => /^\d{4}-\d{2}-\d{2}$/.test(val);
      // Only resolve the calendar zone when a timed datetime is actually changing.
      const timeZone = (params.startDateTime || params.endDateTime)
        ? params.timeZone ?? await resolveTimeZone(context, params.calendarId)
        : params.timeZone;
      if (params.startDateTime) {
        updateData.start = isAllDay(params.startDateTime)
          ? { date: params.startDateTime }
          : { dateTime: params.startDateTime, timeZone };
      }
      if (params.endDateTime) {
        updateData.end = isAllDay(params.endDateTime)
          ? { date: params.endDateTime }
          : { dateTime: params.endDateTime, timeZone };
      }

      const event = await calendar.updateEvent(
        { accessToken: context.accessToken },
        params.calendarId,
        params.eventId,
        updateData
      );
      return {
        id: event.id,
        summary: event.summary,
        start: event.start,
        end: event.end,
        link: event.htmlLink,
      };
    },
  },

  delete_event: {
    product: 'calendar' as const,
    scopes: ['https://www.googleapis.com/auth/calendar.events'],
    description: 'Delete a calendar event',
    parameters: z.object({
      calendarId: z.string().default('primary').describe('Calendar ID'),
      eventId: z.string().describe('Event ID to delete'),
    }),
    execute: async (context: ToolContext, params: { calendarId: string; eventId: string }) => {
      await calendar.deleteEvent({ accessToken: context.accessToken }, params.calendarId, params.eventId);
      return { success: true, message: `Event ${params.eventId} deleted` };
    },
  },
};

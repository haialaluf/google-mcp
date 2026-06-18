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

// Key under extendedProperties.private that records the contact who owns an
// event. The calendar is shared across all of an org's contacts, so this stamp
// is what scopes read/modify to a single contact.
const OWNER_KEY = 'acrm_contact_id';

// Reject the operation unless the target event belongs to the current contact.
// Skipped for trusted internal/operator calls (no contactId). Events with no
// owner stamp (created directly in Google Calendar, or before this guard) are
// treated as not owned by any contact and therefore not modifiable by one.
async function assertEventOwnership(
  context: ToolContext,
  calendarId: string,
  eventId: string
): Promise<void> {
  if (!context.contactId) return;

  const event = await calendar.getEvent({ accessToken: context.accessToken }, calendarId, eventId);

  if (event.extendedProperties?.private?.[OWNER_KEY] !== context.contactId) {
    // Same message whether the event belongs to another contact or has no
    // stamp — avoids leaking whether/whose event it is.
    throw new Error('This event does not belong to you and cannot be modified.');
  }
}

// Resolve the calendar's IANA timezone (e.g. 'Asia/Jerusalem'), defaulting to UTC.
async function resolveTimeZone(context: ToolContext, calendarId: string): Promise<string> {
  const { items } = await calendar.listCalendars({ accessToken: context.accessToken });
  const entry = calendarId === 'primary'
    ? items.find((c) => c.primary) ?? items[0]
    : items.find((c) => c.id === calendarId) ?? items.find((c) => c.primary);
  return entry?.timeZone ?? 'UTC';
}

const emailSchema = z.email();

// Keep only the entries that are valid email addresses. Returns undefined when
// nothing valid remains so the event is created without an attendees field.
function validAttendees(attendees?: string[]): string[] | undefined {
  const valid = attendees
    ?.map((a) => a.trim())
    .filter((a) => emailSchema.safeParse(a).success);
  return valid?.length ? valid : undefined;
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
    scopes: [
      'https://www.googleapis.com/auth/calendar.events.freebusy',
      'https://www.googleapis.com/auth/calendar.calendarlist.readonly',
    ],
    description: 'Check calendar availability (busy/free times) without exposing event details. Query a wider time range than the user requested (e.g. the surrounding days or week, not just the exact slot) so that if the requested time is busy you can offer nearby free alternatives in your response.',
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

      // Google returns busy intervals in UTC ('...Z'). Re-render them in the
      // calendar's timezone so the times match the reported `timeZone` instead
      // of silently mixing zones (e.g. 10:00Z is 13:00 in Asia/Jerusalem).
      const busy = calendarData.busy.map((slot) => ({
        start: dayjs(slot.start).tz(timeZone).format(),
        end: dayjs(slot.end).tz(timeZone).format(),
      }));

      return { timeZone, busy };
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
        // Contact-scoped calls only ever see their own events; operator calls
        // (no contactId) see everything.
        privateExtendedProperty: context.contactId ? `${OWNER_KEY}=${context.contactId}` : undefined,
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
      attendees: z.array(z.string()).optional().describe('Optional list of attendee email addresses. Only pass real email addresses — never phone numbers or names. Omit if no email is known.'),
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
          attendees: validAttendees(params.attendees)?.map((email) => ({ email })),
          guestsCanModify: false,
          guestsCanInviteOthers: false,
          guestsCanSeeOtherGuests: false,
          colorId: params.color ? EVENT_COLORS[params.color] : undefined,
          // Stamp ownership so this contact (and only this contact) can later
          // list/update/delete it. Omitted for operator calls.
          extendedProperties: context.contactId
            ? { private: { [OWNER_KEY]: context.contactId } }
            : undefined,
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
      await assertEventOwnership(context, params.calendarId, params.eventId);

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
      await assertEventOwnership(context, params.calendarId, params.eventId);

      await calendar.deleteEvent({ accessToken: context.accessToken }, params.calendarId, params.eventId);
      return { success: true, message: `Event ${params.eventId} deleted` };
    },
  },
};

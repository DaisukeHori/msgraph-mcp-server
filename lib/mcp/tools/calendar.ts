import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import {
  graphGet,
  graphPost,
  graphPatch,
  graphDelete,
  truncateResponse,
  handleToolError,
  GraphPagedResponse,
} from "@/lib/msgraph/graph-client";
import { CalendarEvent } from "@/lib/msgraph/types";
import { DEFAULT_PAGE_SIZE } from "@/lib/config";

export function registerCalendarTools(server: McpServer): void {
  // -------------------------------------------------------
  // calendar_list_events
  // -------------------------------------------------------
  server.registerTool(
    "calendar_list_events",
    {
      title: "List Calendar Events",
      description: `List calendar events for the signed-in user.
Can filter by date range, search, etc.

Args:
  - start_datetime: Start of date range (ISO 8601, e.g. "2025-03-01T00:00:00Z")
  - end_datetime: End of date range (ISO 8601)
  - search: Search query
  - filter: OData filter
  - top: Number of events (1-50, default 25)
  - orderby: Sort order (default "start/dateTime")
  - calendar_id: Specific calendar ID (default: primary)

Returns: List of events with subject, start/end, location, organizer, attendees`,
      inputSchema: {
        start_datetime: z.string().optional().describe("Start datetime ISO 8601"),
        end_datetime: z.string().optional().describe("End datetime ISO 8601"),
        search: z.string().optional().describe("Search query"),
        filter: z.string().optional().describe("OData filter"),
        top: z.number().int().min(1).max(50).default(DEFAULT_PAGE_SIZE).describe("Number of events"),
        orderby: z.string().default("start/dateTime").describe("Sort order"),
        calendar_id: z.string().optional().describe("Calendar ID"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        // If date range provided, use calendarView
        if (params.start_datetime && params.end_datetime) {
          const calPath = params.calendar_id
            ? `/me/calendars/${params.calendar_id}/calendarView`
            : "/me/calendarView";
          const data = await graphGet<GraphPagedResponse<CalendarEvent>>(calPath, {
            startDateTime: params.start_datetime,
            endDateTime: params.end_datetime,
            $top: params.top,
            $orderby: params.orderby,
            $select: "id,subject,bodyPreview,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeetingUrl,isAllDay,showAs,importance",
          });

          return {
            content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: data.value.length, events: data.value }, null, 2)) }],
          };
        }

        // Otherwise use events endpoint
        const evtPath = params.calendar_id
          ? `/me/calendars/${params.calendar_id}/events`
          : "/me/events";
        const queryParams: Record<string, string | number | boolean | undefined> = {
          $top: params.top,
          $orderby: params.orderby,
          $select: "id,subject,bodyPreview,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeetingUrl,isAllDay,showAs,importance",
        };
        if (params.search) queryParams.$search = `"${params.search}"`;
        if (params.filter) queryParams.$filter = params.filter;

        const data = await graphGet<GraphPagedResponse<CalendarEvent>>(evtPath, queryParams);

        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify({ count: data.value.length, events: data.value }, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // calendar_get_event
  // -------------------------------------------------------
  server.registerTool(
    "calendar_get_event",
    {
      title: "Get Calendar Event",
      description: `Get a specific calendar event by ID with full details.

Args:
  - event_id (required): Event ID

Returns: Full event details including body, attendees, recurrence`,
      inputSchema: {
        event_id: z.string().min(1).describe("Event ID"),
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const event = await graphGet<CalendarEvent>(`/me/events/${params.event_id}`);
        return {
          content: [{ type: "text", text: truncateResponse(JSON.stringify(event, null, 2)) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // calendar_create_event
  // -------------------------------------------------------
  server.registerTool(
    "calendar_create_event",
    {
      title: "Create Calendar Event",
      description: `Create a new calendar event.

Args:
  - subject (required): Event subject/title
  - start_datetime (required): Start datetime (ISO 8601, e.g. "2025-04-01T09:00:00")
  - end_datetime (required): End datetime (ISO 8601)
  - timezone: Timezone (default "Asia/Tokyo")
  - body: Event body/description
  - body_type: "HTML" or "Text" (default "HTML")
  - location: Location display name
  - attendees: Array of {email, name?, type?} (type: "required"|"optional")
  - is_online_meeting: Create as online meeting (default false)
  - is_all_day: All-day event (default false)
  - show_as: "free"|"tentative"|"busy"|"oof"|"workingElsewhere" (default "busy")
  - reminder_minutes: Reminder before event in minutes

Returns: Created event details`,
      inputSchema: {
        subject: z.string().min(1).describe("Event subject"),
        start_datetime: z.string().min(1).describe("Start datetime ISO 8601"),
        end_datetime: z.string().min(1).describe("End datetime ISO 8601"),
        timezone: z.string().default("Asia/Tokyo").describe("Timezone"),
        body: z.string().optional().describe("Event body"),
        body_type: z.enum(["HTML", "Text"]).default("HTML").describe("Body type"),
        location: z.string().optional().describe("Location"),
        attendees: z
          .array(
            z.object({
              email: z.string().email(),
              name: z.string().optional(),
              type: z.enum(["required", "optional"]).default("required"),
            })
          )
          .optional()
          .describe("Attendees"),
        is_online_meeting: z.boolean().default(false).describe("Online meeting"),
        is_all_day: z.boolean().default(false).describe("All-day event"),
        show_as: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere"]).default("busy").describe("Show as"),
        reminder_minutes: z.number().int().min(0).optional().describe("Reminder minutes before"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const eventBody: Record<string, unknown> = {
          subject: params.subject,
          start: { dateTime: params.start_datetime, timeZone: params.timezone },
          end: { dateTime: params.end_datetime, timeZone: params.timezone },
          isOnlineMeeting: params.is_online_meeting,
          isAllDay: params.is_all_day,
          showAs: params.show_as,
        };
        if (params.body) {
          eventBody.body = { contentType: params.body_type, content: params.body };
        }
        if (params.location) {
          eventBody.location = { displayName: params.location };
        }
        if (params.attendees) {
          eventBody.attendees = params.attendees.map((a) => ({
            emailAddress: { address: a.email, name: a.name },
            type: a.type,
          }));
        }
        if (params.reminder_minutes !== undefined) {
          eventBody.isReminderOn = true;
          eventBody.reminderMinutesBeforeStart = params.reminder_minutes;
        }

        const created = await graphPost<CalendarEvent>("/me/events", eventBody);

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: created.id, subject: created.subject, webLink: created.webLink }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // calendar_update_event
  // -------------------------------------------------------
  server.registerTool(
    "calendar_update_event",
    {
      title: "Update Calendar Event",
      description: `Update an existing calendar event. Only specified fields are updated.

Args:
  - event_id (required): Event ID to update
  - subject: New subject
  - start_datetime / end_datetime / timezone: Updated times
  - body / body_type: Updated body
  - location: Updated location
  - show_as: Updated show-as status
  - is_online_meeting: Updated online meeting flag

Returns: Updated event`,
      inputSchema: {
        event_id: z.string().min(1).describe("Event ID"),
        subject: z.string().optional().describe("New subject"),
        start_datetime: z.string().optional().describe("New start datetime"),
        end_datetime: z.string().optional().describe("New end datetime"),
        timezone: z.string().default("Asia/Tokyo").describe("Timezone"),
        body: z.string().optional().describe("New body"),
        body_type: z.enum(["HTML", "Text"]).default("HTML").describe("Body type"),
        location: z.string().optional().describe("New location"),
        show_as: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere"]).optional().describe("Show as"),
        is_online_meeting: z.boolean().optional().describe("Online meeting"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        const updates: Record<string, unknown> = {};
        if (params.subject) updates.subject = params.subject;
        if (params.start_datetime) updates.start = { dateTime: params.start_datetime, timeZone: params.timezone };
        if (params.end_datetime) updates.end = { dateTime: params.end_datetime, timeZone: params.timezone };
        if (params.body) updates.body = { contentType: params.body_type, content: params.body };
        if (params.location) updates.location = { displayName: params.location };
        if (params.show_as) updates.showAs = params.show_as;
        if (params.is_online_meeting !== undefined) updates.isOnlineMeeting = params.is_online_meeting;

        const updated = await graphPatch<CalendarEvent>(`/me/events/${params.event_id}`, updates);

        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, id: updated.id, subject: updated.subject }, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );

  // -------------------------------------------------------
  // calendar_delete_event
  // -------------------------------------------------------
  server.registerTool(
    "calendar_delete_event",
    {
      title: "Delete Calendar Event",
      description: `Delete a calendar event.

Args:
  - event_id (required): Event ID to delete

Returns: Confirmation`,
      inputSchema: {
        event_id: z.string().min(1).describe("Event ID"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params) => {
      try {
        await graphDelete(`/me/events/${params.event_id}`);
        return {
          content: [{ type: "text", text: JSON.stringify({ success: true, message: "Event deleted" }) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleToolError(error) }] };
      }
    }
  );
}

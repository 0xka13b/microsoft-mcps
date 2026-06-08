import { z } from "zod";
import { defineTool } from "@microsoft-mcp/core";
import { escapeKql, validateId } from "@microsoft-mcp/validation";

export const tools = [
  defineTool({
    name: "me",
    description: "Get the signed-in user's profile (id, displayName, mail).",
    inputSchema: {},
    confirmationPolicy: "never",
    handler: ({ graph }) =>
      graph.request("GET", "/me", undefined, { $select: "id,displayName,mail" }),
  }),

  defineTool({
    name: "list_events",
    description:
      "List calendar events within a date range. Recurring events are expanded.",
    inputSchema: {
      days_ahead: z.number().int().optional().describe("Days into the future. Defaults to 7."),
      days_back: z.number().int().optional().describe("Days into the past. Defaults to 0."),
      include_details: z
        .boolean()
        .optional()
        .describe("Include event body and attendees. Defaults to true."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const now = new Date();
      const start = new Date(now);
      start.setDate(start.getDate() - (params.days_back ?? 0));
      const end = new Date(now);
      end.setDate(end.getDate() + (params.days_ahead ?? 7));

      const select = [
        "id",
        "subject",
        "start",
        "end",
        "location",
        "organizer",
        "isCancelled",
        "isOnlineMeeting",
        "webLink",
      ];
      if (params.include_details !== false) {
        select.push("body", "attendees", "onlineMeeting");
      }

      const data = await graph.request("GET", "/me/calendarView", undefined, {
        startDateTime: start.toISOString(),
        endDateTime: end.toISOString(),
        $select: select.join(","),
        $orderby: "start/dateTime asc",
        $top: "100",
      });
      return data.value ?? [];
    },
  }),

  defineTool({
    name: "get_event",
    description: "Get a single calendar event by ID.",
    inputSchema: {
      event_id: z.string().describe("Event ID"),
    },
    confirmationPolicy: "never",
    handler: ({ graph }, params) => {
      validateId(params.event_id, "event_id");
      return graph.request("GET", `/me/events/${params.event_id}`);
    },
  }),

  defineTool({
    name: "create_event",
    description: "Create a new calendar event.",
    inputSchema: {
      subject: z.string().describe("Event title"),
      start: z
        .string()
        .describe("Start datetime in ISO 8601 format, e.g. '2024-01-15T10:00:00'"),
      end: z
        .string()
        .describe("End datetime in ISO 8601 format, e.g. '2024-01-15T11:00:00'"),
      timezone: z
        .string()
        .optional()
        .describe("Timezone for start/end, e.g. 'America/New_York'. Defaults to UTC."),
      location: z.string().optional().describe("Event location display name"),
      body: z.string().optional().describe("Event description"),
      body_type: z
        .enum(["text", "html"])
        .optional()
        .describe("Body content type. Defaults to text."),
      attendees: z.array(z.string()).optional().describe("Attendee email addresses"),
      is_online_meeting: z.boolean().optional().describe("Create as an online meeting"),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) => {
      const tz = params.timezone ?? "UTC";
      const event: any = {
        subject: params.subject,
        start: { dateTime: params.start, timeZone: tz },
        end: { dateTime: params.end, timeZone: tz },
      };

      if (params.location) event.location = { displayName: params.location };
      if (params.body) {
        event.body = {
          contentType: params.body_type === "html" ? "HTML" : "Text",
          content: params.body,
        };
      }
      if (params.attendees?.length) {
        event.attendees = params.attendees.map((email) => ({
          emailAddress: { address: email },
          type: "required",
        }));
      }
      if (params.is_online_meeting) event.isOnlineMeeting = true;

      return graph.request("POST", "/me/events", event);
    },
  }),

  defineTool({
    name: "update_event",
    description: "Update calendar event properties.",
    inputSchema: {
      event_id: z.string().describe("Event ID to update"),
      updates: z
        .record(z.string(), z.unknown())
        .describe(
          'Properties to update, e.g. {"subject": "New Title", "location": {"displayName": "Room 1"}}',
        ),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) => {
      validateId(params.event_id, "event_id");
      return graph.request("PATCH", `/me/events/${params.event_id}`, params.updates);
    },
  }),

  defineTool({
    name: "delete_event",
    description:
      "Delete a calendar event. If the user is the organizer, cancellation notices are sent to attendees.",
    inputSchema: {
      event_id: z.string().describe("Event ID to delete"),
      send_cancellation: z
        .boolean()
        .optional()
        .describe("Send cancellation email to attendees (organizer only). Defaults to true."),
    },
    confirmationPolicy: "always",
    handler: async ({ graph, log }, params) => {
      validateId(params.event_id, "event_id");
      if (params.send_cancellation !== false) {
        try {
          await graph.request("POST", `/me/events/${params.event_id}/cancel`, { comment: "" });
          return { success: true };
        } catch (err: any) {
          if (err.status !== 403) throw err;
          log.info("cancel skipped, not organizer — falling through to DELETE", {
            event_id: params.event_id,
          });
        }
      }
      await graph.request("DELETE", `/me/events/${params.event_id}`);
      return { success: true };
    },
  }),

  defineTool({
    name: "respond_event",
    description: "Accept, decline, or tentatively accept a calendar invitation.",
    inputSchema: {
      event_id: z.string().describe("Event ID to respond to"),
      response: z
        .enum(["accept", "decline", "tentativelyAccept"])
        .optional()
        .describe("Response type. Defaults to accept."),
      message: z.string().optional().describe("Optional message to include with the response"),
      send_response: z
        .boolean()
        .optional()
        .describe("Send response email to organizer. Defaults to true."),
    },
    confirmationPolicy: "always",
    handler: async ({ graph }, params) => {
      validateId(params.event_id, "event_id");
      const action = params.response ?? "accept";
      const body: any = { sendResponse: params.send_response ?? true };
      if (params.message) body.comment = params.message;

      await graph.request("POST", `/me/events/${params.event_id}/${action}`, body);
      return { success: true };
    },
  }),

  defineTool({
    name: "check_availability",
    description:
      "Check free/busy schedule for the current user and optional additional attendees within a time window.",
    inputSchema: {
      start: z.string().describe("Start datetime in ISO 8601 format"),
      end: z.string().describe("End datetime in ISO 8601 format"),
      timezone: z.string().optional().describe("Timezone for start/end. Defaults to UTC."),
      attendees: z
        .array(z.string())
        .optional()
        .describe(
          "Additional attendee email addresses to check (current user is always included)",
        ),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const tz = params.timezone ?? "UTC";

      const me = await graph.request("GET", "/me", undefined, { $select: "mail" });
      const myEmail: string = me.mail ?? me.userPrincipalName ?? "";

      const schedules = [
        myEmail,
        ...(params.attendees ?? []).filter((e) => e !== myEmail),
      ].filter(Boolean);

      return graph.request("POST", "/me/calendar/getSchedule", {
        schedules,
        startTime: { dateTime: params.start, timeZone: tz },
        endTime: { dateTime: params.end, timeZone: tz },
        availabilityViewInterval: 30,
      });
    },
  }),

  defineTool({
    name: "search_events",
    description: "Search calendar events by keyword.",
    inputSchema: {
      query: z.string().describe("Search query string"),
      limit: z.number().int().min(1).max(100).optional().describe("Max results. Defaults to 50."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const data = await graph.request("GET", "/me/events", undefined, {
        $search: `"${escapeKql(params.query)}"`,
        $top: String(params.limit ?? 50),
        $select: "id,subject,start,end,location,organizer,isCancelled,webLink",
      });
      return data.value ?? [];
    },
  }),
];

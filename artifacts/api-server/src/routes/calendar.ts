import { and, desc, eq, gte, isNull, lte } from "drizzle-orm";
import { Router, type IRouter, type Request } from "express";
import { randomUUID } from "node:crypto";

import {
  activityEntriesTable,
  automationJobsTable,
  calendarEventsTable,
  db,
  googleAccountsTable,
  remindersTable,
  type CalendarEventRecord,
  type GoogleAccount,
} from "@workspace/db";
import {
  CreateEventBody,
  CreateEventResponse,
  DeleteEventParams,
  GetEventParams,
  GetEventResponse,
  ListCalendarsResponse,
  ListEventsQueryParams,
  ListEventsResponse,
  ListRemindersResponse,
  UpdateEventBody,
  UpdateEventParams,
  UpdateEventResponse,
  UpdateReminderBody,
  UpdateReminderParams,
  UpdateReminderResponse,
  ListJobsQueryParams,
  ListJobsResponse,
  RetryJobParams,
  RetryJobResponse,
  ListActivityResponse,
} from "@workspace/api-zod";

import { requireGoogleAccount } from "../middlewares/google-account";
import {
  buildGoogleEventPayload,
  calendarApi,
  googleEventToDto,
  mapReminder,
  mapStoredEvent,
  prepareReminders,
  recurrenceFrequency,
  validateEventInput,
  type EventInput,
} from "../lib/calendar-service";
import { persistGoogleEvent, replaceEventReminders } from "../lib/event-persistence";

const router: IRouter = Router();
router.use(requireGoogleAccount);

function getAccount(req: Request): GoogleAccount {
  if (!req.googleAccount) throw new Error("Google account was not loaded.");
  return req.googleAccount;
}

function responseError(error: unknown): { status?: number; message: string } {
  const candidate = error as {
    code?: number | string;
    status?: number;
    response?: { status?: number };
    message?: string;
  };
  const status =
    typeof candidate?.response?.status === "number"
      ? candidate.response.status
      : typeof candidate?.status === "number"
        ? candidate.status
        : typeof candidate?.code === "number"
          ? candidate.code
          : undefined;
  return { status, message: candidate?.message ?? "Unknown error" };
}

function sendGoogleError(req: Request, res: import("express").Response, error: unknown): void {
  const details = responseError(error);
  req.log.error({ googleStatus: details.status, message: details.message }, "Google Calendar request failed");
  if (details.status === 401) {
    res.status(401).json({ error: "Google authorization expired. Reconnect your account to continue." });
  } else if (details.status === 403) {
    res.status(403).json({ error: "Google has not granted the permissions needed for this action." });
  } else {
    res.status(502).json({ error: "Google Calendar could not complete the request. Please try again." });
  }
}

function isUuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}

async function remindersForEvents(eventIds: string[]) {
  if (!eventIds.length) return [];
  return db
    .select()
    .from(remindersTable)
    .where(
      and(
        ...eventIds.map((eventId) => eq(remindersTable.eventId, eventId)),
      ),
    );
}

export async function loadEventDtos(
  account: GoogleAccount,
  options: { from?: string; to?: string; search?: string } = {},
): Promise<Record<string, unknown>[]> {
  const api = await calendarApi(account);
  const now = new Date();
  const timeMin = options.from ?? new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000).toISOString();
  const eventResponse = await api.events.list({
    calendarId: "primary",
    timeMin,
    timeMax: options.to,
    q: options.search || undefined,
    singleEvents: true,
    orderBy: "startTime",
    maxResults: 250,
  });

  const rows = await db
    .select()
    .from(calendarEventsTable)
    .where(
      and(
        eq(calendarEventsTable.googleAccountId, account.id),
        isNull(calendarEventsTable.cancelledAt),
        gte(calendarEventsTable.startAt, new Date(timeMin)),
        options.to ? lte(calendarEventsTable.startAt, new Date(options.to)) : undefined,
      ),
    );
  const eventIds = rows.map((row) => row.id);
  const reminders = await remindersForEvents(eventIds);
  const remindersByEventId = new Map<string, typeof reminders>();
  for (const reminder of reminders) {
    const list = remindersByEventId.get(reminder.eventId) ?? [];
    list.push(reminder);
    remindersByEventId.set(reminder.eventId, list);
  }
  const byGoogleId = new Map(rows.map((row) => [row.googleEventId, row]));
  const foundIds = new Set<string>();
  const dtos: Record<string, unknown>[] = [];

  for (const event of eventResponse.data.items ?? []) {
    if (!event.id) continue;
    foundIds.add(event.id);
    const local = byGoogleId.get(event.id);
    const dto = googleEventToDto(
      event,
      "primary",
      local,
      local ? remindersByEventId.get(local.id) ?? [] : [],
    );
    if (dto) dtos.push(dto);
  }

  for (const row of rows) {
    if (row.calendarId === "primary" || foundIds.has(row.googleEventId)) continue;
    dtos.push(mapStoredEvent(row, remindersByEventId.get(row.id) ?? []));
  }

  return dtos
    .filter((event) => {
      const startsAt = new Date(String(event.startAt)).getTime();
      if (options.from && startsAt < new Date(options.from).getTime()) return false;
      if (options.to && startsAt > new Date(options.to).getTime()) return false;
      if (!options.search) return true;
      const searchable = [
        event.title,
        event.description,
        ...(Array.isArray(event.guestEmails) ? event.guestEmails : []),
      ]
        .join(" ")
        .toLowerCase();
      return searchable.includes(options.search.toLowerCase());
    })
    .sort((a, b) => new Date(String(a.startAt)).getTime() - new Date(String(b.startAt)).getTime());
}

router.get("/calendars", async (req, res): Promise<void> => {
  try {
    const api = await calendarApi(getAccount(req));
    const result = await api.calendarList.list({ maxResults: 100 });
    const calendars = (result.data.items ?? []).map((calendar) => ({
      id: calendar.id ?? "",
      name: calendar.summary ?? calendar.id ?? "Calendar",
      primary: Boolean(calendar.primary),
      timeZone: calendar.timeZone ?? "UTC",
    }));
    res.json(ListCalendarsResponse.parse(calendars));
  } catch (error) {
    sendGoogleError(req, res, error);
  }
});

router.get("/events", async (req, res): Promise<void> => {
  const query = ListEventsQueryParams.safeParse(req.query);
  if (!query.success) {
    res.status(400).json({ error: query.error.message });
    return;
  }
  try {
    const events = await loadEventDtos(getAccount(req), {
      from: query.data.from,
      to: query.data.to,
      search: query.data.search,
    });
    res.json(ListEventsResponse.parse(events));
  } catch (error) {
    sendGoogleError(req, res, error);
  }
});

router.post("/events", async (req, res): Promise<void> => {
  const parsed = CreateEventBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const invalid = validateEventInput(parsed.data);
  if (invalid) {
    res.status(400).json({ error: invalid });
    return;
  }
  let created: import("googleapis").calendar_v3.Schema$Event | undefined;
  const account = getAccount(req);
  try {
    const prepared = prepareReminders(
      new Date(parsed.data.startAt),
      parsed.data.timeZone,
      parsed.data.reminders,
    );
    const api = await calendarApi(account);
    const response = await api.events.insert({
      calendarId: parsed.data.calendarId,
      sendUpdates: "all",
      conferenceDataVersion: parsed.data.createMeet ? 1 : 0,
      requestBody: buildGoogleEventPayload(parsed.data, {
        googleOverrides: prepared.googleOverrides,
      }),
    });
    created = response.data;
    const saved = await persistGoogleEvent({
      accountId: account.id,
      eventInput: parsed.data,
      googleEvent: response.data,
      preparedReminders: prepared.rows,
    });
    res.status(201).json(
      CreateEventResponse.parse(mapStoredEvent(saved.event, saved.reminders)),
    );
  } catch (error) {
    if (created?.id) {
      try {
        await (await calendarApi(account)).events.delete({
          calendarId: parsed.data.calendarId,
          eventId: created.id,
          sendUpdates: "all",
        });
      } catch (cleanupError) {
        req.log.error(
          { message: responseError(cleanupError).message, googleEventId: created.id },
          "Could not clean up an event after a local save failure",
        );
      }
    }
    if (error instanceof Error && error.message.startsWith("Google Calendar pop-up")) {
      res.status(400).json({ error: error.message });
      return;
    }
    sendGoogleError(req, res, error);
  }
});

async function findLocalEvent(accountId: string, id: string): Promise<CalendarEventRecord | undefined> {
  if (!isUuid(id)) return undefined;
  const [event] = await db
    .select()
    .from(calendarEventsTable)
    .where(
      and(
        eq(calendarEventsTable.id, id),
        eq(calendarEventsTable.googleAccountId, accountId),
        isNull(calendarEventsTable.cancelledAt),
      ),
    )
    .limit(1);
  return event;
}

router.get("/events/:eventId", async (req, res): Promise<void> => {
  const params = GetEventParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const account = getAccount(req);
  try {
    const local = await findLocalEvent(account.id, params.data.eventId);
    if (local) {
      const reminders = await db
        .select()
        .from(remindersTable)
        .where(eq(remindersTable.eventId, local.id))
        .orderBy(remindersTable.scheduledAt);
      res.json(GetEventResponse.parse(mapStoredEvent(local, reminders)));
      return;
    }
    const api = await calendarApi(account);
    const result = await api.events.get({
      calendarId: "primary",
      eventId: params.data.eventId,
    });
    const dto = googleEventToDto(result.data, "primary", undefined, []);
    if (!dto) {
      res.status(404).json({ error: "Event not found." });
      return;
    }
    res.json(GetEventResponse.parse(dto));
  } catch (error) {
    const details = responseError(error);
    if (details.status === 404) {
      res.status(404).json({ error: "Event not found." });
      return;
    }
    sendGoogleError(req, res, error);
  }
});

router.patch("/events/:eventId", async (req, res): Promise<void> => {
  const params = UpdateEventParams.safeParse(req.params);
  const parsed = UpdateEventBody.safeParse(req.body);
  if (!params.success || !parsed.success) {
    res.status(400).json({
      error: params.success ? parsed.error.message : params.error.message,
    });
    return;
  }

  const account = getAccount(req);
  try {
    const api = await calendarApi(account);
    let local = await findLocalEvent(account.id, params.data.eventId);
    let calendarId = local?.calendarId ?? "primary";
    let googleEventId = local?.googleEventId ?? params.data.eventId;
    const currentResponse = await api.events.get({ calendarId, eventId: googleEventId });
    const current = currentResponse.data;
    const [storedReminders] = local
      ? [await db.select().from(remindersTable).where(eq(remindersTable.eventId, local.id))]
      : [[]];
    const reminderInputs = (parsed.data.reminders ??
      storedReminders.map((reminder) => reminder.configuration)) as EventInput["reminders"];
    const fullInput: EventInput = {
      title: parsed.data.title ?? local?.title ?? current.summary ?? "(Untitled event)",
      startAt:
        parsed.data.startAt ??
        current.start?.dateTime ??
        (current.start?.date ? `${current.start.date}T00:00:00.000Z` : ""),
      endAt:
        parsed.data.endAt ??
        current.end?.dateTime ??
        (current.end?.date ? `${current.end.date}T00:00:00.000Z` : ""),
      allDay: parsed.data.allDay ?? local?.allDay ?? Boolean(current.start?.date),
      timeZone: parsed.data.timeZone ?? local?.timeZone ?? current.start?.timeZone ?? "UTC",
      description: parsed.data.description ?? local?.description ?? current.description ?? "",
      guestEmails:
        parsed.data.guestEmails ??
        local?.guestEmails ??
        (current.attendees ?? []).map((attendee) => attendee.email ?? "").filter(Boolean),
      calendarId: parsed.data.calendarId ?? calendarId,
      repeat: parsed.data.repeat ?? (local?.repeat as EventInput["repeat"] | undefined) ?? recurrenceFrequency(current),
      createMeet:
        parsed.data.createMeet ??
        Boolean(local?.meetUrl || current.conferenceData?.createRequest),
      reminders: reminderInputs,
    };
    const invalid = validateEventInput(fullInput);
    if (invalid) {
      res.status(400).json({ error: invalid });
      return;
    }
    const prepared = prepareReminders(
      new Date(fullInput.startAt),
      fullInput.timeZone,
      fullInput.reminders,
    );
    const paused = parsed.data.automationPaused ?? local?.automationPaused ?? false;
    if (fullInput.calendarId !== calendarId) {
      const moved = await api.events.move({
        calendarId,
        eventId: googleEventId,
        destination: fullInput.calendarId,
        sendUpdates: "all",
      });
      googleEventId = moved.data.id ?? googleEventId;
      calendarId = fullInput.calendarId;
    }

    const requestBody = buildGoogleEventPayload(fullInput, {
      googleOverrides: paused ? [] : prepared.googleOverrides,
    });
    const conferenceData = requestBody.conferenceData;
    delete requestBody.conferenceData;
    if (parsed.data.createMeet === true && !local?.meetUrl && !current.conferenceData) {
      requestBody.conferenceData = conferenceData;
    }
    const updatedResponse = await api.events.patch({
      calendarId,
      eventId: googleEventId,
      sendUpdates: "all",
      conferenceDataVersion: parsed.data.createMeet ? 1 : 0,
      requestBody,
    });
    const updatedGoogleEvent = updatedResponse.data;
    const eventValues = {
      googleAccountId: account.id,
      googleEventId,
      calendarId,
      title: fullInput.title.trim(),
      startAt: new Date(fullInput.startAt),
      endAt: new Date(fullInput.endAt),
      allDay: fullInput.allDay,
      timeZone: fullInput.timeZone,
      description: fullInput.description?.trim() || null,
      guestEmails: fullInput.guestEmails.map((email) => email.trim().toLowerCase()),
      repeat: fullInput.repeat,
      automationPaused: paused,
      meetUrl:
        updatedGoogleEvent.conferenceData?.entryPoints?.find((point) => point.entryPointType === "video")?.uri ??
        local?.meetUrl ??
        null,
      meetStatus:
        updatedGoogleEvent.conferenceData?.createRequest?.status?.statusCode === "failure"
          ? "failed"
          : updatedGoogleEvent.conferenceData?.entryPoints?.some((point) => point.entryPointType === "video")
            ? "ready"
            : fullInput.createMeet
              ? "pending"
              : "not_requested",
      updatedAt: new Date(),
      cancelledAt: null,
    };
    if (local) {
      const [saved] = await db
        .update(calendarEventsTable)
        .set(eventValues)
        .where(eq(calendarEventsTable.id, local.id))
        .returning();
      local = saved;
    } else {
      const [saved] = await db
        .insert(calendarEventsTable)
        .values(eventValues)
        .onConflictDoUpdate({
          target: [
            calendarEventsTable.googleAccountId,
            calendarEventsTable.calendarId,
            calendarEventsTable.googleEventId,
          ],
          set: eventValues,
        })
        .returning();
      local = saved;
    }
    const updatedReminders = await replaceEventReminders({
      accountId: account.id,
      event: local,
      reminders: fullInput.reminders,
      preparedReminders: prepared.rows,
      paused,
    });
    if (paused) {
      await db
        .update(remindersTable)
        .set({ enabled: false, status: "skipped", updatedAt: new Date() })
        .where(eq(remindersTable.eventId, local.id));
    }
    if (local.meetStatus === "pending") {
      await db
        .insert(automationJobsTable)
        .values({
          googleAccountId: account.id,
          kind: "conference_check",
          status: "queued",
          eventId: local.id,
          idempotencyKey: `conference-check:${local.id}`,
          payload: { calendarId, googleEventId },
          scheduledAt: new Date(Date.now() + 15_000),
        })
        .onConflictDoNothing();
    }
    await db.insert(activityEntriesTable).values({
      googleAccountId: account.id,
      kind: "event_updated",
      message: `Updated “${local.title}” in Google Calendar.`,
      status: "success",
    });
    res.json(UpdateEventResponse.parse(mapStoredEvent(local, updatedReminders)));
  } catch (error) {
    if (error instanceof Error && error.message.startsWith("Google Calendar pop-up")) {
      res.status(400).json({ error: error.message });
      return;
    }
    sendGoogleError(req, res, error);
  }
});

router.delete("/events/:eventId", async (req, res): Promise<void> => {
  const params = DeleteEventParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const account = getAccount(req);
  try {
    const local = await findLocalEvent(account.id, params.data.eventId);
    const calendarId = local?.calendarId ?? "primary";
    const googleEventId = local?.googleEventId ?? params.data.eventId;
    try {
      await (await calendarApi(account)).events.delete({
        calendarId,
        eventId: googleEventId,
        sendUpdates: "all",
      });
    } catch (error) {
      if (responseError(error).status !== 404) throw error;
    }
    if (local) {
      await db
        .update(calendarEventsTable)
        .set({ cancelledAt: new Date(), automationPaused: true, updatedAt: new Date() })
        .where(eq(calendarEventsTable.id, local.id));
      await db
        .update(remindersTable)
        .set({ enabled: false, status: "skipped", updatedAt: new Date() })
        .where(eq(remindersTable.eventId, local.id));
    }
    await db.insert(activityEntriesTable).values({
      googleAccountId: account.id,
      kind: "event_cancelled",
      message: "Cancelled an event in Google Calendar.",
      status: "success",
    });
    res.sendStatus(204);
  } catch (error) {
    sendGoogleError(req, res, error);
  }
});

router.get("/reminders", async (req, res): Promise<void> => {
  try {
    const account = getAccount(req);
    const reminders = await db
      .select({ reminder: remindersTable, eventTitle: calendarEventsTable.title })
      .from(remindersTable)
      .innerJoin(calendarEventsTable, eq(remindersTable.eventId, calendarEventsTable.id))
      .where(
        and(
          eq(remindersTable.googleAccountId, account.id),
          isNull(calendarEventsTable.cancelledAt),
        ),
      )
      .orderBy(desc(remindersTable.scheduledAt))
      .limit(200);
    res.json(
      ListRemindersResponse.parse(
        reminders.map(({ reminder, eventTitle }) => mapReminder(reminder, eventTitle)),
      ),
    );
  } catch (error) {
    req.log.error({ message: responseError(error).message }, "Could not list reminders");
    res.status(500).json({ error: "Reminders could not be loaded." });
  }
});

router.patch("/reminders/:reminderId", async (req, res): Promise<void> => {
  const params = UpdateReminderParams.safeParse(req.params);
  const body = UpdateReminderBody.safeParse(req.body);
  if (!params.success || !body.success) {
    res.status(400).json({ error: params.success ? body.error.message : params.error.message });
    return;
  }
  try {
    const account = getAccount(req);
    const [reminder] = await db
      .select()
      .from(remindersTable)
      .where(
        and(
          eq(remindersTable.id, params.data.reminderId),
          eq(remindersTable.googleAccountId, account.id),
        ),
      )
      .limit(1);
    if (!reminder) {
      res.status(404).json({ error: "Reminder not found." });
      return;
    }
    const [event] = await db
      .select()
      .from(calendarEventsTable)
      .where(eq(calendarEventsTable.id, reminder.eventId))
      .limit(1);
    const [updated] = await db
      .update(remindersTable)
      .set({
        enabled: body.data.enabled,
        status: body.data.enabled ? "scheduled" : "skipped",
        updatedAt: new Date(),
      })
      .where(eq(remindersTable.id, reminder.id))
      .returning();
    const response = mapReminder(updated, event?.title ?? "Calendar event");
    res.json(UpdateReminderResponse.parse(response));
  } catch (error) {
    req.log.error({ message: responseError(error).message }, "Could not update reminder");
    res.status(500).json({ error: "Reminder settings could not be saved." });
  }
});

router.get("/jobs", async (req, res): Promise<void> => {
  const query = ListJobsQueryParams.safeParse(req.query);
  if (!query.success) {
    res.status(400).json({ error: query.error.message });
    return;
  }
  try {
    const account = getAccount(req);
    const jobs = await db
      .select()
      .from(automationJobsTable)
      .where(
        and(
          eq(automationJobsTable.googleAccountId, account.id),
          query.data.status ? eq(automationJobsTable.status, query.data.status) : undefined,
        ),
      )
      .orderBy(desc(automationJobsTable.createdAt))
      .limit(100);
    res.json(
      ListJobsResponse.parse(
        jobs.map((job) => ({
          id: job.id,
          kind: job.kind,
          status: job.status,
          eventId: job.eventId,
          createdAt: job.createdAt.toISOString(),
          attempts: job.attempts,
          error: job.error,
        })),
      ),
    );
  } catch (error) {
    req.log.error({ message: responseError(error).message }, "Could not list automation jobs");
    res.status(500).json({ error: "Automation history could not be loaded." });
  }
});

router.post("/jobs/:jobId/retry", async (req, res): Promise<void> => {
  const params = RetryJobParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  try {
    const account = getAccount(req);
    const [job] = await db
      .select()
      .from(automationJobsTable)
      .where(
        and(
          eq(automationJobsTable.id, params.data.jobId),
          eq(automationJobsTable.googleAccountId, account.id),
        ),
      )
      .limit(1);
    if (!job) {
      res.status(404).json({ error: "Automation job not found." });
      return;
    }
    if (job.status !== "failed") {
      res.status(409).json({ error: "Only failed jobs can be retried." });
      return;
    }
    const [retried] = await db
      .update(automationJobsTable)
      .set({
        status: "queued",
        scheduledAt: new Date(),
        attempts: job.attempts + 1,
        error: null,
      })
      .where(eq(automationJobsTable.id, job.id))
      .returning();
    res.status(202).json(
      RetryJobResponse.parse({
        id: retried.id,
        kind: retried.kind,
        status: retried.status,
        eventId: retried.eventId,
        createdAt: retried.createdAt.toISOString(),
        attempts: retried.attempts,
        error: retried.error,
      }),
    );
  } catch (error) {
    req.log.error({ message: responseError(error).message }, "Could not retry automation job");
    res.status(500).json({ error: "The job could not be retried." });
  }
});

router.get("/activity", async (req, res): Promise<void> => {
  try {
    const account = getAccount(req);
    const activity = await db
      .select()
      .from(activityEntriesTable)
      .where(eq(activityEntriesTable.googleAccountId, account.id))
      .orderBy(desc(activityEntriesTable.createdAt))
      .limit(100);
    res.json(
      ListActivityResponse.parse(
        activity.map((entry) => ({
          id: entry.id,
          kind: entry.kind,
          message: entry.message,
          createdAt: entry.createdAt.toISOString(),
          status: entry.status,
        })),
      ),
    );
  } catch (error) {
    req.log.error({ message: responseError(error).message }, "Could not load activity");
    res.status(500).json({ error: "Automation activity could not be loaded." });
  }
});

export default router;

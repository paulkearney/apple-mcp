import { runAppleScript } from 'run-applescript';
import { runJxa } from './jxa';

/**
 * Patched calendar backend.
 *
 * Upstream getEvents/searchEvents never queried Calendar at all — they built a
 * hardcoded placeholder event titled "No events available - Calendar operations
 * too slow" and returned it, and the Array.isArray() bug then discarded even
 * that. openEvent merely activated the app.
 *
 * Calendar.app's AppleScript interface really is slow (a 7-day query over 20
 * calendars took >55s here). The fix is to bypass it: EventKit answers the same
 * query in ~0.12s via JXA's ObjC bridge. Writes still use AppleScript, which
 * worked and isn't latency-sensitive.
 */

interface CalendarEvent {
    id: string;
    title: string;
    location: string | null;
    notes: string | null;
    startDate: string | null;
    endDate: string | null;
    calendarName: string;
    isAllDay: boolean;
    url: string | null;
}

const CONFIG = {
    TIMEOUT_MS: 60000,
    MAX_EVENTS: 50,
};

/** EventKit rejects predicates spanning more than ~4 years. */
const MAX_RANGE_DAYS = 1400;

async function checkCalendarAccess(): Promise<boolean> {
    try {
        // Must actually *request* access, not just read: calendarsForEntityType
        // returns 0 when unauthorized without ever prompting, which looks
        // identical to an empty calendar. Requesting triggers the TCC prompt
        // once, after which the grant is recorded and this returns instantly.
        const ok = await runJxa<boolean>(
            `ObjC.import('EventKit');
             ObjC.import('Foundation');
             const store = $.EKEventStore.alloc.init;
             // Fast path: if the grant is already in place this returns a
             // non-zero count immediately, so no run loop is entered.
             let n = parseInt(String(store.calendarsForEntityType(0).count), 10) || 0;
             if (n === 0) {
               // Unauthorized reads return 0 without prompting, which is
               // indistinguishable from "no calendars" -- so ask explicitly.
               let done = false;
               if (typeof store.requestFullAccessToEventsWithCompletion === 'function') {
                 store.requestFullAccessToEventsWithCompletion(function (g, err) { done = true; });
               } else {
                 store.requestAccessToEntityTypeCompletion(0, function (g, err) { done = true; });
               }
               const deadline = $.NSDate.dateWithTimeIntervalSinceNow(45);
               while (!done && $.NSDate.date.compare(deadline) < 0) {
                 $.NSRunLoop.currentRunLoop.runModeBeforeDate($.NSDefaultRunLoopMode, $.NSDate.dateWithTimeIntervalSinceNow(0.05));
               }
               n = parseInt(String($.EKEventStore.alloc.init.calendarsForEntityType(0).count), 10) || 0;
             }
             JSON.stringify(n > 0);`,
            60000,
        );
        return ok === true;
    } catch (error) {
        console.error(
            `Cannot access Calendar: ${error instanceof Error ? error.message : String(error)}`,
        );
        return false;
    }
}

async function requestCalendarAccess(): Promise<{ hasAccess: boolean; message: string }> {
    if (await checkCalendarAccess()) {
        return { hasAccess: true, message: 'Calendar access is already granted.' };
    }
    return {
        hasAccess: false,
        message:
            'Calendar access is required but not granted. Please:\n' +
            '1. Open System Settings > Privacy & Security > Calendars\n' +
            '2. Enable access for your terminal/app\n' +
            '3. Restart the app and try again',
    };
}

function clampRange(fromDate?: string, toDate?: string): { start: Date; end: Date } {
    const start = fromDate ? new Date(fromDate) : new Date();
    let end = toDate ? new Date(toDate) : new Date(start.getTime() + 7 * 86400000);
    if (isNaN(start.getTime())) throw new Error(`Invalid fromDate: ${fromDate}`);
    if (isNaN(end.getTime())) throw new Error(`Invalid toDate: ${toDate}`);
    const maxEnd = new Date(start.getTime() + MAX_RANGE_DAYS * 86400000);
    if (end > maxEnd) end = maxEnd;
    return { start, end };
}

/** Pulls events in [start, end) straight from EventKit and returns them as JSON. */
async function fetchEvents(start: Date, end: Date, limit: number): Promise<CalendarEvent[]> {
    const script = `
ObjC.import('EventKit');
const store = $.EKEventStore.alloc.init;
const start = $.NSDate.dateWithTimeIntervalSince1970(${Math.floor(start.getTime() / 1000)});
const end   = $.NSDate.dateWithTimeIntervalSince1970(${Math.floor(end.getTime() / 1000)});
const pred = store.predicateForEventsWithStartDateEndDateCalendars(start, end, $());
const events = store.eventsMatchingPredicate(pred);
const n = parseInt(String(events.count), 10) || 0;
const iso = function (d) {
  if (!d || d.js === undefined) { try { return new Date(ObjC.unwrap(d.description)).toISOString(); } catch (e) { return null; } }
  try { return new Date(d.js).toISOString(); } catch (e) { return null; }
};
const out = [];
for (let i = 0; i < n && out.length < ${limit}; i++) {
  const e = events.objectAtIndex(i);
  let loc = null, notes = null, url = null;
  try { loc = ObjC.unwrap(e.location) || null; } catch (er) {}
  try { notes = ObjC.unwrap(e.notes) || null; } catch (er) {}
  try { url = e.URL.js ? String(ObjC.unwrap(e.URL.absoluteString)) : null; } catch (er) {}
  out.push({
    id: String(ObjC.unwrap(e.eventIdentifier) || ''),
    title: String(ObjC.unwrap(e.title) || 'Untitled Event'),
    location: loc,
    notes: notes,
    startDate: iso(e.startDate),
    endDate: iso(e.endDate),
    calendarName: String(ObjC.unwrap(e.calendar.title) || 'Unknown Calendar'),
    isAllDay: e.allDay === true,
    url: url
  });
}
JSON.stringify(out);
`;
    return (await runJxa<CalendarEvent[]>(script, CONFIG.TIMEOUT_MS)) || [];
}

async function getEvents(limit = 10, fromDate?: string, toDate?: string): Promise<CalendarEvent[]> {
    try {
        const access = await requestCalendarAccess();
        if (!access.hasAccess) throw new Error(access.message);
        const { start, end } = clampRange(fromDate, toDate);
        const events = await fetchEvents(start, end, Math.min(limit, CONFIG.MAX_EVENTS));
        // EventKit returns recurrence instances unsorted across calendars.
        return events.sort((a, b) => String(a.startDate).localeCompare(String(b.startDate)));
    } catch (error) {
        console.error(
            `Error getting events: ${error instanceof Error ? error.message : String(error)}`,
        );
        return [];
    }
}

async function searchEvents(
    searchText: string,
    limit = 10,
    fromDate?: string,
    toDate?: string,
): Promise<CalendarEvent[]> {
    try {
        const access = await requestCalendarAccess();
        if (!access.hasAccess) throw new Error(access.message);
        if (!searchText || !searchText.trim()) return [];

        // Default to a wider window for search than for a plain listing.
        const start = fromDate ? new Date(fromDate) : new Date();
        const end = toDate ? new Date(toDate) : new Date(start.getTime() + 30 * 86400000);
        const range = clampRange(start.toISOString(), end.toISOString());

        const q = searchText.toLowerCase().trim();
        const all = await fetchEvents(range.start, range.end, CONFIG.MAX_EVENTS * 10);
        return all
            .filter((e) =>
                [e.title, e.location, e.notes]
                    .filter(Boolean)
                    .some((f) => String(f).toLowerCase().includes(q)),
            )
            .sort((a, b) => String(a.startDate).localeCompare(String(b.startDate)))
            .slice(0, Math.min(limit, CONFIG.MAX_EVENTS));
    } catch (error) {
        console.error(
            `Error searching events: ${error instanceof Error ? error.message : String(error)}`,
        );
        return [];
    }
}

/** Unchanged from upstream — the AppleScript create path worked. */
async function createEvent(
    title: string,
    startDate: string,
    endDate: string,
    location?: string,
    notes?: string,
    isAllDay = false,
    calendarName?: string,
): Promise<{ success: boolean; message: string; eventId?: string }> {
    try {
        const accessResult = await requestCalendarAccess();
        if (!accessResult.hasAccess) return { success: false, message: accessResult.message };
        if (!title.trim()) return { success: false, message: 'Event title cannot be empty' };
        if (!startDate || !endDate)
            return { success: false, message: 'Start date and end date are required' };

        const start = new Date(startDate);
        const end = new Date(endDate);
        if (isNaN(start.getTime()) || isNaN(end.getTime()))
            return {
                success: false,
                message: 'Invalid date format. Please use ISO format (YYYY-MM-DDTHH:mm:ss.sssZ)',
            };
        if (end <= start) return { success: false, message: 'End date must be after start date' };

        const targetCalendar = calendarName || 'Calendar';
        const script = `
tell application "Calendar"
    set startDate to date "${start.toLocaleString()}"
    set endDate to date "${end.toLocaleString()}"
    set targetCal to null
    try
        set targetCal to calendar "${targetCalendar.replace(/"/g, '\\"')}"
    on error
        set targetCal to first calendar
    end try
    tell targetCal
        set newEvent to make new event with properties {summary:"${title.replace(/"/g, '\\"')}", start date:startDate, end date:endDate, allday event:${isAllDay}}
        ${location ? `set location of newEvent to "${location.replace(/"/g, '\\"')}"` : ''}
        ${notes ? `set description of newEvent to "${notes.replace(/"/g, '\\"')}"` : ''}
        return uid of newEvent
    end tell
end tell`;

        const eventId = (await runAppleScript(script)) as string;
        return { success: true, message: `Event "${title}" created successfully.`, eventId };
    } catch (error) {
        return {
            success: false,
            message: `Error creating event: ${error instanceof Error ? error.message : String(error)}`,
        };
    }
}

/** Resolves the id through EventKit, then opens it via Calendar's URL scheme. */
async function openEvent(eventId: string): Promise<{ success: boolean; message: string }> {
    try {
        const access = await requestCalendarAccess();
        if (!access.hasAccess) return { success: false, message: access.message };
        if (!eventId || !eventId.trim())
            return { success: false, message: 'Event ID is required' };

        const found = await runJxa<{ ok: boolean; title?: string }>(
            `ObjC.import('EventKit');
             const store = $.EKEventStore.alloc.init;
             const ev = store.eventWithIdentifier(${JSON.stringify(eventId)});
             JSON.stringify(ev.js === undefined && !ev ? {ok:false} : {ok:true, title:String(ObjC.unwrap(ev.title)||'')});`,
            30000,
        ).catch((): { ok: boolean; title?: string } => ({ ok: false }));

        if (!found || !found.ok) {
            return { success: false, message: `No event found with ID: ${eventId}` };
        }

        await runAppleScript(
            `open location "ical://ekevent/${eventId.replace(/"/g, '')}?method=show&options=more"`,
        );
        return { success: true, message: `Opened event: ${found.title || eventId}` };
    } catch (error) {
        return {
            success: false,
            message: `Error opening event: ${error instanceof Error ? error.message : String(error)}`,
        };
    }
}

const calendar = {
    searchEvents,
    openEvent,
    getEvents,
    createEvent,
    requestCalendarAccess,
};

export default calendar;

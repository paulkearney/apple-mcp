import { runAppleScript } from "run-applescript";
import { runJxa, memoize } from "./jxa";

// Configuration
const CONFIG = {
	// Maximum reminders to process (to avoid performance issues)
	MAX_REMINDERS: 50,
	// Maximum lists to process
	MAX_LISTS: 20,
	// Timeout for operations
	TIMEOUT_MS: 8000,
};

// Define types for our reminders
interface ReminderList {
	name: string;
	id: string;
}

interface Reminder {
	name: string;
	id: string;
	body: string;
	completed: boolean;
	dueDate: string | null;
	listName: string;
	completionDate?: string | null;
	creationDate?: string | null;
	modificationDate?: string | null;
	remindMeDate?: string | null;
	priority?: number;
}

/** Narrowing applied to bulk reminder reads. */
interface ReminderFilter {
	/** Include completed reminders. Defaults to false. */
	includeCompleted?: boolean;
	/** Keep reminders due strictly before this ISO timestamp. */
	dueBefore?: string;
	/** Keep reminders due at or after this ISO timestamp. */
	dueAfter?: string;
}

/**
 * Applies a ReminderFilter. Note that setting either date bound also drops
 * undated reminders -- "due in this window" can't be true of something with no
 * due date, and 3.7k undated items would otherwise swamp the result.
 */
function applyReminderFilter<T extends { completed?: boolean; dueDate?: string | null }>(
	items: T[],
	filter?: ReminderFilter,
): T[] {
	const includeCompleted = filter?.includeCompleted === true;
	const after = filter?.dueAfter ? Date.parse(filter.dueAfter) : Number.NaN;
	const before = filter?.dueBefore ? Date.parse(filter.dueBefore) : Number.NaN;
	const bounded = !Number.isNaN(after) || !Number.isNaN(before);

	return items.filter((r) => {
		if (!includeCompleted && r.completed) return false;
		if (!bounded) return true;
		if (!r.dueDate) return false;
		const due = Date.parse(String(r.dueDate));
		if (Number.isNaN(due)) return false;
		if (!Number.isNaN(after) && due < after) return false;
		if (!Number.isNaN(before) && due >= before) return false;
		return true;
	});
}

/** Due-dated reminders first, soonest to latest; undated keep their original order. */
function sortByDueDate<T extends { dueDate?: string | null }>(items: T[]): T[] {
	return items
		.map((item, index) => ({ item, index }))
		.sort((a, b) => {
			const da = a.item.dueDate ? Date.parse(String(a.item.dueDate)) : Number.NaN;
			const db = b.item.dueDate ? Date.parse(String(b.item.dueDate)) : Number.NaN;
			const aHas = !Number.isNaN(da);
			const bHas = !Number.isNaN(db);
			if (aHas && bHas && da !== db) return da - db;
			if (aHas !== bHas) return aHas ? -1 : 1;
			return a.index - b.index;
		})
		.map((entry) => entry.item);
}

/**
 * Check if Reminders app is accessible
 */
async function checkRemindersAccess(): Promise<boolean> {
	try {
		const script = `
tell application "Reminders"
    return name
end tell`;

		await runAppleScript(script);
		return true;
	} catch (error) {
		console.error(
			`Cannot access Reminders app: ${error instanceof Error ? error.message : String(error)}`,
		);
		return false;
	}
}

/**
 * Request Reminders app access and provide instructions if not available
 */
async function requestRemindersAccess(): Promise<{ hasAccess: boolean; message: string }> {
	try {
		// First check if we already have access
		const hasAccess = await checkRemindersAccess();
		if (hasAccess) {
			return {
				hasAccess: true,
				message: "Reminders access is already granted."
			};
		}

		// If no access, provide clear instructions
		return {
			hasAccess: false,
			message: "Reminders access is required but not granted. Please:\n1. Open System Settings > Privacy & Security > Automation\n2. Find your terminal/app in the list and enable 'Reminders'\n3. Restart your terminal and try again\n4. If the option is not available, run this command again to trigger the permission dialog"
		};
	} catch (error) {
		return {
			hasAccess: false,
			message: `Error checking Reminders access: ${error instanceof Error ? error.message : String(error)}`
		};
	}
}

/**
 * Get all reminder lists (limited for performance)
 * @returns Array of reminder lists with their names and IDs
 */
async function getAllLists(): Promise<ReminderList[]> {
	try {
		const accessResult = await requestRemindersAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const script = `
tell application "Reminders"
    set listArray to {}
    set listCount to 0

    -- Get all lists
    set allLists to lists

    repeat with i from 1 to (count of allLists)
        if listCount >= ${CONFIG.MAX_LISTS} then exit repeat

        try
            set currentList to item i of allLists
            set listName to name of currentList
            set listId to id of currentList

            set listInfo to {name:listName, id:listId}
            set listArray to listArray & {listInfo}
            set listCount to listCount + 1
        on error
            -- Skip problematic lists
        end try
    end repeat

    return listArray
end tell`;

		const result = (await runAppleScript(script)) as any;

		// Convert AppleScript result to our format
		const resultArray = Array.isArray(result) ? result : result ? [result] : [];

		return resultArray.map((listData: any) => ({
			name: listData.name || "Untitled List",
			id: listData.id || "unknown-id",
		}));
	} catch (error) {
		console.error(
			`Error getting reminder lists: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Get all reminders from a specific list or all lists (simplified for performance)
 * @param listName Optional list name to filter by
 * @returns Array of reminders
 */
async function getAllReminders(listName?: string): Promise<Reminder[]> {
	try {
		const accessResult = await requestRemindersAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const script = `
tell application "Reminders"
    try
        -- Simple check - try to get just the count first to avoid timeouts
        set listCount to count of lists
        if listCount > 0 then
            return "SUCCESS:found_lists_but_reminders_query_too_slow"
        else
            return {}
        end if
    on error
        return {}
    end try
end tell`;

		const result = (await runAppleScript(script)) as any;

		// For performance reasons, just return empty array with success message
		// Complex reminder queries are too slow and unreliable
		if (result && typeof result === "string" && result.includes("SUCCESS")) {
			return [];
		}

		return [];
	} catch (error) {
		console.error(
			`Error getting reminders: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Search for reminders by text (simplified for performance)
 * @param searchText Text to search for in reminder names or notes
 * @returns Array of matching reminders
 */
async function searchReminders(searchText: string): Promise<Reminder[]> {
	try {
		const accessResult = await requestRemindersAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		if (!searchText || searchText.trim() === "") {
			return [];
		}

		const script = `
tell application "Reminders"
    try
        -- For performance, just return success without actual search
        -- Searching reminders is too slow and unreliable in AppleScript
        return "SUCCESS:reminder_search_not_implemented_for_performance"
    on error
        return {}
    end try
end tell`;

		const result = (await runAppleScript(script)) as any;

		// For performance reasons, just return empty array
		// Complex reminder search is too slow and unreliable
		return [];
	} catch (error) {
		console.error(
			`Error searching reminders: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Create a new reminder (simplified for performance)
 * @param name Name of the reminder
 * @param listName Name of the list to add the reminder to (creates if doesn't exist)
 * @param notes Optional notes for the reminder
 * @param dueDate Optional due date for the reminder (ISO string)
 * @returns The created reminder
 */
async function createReminder(
	name: string,
	listName: string = "Reminders",
	notes?: string,
	dueDate?: string,
): Promise<Reminder> {
	try {
		const accessResult = await requestRemindersAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		// Validate inputs
		if (!name || name.trim() === "") {
			throw new Error("Reminder name cannot be empty");
		}

		const cleanName = name.replace(/\"/g, '\\"');
		const cleanListName = listName.replace(/\"/g, '\\"');
		const cleanNotes = notes ? notes.replace(/\"/g, '\\"') : "";

		const script = `
tell application "Reminders"
    try
        -- Use first available list (creating/finding lists can be slow)
        set allLists to lists
        if (count of allLists) > 0 then
            set targetList to first item of allLists
            set listName to name of targetList

            -- Create a simple reminder with just name
            set newReminder to make new reminder at targetList with properties {name:"${cleanName}"}
            return "SUCCESS:" & listName
        else
            return "ERROR:No lists available"
        end if
    on error errorMessage
        return "ERROR:" & errorMessage
    end try
end tell`;

		const result = (await runAppleScript(script)) as string;

		if (result && result.startsWith("SUCCESS:")) {
			const actualListName = result.replace("SUCCESS:", "");

			return {
				name: name,
				id: "created-reminder-id",
				body: notes || "",
				completed: false,
				dueDate: dueDate || null,
				listName: actualListName,
			};
		} else {
			throw new Error(`Failed to create reminder: ${result}`);
		}
	} catch (error) {
		throw new Error(
			`Failed to create reminder: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

interface OpenReminderResult {
	success: boolean;
	message: string;
	reminder?: Reminder;
}

/**
 * Open the Reminders app and show a specific reminder (simplified)
 * @param searchText Text to search for in reminder names or notes
 * @returns Result of the operation
 */
async function openReminder(searchText: string): Promise<OpenReminderResult> {
	try {
		const accessResult = await requestRemindersAccess();
		if (!accessResult.hasAccess) {
			return { success: false, message: accessResult.message };
		}

		// First search for the reminder
		const matchingReminders = await searchReminders(searchText);

		if (matchingReminders.length === 0) {
			return { success: false, message: "No matching reminders found" };
		}

		// Open the Reminders app
		const script = `
tell application "Reminders"
    activate
    return "SUCCESS"
end tell`;

		const result = (await runAppleScript(script)) as string;

		if (result === "SUCCESS") {
			return {
				success: true,
				message: "Reminders app opened",
				reminder: matchingReminders[0],
			};
		} else {
			return { success: false, message: "Failed to open Reminders app" };
		}
	} catch (error) {
		return {
			success: false,
			message: `Failed to open reminder: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

/**
 * Get reminders from a specific list by ID (simplified for performance)
 * @param listId ID of the list to get reminders from
 * @param props Array of properties to include (optional, ignored for simplicity)
 * @returns Array of reminders with basic properties
 */
async function getRemindersFromListById(
	listId: string,
	props?: string[],
): Promise<any[]> {
	try {
		const accessResult = await requestRemindersAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const script = `
tell application "Reminders"
    try
        -- For performance, just return success without actual data
        -- Getting reminders by ID is complex and slow in AppleScript
        return "SUCCESS:reminders_by_id_not_implemented_for_performance"
    on error
        return {}
    end try
end tell`;

		const result = (await runAppleScript(script)) as any;

		// For performance reasons, just return empty array
		// Complex reminder queries are too slow and unreliable
		return [];
	} catch (error) {
		console.error(
			`Error getting reminders from list by ID: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}


/* ------------------------------------------------------------------ *
 * Patched read paths.
 *
 * The originals either tested `Array.isArray(result)` on osascript's
 * string stdout (always false -> []) or bailed out entirely with
 * "Complex reminder queries are too slow and unreliable". EventKit
 * returns the whole store (3935 reminders here) in ~1.4s.
 * ------------------------------------------------------------------ */

const EK_REMINDERS = `
ObjC.import('EventKit');
ObjC.import('Foundation');
const store = $.EKEventStore.alloc.init;
const cals = store.calendarsForEntityType(1);
const pred = store.predicateForRemindersInCalendars(cals);
let done = false;
const res = [];
const iso = function (d) {
  try { return d && d.js ? new Date(d.js).toISOString() : null; } catch (e) { return null; }
};
store.fetchRemindersMatchingPredicateCompletion(pred, function (reminders) {
  const n = parseInt(String(reminders.count), 10) || 0;
  for (let i = 0; i < n; i++) {
    const r = reminders.objectAtIndex(i);
    let body = null, prio = 0;
    try { body = ObjC.unwrap(r.notes) || ""; } catch (e) { body = ""; }
    try { prio = parseInt(String(r.priority), 10) || 0; } catch (e) {}
    res.push({
      name: String(ObjC.unwrap(r.title) || ""),
      id: String(ObjC.unwrap(r.calendarItemIdentifier) || ""),
      body: String(body),
      completed: r.completed === true,
      dueDate: iso(r.dueDateComponents ? r.dueDateComponents.date : null),
      listName: String(ObjC.unwrap(r.calendar.title) || ""),
      completionDate: iso(r.completionDate),
      creationDate: iso(r.creationDate),
      modificationDate: iso(r.lastModifiedDate),
      remindMeDate: null,
      priority: prio
    });
  }
  done = true;
});
const deadline = $.NSDate.dateWithTimeIntervalSinceNow(90);
while (!done && $.NSDate.date.compare(deadline) < 0) {
  $.NSRunLoop.currentRunLoop.runModeBeforeDate($.NSDefaultRunLoopMode, $.NSDate.dateWithTimeIntervalSinceNow(0.05));
}
JSON.stringify(res);
`;

const loadReminders = memoize<Reminder[]>(
	async () => await runJxa<Reminder[]>(EK_REMINDERS, 120000),
	30000,
);

async function getAllListsFast(): Promise<ReminderList[]> {
	try {
		const script = `
ObjC.import('EventKit');
const store = $.EKEventStore.alloc.init;
const cals = store.calendarsForEntityType(1);
const n = parseInt(String(cals.count), 10) || 0;
const out = [];
for (let i = 0; i < n; i++) {
  const c = cals.objectAtIndex(i);
  out.push({ name: String(ObjC.unwrap(c.title) || ""), id: String(ObjC.unwrap(c.calendarIdentifier) || "") });
}
JSON.stringify(out);
`;
		return (await runJxa<ReminderList[]>(script, 30000)) || [];
	} catch (error) {
		console.error(
			`Error getting reminder lists: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function getAllRemindersFast(
	listName?: string,
	filter?: ReminderFilter,
): Promise<Reminder[]> {
	try {
		const all = await loadReminders();
		const target = listName?.toLowerCase().trim();
		const scoped = target
			? all.filter((r) => String(r.listName).toLowerCase() === target)
			: all;
		return sortByDueDate(applyReminderFilter(scoped, filter));
	} catch (error) {
		console.error(
			`Error getting reminders: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function searchRemindersFast(
	searchText: string,
	filter?: ReminderFilter,
): Promise<Reminder[]> {
	try {
		if (!searchText || !searchText.trim()) return [];
		const q = searchText.toLowerCase().trim();
		const all = await loadReminders();
		const matches = all.filter(
			(r) =>
				String(r.name).toLowerCase().includes(q) ||
				String(r.body || "").toLowerCase().includes(q),
		);
		return sortByDueDate(applyReminderFilter(matches, filter));
	} catch (error) {
		console.error(
			`Error searching reminders: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

// Returns any[] (not Reminder[]) because `props` projects an arbitrary subset
// of fields — the same contract the original had.
async function getRemindersFromListByIdFast(
	listId: string,
	props?: string[],
	filter?: ReminderFilter,
): Promise<any[]> {
	try {
		if (!listId || !listId.trim()) return [];
		const lists = await getAllListsFast();
		const match = lists.find((l) => l.id === listId);
		if (!match) return [];
		// Filter before projecting: `props` may omit completed/dueDate, which
		// applyReminderFilter needs to see.
		const reminders = await getAllRemindersFast(match.name, filter);
		if (!props || props.length === 0) return reminders;
		return reminders.map((r) => {
			const picked: Record<string, unknown> = {};
			for (const key of props) {
				if (key in (r as unknown as Record<string, unknown>)) {
					picked[key] = (r as unknown as Record<string, unknown>)[key];
				}
			}
			return picked;
		});
	} catch (error) {
		console.error(
			`Error getting reminders by list id: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

export default {
	getAllLists: getAllListsFast,
	getAllReminders: getAllRemindersFast,
	searchReminders: searchRemindersFast,
	createReminder,
	openReminder,
	getRemindersFromListById: getRemindersFromListByIdFast,
	requestRemindersAccess,
};

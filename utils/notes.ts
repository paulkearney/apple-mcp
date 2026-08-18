import { runAppleScript } from "run-applescript";
import { runJxa, memoize } from "./jxa";

// Configuration
const CONFIG = {
	// Maximum notes to process (to avoid performance issues)
	MAX_NOTES: 50,
	// Maximum content length for previews
	MAX_CONTENT_PREVIEW: 200,
	// Timeout for operations
	TIMEOUT_MS: 8000,
};

type Note = {
	name: string;
	content: string;
	creationDate?: Date;
	modificationDate?: Date;
};

type CreateNoteResult = {
	success: boolean;
	note?: Note;
	message?: string;
	folderName?: string;
	usedDefaultFolder?: boolean;
};

/**
 * Check if Notes app is accessible
 */
async function checkNotesAccess(): Promise<boolean> {
	try {
		const script = `
tell application "Notes"
    return name
end tell`;

		await runAppleScript(script);
		return true;
	} catch (error) {
		console.error(
			`Cannot access Notes app: ${error instanceof Error ? error.message : String(error)}`,
		);
		return false;
	}
}

/**
 * Request Notes app access and provide instructions if not available
 */
async function requestNotesAccess(): Promise<{ hasAccess: boolean; message: string }> {
	try {
		// First check if we already have access
		const hasAccess = await checkNotesAccess();
		if (hasAccess) {
			return {
				hasAccess: true,
				message: "Notes access is already granted."
			};
		}

		// If no access, provide clear instructions
		return {
			hasAccess: false,
			message: "Notes access is required but not granted. Please:\n1. Open System Settings > Privacy & Security > Automation\n2. Find your terminal/app in the list and enable 'Notes'\n3. Restart your terminal and try again\n4. If the option is not available, run this command again to trigger the permission dialog"
		};
	} catch (error) {
		return {
			hasAccess: false,
			message: `Error checking Notes access: ${error instanceof Error ? error.message : String(error)}`
		};
	}
}

/**
 * Get all notes from Notes app (limited for performance)
 */
async function getAllNotes(): Promise<Note[]> {
	try {
		const accessResult = await requestNotesAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		const script = `
tell application "Notes"
    set notesList to {}
    set noteCount to 0

    -- Get all notes from all folders
    set allNotes to notes

    repeat with i from 1 to (count of allNotes)
        if noteCount >= ${CONFIG.MAX_NOTES} then exit repeat

        try
            set currentNote to item i of allNotes
            set noteName to name of currentNote
            set noteContent to plaintext of currentNote

            -- Limit content for preview
            if (length of noteContent) > ${CONFIG.MAX_CONTENT_PREVIEW} then
                set noteContent to (characters 1 thru ${CONFIG.MAX_CONTENT_PREVIEW} of noteContent) as string
                set noteContent to noteContent & "..."
            end if

            set noteInfo to {name:noteName, content:noteContent}
            set notesList to notesList & {noteInfo}
            set noteCount to noteCount + 1
        on error
            -- Skip problematic notes
        end try
    end repeat

    return notesList
end tell`;

		const result = (await runAppleScript(script)) as any;

		// Convert AppleScript result to our format
		const resultArray = Array.isArray(result) ? result : result ? [result] : [];

		return resultArray.map((noteData: any) => ({
			name: noteData.name || "Untitled Note",
			content: noteData.content || "",
			creationDate: undefined,
			modificationDate: undefined,
		}));
	} catch (error) {
		console.error(
			`Error getting all notes: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Find notes by search text
 */
async function findNote(searchText: string): Promise<Note[]> {
	try {
		const accessResult = await requestNotesAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}

		if (!searchText || searchText.trim() === "") {
			return [];
		}

		const searchTerm = searchText.toLowerCase();

		const script = `
tell application "Notes"
    set matchedNotes to {}
    set noteCount to 0
    set searchTerm to "${searchTerm}"

    -- Get all notes and search through them
    set allNotes to notes

    repeat with i from 1 to (count of allNotes)
        if noteCount >= ${CONFIG.MAX_NOTES} then exit repeat

        try
            set currentNote to item i of allNotes
            set noteName to name of currentNote
            set noteContent to plaintext of currentNote

            -- Simple case-insensitive search in name and content
            if (noteName contains searchTerm) or (noteContent contains searchTerm) then
                -- Limit content for preview
                if (length of noteContent) > ${CONFIG.MAX_CONTENT_PREVIEW} then
                    set noteContent to (characters 1 thru ${CONFIG.MAX_CONTENT_PREVIEW} of noteContent) as string
                    set noteContent to noteContent & "..."
                end if

                set noteInfo to {name:noteName, content:noteContent}
                set matchedNotes to matchedNotes & {noteInfo}
                set noteCount to noteCount + 1
            end if
        on error
            -- Skip problematic notes
        end try
    end repeat

    return matchedNotes
end tell`;

		const result = (await runAppleScript(script)) as any;

		// Convert AppleScript result to our format
		const resultArray = Array.isArray(result) ? result : result ? [result] : [];

		return resultArray.map((noteData: any) => ({
			name: noteData.name || "Untitled Note",
			content: noteData.content || "",
			creationDate: undefined,
			modificationDate: undefined,
		}));
	} catch (error) {
		console.error(
			`Error finding notes: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/**
 * Create a new note
 */
async function createNote(
	title: string,
	body: string,
	folderName: string = "Claude",
): Promise<CreateNoteResult> {
	try {
		const accessResult = await requestNotesAccess();
		if (!accessResult.hasAccess) {
			return {
				success: false,
				message: accessResult.message,
			};
		}

		// Validate inputs
		if (!title || title.trim() === "") {
			return {
				success: false,
				message: "Note title cannot be empty",
			};
		}

		// Keep the body as-is to preserve original formatting
		// Notes.app handles markdown and formatting natively
		const formattedBody = body.trim();

		// Use file-based approach for complex content to avoid AppleScript string issues
		const tmpFile = `/tmp/note-content-${Date.now()}.txt`;
		const fs = require("fs");

		// Write content to temporary file to avoid AppleScript escaping issues
		fs.writeFileSync(tmpFile, formattedBody, "utf8");

		const script = `
tell application "Notes"
    set targetFolder to null
    set folderFound to false
    set actualFolderName to "${folderName}"

    -- Try to find the specified folder
    try
        set allFolders to folders
        repeat with currentFolder in allFolders
            if name of currentFolder is "${folderName}" then
                set targetFolder to currentFolder
                set folderFound to true
                exit repeat
            end if
        end repeat
    on error
        -- Folders might not be accessible
    end try

    -- If folder not found and it's a test folder, try to create it
    if not folderFound and ("${folderName}" is "Claude" or "${folderName}" is "Test-Claude") then
        try
            make new folder with properties {name:"${folderName}"}
            -- Try to find it again
            set allFolders to folders
            repeat with currentFolder in allFolders
                if name of currentFolder is "${folderName}" then
                    set targetFolder to currentFolder
                    set folderFound to true
                    set actualFolderName to "${folderName}"
                    exit repeat
                end if
            end repeat
        on error
            -- Folder creation failed, use default
            set actualFolderName to "Notes"
        end try
    end if

    -- Read content from file to preserve formatting
    set noteContent to read file POSIX file "${tmpFile}" as «class utf8»

    -- Create the note with proper content
    if folderFound and targetFolder is not null then
        -- Create note in specified folder
        make new note at targetFolder with properties {name:"${title.replace(/"/g, '\\"')}", body:noteContent}
        return "SUCCESS:" & actualFolderName & ":false"
    else
        -- Create note in default location
        make new note with properties {name:"${title.replace(/"/g, '\\"')}", body:noteContent}
        return "SUCCESS:Notes:true"
    end if
end tell`;

		const result = (await runAppleScript(script)) as string;

		// Clean up temporary file
		try {
			fs.unlinkSync(tmpFile);
		} catch (e) {
			// Ignore cleanup errors
		}

		// Parse the result string format: "SUCCESS:folderName:usedDefault"
		if (result && typeof result === "string" && result.startsWith("SUCCESS:")) {
			const parts = result.split(":");
			const folderName = parts[1] || "Notes";
			const usedDefaultFolder = parts[2] === "true";

			return {
				success: true,
				note: {
					name: title,
					content: formattedBody,
				},
				folderName: folderName,
				usedDefaultFolder: usedDefaultFolder,
			};
		} else {
			return {
				success: false,
				message: `Failed to create note: ${result || "No result from AppleScript"}`,
			};
		}
	} catch (error) {
		return {
			success: false,
			message: `Failed to create note: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

/**
 * Get notes from a specific folder
 */
async function getNotesFromFolder(
	folderName: string,
): Promise<{ success: boolean; notes?: Note[]; message?: string }> {
	try {
		const accessResult = await requestNotesAccess();
		if (!accessResult.hasAccess) {
			return {
				success: false,
				message: accessResult.message,
			};
		}

		const script = `
tell application "Notes"
    set notesList to {}
    set noteCount to 0
    set folderFound to false

    -- Try to find the specified folder
    try
        set allFolders to folders
        repeat with currentFolder in allFolders
            if name of currentFolder is "${folderName}" then
                set folderFound to true

                -- Get notes from this folder
                set folderNotes to notes of currentFolder

                repeat with i from 1 to (count of folderNotes)
                    if noteCount >= ${CONFIG.MAX_NOTES} then exit repeat

                    try
                        set currentNote to item i of folderNotes
                        set noteName to name of currentNote
                        set noteContent to plaintext of currentNote

                        -- Limit content for preview
                        if (length of noteContent) > ${CONFIG.MAX_CONTENT_PREVIEW} then
                            set noteContent to (characters 1 thru ${CONFIG.MAX_CONTENT_PREVIEW} of noteContent) as string
                            set noteContent to noteContent & "..."
                        end if

                        set noteInfo to {name:noteName, content:noteContent}
                        set notesList to notesList & {noteInfo}
                        set noteCount to noteCount + 1
                    on error
                        -- Skip problematic notes
                    end try
                end repeat

                exit repeat
            end if
        end repeat
    on error
        -- Handle folder access errors
    end try

    if not folderFound then
        return "ERROR:Folder not found"
    end if

    return "SUCCESS:" & (count of notesList)
end tell`;

		const result = (await runAppleScript(script)) as any;

		// Simple success/failure check based on string result
		if (result && typeof result === "string") {
			if (result.startsWith("ERROR:")) {
				return {
					success: false,
					message: result.replace("ERROR:", ""),
				};
			} else if (result.startsWith("SUCCESS:")) {
				// For now, just return success - the actual notes are complex to parse from AppleScript
				return {
					success: true,
					notes: [], // Return empty array for simplicity
				};
			}
		}

		// If we get here, assume folder was found but no notes
		return {
			success: true,
			notes: [],
		};
	} catch (error) {
		return {
			success: false,
			message: `Failed to get notes from folder: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

/**
 * Get recent notes from a specific folder
 */
async function getRecentNotesFromFolder(
	folderName: string,
	limit: number = 5,
): Promise<{ success: boolean; notes?: Note[]; message?: string }> {
	try {
		// For simplicity, just get notes from folder (they're typically in recent order)
		const result = await getNotesFromFolder(folderName);

		if (result.success && result.notes) {
			return {
				success: true,
				notes: result.notes.slice(0, Math.min(limit, result.notes.length)),
			};
		}

		return result;
	} catch (error) {
		return {
			success: false,
			message: `Failed to get recent notes from folder: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

/**
 * Get notes by date range (simplified implementation)
 */
async function getNotesByDateRange(
	folderName: string,
	fromDate?: string,
	toDate?: string,
	limit: number = 20,
): Promise<{ success: boolean; notes?: Note[]; message?: string }> {
	try {
		// For simplicity, just return notes from folder
		// Date filtering is complex and unreliable in AppleScript
		const result = await getNotesFromFolder(folderName);

		if (result.success && result.notes) {
			return {
				success: true,
				notes: result.notes.slice(0, Math.min(limit, result.notes.length)),
			};
		}

		return result;
	} catch (error) {
		return {
			success: false,
			message: `Failed to get notes by date range: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}


/* ------------------------------------------------------------------ *
 * Patched read paths.
 *
 * The originals ran AppleScript and then tested `Array.isArray(result)`
 * on osascript's string stdout, which is never an array — so every list
 * silently collapsed to []. These fetch all notes in one bulk JXA call
 * (~0.5s for 237 notes) and filter in JS.
 * ------------------------------------------------------------------ */

type RawNote = {
	name: string;
	content: string;
	folder: string;
	creationDate: string | null;
	modificationDate: string | null;
};

const loadAllNotes = memoize<RawNote[]>(async () => {
	const script = `
// Bulk "notes.container.name()" returns nulls, so walk folders to keep the
// note -> folder association while still using one bulk fetch per folder.
const N = Application("Notes");
const folders = N.folders;
const fnames = folders.name();
const out = [];
for (let i = 0; i < fnames.length; i++) {
  const fld = folders[i];
  let names = [], bodies = [], created = [], modified = [];
  try { names = fld.notes.name(); } catch (e) { continue; }
  try { bodies = fld.notes.plaintext(); } catch (e) {}
  try { created = fld.notes.creationDate(); } catch (e) {}
  try { modified = fld.notes.modificationDate(); } catch (e) {}
  for (let j = 0; j < names.length; j++) {
    out.push({
      name: String(names[j] || "Untitled"),
      content: String(bodies[j] || ""),
      folder: String(fnames[i] || ""),
      creationDate: created[j] ? new Date(created[j]).toISOString() : null,
      modificationDate: modified[j] ? new Date(modified[j]).toISOString() : null
    });
  }
}
JSON.stringify(out);
`;
	return await runJxa<RawNote[]>(script, 120000);
}, 60000);

function toNote(r: RawNote): Note {
	return {
		name: r.name,
		content: r.content,
		creationDate: r.creationDate ? new Date(r.creationDate) : undefined,
		modificationDate: r.modificationDate ? new Date(r.modificationDate) : undefined,
	};
}

async function getAllNotesFast(): Promise<Note[]> {
	try {
		const access = await requestNotesAccess();
		if (!access.hasAccess) throw new Error(access.message);
		return (await loadAllNotes()).map(toNote);
	} catch (error) {
		console.error(
			`Error getting notes: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function findNoteFast(searchText: string): Promise<Note[]> {
	try {
		if (!searchText || !searchText.trim()) return [];
		const q = searchText.toLowerCase().trim();
		const all = await loadAllNotes();
		// Title matches first, then body matches.
		const byTitle = all.filter((n) => n.name.toLowerCase().includes(q));
		const byBody = all.filter(
			(n) => !n.name.toLowerCase().includes(q) && n.content.toLowerCase().includes(q),
		);
		return [...byTitle, ...byBody].map(toNote);
	} catch (error) {
		console.error(
			`Error finding note: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function getNotesFromFolderFast(
	folderName: string,
): Promise<{ success: boolean; notes?: Note[]; message?: string }> {
	try {
		if (!folderName || !folderName.trim()) {
			return { success: false, message: "Folder name is required" };
		}
		const target = folderName.toLowerCase().trim();
		const all = await loadAllNotes();
		const hits = all.filter((n) => n.folder.toLowerCase() === target);
		if (hits.length === 0) {
			const known = [...new Set(all.map((n) => n.folder).filter(Boolean))];
			return {
				success: false,
				message: `No notes found in folder "${folderName}". Available folders: ${known.join(", ")}`,
			};
		}
		return { success: true, notes: hits.map(toNote) };
	} catch (error) {
		return {
			success: false,
			message: `Error getting notes from folder: ${error instanceof Error ? error.message : String(error)}`,
		};
	}
}

async function getRecentNotesFromFolderFast(
	folderName: string,
	limit: number = 5,
): Promise<{ success: boolean; notes?: Note[]; message?: string }> {
	const result = await getNotesFromFolderFast(folderName);
	if (!result.success || !result.notes) return result;
	const sorted = [...result.notes].sort(
		(a, b) => (b.modificationDate?.getTime() ?? 0) - (a.modificationDate?.getTime() ?? 0),
	);
	return { success: true, notes: sorted.slice(0, limit) };
}

async function getNotesByDateRangeFast(
	folderName: string,
	fromDate?: string,
	toDate?: string,
	limit: number = 20,
): Promise<{ success: boolean; notes?: Note[]; message?: string }> {
	const result = await getNotesFromFolderFast(folderName);
	if (!result.success || !result.notes) return result;
	const from = fromDate ? new Date(fromDate).getTime() : Number.NEGATIVE_INFINITY;
	const to = toDate ? new Date(toDate).getTime() : Number.POSITIVE_INFINITY;
	const inRange = result.notes.filter((n) => {
		const t = n.modificationDate?.getTime() ?? n.creationDate?.getTime();
		return t !== undefined && t >= from && t <= to;
	});
	return { success: true, notes: inRange.slice(0, limit) };
}

export default {
	getAllNotes: getAllNotesFast,
	findNote: findNoteFast,
	createNote,
	getNotesFromFolder: getNotesFromFolderFast,
	getRecentNotesFromFolder: getRecentNotesFromFolderFast,
	getNotesByDateRange: getNotesByDateRangeFast,
	requestNotesAccess,
};

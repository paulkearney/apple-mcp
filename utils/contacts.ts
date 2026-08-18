import { runJxa, memoize } from "./jxa";

/**
 * Patched contacts backend.
 *
 * Upstream walked every contact with AppleScript, spawning a *shell process per
 * contact* (`do shell script "echo ... | tr ..."`) just to lowercase a name, and
 * appended to an AppleScript list (O(n^2) copying). On a 1771-contact address
 * book that never completed. This fetches every name/phone/email in a single
 * bulk JXA round-trip (~5s) and does all matching in JavaScript.
 */

export type Contact = { name: string; phones: string[]; emails: string[] };

const CACHE_TTL_MS = 60_000;

async function checkContactsAccess(): Promise<boolean> {
	try {
		await runJxa<number>(
			`JSON.stringify(Application("Contacts").people.length)`,
			15000,
		);
		return true;
	} catch (error) {
		console.error(
			`Cannot access Contacts app: ${error instanceof Error ? error.message : String(error)}`,
		);
		return false;
	}
}

async function requestContactsAccess(): Promise<{ hasAccess: boolean; message: string }> {
	if (await checkContactsAccess()) {
		return { hasAccess: true, message: "Contacts access is already granted." };
	}
	return {
		hasAccess: false,
		message:
			"Contacts access is required but not granted. Please:\n" +
			"1. Open System Settings > Privacy & Security > Automation\n" +
			"2. Find your terminal/app in the list and enable 'Contacts'\n" +
			"3. Restart the app and try again",
	};
}

/** One Apple Event per property across the whole address book, not per contact. */
const loadContacts = memoize<Contact[]>(async () => {
	const access = await requestContactsAccess();
	if (!access.hasAccess) throw new Error(access.message);

	const script = `
const C = Application("Contacts");
const names = C.people.name();
let phones = [], emails = [];
try { phones = C.people.phones.value(); } catch (e) {}
try { emails = C.people.emails.value(); } catch (e) {}
const out = [];
for (let i = 0; i < names.length; i++) {
  const nm = names[i];
  if (!nm) continue;
  const p = (phones[i] || []).filter(function (x) { return x; });
  const e = (emails[i] || []).filter(function (x) { return x; });
  out.push({ name: String(nm), phones: p.map(String), emails: e.map(String) });
}
JSON.stringify(out);
`;
	return await runJxa<Contact[]>(script, 120000);
}, CACHE_TTL_MS);

async function getAllNumbers(): Promise<{ [key: string]: string[] }> {
	try {
		const contacts = await loadContacts();
		const out: { [key: string]: string[] } = {};
		for (const c of contacts) {
			if (c.phones.length === 0) continue;
			// Merge rather than overwrite: duplicate display names are common.
			out[c.name] = out[c.name] ? [...new Set([...out[c.name], ...c.phones])] : c.phones;
		}
		return out;
	} catch (error) {
		console.error(
			`Error getting all contacts: ${error instanceof Error ? error.message : String(error)}`,
		);
		return {};
	}
}

async function getAllContacts(): Promise<Contact[]> {
	try {
		return await loadContacts();
	} catch (error) {
		console.error(
			`Error getting all contacts: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/** Strip emoji/symbols and collapse whitespace so "Kerry ❤️" matches "kerry". */
function cleanName(name: string): string {
	return name
		.toLowerCase()
		.replace(
			/[\u{1F600}-\u{1F64F}]|[\u{1F300}-\u{1F5FF}]|[\u{1F680}-\u{1F6FF}]|[\u{1F1E0}-\u{1F1FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{FE00}-\u{FE0F}]/gu,
			"",
		)
		.replace(/\s+/g, " ")
		.trim();
}

/** Ordered best-match-first strategies; first strategy with any hit wins. */
function rankMatches(contacts: Contact[], query: string): Contact[] {
	const q = cleanName(query);
	if (!q) return [];

	const strategies: Array<(n: string) => boolean> = [
		(n) => n === q,
		(n) => n.startsWith(q),
		(n) => n.split(" ")[0] === q,
		(n) => n.split(" ").slice(-1)[0] === q,
		(n) => n.includes(q),
		(n) => n.split(" ").some((w) => w.startsWith(q)),
		(n) => q.includes(n) && n.length > 2,
	];

	for (const match of strategies) {
		const hits = contacts.filter((c) => match(cleanName(c.name)));
		if (hits.length > 0) return hits;
	}
	return [];
}

async function findNumber(name: string): Promise<string[]> {
	try {
		if (!name || name.trim() === "") return [];
		const contacts = await loadContacts();
		const hits = rankMatches(
			contacts.filter((c) => c.phones.length > 0),
			name,
		);
		if (hits.length === 0) return [];
		return hits[0].phones.filter((p) => p && p.trim() !== "");
	} catch (error) {
		console.error(
			`Error finding contact: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/** Full match list, so a tool can show every "Kerry" instead of guessing one. */
async function searchContacts(name: string): Promise<Contact[]> {
	try {
		if (!name || name.trim() === "") return [];
		return rankMatches(await loadContacts(), name);
	} catch (error) {
		console.error(
			`Error searching contacts: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

function normalizePhone(p: string): string {
	return p.replace(/[^0-9+]/g, "");
}

async function findContactByPhone(phoneNumber: string): Promise<string | null> {
	try {
		if (!phoneNumber || phoneNumber.trim() === "") return null;
		const search = normalizePhone(phoneNumber);
		// Compare on the last 10 digits: handles +1 / country-code mismatches.
		const tail = (s: string) => s.replace(/\D/g, "").slice(-10);
		const target = tail(search);
		if (!target) return null;

		const contacts = await loadContacts();
		for (const c of contacts) {
			if (c.phones.some((p) => tail(p) === target)) return c.name;
		}
		return null;
	} catch (error) {
		console.error(
			`Error finding contact by phone: ${error instanceof Error ? error.message : String(error)}`,
		);
		return null;
	}
}

export default {
	getAllNumbers,
	getAllContacts,
	findNumber,
	searchContacts,
	findContactByPhone,
	requestContactsAccess,
};

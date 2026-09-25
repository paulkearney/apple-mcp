import fs from "node:fs";
import { runAppleScript } from "run-applescript";
import { runJxa } from "./jxa";

/**
 * Patched mail backend.
 *
 * Upstream was largely non-functional:
 *   - getUnreadMails/searchMails ran a script and then `return []` unconditionally
 *     (with a comment saying real parsing was "complex").
 *   - getAccounts' AppleScript returned the literal string {"Default Account"},
 *     and the caller tested Array.isArray() on osascript's string stdout, so it
 *     always yielded [].
 *
 * This version emits JSON from JXA and parses it. It also avoids `whose(...)`
 * clauses, which Mail evaluates message-by-message (37s on a 68k-message inbox);
 * bulk property access pulls the same data in ~1s, and filtering happens in JS.
 */

const CONFIG = {
	MAX_EMAILS: 20,
	MAX_CONTENT_PREVIEW: 300,
	// Message bodies cost ~5s each over Exchange, so a 10-message listing can
	// blow past the MCP client's own request timeout. Bound the body fetching
	// and mark whatever we had to skip, rather than returning nothing at all.
	CONTENT_BUDGET_SECS: 15,
	TIMEOUT_MS: 120000,
};

interface EmailMessage {
	subject: string;
	sender: string;
	dateSent: string;
	content: string;
	isRead: boolean;
	mailbox: string;
}

/** Escapes a string for embedding in a JXA source literal. */
function jsLit(s: string): string {
	return JSON.stringify(String(s));
}

/** Escapes a string for embedding in a double-quoted AppleScript literal. */
function asLit(s: string): string {
	return String(s).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

/**
 * JXA helpers shared by the mailbox-addressed operations (move, createMailbox).
 *
 * Mail flattens `account.mailboxes` across every nesting depth (a folder
 * "Clients/Regus/Old" shows up as three entries named "Clients", "Regus" and
 * "Old"), while `mailbox.mailboxes` lists only direct children. Paths are
 * therefore resolved by matching the first segment at the top level and each
 * later segment among the previous mailbox's children. Matching is
 * case-insensitive so "inbox" finds "INBOX".
 */
const MAILBOX_HELPERS = `
function findAccount(M, name) {
  const accts = M.accounts;
  const names = accts.name();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i]) === name) return accts[i];
  }
  throw new Error("No such account: " + name);
}
// Reading a mailbox-only property off an account container throws, which is
// the only reliable way to tell "top-level mailbox" from "nested mailbox".
function isMailbox(x) {
  try { x.unreadCount(); return true; } catch (e) { return false; }
}
function splitPath(path) {
  return String(path).split("/").map(function (s) { return s.trim(); }).filter(Boolean);
}
function lower(s) { return String(s).toLowerCase(); }
// Returns the mailbox at \`path\` inside \`acct\`, or null if absent. The
// candidates are the mailboxes named like the last segment; each one's
// container chain is walked upwards against the remaining segments and must
// end at the account. Walking up (not down from the account) matters because
// Mail does not enumerate an IMAP parent folder that exists only as a path
// prefix, even though a child's container() still returns it. With
// \`lenient\`, a path that matches the tail of exactly one deeper path (e.g.
// "Regus/Old" for "Clients/Regus/Old") resolves as well.
function resolveMailbox(acct, path, lenient) {
  const segs = splitPath(path).map(lower);
  if (segs.length === 0) throw new Error("Mailbox path is empty");
  const all = acct.mailboxes;
  const names = all.name();
  const leaf = segs[segs.length - 1];
  const exact = [];
  const partial = [];
  for (let i = 0; i < names.length; i++) {
    if (lower(names[i]) !== leaf) continue;
    let cur = all[i];
    let k = segs.length - 1;
    let matched = true;
    while (k > 0) {
      let parent = null;
      try { parent = cur.container(); } catch (e) { parent = null; }
      if (!parent || !isMailbox(parent) || lower(parent.name()) !== segs[k - 1]) {
        matched = false;
        break;
      }
      cur = parent;
      k--;
    }
    if (!matched) continue;
    let top = null;
    try { top = cur.container(); } catch (e) { top = null; }
    if (top && isMailbox(top)) partial.push(i); else exact.push(i);
  }
  if (exact.length > 0) return all[exact[0]];
  if (lenient && partial.length === 1) return all[partial[0]];
  if (lenient && partial.length > 1) {
    throw new Error("Mailbox path '" + path + "' is ambiguous in account "
      + String(acct.name()) + "; give its full path from the top level");
  }
  return null;
}
`;

async function checkMailAccess(): Promise<boolean> {
	try {
		await runJxa<string[]>(`JSON.stringify(Application("Mail").accounts.name())`, 20000);
		return true;
	} catch (error) {
		console.error(
			`Cannot access Mail app: ${error instanceof Error ? error.message : String(error)}`,
		);
		return false;
	}
}

async function requestMailAccess(): Promise<{ hasAccess: boolean; message: string }> {
	if (await checkMailAccess()) {
		return { hasAccess: true, message: "Mail access is already granted." };
	}
	return {
		hasAccess: false,
		message:
			"Mail access is required but not granted. Please:\n" +
			"1. Open System Settings > Privacy & Security > Automation\n" +
			"2. Enable 'Mail' for your terminal/app\n" +
			"3. Make sure Mail is running and has at least one account\n" +
			"4. Restart the app and try again",
	};
}

/**
 * Shared collector: bulk-reads metadata for the inbox, selects indices with
 * `pick`, then fetches content only for the messages actually returned.
 */
function collectorScript(pickBody: string, limit: number, account?: string): string {
	// Scoping to one account uses that account's INBOX rather than the unified
	// inbox. Never use a `whose(...)` clause here: Mail evaluates it per message
	// (37s on one 68k mailbox), which is what wedged Mail.app before.
	const boxExpr = account
		? `(function () {
             const accts = M.accounts;
             const names = accts.name();
             let acct = null;
             for (let i = 0; i < names.length; i++) {
               if (String(names[i]) === ${jsLit(account)}) { acct = accts[i]; break; }
             }
             if (!acct) throw new Error("No such account: " + ${jsLit(account)});
             const boxes = acct.mailboxes;
             const bn = boxes.name();
             for (let i = 0; i < bn.length; i++) {
               const nm = String(bn[i]).toLowerCase();
               if (nm === "inbox") return boxes[i];
             }
             return acct.mailboxes[0];
           })()`
		: "M.inbox";
	return `
ObjC.import('Foundation');
const M = Application("Mail");
const box = ${boxExpr};
const msgs = box.messages;
const read = msgs.readStatus();
const subj = msgs.subject();
const send = msgs.sender();
const date = msgs.dateSent();
const n = read.length;
const picked = [];
${pickBody}
// The unified inbox is grouped by ACCOUNT, not globally sorted by date: each
// account's block is internally newest-first, but block order is arbitrary.
// Taking the first N matches therefore returns whichever account sorts first
// (burying today's mail from a busier account behind another's stale mail).
// Sort by date descending before truncating so "latest"/"unread" mean it.
picked.sort(function (a, b) { return (date[b] || 0) - (date[a] || 0); });
const out = [];
const contentDeadline = $.NSDate.dateWithTimeIntervalSinceNow(${CONFIG.CONTENT_BUDGET_SECS});
for (let k = 0; k < picked.length && k < ${limit}; k++) {
  const i = picked[k];
  let content = "";
  if ($.NSDate.date.compare(contentDeadline) >= 0) {
    content = "[Body not fetched: content budget exceeded]";
  } else {
    try {
      content = String(msgs[i].content() || "");
    } catch (e) { content = "[Content not available]"; }
  }
  if (content.length > ${CONFIG.MAX_CONTENT_PREVIEW}) {
    content = content.slice(0, ${CONFIG.MAX_CONTENT_PREVIEW}) + "...";
  }
  out.push({
    subject: String(subj[i] || "(no subject)"),
    sender: String(send[i] || ""),
    dateSent: String(date[i] || ""),
    content: content,
    isRead: read[i] === true,
    mailbox: ${account ? JSON.stringify(account + " - Inbox") : '"Inbox"'}
  });
}
JSON.stringify(out);
`;
}

async function getUnreadMails(limit = 10, account?: string): Promise<EmailMessage[]> {
	try {
		const access = await requestMailAccess();
		if (!access.hasAccess) throw new Error(access.message);
		const max = Math.min(limit, CONFIG.MAX_EMAILS);
		// Inbox is ordered newest-first, so a forward scan yields newest unread.
		const pick = `for (let i = 0; i < n; i++) { if (read[i] === false) picked.push(i); }`;
		return await runJxa<EmailMessage[]>(collectorScript(pick, max, account), CONFIG.TIMEOUT_MS);
	} catch (error) {
		console.error(
			`Error getting unread emails: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function searchMails(searchTerm: string, limit = 10, account?: string): Promise<EmailMessage[]> {
	try {
		const access = await requestMailAccess();
		if (!access.hasAccess) throw new Error(access.message);
		if (!searchTerm || !searchTerm.trim()) return [];
		const max = Math.min(limit, CONFIG.MAX_EMAILS);
		const pick = `
const q = ${jsLit(searchTerm.toLowerCase().trim())};
for (let i = 0; i < n; i++) {
  const s = String(subj[i] || "").toLowerCase();
  const f = String(send[i] || "").toLowerCase();
  if (s.indexOf(q) !== -1 || f.indexOf(q) !== -1) picked.push(i);
}`;
		return await runJxa<EmailMessage[]>(collectorScript(pick, max, account), CONFIG.TIMEOUT_MS);
	} catch (error) {
		console.error(
			`Error searching emails: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function getLatestMails(account?: string, limit = 10): Promise<EmailMessage[]> {
	try {
		const access = await requestMailAccess();
		if (!access.hasAccess) throw new Error(access.message);
		const max = Math.min(limit, CONFIG.MAX_EMAILS);
		const pick = `for (let i = 0; i < n; i++) { picked.push(i); }`;
		return await runJxa<EmailMessage[]>(collectorScript(pick, max, account), CONFIG.TIMEOUT_MS);
	} catch (error) {
		console.error(
			`Error getting latest emails: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function getAccounts(): Promise<string[]> {
	try {
		const access = await requestMailAccess();
		if (!access.hasAccess) throw new Error(access.message);
		const names = await runJxa<string[]>(
			`JSON.stringify(Application("Mail").accounts.name())`,
			30000,
		);
		return (names || []).filter((n) => typeof n === "string" && n.trim() !== "");
	} catch (error) {
		console.error(
			`Error getting accounts: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function getMailboxes(): Promise<string[]> {
	try {
		const access = await requestMailAccess();
		if (!access.hasAccess) throw new Error(access.message);
		const names = await runJxa<string[]>(
			`JSON.stringify(Application("Mail").mailboxes.name())`,
			60000,
		);
		return (names || []).filter((n) => typeof n === "string" && n.trim() !== "");
	} catch (error) {
		console.error(
			`Error getting mailboxes: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

async function getMailboxesForAccount(accountName: string): Promise<string[]> {
	try {
		const access = await requestMailAccess();
		if (!access.hasAccess) throw new Error(access.message);
		if (!accountName || !accountName.trim()) return [];
		const script = `
const M = Application("Mail");
const target = ${jsLit(accountName)};
let out = [];
const accts = M.accounts;
for (let i = 0; i < accts.length; i++) {
  if (String(accts[i].name()) === target) { out = accts[i].mailboxes.name().map(String); break; }
}
JSON.stringify(out);
`;
		const names = await runJxa<string[]>(script, 60000);
		return (names || []).filter((n) => typeof n === "string" && n.trim() !== "");
	} catch (error) {
		console.error(
			`Error getting mailboxes for account: ${error instanceof Error ? error.message : String(error)}`,
		);
		return [];
	}
}

/** Unchanged from upstream — the send path already worked. */
async function sendMail(
	to: string,
	subject: string,
	body: string,
	cc?: string,
	bcc?: string,
): Promise<string | undefined> {
	try {
		const accessResult = await requestMailAccess();
		if (!accessResult.hasAccess) {
			throw new Error(accessResult.message);
		}
		if (!to || !to.trim()) throw new Error("To address is required");
		if (!subject || !subject.trim()) throw new Error("Subject is required");
		if (!body || !body.trim()) throw new Error("Email body is required");

		const tmpFile = `/tmp/email-body-${Date.now()}.txt`;
		const fs = require("fs");
		fs.writeFileSync(tmpFile, body.trim(), "utf8");

		const script = `
tell application "Mail"
    activate
    set emailBody to read file POSIX file "${tmpFile}" as «class utf8»
    set newMessage to make new outgoing message with properties {subject:"${subject.replace(/"/g, '\\"')}", content:emailBody, visible:true}
    tell newMessage
        make new to recipient with properties {address:"${to.replace(/"/g, '\\"')}"}
        ${cc ? `make new cc recipient with properties {address:"${cc.replace(/"/g, '\\"')}"}` : ""}
        ${bcc ? `make new bcc recipient with properties {address:"${bcc.replace(/"/g, '\\"')}"}` : ""}
    end tell
    send newMessage
    return "SUCCESS"
end tell`;

		const result = (await runAppleScript(script)) as string;
		try { fs.unlinkSync(tmpFile); } catch (e) { /* ignore */ }

		if (result === "SUCCESS") {
			return `Email sent to ${to} with subject "${subject}"`;
		}
		throw new Error("Failed to send email");
	} catch (error) {
		console.error(
			`Error sending email: ${error instanceof Error ? error.message : String(error)}`,
		);
		throw new Error(
			`Error sending email: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

interface MoveMailOptions {
	account: string;
	subject: string;
	sender: string;
	/** When set, only a message with this read status matches. */
	isRead?: boolean;
	destinationMailbox: string;
	/** Mailbox path to search; defaults to the account's inbox. */
	sourceMailbox?: string;
}

interface MoveMailResult {
	subject: string;
	sender: string;
	dateSent: string;
	isRead: boolean;
	/** Number of messages that matched; the newest one was moved. */
	matched: number;
	source: string;
	destination: string;
}

/**
 * Moves the newest message in the account's inbox (or `sourceMailbox`) whose
 * subject equals `subject`, whose sender contains `sender`, and whose read
 * status matches `isRead` when given, into `destinationMailbox`.
 */
async function moveMail(opts: MoveMailOptions): Promise<MoveMailResult> {
	const access = await requestMailAccess();
	if (!access.hasAccess) throw new Error(access.message);
	if (!opts.account || !opts.account.trim()) throw new Error("Account is required");
	if (!opts.subject || !opts.subject.trim()) throw new Error("Subject is required");
	if (!opts.sender || !opts.sender.trim()) throw new Error("Sender is required");
	if (!opts.destinationMailbox || !opts.destinationMailbox.trim()) {
		throw new Error("Destination mailbox is required");
	}
	const source = (opts.sourceMailbox || "INBOX").trim();
	const script = `
${MAILBOX_HELPERS}
const M = Application("Mail");
const acct = findAccount(M, ${jsLit(opts.account)});
const src = resolveMailbox(acct, ${jsLit(source)}, true);
if (!src) {
  throw new Error("No mailbox " + ${jsLit(source)} + " in account " + ${jsLit(opts.account)});
}
const dest = resolveMailbox(acct, ${jsLit(opts.destinationMailbox.trim())}, true);
if (!dest) {
  throw new Error("No mailbox " + ${jsLit(opts.destinationMailbox.trim())} + " in account "
    + ${jsLit(opts.account)} + " (use createMailbox first)");
}
// Bulk property reads, never a whose() clause: see collectorScript.
const msgs = src.messages;
const read = msgs.readStatus();
const subj = msgs.subject();
const send = msgs.sender();
const date = msgs.dateSent();
const wantSubj = ${jsLit(opts.subject.trim().toLowerCase())};
const wantSender = ${jsLit(opts.sender.trim().toLowerCase())};
const wantRead = ${opts.isRead === undefined ? "null" : String(opts.isRead)};
const matches = [];
for (let i = 0; i < read.length; i++) {
  if (String(subj[i] || "").trim().toLowerCase() !== wantSubj) continue;
  if (String(send[i] || "").toLowerCase().indexOf(wantSender) === -1) continue;
  if (wantRead !== null && (read[i] === true) !== wantRead) continue;
  matches.push(i);
}
if (matches.length === 0) {
  throw new Error("No " + (wantRead === null ? "" : wantRead ? "read " : "unread ")
    + "message from '" + ${jsLit(opts.sender)} + "' with subject '" + ${jsLit(opts.subject)}
    + "' in " + ${jsLit(source)} + " of account " + ${jsLit(opts.account)});
}
matches.sort(function (a, b) { return (date[b] || 0) - (date[a] || 0); });
const i = matches[0];
const result = {
  subject: String(subj[i] || ""),
  sender: String(send[i] || ""),
  dateSent: String(date[i] || ""),
  isRead: read[i] === true,
  matched: matches.length,
  source: ${jsLit(source)},
  destination: ${jsLit(opts.destinationMailbox.trim())}
};
M.move(msgs[i], { to: dest });
JSON.stringify(result);
`;
	try {
		return await runJxa<MoveMailResult>(script, CONFIG.TIMEOUT_MS);
	} catch (error) {
		const msg = error instanceof Error ? error.message : String(error);
		console.error(`Error moving email: ${msg}`);
		throw new Error(`Error moving email: ${msg}`);
	}
}

interface CreateMailboxResult {
	account: string;
	path: string;
	/** False when the mailbox already existed. */
	created: boolean;
	/** Set when Mail accepted the request but the result is not what was asked. */
	note?: string;
}

/**
 * Creates the mailbox at a slash-separated `path` (e.g. "Clients/Acme") in
 * `account`, including any missing parents.
 *
 * Mail's `make new mailbox` accepts the whole path as the name, and it is the
 * only form Mail honours: making a mailbox inside another mailbox fails with
 * "AppleEvent handler failed" on local, iCloud and Exchange accounts alike.
 * Server-backed accounts create asynchronously, so the result is polled for.
 * Exchange accounts place the new mailbox under the right parent but keep the
 * literal slash path as its name; that case is reported in `note`.
 */
async function createMailbox(account: string, path: string): Promise<CreateMailboxResult> {
	const access = await requestMailAccess();
	if (!access.hasAccess) throw new Error(access.message);
	if (!account || !account.trim()) throw new Error("Account is required");
	const segments = (path || "")
		.split("/")
		.map((s) => s.trim())
		.filter(Boolean);
	if (segments.length === 0) throw new Error("Mailbox path is required");
	const normalized = segments.join("/");
	const script = `
${MAILBOX_HELPERS}
const M = Application("Mail");
const acct = findAccount(M, ${jsLit(account)});
const path = ${jsLit(normalized)};
const result = { account: ${jsLit(account)}, path: path, created: false };
if (!resolveMailbox(acct, path, false)) {
  acct.mailboxes.push(M.Mailbox({ name: path }));
  result.created = true;
  let found = null;
  for (let t = 0; t < 20 && !found; t++) {
    delay(0.5);
    found = resolveMailbox(acct, path, false);
  }
  if (!found) {
    const names = acct.mailboxes.name();
    let literal = false;
    for (let i = 0; i < names.length; i++) {
      if (String(names[i]) === path) { literal = true; break; }
    }
    result.note = literal
      ? "This account kept the literal name '" + path + "' for the new mailbox instead of "
        + "nesting it; rename it in Mail if needed"
      : "Mail accepted the request but the mailbox is not visible yet; "
        + "it may take a moment to sync";
  }
}
JSON.stringify(result);
`;
	try {
		return await runJxa<CreateMailboxResult>(script, CONFIG.TIMEOUT_MS);
	} catch (error) {
		const msg = error instanceof Error ? error.message : String(error);
		console.error(`Error creating mailbox: ${msg}`);
		throw new Error(`Error creating mailbox: ${msg}`);
	}
}

/**
 * Saves a new message to the Drafts mailbox of `account` without sending it.
 * The sender is set from the account so Mail files the draft under it.
 */
async function saveDraft(
	account: string,
	to: string,
	subject: string,
	body: string,
	cc?: string,
	bcc?: string,
): Promise<string> {
	const access = await requestMailAccess();
	if (!access.hasAccess) throw new Error(access.message);
	if (!account || !account.trim()) throw new Error("Account is required");
	if (!to || !to.trim()) throw new Error("To address is required");
	if (!subject || !subject.trim()) throw new Error("Subject is required");
	if (!body || !body.trim()) throw new Error("Email body is required");

	const tmpFile = `/tmp/email-draft-${Date.now()}.txt`;
	fs.writeFileSync(tmpFile, body.trim(), "utf8");

	const script = `
tell application "Mail"
    set theAccount to account "${asLit(account)}"
    set acctAddresses to email addresses of theAccount
    if (count of acctAddresses) is 0 then error "Account has no email address"
    set senderLine to (full name of theAccount) & " <" & (item 1 of acctAddresses) & ">"
    set emailBody to read file POSIX file "${tmpFile}" as «class utf8»
    set newMessage to make new outgoing message with properties {sender:senderLine, subject:"${asLit(subject)}", content:emailBody, visible:false}
    tell newMessage
        make new to recipient with properties {address:"${asLit(to)}"}
        ${cc ? `make new cc recipient with properties {address:"${asLit(cc)}"}` : ""}
        ${bcc ? `make new bcc recipient with properties {address:"${asLit(bcc)}"}` : ""}
    end tell
    save newMessage
    return "SUCCESS"
end tell`;

	try {
		const result = (await runAppleScript(script)) as string;
		if (result !== "SUCCESS") throw new Error("Failed to save draft");
		return `Draft to ${to} with subject "${subject}" saved in account "${account}"`;
	} catch (error) {
		const msg = error instanceof Error ? error.message : String(error);
		console.error(`Error saving draft: ${msg}`);
		throw new Error(`Error saving draft: ${msg}`);
	} finally {
		try { fs.unlinkSync(tmpFile); } catch (e) { /* ignore */ }
	}
}

export default {
	getUnreadMails,
	searchMails,
	sendMail,
	moveMail,
	createMailbox,
	saveDraft,
	getMailboxes,
	getAccounts,
	getMailboxesForAccount,
	getLatestMails,
	requestMailAccess,
};

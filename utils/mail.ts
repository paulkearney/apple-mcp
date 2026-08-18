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

export default {
	getUnreadMails,
	searchMails,
	sendMail,
	getMailboxes,
	getAccounts,
	getMailboxesForAccount,
	getLatestMails,
	requestMailAccess,
};

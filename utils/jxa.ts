import { execFile } from "node:child_process";

/**
 * Runs a JXA (JavaScript for Automation) script and parses its stdout as JSON.
 *
 * The upstream `runAppleScript` helper returns osascript's *human-readable*
 * stdout, which is always a string. Callers then tested `Array.isArray(result)`
 * and silently fell through to `[]`, so every list-returning tool reported "no
 * data found". Emitting JSON and parsing it here removes that whole class of bug.
 */
export async function runJxa<T>(script: string, timeoutMs = 60000): Promise<T> {
	const stdout = await new Promise<string>((resolve, reject) => {
		execFile(
			"osascript",
			["-l", "JavaScript", "-e", script],
			{ timeout: timeoutMs, maxBuffer: 64 * 1024 * 1024 },
			(err, out, stderr) => {
				if (err) {
					const detail = String(stderr || "").trim() || err.message;
					reject(new Error(detail));
					return;
				}
				resolve(out);
			},
		);
	});

	const trimmed = stdout.trim();
	if (trimmed === "") return undefined as T;
	try {
		return JSON.parse(trimmed) as T;
	} catch {
		throw new Error(`Expected JSON from JXA script, got: ${trimmed.slice(0, 200)}`);
	}
}

/** Simple time-boxed memo so repeated tool calls don't re-pay a full bulk fetch. */
export function memoize<T>(fn: () => Promise<T>, ttlMs: number): () => Promise<T> {
	let at = 0;
	let cached: Promise<T> | null = null;
	return () => {
		const now = Date.now();
		if (!cached || now - at > ttlMs) {
			at = now;
			cached = fn().catch((e) => {
				cached = null; // don't cache failures
				throw e;
			});
		}
		return cached;
	};
}

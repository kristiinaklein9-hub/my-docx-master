import { t as loadDocx } from "./_shared/load.js";
import { a as truncate, n as formatComputedRPrParts, r as pad, t as formatComputedPPrParts } from "./_shared/format.js";

//#region skill/tools/inspect-style.ts
async function main() {
	const file = process.argv[2];
	const handle = process.argv[3];
	if (!file || !handle) {
		console.error("Usage: node scripts/inspect_style.js <docx-path> <fingerprint-label-or-hash>");
		process.exit(1);
	}
	try {
		const doc = await loadDocx(file);
		const sum = doc.summary.find((s) => s.label === handle || s.hash === handle);
		if (!sum) {
			const letters = doc.summary.map((s) => s.label).join(", ");
			const hashes = doc.summary.map((s) => s.hash).join(", ");
			console.error(`Fingerprint "${handle}" not found.\n  Available letters: [${letters}]\n  Available hashes:  [${hashes}]`);
			process.exit(1);
		}
		const label = sum.label;
		const matches = doc.paragraphs.filter((p) => p.fingerprint === label);
		const styleIds = Array.from(new Set(matches.map((p) => p.styleId)));
		const out = [];
		out.push(`Fingerprint ${label}: ${sum.description}`);
		if (matches.length > 0) {
			const first = matches[0];
			out.push(`Computed rPr: ${formatRPr(first)}`);
			out.push(`Computed pPr: ${formatPPr(first, matches)}`);
		}
		out.push(`Referenced pStyles: [${styleIds.map((s) => `"${s}"`).join(", ")}]`);
		out.push(`Occurrences: ${matches.length}`);
		out.push("");
		const limit = Math.min(20, matches.length);
		for (let i = 0; i < limit; i++) {
			const p = matches[i];
			out.push(`  #${pad(p.index)} [${p.fingerprint}]  "${truncate(p.text, 40)}"`);
		}
		if (matches.length > limit) out.push(`  ... (showing first ${limit}, total ${matches.length})`);
		console.log(out.join("\n"));
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
function formatRPr(p) {
	const r = p.rPr;
	const parts = formatComputedRPrParts(r, {
		truthyToggles: true,
		filterAutoColor: true
	});
	if (r.underline) parts.push(`underline: ${r.underline}`);
	return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`;
}
function formatPPr(p, all) {
	const parts = formatComputedPPrParts(p.pPr, { numIdDisplay: resolveNumId(all) });
	return parts.length === 0 ? "{}" : `{ ${parts.join(", ")} }`;
}
function resolveNumId(all) {
	const seen = /* @__PURE__ */ new Set();
	let hasUnset = false;
	for (const p of all) if (p.pPr.numId !== void 0) seen.add(p.pPr.numId);
	else hasUnset = true;
	if (seen.size === 0) return null;
	if (seen.size === 1 && !hasUnset) return Array.from(seen)[0];
	return "mixed";
}
main();

//#endregion
export {  };
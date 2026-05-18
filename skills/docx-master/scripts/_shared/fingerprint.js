import { createHash } from "node:crypto";

//#region lib/parse/fingerprint.ts
var Fingerprinter = class {
	assign(paragraphs, styleResolver) {
		const counts = /* @__PURE__ */ new Map();
		const samples = /* @__PURE__ */ new Map();
		const totalTextLen = /* @__PURE__ */ new Map();
		const styleIdCounts = /* @__PURE__ */ new Map();
		for (const p of paragraphs) {
			const hash = makeHash(p);
			counts.set(hash, (counts.get(hash) || 0) + 1);
			if (!samples.has(hash)) samples.set(hash, p);
			totalTextLen.set(hash, (totalTextLen.get(hash) ?? 0) + p.text.trim().length);
			if (p.styleId) {
				let inner = styleIdCounts.get(hash);
				if (!inner) {
					inner = /* @__PURE__ */ new Map();
					styleIdCounts.set(hash, inner);
				}
				inner.set(p.styleId, (inner.get(p.styleId) ?? 0) + 1);
			}
		}
		const sorted = Array.from(counts.entries()).sort((a, b) => {
			if (b[1] !== a[1]) return b[1] - a[1];
			return a[0].localeCompare(b[0]);
		});
		const labels = /* @__PURE__ */ new Map();
		const hashes = /* @__PURE__ */ new Map();
		const summary = [];
		for (let i = 0; i < sorted.length; i++) {
			const [hash, count] = sorted[i];
			const label = letterLabel(i);
			const contentHash = shortHash(hash);
			labels.set(hash, label);
			hashes.set(hash, contentHash);
			const sample = samples.get(hash);
			const avgTextLength = Math.round((totalTextLen.get(hash) ?? 0) / count);
			const dominant = pickDominantStyle(styleIdCounts.get(hash) ?? /* @__PURE__ */ new Map(), count);
			const boundStyleName = dominant && styleResolver ? styleResolver.getStyleDefinition(dominant)?.name : void 0;
			summary.push({
				label,
				hash: contentHash,
				description: describe(sample),
				count,
				rawFingerprint: hash,
				avgTextLength,
				boundStyleId: dominant,
				boundStyleName
			});
		}
		for (const p of paragraphs) {
			const h = makeHash(p);
			p.fingerprint = labels.get(h) || "?";
		}
		return {
			labels,
			hashes,
			summary
		};
	}
};
/** Return the styleId used by ≥80% of paragraphs sharing this fingerprint, or
* undefined when the binding is split — including the "mostly unstyled"
* case (because paragraphs without an explicit pStyle don't contribute to
* `styleCounts`, so their share is implicitly counted against the total).
* Skips "Normal" because every Word doc inherits it by default; surfacing
* it adds no signal beyond "no custom binding". */
function pickDominantStyle(styleCounts, total) {
	let bestId;
	let bestCount = 0;
	for (const [id, n] of styleCounts) if (n > bestCount) {
		bestId = id;
		bestCount = n;
	}
	if (bestId === void 0) return void 0;
	if (bestId === "Normal") return void 0;
	if (bestCount / total < .8) return void 0;
	return bestId;
}
function makeHash(p) {
	const r = p.rPr;
	const pp = p.pPr;
	return `${r.fontAscii || r.fontHAnsi || r.fontEastAsia || "?"}|${r.size !== void 0 ? String(r.size) : "?"}|${(r.bold ? "B" : "") + (r.italic ? "I" : "") + (r.underline ? "U" : "") + (r.caps ? "C" : "")}|${r.color && r.color !== "auto" ? r.color : ""}|${pp.alignment || ""}|${pp.firstLineIndent || pp.firstLineIndentChars ? "1stInd" : ""}|${pp.numId ? "L" : ""}`;
}
function describe(p) {
	const r = p.rPr;
	const pp = p.pPr;
	const parts = [];
	const font = r.fontEastAsia || r.fontAscii || r.fontHAnsi;
	if (font) parts.push(font);
	if (r.size !== void 0) parts.push(`${r.size / 2}pt`);
	if (r.bold) parts.push("Bold");
	if (r.italic) parts.push("Italic");
	if (r.underline) parts.push("Underline");
	if (r.caps) parts.push("Caps");
	if (r.color && r.color !== "auto") parts.push(`#${r.color}`);
	if (pp.alignment) parts.push(capitalize(pp.alignment));
	if (pp.firstLineIndent || pp.firstLineIndentChars) parts.push("1stIndent");
	if (pp.outlineLevel !== void 0) parts.push(`Heading-${pp.outlineLevel + 1}`);
	else if (pp.numId) parts.push("List");
	return parts.join(" ") || "(no formatting)";
}
function capitalize(s) {
	return s.charAt(0).toUpperCase() + s.slice(1);
}
/**
* Content-derived stable identifier — first 6 hex chars of SHA-256 of the
* rawFingerprint string. Same visual fingerprint across docs / edits maps
* to the same hash, so configs that reference fingerprints by hash survive
* doc changes that would shuffle the frequency-sorted letter labels.
*/
function shortHash(raw) {
	return createHash("sha256").update(raw).digest("hex").slice(0, 6);
}
function letterLabel(i) {
	let n = i;
	let out = "";
	do {
		out = String.fromCharCode(65 + n % 26) + out;
		n = Math.floor(n / 26) - 1;
	} while (n >= 0);
	return out;
}

//#endregion
export { Fingerprinter as t };
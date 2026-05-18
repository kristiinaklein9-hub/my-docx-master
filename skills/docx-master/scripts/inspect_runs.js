import { a as firstChildNS, d as textContent, f as wAttr, h as NS, o as getChildren, s as getChildrenNS } from "./_shared/xml-utils.js";
import { t as summarizeTable } from "./_shared/table-classifier.js";
import { t as loadDocx } from "./_shared/load.js";
import { r as pad } from "./_shared/format.js";

//#region skill/tools/inspect-runs.ts
/**
* inspect_runs <docx> <paraIndex>
*
* Dumps every run inside a paragraph: text, rPr, and a "run-level diversity"
* summary that pinpoints which character properties differ across runs
* (intentional inline emphasis) vs. which are uniform (style-controllable).
*
* This tool exists because `inspect_range` only shows the paragraph's
* computed/dominant rPr — it cannot reveal mixed-format paragraphs like
*   <w:r><w:b/>幻觉与安全性</w:r><w:r>与传统软件不同…</w:r>
* where the leading bold phrase is a separate run from the non-bold body.
* For that, you need this tool.
*/
async function main() {
	const file = process.argv[2];
	const idxArg = process.argv[3];
	if (!file || !idxArg) {
		console.error("Usage: node scripts/inspect_runs.js <docx-path> <paragraph-index>");
		process.exit(1);
	}
	const targetIdx = parseInt(idxArg, 10);
	if (isNaN(targetIdx) || targetIdx < 1) {
		console.error("Invalid paragraph index (must be 1-based positive integer)");
		process.exit(1);
	}
	try {
		const doc = await loadDocx(file);
		const para = doc.paragraphs.find((p) => p.index === targetIdx);
		if (!para) {
			const max = doc.paragraphs.length;
			const closest = doc.paragraphs.reduce((best, p) => Math.abs(p.index - targetIdx) < Math.abs(best - targetIdx) ? p.index : best, doc.paragraphs[0]?.index ?? 0);
			console.error(`Paragraph #${targetIdx} not found. Document has ${max} indexed paragraphs (1..${doc.paragraphs[max - 1]?.index ?? 0}). Closest: #${closest}.`);
			console.error("Note: paragraphs inside data/form tables are not indexed and cannot be referenced.");
			process.exit(1);
		}
		const pEl = findParagraphElement(doc.documentDoc, targetIdx);
		if (!pEl) {
			console.error(`Internal error: paragraph #${targetIdx} index lookup failed`);
			process.exit(1);
		}
		const runs = extractRuns(pEl);
		console.log(renderReport(targetIdx, para.fingerprint, para.text, para.styleId, runs));
	} catch (err) {
		console.error(`Error: ${err.message}`);
		process.exit(1);
	}
}
/**
* Walks the body in the same order as DocumentParser to assign indices,
* returning the <w:p> element that matches the target paragraph index.
* Includes paragraphs inside layout tables; skips data/form tables.
*/
function findParagraphElement(doc, targetIdx) {
	const body = firstChildNS(doc.documentElement, NS.w, "body");
	if (!body) return null;
	let counter = 0;
	let found = null;
	const visit = (parent) => {
		if (found) return;
		for (const child of getChildren(parent)) {
			if (found) return;
			if (child.namespaceURI !== NS.w) continue;
			if (child.localName === "p") {
				counter++;
				if (counter === targetIdx) {
					found = child;
					return;
				}
			} else if (child.localName === "tbl") {
				if (summarizeTable(child).classification === "layout") for (const tr of getChildrenNS(child, NS.w, "tr")) for (const tc of getChildrenNS(tr, NS.w, "tc")) {
					visit(tc);
					if (found) return;
				}
			}
		}
	};
	visit(body);
	return found;
}
function extractRuns(pEl) {
	const w = NS.w;
	const out = [];
	let runIdx = 0;
	for (const r of getChildrenNS(pEl, w, "r")) {
		runIdx++;
		let text = "";
		let hasDrawing = false;
		let hasMath = false;
		let hasBreak = false;
		let hasField = false;
		for (const c of getChildren(r)) if (c.namespaceURI === w) {
			if (c.localName === "t") text += textContent(c);
			else if (c.localName === "tab") text += "	";
			else if (c.localName === "br") hasBreak = true;
			else if (c.localName === "drawing") hasDrawing = true;
			else if (c.localName === "fldChar" || c.localName === "instrText") hasField = true;
		} else if (c.namespaceURI === NS.m) hasMath = true;
		const rPrEl = firstChildNS(r, w, "rPr");
		const rPr = rPrEl ? serializeRPr(rPrEl) : {};
		out.push({
			index: runIdx,
			text,
			rPr,
			hasDrawing,
			hasMath,
			hasBreak,
			hasField
		});
	}
	return out;
}
/**
* Convert <w:rPr>'s children into a flat property → string map. Boolean
* toggles ("b", "i", ...) collapse to "on" / "off" so they compare cleanly
* across runs. Font / size / color preserve their `w:val` (and rFonts'
* ascii/hAnsi/eastAsia attrs).
*/
function serializeRPr(rPr) {
	const w = NS.w;
	const out = {};
	for (const c of getChildren(rPr)) {
		if (c.namespaceURI !== w) continue;
		const name = c.localName;
		out[name] = signatureForRPrChild(c, name);
	}
	return out;
}
function signatureForRPrChild(el, name) {
	if (name === "rFonts") return [
		"ascii",
		"hAnsi",
		"eastAsia",
		"cs"
	].map((a) => `${a}=${wAttr(el, a) ?? ""}`).filter((p) => !p.endsWith("=")).join(" ");
	if (TOGGLE_PROPS.has(name)) {
		const v = wAttr(el, "val");
		return v === "0" || v === "false" ? "off" : "on";
	}
	return wAttr(el, "val") ?? "(present)";
}
const TOGGLE_PROPS = new Set([
	"b",
	"bCs",
	"i",
	"iCs",
	"caps",
	"smallCaps",
	"strike",
	"dstrike",
	"vanish",
	"snapToGrid",
	"noProof",
	"outline",
	"shadow",
	"emboss",
	"imprint"
]);
const ROLE_HINT_PROPS = [
	"rFonts",
	"sz",
	"szCs",
	"b",
	"bCs",
	"i",
	"iCs",
	"color",
	"u",
	"highlight",
	"strike",
	"vertAlign"
];
function renderReport(paraIdx, fingerprint, fullText, styleId, runs) {
	const lines = [];
	lines.push(`#${pad(paraIdx)} [${fingerprint}]  pStyle="${styleId}"  ${runs.length} run${runs.length === 1 ? "" : "s"}`);
	const truncated = fullText.length > 80 ? fullText.slice(0, 77) + "…" : fullText;
	lines.push(`  text: ${JSON.stringify(truncated)}`);
	lines.push("");
	if (runs.length === 0) {
		lines.push("  (paragraph has no runs — likely image-only or empty)");
		return lines.join("\n");
	}
	const idxWidth = String(runs.length).length;
	for (const r of runs) {
		const tag = [];
		if (r.hasDrawing) tag.push("[drawing]");
		if (r.hasMath) tag.push("[math]");
		if (r.hasBreak) tag.push("[break]");
		if (r.hasField) tag.push("[field]");
		const tagStr = tag.length > 0 ? " " + tag.join(" ") : "";
		const idx = String(r.index).padStart(idxWidth, " ");
		const txt = r.text.length > 60 ? r.text.slice(0, 57) + "…" : r.text;
		lines.push(`  run ${idx}: text=${JSON.stringify(txt)}${tagStr}`);
		const rPrStr = formatRPrMap(r.rPr);
		lines.push(`         rPr=${rPrStr}`);
	}
	lines.push("");
	const propValues = /* @__PURE__ */ new Map();
	for (const name of ROLE_HINT_PROPS) propValues.set(name, /* @__PURE__ */ new Set());
	for (const r of runs) for (const name of ROLE_HINT_PROPS) propValues.get(name).add(r.rPr[name] ?? "<absent>");
	const mixed = [];
	const uniform = [];
	for (const [name, vals] of propValues) if (vals.size > 1) {
		const display = Array.from(vals).map((v) => v === "<absent>" ? "—" : v).join(" / ");
		mixed.push(`${name}: ${display}`);
	} else {
		const v = Array.from(vals)[0];
		if (v !== "<absent>") uniform.push(`${name}=${v}`);
	}
	lines.push(`  Run-level diversity:`);
	lines.push(`    mixed (preserved on restyle): ${mixed.length === 0 ? "(none)" : mixed.join(", ")}`);
	lines.push(`    uniform (strippable):         ${uniform.length === 0 ? "(none)" : uniform.join(", ")}`);
	return lines.join("\n");
}
function formatRPrMap(rPr) {
	const keys = Object.keys(rPr);
	if (keys.length === 0) return "{}";
	const parts = [];
	for (const k of keys) {
		const v = rPr[k];
		if (v === "on") parts.push(k);
		else if (v === "off") parts.push(`${k}(off)`);
		else if (v === "(present)") parts.push(k);
		else parts.push(`${k}=${v}`);
	}
	return `{ ${parts.join(", ")} }`;
}
main();

//#endregion
export {  };
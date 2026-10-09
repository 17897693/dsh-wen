import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

/**
 * Bundled `dsh-wen` skill provider.
 *
 * Exposes the office_* tool plugin (index.js) as a discoverable skill, so every
 * session sees it in the turn-0 catalog instead of only having four anonymous
 * tools to guess about.
 *
 * @module dsh-wen/skill
 */

// Registry-UNIQUE provider id. It must NOT be "dsh-wen": the desktop build
// ships @deepseek-ai/dsh-skill-office (the bundled `desktop-office`, mounted by
// the desktop host), which already registers a skills provider under the name
// "dsh-office" in the same (global) layer — and SkillsRegistry.registerProvider
// throws on a duplicate name, so this whole row failed to mount on desktop with
// "a skill provider named \"dsh-office\" is already registered".
// The SKILL name below stays "dsh-wen", so `skill("dsh-wen")` is unchanged.
const PROVIDER_NAME = "dsh-wen-tools";
const SKILL_BODY_URL = new URL("./SKILL.md", import.meta.url);
const RESOURCE_BASE = {
	kind: "directory",
	path: fileURLToPath(new URL("./", import.meta.url)),
};

const CANDIDATE = {
	name: "dsh-wen",
	description:
		"Read, create, edit, and convert office documents with the office_read / office_create / office_edit / office_convert tools: Word (.docx .doc .wps), Excel (.xlsx .xls .csv .tsv), PowerPoint (.pptx .ppt .dps), PDF (text layer + scanned pages via built-in OCR), OpenDocument (.odt .ods .odp), HTML (.html .htm), Markdown and plain text (.md .txt .json .rtf). Use for ANY office or PDF file the user attaches, mentions, or asks to produce — before reaching for Python libraries, CLI converters, or third-party converters. 办公文档（Word/Excel/PPT/PDF）读写、编辑、格式转换一律用本技能。",
	invocation: {
		modelInvocable: true,
		userInvocable: true,
	},
	provider: PROVIDER_NAME,
	source: "custom",
	resourceBase: RESOURCE_BASE,
	rank: 600,
	locator: SKILL_BODY_URL,
};

const provider = {
	name: PROVIDER_NAME,
	list: () => Promise.resolve([CANDIDATE]),
	async get(_candidate) {
		return {
			name: CANDIDATE.name,
			description: CANDIDATE.description,
			invocation: CANDIDATE.invocation,
			provider: CANDIDATE.provider,
			source: CANDIDATE.source,
			resourceBase: RESOURCE_BASE,
			// Read per load, so editing SKILL.md takes effect without remounting.
			content: await readFile(SKILL_BODY_URL, "utf8"),
		};
	},
};

/** Cordis plugin name. */
export const name = "skill-office";
/** Service required by this provider. */
export const inject = ["skills"];

/** Register the bundled `dsh-wen` provider on `ctx.skills`. */
export function apply(ctx) {
	ctx.skills.registerProvider(() => provider);
}

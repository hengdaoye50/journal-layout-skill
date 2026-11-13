/**
 * journal-layout DSH 插件：把「中文期刊投稿版式排版」这份 skill 挂进 Harness 的技能目录。
 *
 * 做法与 @deepseek-ai/dsh-skill-office 相同：注册一个 bundled skill provider，
 * 由它把 SKILL.md 的正文交给技能注册表，并声明资源目录（scripts/、references/），
 * 让 Agent 能用文件工具直接读、能按 SKILL.md 里的命令去跑 Python 脚本。
 *
 * 技能本体（SKILL.md / scripts / references）仍然只维护一份：优先用包内 assets/，
 * 没有就回落到插件目录的上一级（本仓库的布局：仓库根就是技能目录）。
 */
import { existsSync, readFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

/** 与 @deepseek-ai/dsh-skill 的 BUNDLED_SKILL_RANK 一致（内置技能层的排序权重）。 */
const BUNDLED_SKILL_RANK = 600;
const SKILL_NAME = "journal-layout";
const PROVIDER_NAME = "journal-layout";

/**
 * 技能根目录：包内 assets/journal-layout 优先，否则回落到仓库根（插件目录的上一级）。
 * 本地目录安装时 install_bundle 链接的是检出目录，因此上一级就是仓库根。
 */
function resolveSkillRoot() {
	const here = dirname(fileURLToPath(import.meta.url));
	const candidates = [
		join(here, "..", "assets", SKILL_NAME),
		join(here, "..", ".."),
	];
	for (const dir of candidates) {
		if (existsSync(join(dir, "SKILL.md"))) return resolve(dir);
	}
	throw new Error(
		`journal-layout: 找不到 SKILL.md。已尝试：${candidates.join(" , ")}`,
	);
}

/** 只取 frontmatter 的 description（不引入 YAML 依赖，够用且零依赖）。 */
function parseSkill(raw, path) {
	raw = raw.replace(/^\uFEFF/u, "");      // 容忍 UTF-8 BOM：否则 ^--- 匹配不到 frontmatter
	const frontmatter = /^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/u.exec(raw);
	if (frontmatter?.[1] === undefined) {
		throw new Error(`${SKILL_NAME}: ${path} 缺少 YAML frontmatter`);
	}
	const meta = frontmatter[1];
	const match = /^description:\s*(.+)$/mu.exec(meta);
	if (!match) throw new Error(`${SKILL_NAME}: ${path} 缺少 description`);
	let description = match[1].trim();
	if (
		(description.startsWith('"') && description.endsWith('"')) ||
		(description.startsWith("'") && description.endsWith("'"))
	) {
		description = description.slice(1, -1);
	}
	return { description, content: raw.slice(frontmatter[0].length).trim() };
}

/** 技能正文后追加运行环境说明：资源绝对路径、依赖、已知限制。 */
function runtimeNotes(skillRoot) {
	const scripts = join(skillRoot, "scripts").split("\\").join("/");
	const references = join(skillRoot, "references").split("\\").join("/");
	const info = {
		skillRoot: skillRoot.split("\\").join("/"),
		scripts,
		references,
		python: "需要 python-docx、lxml、Pillow；抽页预览另需 PyMuPDF 与 LibreOffice",
	};
	return [
		"",
		"## 本机已安装的 skill 资源（绝对路径）",
		"",
		"SKILL.md 里的命令请用下面这些绝对路径，按你的 shell 语法调用。",
		"",
		"```json",
		JSON.stringify({ "journal-layout": info }, undefined, 2),
		"```",
		"",
		"注意：`render_preview.py` 在 Windows 上给 LibreOffice 传 `-env:UserInstallation=`",
		"指向工作区内可写目录；`journal_layout.py --journal china-admin` 默认 `columns_new_page`，",
		"因为 LibreOffice 在“页中分栏 + 大量页下注”时会死循环。",
	].join("\n");
}

function createProvider(skillRoot) {
	const path = join(skillRoot, "SKILL.md");
	const { description } = parseSkill(readFileSync(path, "utf8"), path);
	const candidate = {
		name: SKILL_NAME,
		description,
		invocation: { modelInvocable: true, userInvocable: true },
		provider: PROVIDER_NAME,
		source: "bundled",
		rank: BUNDLED_SKILL_RANK,
		resourceBase: {
			kind: "directory",
			path: skillRoot.split("\\").join("/"),
		},
		locator: path,
	};
	return {
		name: PROVIDER_NAME,
		list: () => Promise.resolve([candidate]),
		async get(_candidate, options) {
			const { rank: _rank, locator, ...summary } = candidate;
			const body = parseSkill(
				await readFile(locator, { encoding: "utf8", signal: options.signal }),
				locator,
			).content;
			return { ...summary, content: body + runtimeNotes(skillRoot) };
		},
	};
}

/** Cordis 插件入口。 */
export function apply(ctx) {
	const skillRoot = resolveSkillRoot();
	const provider = createProvider(skillRoot);
	ctx.skills.registerProvider(() => provider);
}

export const name = "journal-layout-skill";
export const inject = ["skills"];

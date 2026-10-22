/**
 * pi-pptx — run deck build scripts, then inspect and visually verify the result.
 *
 * The tool surface is deliberately small: a deck is built by a re-runnable
 * Python script (`build_deck.py`) that imports `pi_pptx`, and the agent's loop
 * is `write script -> run -> look`. So there is one tool to run the script
 * (`ppt_run_deck`, which injects the interpreter and the backend path) and two
 * tools to look (`ppt_inspect` for geometry, `ppt_shot` for pixels), plus
 * `ppt_doctor` for the font substitution that decides line-width margins.
 *
 * Editing a PPTX shape-by-shape is not a tool call: a deck with 500 shapes is
 * rebuilt from the template in seconds, the script *is* the transaction, and a
 * Python loop beats 500 JSON ops. The op layer stays importable in
 * `../python/pi_pptx` (`from pi_pptx import Deck`).
 *
 * The rules and examples are NOT in the system prompt. They live in `../skills`
 * and are registered once at startup, so Pi only advertises each skill's name,
 * description and absolute path; the full text is read when a task matches.
 * Nothing is injected or swapped while a session runs.
 */
import { execFile } from "node:child_process";
import { copyFile, mkdir, readFile, readdir, stat, unlink } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Type } from "@earendil-works/pi-ai";
import {
	defineTool,
	withFileMutationQueue,
	type ExtensionAPI,
} from "@earendil-works/pi-coding-agent";

const PYTHON = process.env.PI_PPTKIT_PYTHON || "python3";
// Resolved from this file, never from the development checkout: the plugin is
// published as an npm package and `python/` ships next to `extensions/`.
const PYTHON_DIR = fileURLToPath(new URL("../python", import.meta.url));
// Skills carry the rules and examples, and are advertised by path. Resolved the
// same way as python/: dev checkout, local install and npm package all work.
const SKILLS_DIR = fileURLToPath(new URL("../skills", import.meta.url));
const KIT = process.env.PI_PPTKIT_PY || path.join(PYTHON_DIR, "pptkit.py");

const MAX_SHOT_PAGES = 6;
// A montage is ONE image, so the per-image token cost does not grow with it.
const MAX_MONTAGE_PAGES = 60;
// Below this the contact-sheet cells cannot be read for proofreading.
const MONTAGE_LEGIBLE_DPI = 100;

const STDOUT_TAIL_LINES = 40;
const STDOUT_TAIL_BYTES = 16 * 1024;
const TRACEBACK_TAIL_LINES = 30;
const TRACEBACK_TAIL_BYTES = 8 * 1024;

type KitResult = Record<string, any>;

const PPT_PATH = Type.String({
	description: "PPTX file path. Relative paths resolve from the working directory.",
});

function resolveFromCwd(cwd: string, file: string): string {
	const expanded = file.replace(/^~(?=$|\/)/, os.homedir());
	return path.isAbsolute(expanded) ? expanded : path.resolve(cwd, expanded);
}

function kitError(args: string[], error: unknown, stdout: string, stderr: string): Error {
	const detail = (stderr || stdout || (error as Error)?.message || String(error)).trim();
	const suffix = detail.length > 4000 ? `${detail.slice(0, 4000)}\n…truncated` : detail;
	return new Error(`pptkit ${args[0]} failed: ${suffix}`);
}

function execKit(args: string[], cwd: string, signal?: AbortSignal): Promise<KitResult> {
	return new Promise((resolve, reject) => {
		execFile(
			PYTHON,
			[KIT, ...args],
			{
				cwd,
				signal,
				timeout: 300_000,
				maxBuffer: 32 * 1024 * 1024,
				encoding: "utf8",
			},
			(error, stdout, stderr) => {
				if (error) {
					reject(kitError(args, error, String(stdout ?? ""), String(stderr ?? "")));
					return;
				}
				try {
					resolve(JSON.parse(String(stdout)) as KitResult);
				} catch {
					reject(
						new Error(`pptkit ${args[0]} returned non-JSON output: ${String(stdout).slice(0, 1000)}`),
					);
				}
			},
		);
	});
}

function jsonResult(data: KitResult, summary?: string) {
	const text = JSON.stringify(data);
	return {
		content: [{ type: "text" as const, text: summary ? `${summary}\n${text}` : text }],
		details: data,
	};
}

async function statOrNull(file: string) {
	try {
		return await stat(file);
	} catch {
		return null;
	}
}

/**
 * One-line pointer to the flags. The boxes, ids and reasons live in
 * `overflow[]` / `overlaps[]` / `suppressed[]`; repeating them as prose made
 * the same warnings appear twice in every reply.
 */
function shotWarnings(data: KitResult): string {
	const parts: string[] = [];
	for (const shot of data.shots ?? []) {
		const flags = [
			...(shot.overflow ?? []).length ? [`overflow ${shot.overflow.map((i: any) => i.shape).join("/")}`] : [],
			...(shot.overlaps ?? []).length
				? [`overlap ${shot.overlaps.map((i: any) => `${i.a}x${i.b}`).join("/")}`]
				: [],
		];
		if (flags.length) parts.push(`slide ${shot.slide}: ${flags.join(" ")}`);
	}
	const lines: string[] = [];
	if (parts.length) {
		lines.push(
			`WARNING text overflow / shape overlap — ${parts.join("; ")}. Geometry is in overflow[]/overlaps[]; suppressed[] lists the sides and pairs that anchor/container rules skipped (not fixed).`,
		);
	}
	if (data.expectation_failed) {
		lines.push(
			`EXPECTATION FAILED: flags on slide(s) ${(data.flagged_slides ?? []).join(", ")} — fix them before shipping.`,
		);
	}
	return lines.join("\n");
}

// ---------------------------------------------------------------- read tools

const inspectTool = defineTool({
	name: "ppt_inspect",
	label: "PPT Inspect",
	promptSnippet: "Inspect PPTX layouts, shapes (names, paths, inch boxes, text fit), tables, notes and fonts.",
	description: `Read a PPTX: slide size/count, every layout (index, name, placeholders, and the non-placeholder shapes each layout draws — logos, rules, full-bleed backgrounds marked \`inherited: true\`), and for the requested slides the full shape tree with names, recursive paths ("组合 7/矩形 9"), inch boxes, text, tables, pictures and notes.

Every text shape also carries \`text_fit\`: \`est_lines\`, \`capacity_lines\`, \`needed_in\`, \`box_h_in\`, \`slack_in\`, \`fits\`. That is the "will this fit" answer *before* drawing — the expensive path (render, look, fix) is only needed when \`fits: false\`.
\`font_substitutions\` comes along too: a substituted face is ~5% wider, which is the margin every box has to leave.

Modes: \`full\` (default, shape tree), \`summary\` (per-slide title/counts), \`text\` (flat {shape, id, text} per page — proofread the product for typos and terminology, instead of grepping the build script).`,
	parameters: Type.Object({
		path: PPT_PATH,
		slides: Type.Optional(
			Type.String({ description: "Slide selection, e.g. '1' or '1,3' or '2-4'. Default: all." }),
		),
		mode: Type.Optional(
			Type.Union([Type.Literal("full"), Type.Literal("summary"), Type.Literal("text")], {
				description: "full (default) = shape tree; summary = per-slide overview; text = flat text rows only.",
			}),
		),
		depth: Type.Optional(
			Type.Number({
				description: "Group recursion depth. 0 flattens groups to one entry. Default: full depth.",
				minimum: 0,
				maximum: 20,
			}),
		),
	}),
	annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["inspect", params.path];
		if (params.slides) args.push("--slides", params.slides);
		if (params.mode) args.push("--mode", params.mode);
		if (params.depth != null) args.push("--depth", String(params.depth));
		const data = await execKit(args, ctx.cwd, signal);
		const summary = `${data.slide_count} slides, ${data.slides?.length ?? 0} inspected (${data.mode ?? "full"}), ${data.layouts?.length ?? 0} layouts, canvas ${data.slide_size_in?.[0]}x${data.slide_size_in?.[1]} in`;
		return jsonResult(data, summary);
	},
});

const shotTool = defineTool({
	name: "ppt_shot",
	label: "PPT Screenshot",
	promptSnippet: "Render PPTX slides to PNG images (with overflow/overlap warnings) so you can see the result.",
	description: `Render slides of a PPTX to PNG and return the images plus a JSON report. This is how you SEE a deck: use it to check style against a template's sample pages and to self-check every build (overflow, wrapping, overlap, font size).

- Returns at most ${MAX_SHOT_PAGES} images per call. With \`montage\` it returns ONE contact sheet and can cover up to ${MAX_MONTAGE_PAGES} slides, so a whole deck fits in one call; the rest is reported in \`omitted_slides\`.
- \`dpi\`: 80 for rhythm, 110 (default) for checking, 150 for fine detail. A contact sheet cell is ~640px wide — at that size 12pt footnotes are illegible, so proofread at ${MONTAGE_LEGIBLE_DPI}+ dpi on a single page.
- \`montage\`: one contact sheet (e.g. 3 columns) instead of N separate images — best for reviewing pitch and pagination.
- \`region\`: crop [x,y,w,h] in inches, globally or per page: {"3": [x,y,w,h], "default": [...]}.
- \`overflow\`: "loose" (default; skips left/right checks on centred text), "strict" (report every side), "off".
- \`ignore\`: skip shapes by name, "#<id>" or "/regex/" — stable across rebuilds.
- \`density: true\` adds a per-shape table (\`rendered_lines\`, \`needed_in\`, \`box_h_in\`, \`slack_in\`, \`tight\`). Text that does not fit is always one shape that is too short, so this names it directly instead of a page-level average.
- \`compareTo\`: another PPTX to diff against, per page as \`changed_px_ratio\`. This is the only reliable way to see "changing A broke B".
- \`expect: "clean"\`: sets \`clean\` / \`expectation_failed\` so a build gate can be automated.
- A side of the box is only reported when the text anchor allows it: top-anchored first lines never "overflow" upwards. Template title leading (a top spill under 5pt on every page) is not reported at all.
- Report fields: \`shots[].overflow\` (with \`rendered_lines\`/\`needed_in\`/\`slack_in\`), \`shots[].overlaps\`, \`suppressed[]\`, \`flagged_slides\`, \`density\`, \`compared\`, \`font_substitutions\`, \`cache_hit\`, \`pdf\`.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slides: Type.Optional(
			Type.String({ description: "Slides to render, e.g. '2' or '1,3' or '1-60'. Default: from slide 1." }),
		),
		dpi: Type.Optional(Type.Number({ description: "Raster DPI. Default 110.", minimum: 40, maximum: 300 })),
		montage: Type.Optional(
			Type.Number({ description: "Output a contact sheet with this many columns. Default: off.", minimum: 1, maximum: 6 }),
		),
		region: Type.Optional(
			Type.Union([
				Type.Array(Type.Number(), {
					description: "Crop every rendered slide to [x,y,w,h] in inches.",
					minItems: 4,
					maxItems: 4,
				}),
				Type.Record(Type.String(), Type.Array(Type.Number()), {
					description: 'Per-page crops, e.g. {"3": [x,y,w,h], "default": [x,y,w,h]}.',
				}),
			]),
		),
		density: Type.Optional(
			Type.Boolean({ description: "Include the per-shape rendered-lines/needed-height table. Default false." }),
		),
		overflow: Type.Optional(
			Type.Union([Type.Literal("loose"), Type.Literal("strict"), Type.Literal("off")], {
				description: "Overflow heuristic mode. Default loose.",
			}),
		),
		ignore: Type.Optional(
			Type.Array(Type.String(), {
				description: "Shapes to skip in overflow/overlap checks: name, '#<id>' or '/regex/'.",
			}),
		),
		compareTo: Type.Optional(
			Type.String({ description: "Previous PPTX to diff against (per-page changed pixel ratio)." }),
		),
		expect: Type.Optional(
			Type.Union([Type.Literal("clean")], {
				description: 'expect="clean" marks the result failed when any page is flagged.',
			}),
		),
	}),
	annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const limit = params.montage ? MAX_MONTAGE_PAGES : MAX_SHOT_PAGES;
		const args = ["shot", params.path, "--dpi", String(params.dpi ?? 110), "--max-pages", String(limit)];
		if (params.slides) args.push("--slides", params.slides);
		if (params.montage != null) args.push("--montage", String(params.montage));
		if (params.region) {
			args.push("--region", Array.isArray(params.region) ? params.region.join(",") : JSON.stringify(params.region));
		}
		if (params.density) args.push("--density");
		if (params.overflow) args.push("--overflow", params.overflow);
		if (params.ignore?.length) args.push("--ignore", params.ignore.join(","));
		if (params.compareTo) args.push("--compare-to", params.compareTo);
		if (params.expect) args.push("--expect", params.expect);
		const data = await execKit(args, ctx.cwd, signal);
		return shotResult(data);
	},
});

async function shotResult(data: KitResult) {
	const content: any[] = [];
	const warning = shotWarnings(data);
	const report = {
		pdf: data.pdf,
		dpi: data.dpi,
		cache_hit: data.cache_hit,
		overflow_mode: data.overflow_mode,
		shot_slides: data.shot_slides,
		flagged_slides: data.flagged_slides,
		shots: data.shots,
		...(data.density ? { density: data.density } : {}),
		...(data.compared ? { compare_to: data.compare_to, compared: data.compared, changed_vs_previous: data.changed_vs_previous } : {}),
		...(data.clean != null ? { clean: data.clean, expectation_failed: data.expectation_failed } : {}),
		...(data.font_substitutions ? { font_substitutions: data.font_substitutions, render_note: data.render_note } : {}),
		...(data.no_changes ? { no_changes: true, hint: data.hint } : {}),
		...(data.montage ? { montage: data.montage } : {}),
		...(data.omitted_slides ? { omitted_slides: data.omitted_slides, hint: data.hint } : {}),
	};
	content.push({
		type: "text" as const,
		text: warning ? `${warning}\n${JSON.stringify(report)}` : JSON.stringify(report),
	});
	if (data.montage) {
		content.push({
			type: "image" as const,
			data: (await readFile(data.montage.file)).toString("base64"),
			mimeType: "image/png",
		});
	} else {
		for (const shot of data.shots ?? []) {
			content.push({
				type: "image" as const,
				data: (await readFile(shot.file)).toString("base64"),
				mimeType: "image/png",
			});
		}
	}
	return {
		content,
		details: {
			pdf: data.pdf,
			dpi: data.dpi,
			cache_hit: data.cache_hit,
			slides: (data.shots ?? []).map((shot: any) => ({
				slide: shot.slide,
				file: shot.file,
				px: shot.px,
				overflow: shot.overflow?.map((item: any) => item.shape) ?? [],
				overlaps: shot.overlaps ?? [],
			})),
			flagged_slides: data.flagged_slides,
			montage: data.montage,
			density: data.density,
			compared: data.compared,
			font_substitutions: data.font_substitutions,
			omitted_slides: data.omitted_slides,
		},
	};
}

const doctorTool = defineTool({
	name: "ppt_doctor",
	label: "PPT Doctor",
	promptSnippet: "Check the PPTX backend (python-pptx, PyMuPDF, LibreOffice, CJK font substitution).",
	description:
		"Check the PPTX backend environment: Python, python-pptx, PyMuPDF, LibreOffice and how CJK template fonts resolve. Run it once before the first build: the font substitutions it reports are a layout parameter (~5% width), not just environment trivia. Also reads the theme fonts when given a deck.",
	parameters: Type.Object({
		path: Type.Optional(Type.String({ description: "Deck whose theme fonts should be resolved." })),
	}),
	annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = params.path ? ["doctor", params.path] : ["doctor"];
		const data = await execKit(args, ctx.cwd, signal);
		const fonts = Object.entries(data.fonts ?? {})
			.map(([font, match]) => `${font} -> ${match}`)
			.join("; ");
		return jsonResult(
			data,
			`python-pptx ${data.python_pptx}, PyMuPDF ${data.pymupdf}, ${data.soffice_version || "no soffice"}; fonts: ${fonts}`,
		);
	},
});

// ------------------------------------------------------------- run the build

/** Nearest git root (fallback: the directory itself). */
async function findRoot(start: string): Promise<string> {
	const origin = path.resolve(start);
	let dir = origin;
	for (;;) {
		if (await statOrNull(path.join(dir, ".git"))) return dir;
		const parent = path.dirname(dir);
		if (parent === dir) return origin;
		dir = parent;
	}
}

function tail(text: string, lines: number, bytes: number): string {
	const split = text.split("\n");
	const sliced = split.slice(-lines).join("\n");
	return sliced.length > bytes ? `…${sliced.slice(-bytes)}` : sliced;
}

function lineOf(line: string): string {
	return line.replace(/^\s+/, "");
}

/** Parse the optional `@@PAGE@@` / `@@DECK@@` / `@@WARN@@` stdout protocol. */
function parseRunStdout(text: string) {
	const pages: number[] = [];
	const warnings: string[] = [];
	const rest: string[] = [];
	let deckPayload: KitResult | undefined;
	for (const raw of text.split("\n")) {
		const line = lineOf(raw);
		if (line.startsWith("@@PAGE@@")) {
			const page = Number(line.slice("@@PAGE@@".length).trim());
			if (Number.isFinite(page)) pages.push(page);
			continue;
		}
		if (line.startsWith("@@WARN@@")) {
			const message = line.slice("@@WARN@@".length).trim();
			if (message) warnings.push(message);
			continue;
		}
		if (line.startsWith("@@DECK@@")) {
			try {
				deckPayload = JSON.parse(line.slice("@@DECK@@".length).trim());
			} catch {
				warnings.push(`@@DECK@@ 不是合法 JSON: ${line.slice(0, 200)}`);
			}
			continue;
		}
		rest.push(raw);
	}
	return { deck: deckPayload, pages: [...new Set(pages)], warnings, rest: rest.join("\n") };
}

/** Pull type/message/file/line out of a Python traceback. */
function parseTraceback(stderr: string) {
	const lines = stderr.trimEnd().split("\n");
	const lastFile = [...lines].reverse().find((line) => /File ".+", line \d+/.test(line));
	const match = lastFile?.match(/File "(.+)", line (\d+)/);
	let type: string | undefined;
	let message: string | undefined;
	for (const line of [...lines].reverse()) {
		const hit = line.match(/^([A-Za-z_][\w.]*(?:Error|Exception|Interrupt|Warning))(?::\s*(.*))?$/);
		if (hit) {
			type = hit[1];
			message = hit[2];
			break;
		}
	}
	if (!type) {
		// `raise SystemExit("…")` and argparse errors print a bare message with no
		// traceback; the last non-empty line is the only structured thing there is.
		const last = [...lines].reverse().find((line) => line.trim());
		if (last) {
			type = "stderr";
			message = last.trim();
		}
	}
	return {
		error: { type, message, file: match?.[1], line: match ? Number(match[2]) : undefined },
		traceback_tail: tail(stderr, TRACEBACK_TAIL_LINES, TRACEBACK_TAIL_BYTES),
	};
}

function slideNumbersIn(text: string): number[] {
	const found: number[] = [];
	for (const hit of text.matchAll(/(?:slide|第)\s*(\d+)|[sS](\d+)\b/g)) {
		const value = Number(hit[1] ?? hit[2]);
		if (Number.isFinite(value) && value > 0) found.push(value);
	}
	return [...new Set(found)];
}

/** Move (not delete) the previous output aside, so `compareTo` keeps working. */
async function stashVersions(out: string, keep: number): Promise<string[]> {
	const versionsDir = path.join(os.tmpdir(), "pi-pptx", "versions");
	await mkdir(versionsDir, { recursive: true });
	const stamp = new Date().toISOString().replace(/[:.]/g, "-");
	const target = path.join(versionsDir, `${path.basename(out, ".pptx")}.${stamp}.pptx`);
	await copyFile(out, target);
	const prefix = `${path.basename(out, ".pptx")}.`;
	const files = (await readdir(versionsDir)).filter((name) => name.startsWith(prefix)).sort();
	while (files.length > keep) {
		const stale = files.shift();
		if (stale) await unlink(path.join(versionsDir, stale)).catch(() => {});
	}
	return [target];
}

const runDeckTool = defineTool({
	name: "ppt_run_deck",
	label: "PPT Run Deck",
	promptSnippet: "Run a build_deck.py that renders a PPTX from a template, then auto-shoot the changed pages.",
	description: `Run the Python script that builds the deck, with the backend path and interpreter already wired, then render what changed. This is the authoring loop: edit \`build_deck.py\`, run it, look.

Injected environment (never hardcode these in the script):
- \`PYTHONPATH\` += the plugin's \`python/\` directory → \`from pi_pptx import Deck\` works with no \`sys.path\` line.
- \`PI_PPTKIT_PY\` = the backend CLI, for scripts that shell out.
- \`DECK_ROOT\` = nearest git root (fallback: working directory); the script runs with cwd = DECK_ROOT, so \`ROOT / "data/x.csv"\` resolves.
- \`DECK_OUT\` / \`DECK_TEMPLATE\` = \`out\` / \`template\` made absolute.
- \`PYTHONDONTWRITEBYTECODE=1\`, \`PI_PPTX_RUN_ID\` = this run's id.
- \`env\` adds anything else; \`python\` overrides the interpreter (default \`PI_PPTKIT_PYTHON\` or python3).

Optional stdout protocol (parse-friendly progress and summaries):
- \`@@PAGE@@ 3\` — which page is being drawn, so a crash says "died on page 3".
- \`@@DECK@@ {"out":…,"slides":22,"shapes":536,…}\` — \`deck.report()\` prints exactly this.
- \`@@WARN@@ slide 7: bottom_pt 17.7\` — self-check findings; with \`shot:"flagged"\` only those pages are rendered.

Failure handling: exit code, error type/message/file/line, \`traceback_tail\` (last ${TRACEBACK_TAIL_LINES} lines), \`stdout_tail\` (last ${STDOUT_TAIL_LINES} lines) — both hard-truncated so a runaway print cannot flood the context. The previous \`out\` file is left in place (\`stale: true\`), so it can still be shot and compared.

\`shot\`: \`"changed"\` (default: pages this run touched), \`"flagged"\` (only pages with \`@@WARN@@\` — the fast iteration mode), \`"all"\`, or \`false\`.`,
	parameters: Type.Object({
		script: Type.String({ description: "Build script path, relative to the working directory." }),
		out: Type.Optional(Type.String({ description: "Deck to write; injected as DECK_OUT and shot afterwards." })),
		template: Type.Optional(Type.String({ description: "Template deck; injected as DECK_TEMPLATE." })),
		args: Type.Optional(
			Type.Array(Type.String(), { description: "Extra argv for the script, e.g. ['--only', '3,7']." }),
		),
		env: Type.Optional(Type.Record(Type.String(), Type.String(), { description: "Extra environment variables." })),
		python: Type.Optional(Type.String({ description: "Interpreter override. Default PI_PPTKIT_PYTHON or python3." })),
		timeout: Type.Optional(Type.Number({ description: "Seconds before the run is killed. Default 600.", minimum: 1 })),
		shot: Type.Optional(
			Type.Union([Type.Literal("changed"), Type.Literal("flagged"), Type.Literal("all"), Type.Literal(false)], {
				description: "Which pages to render afterwards. Default changed.",
			}),
		),
		dpi: Type.Optional(Type.Number({ description: "Screenshot DPI. Default 110.", minimum: 40, maximum: 300 })),
		overflow: Type.Optional(
			Type.Union([Type.Literal("loose"), Type.Literal("strict"), Type.Literal("off")], {
				description: "Overflow mode for the auto-shot. Default loose.",
			}),
		),
		expect: Type.Optional(
			Type.Union([Type.Literal("clean")], { description: 'expect="clean" turns any flag into ok=false.' }),
		),
		keepVersions: Type.Optional(
			Type.Number({
				description: "Keep this many previous outputs in <tmp>/pi-pptx/versions for compareTo. Default 0.",
				minimum: 0,
				maximum: 20,
			}),
		),
		compareTo: Type.Optional(
			Type.String({ description: "Diff the rendered pages against another PPTX (e.g. the stashed previous version)." }),
		),
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const script = resolveFromCwd(ctx.cwd, params.script);
		if (!(await statOrNull(script))) throw new Error(`build script not found: ${script}`);
		const root = await findRoot(ctx.cwd);
		// Paths the caller writes resolve from the working directory (same rule as
		// `ppt_shot`); DECK_ROOT only decides what `ROOT / "..."` means *inside* the
		// script.
		const out = params.out ? resolveFromCwd(ctx.cwd, params.out) : undefined;
		const template = params.template ? resolveFromCwd(ctx.cwd, params.template) : undefined;
		const python = params.python || PYTHON;
		const runId = `run-${Date.now().toString(36)}`;

		const run = async () => {
			const before = out ? await statOrNull(out) : null;
			const versions: string[] = [];
			if (out && before && params.keepVersions) {
				versions.push(...(await stashVersions(out, params.keepVersions)));
			}

			const env: NodeJS.ProcessEnv = {
				...process.env,
				...params.env,
				PYTHONPATH: [PYTHON_DIR, process.env.PYTHONPATH].filter(Boolean).join(path.delimiter),
				PYTHONDONTWRITEBYTECODE: "1",
				PI_PPTKIT_PY: KIT,
				DECK_ROOT: root,
				PI_PPTX_RUN_ID: runId,
			};
			if (out) env.DECK_OUT = out;
			if (template) env.DECK_TEMPLATE = template;

			const started = Date.now();
			const result = await new Promise<{ code: number | null; stdout: string; stderr: string; error?: any }>(
				(resolve) => {
					execFile(
						python,
						[script, ...(params.args ?? [])],
						{
							cwd: root,
							env,
							signal,
							timeout: (params.timeout ?? 600) * 1000,
							maxBuffer: 16 * 1024 * 1024,
							encoding: "utf8",
						},
						(error, stdout, stderr) =>
							resolve({
								code: error?.code != null && typeof error.code === "number" ? error.code : error ? null : 0,
								stdout: String(stdout ?? ""),
								stderr: String(stderr ?? ""),
								error,
							}),
					);
				},
			);
			const duration = (Date.now() - started) / 1000;
			const parsed = parseRunStdout(result.stdout);
			const after = out ? await statOrNull(out) : null;
			const wrote = Boolean(after) && (!before || after!.mtimeMs !== before.mtimeMs || after!.size !== before.size);
			const ok = !result.error;

			// Fallback deck summary: the script may not print @@DECK@@.
			let deck = parsed.deck ?? null;
			if (!deck && out && ok && wrote) {
				try {
					const summary = await execKit(["inspect", out, "--mode", "summary"], root, signal);
					deck = {
						out,
						slides: summary.slide_count,
						shapes: (summary.slides ?? []).reduce((sum: number, s: any) => sum + (s.shape_count ?? 0), 0),
						tables: (summary.slides ?? []).filter((s: any) => s.has_table).length,
						pictures: (summary.slides ?? []).filter((s: any) => s.has_picture).length,
						notes: (summary.slides ?? []).filter((s: any) => s.notes).length,
					};
				} catch {
					deck = null;
				}
			}
			// Pages this run touched. `@@DECK@@` (deck.report()) knows exactly which
			// ones; `@@PAGE@@` markers are the fallback for a script without it.
			const changed: number[] = deck?.changed_slides?.length ? [...deck.changed_slides] : parsed.pages;

			const report: KitResult = {
				ok,
				exit_code: result.code,
				duration_s: Math.round(duration * 10) / 10,
				python,
				script,
				root,
				run_id: runId,
				stdout_tail: tail(parsed.rest, STDOUT_TAIL_LINES, STDOUT_TAIL_BYTES),
				warnings: parsed.warnings,
				changed_slides: changed,
			};
			if (out) report.out = out;
			if (template) report.template = template;
			if (versions.length) report.versions = versions;
			if (deck) report.deck = deck;
			if (out && !ok) report.stale = Boolean(before);
			if (!ok) {
				Object.assign(report, parseTraceback(result.stderr || result.stdout));
				report.timed_out = result.error?.killed === true || result.error?.signal === "SIGTERM";
			}

			// Auto-shot: only what changed (or only what warned) — two round trips
			// per iteration is what made the manual loop slow.
			const mode = params.shot === undefined ? "changed" : params.shot;
			if (!ok || mode === false || !out) {
				report.shot = {
					skipped: !ok ? "build failed" : mode === false ? "shot disabled" : "no out= given to shoot",
				};
				return report;
			}
			let slides: number[] | undefined;
			if (mode === "flagged") {
				slides = [...new Set(parsed.warnings.flatMap((line) => slideNumbersIn(line)))];
				if (!slides.length) {
					report.shot = { skipped: "no @@WARN@@ lines from this run" };
					return report;
				}
			} else if (mode === "changed") {
				slides = changed.filter((page) => page >= 1);
			}
			const shotArgs = ["shot", out, "--dpi", String(params.dpi ?? 110), "--max-pages", String(MAX_SHOT_PAGES)];
			if (slides?.length) shotArgs.push("--slides", slides.join(","));
			if (params.overflow) shotArgs.push("--overflow", params.overflow);
			if (params.compareTo) {
				shotArgs.push("--compare-to", resolveFromCwd(ctx.cwd, params.compareTo));
			}
			try {
				report.shot = await execKit(shotArgs, root, signal);
			} catch (error) {
				report.shot = { error: (error as Error).message.slice(0, 2000) };
			}
			return report;
		};

		const report = out ? await withFileMutationQueue(out, run) : await run();
		if (params.expect === "clean" && (report.shot?.flagged_slides ?? []).length) {
			report.ok = false;
			report.expectation = "clean";
			report.expectation_failed = true;
		}

		const summaryLines = [
			`${report.ok ? "OK" : "FAILED"} ppt_run_deck ${path.relative(root, script) || script}: exit ${report.exit_code ?? "?"} in ${report.duration_s}s`,
		];
		if (report.deck) {
			const deck = report.deck;
			summaryLines.push(
				`deck: ${deck.slides} slides, ${deck.shapes} shapes, ${deck.tables} tables, ${deck.pictures} pictures, ${deck.notes} notes${report.out ? ` -> ${report.out}` : ""}`,
			);
		}
		if (report.versions?.length) {
			summaryLines.push(`previous version kept at ${report.versions.join(", ")} (pass it as compareTo next run)`);
		}
		if (report.stale) summaryLines.push("out still holds the previous version (Deck.save is atomic).");
		if (report.changed_slides?.length) summaryLines.push(`changed slides: ${report.changed_slides.join(", ")}`);
		if (report.warnings?.length) summaryLines.push(`WARNINGS: ${report.warnings.join(" | ")}`);
		if (report.error?.type) {
			const where = report.error.file ? ` at ${report.error.file}:${report.error.line}` : "";
			summaryLines.push(`ERROR ${report.error.type}: ${report.error.message ?? ""}${where}`);
		}
		if (report.shot?.skipped) summaryLines.push(`shot skipped: ${report.shot.skipped}`);
		if (report.shot?.flagged_slides?.length) {
			summaryLines.push(`shot flags on slides ${report.shot.flagged_slides.join(", ")}`);
		}
		if (report.stale && report.shot?.shots?.length === 0) {
			summaryLines.push("the shot above shows the previous build; fix the error before looking at pixels.");
		}

		const content: any[] = [{ type: "text" as const, text: `${summaryLines.join("\n")}\n${JSON.stringify(report)}` }];
		const shot = report.shot;
		if (shot?.shots) {
			if (shot.montage) {
				content.push({
					type: "image" as const,
					data: (await readFile(shot.montage.file)).toString("base64"),
					mimeType: "image/png",
				});
			} else {
				for (const entry of shot.shots) {
					content.push({
						type: "image" as const,
						data: (await readFile(entry.file)).toString("base64"),
						mimeType: "image/png",
					});
				}
			}
		}
		return {
			content,
			details: {
				ok: report.ok,
				exit_code: report.exit_code,
				out: report.out,
				deck: report.deck,
				changed_slides: report.changed_slides,
				warnings: report.warnings,
				shot: shot
					? {
							shots: shot.shots,
							flagged_slides: shot.flagged_slides,
							compared: shot.compared,
							changed_vs_previous: shot.changed_vs_previous,
							error: shot.error,
							skipped: shot.skipped,
						}
					: undefined,
				error: report.error,
				stale: report.stale,
			},
		};
	},
});

// ------------------------------------------------------------- skills + cmd
export default function (pi: ExtensionAPI) {
	for (const tool of [inspectTool, shotTool, doctorTool, runDeckTool]) {
		pi.registerTool(tool);
	}

	// Skills, not a mid-session prompt swap: the deck rules and examples are
	// discovered once at startup, advertised by their own absolute path, and read
	// only when the task matches. Enabling this plugin already means PPT work, so
	// no prompt sniffing is needed.
	pi.on("resources_discover", () => ({ skillPaths: [SKILLS_DIR] }));

	pi.registerCommand("pptx", {
		description: "Check the pptx backend (python-pptx, PyMuPDF, LibreOffice, fonts)",
		handler: async (_args, ctx) => {
			const data = await execKit(["doctor"], ctx.cwd);
			const fonts = Object.entries(data.fonts ?? {})
				.map(([font, match]) => `${font} -> ${match}`)
				.join("\n");
			ctx.ui.notify(
				`pi-pptx backend\npython-pptx ${data.python_pptx} · PyMuPDF ${data.pymupdf} · ${data.soffice_version || "soffice missing"}\n${fonts}`,
				"info",
			);
		},
	});
}

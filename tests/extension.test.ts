/**
 * Extension-level smoke test: load the real extension factory with a stub
 * ExtensionAPI, then call tool.execute() directly (no model involved).
 *
 * Run: node --test tests/extension.test.ts   (Node 22.18+/26 strips types)
 */
import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { after, test } from "node:test";
import { fileURLToPath } from "node:url";

import type { ExtensionAPI, ToolDefinition } from "@earendil-works/pi-coding-agent";

import extension from "../extensions/pi-pptx.ts";

type Handler = (...args: any[]) => any;

const tools = new Map<string, ToolDefinition<any, any, any>>();
const events: Array<{ name: string; handler: Handler }> = [];
const commands: string[] = [];

const stub = {
	registerTool(tool: ToolDefinition) {
		tools.set(tool.name, tool as ToolDefinition<any, any, any>);
	},
	on(name: string, handler: Handler) {
		events.push({ name, handler });
		return () => {};
	},
	registerCommand(name: string) {
		commands.push(name);
	},
} as unknown as ExtensionAPI;

extension(stub);

const dirs: string[] = [];
async function tempDir(): Promise<string> {
	const dir = await mkdtemp(path.join(tmpdir(), "pi-pptx-ext-"));
	dirs.push(dir);
	return dir;
}

// Minimal ExtensionToolContext: the tools only use ctx.cwd.
function ctxFor(cwd: string) {
	return { cwd } as any;
}

async function call(name: string, params: Record<string, unknown>, cwd: string): Promise<any> {
	const tool = tools.get(name);
	assert.ok(tool, `tool ${name} is registered`);
	return tool!.execute("test-call", params, undefined, undefined, ctxFor(cwd));
}

/** Write a build script that only relies on the injected environment. */
async function buildScript(cwd: string, name: string, body: string): Promise<string> {
	const header = [
		"#!/usr/bin/env python3",
		"import os",
		"from pathlib import Path",
		"from pi_pptx import Deck  # PYTHONPATH is injected by ppt_run_deck",
		"",
		'OUT = Path(os.environ["DECK_OUT"])',
		'ROOT = Path(os.environ["DECK_ROOT"])',
		"",
		"",
	].join("\n");
	await writeFile(path.join(cwd, name), `${header}${body.trimEnd()}\n`, "utf8");
	return name;
}

after(async () => {
	await Promise.all(dirs.map((dir) => rm(dir, { recursive: true, force: true })));
});

test("registers only the run/inspect/shot/doctor toolset", () => {
	assert.deepEqual([...tools.keys()].sort(), ["ppt_doctor", "ppt_inspect", "ppt_run_deck", "ppt_shot"]);
	// Editing a shape is a script call, not a tool: the deck is rebuilt instead.
	for (const gone of ["ppt_apply", "ppt_set_text", "ppt_clone", "ppt_add_shape", "ppt_delete"]) {
		assert.equal(tools.has(gone), false, `${gone} must not be registered`);
	}
	assert.deepEqual(commands, ["pptx"]);
	assert.ok(events.some((event) => event.name === "resources_discover"));
	// No per-run system prompt mutation: the rules live in skills/, advertised once
	// at startup, so the prompt is fixed from the first turn.
	assert.equal(events.some((event) => event.name === "before_agent_start"), false);
	assert.equal(tools.get("ppt_shot")!.annotations?.readOnlyHint, true);
	assert.equal(tools.get("ppt_run_deck")!.annotations?.readOnlyHint, false);
});

test("ppt_run_deck injects the backend path, shoots changed pages and reports the deck", async () => {
	const cwd = await tempDir();
	const script = await buildScript(
		cwd,
		"build.py",
		`
D = Deck.create(OUT, overwrite=True)
print("@@PAGE@@ 1", flush=True)
D.add_text(slide=1, text="脚本生成的标题", x=1, y=1, w=8, h=1, size=28)
D.save()
D.report()
`,
	);
	const result = await call("ppt_run_deck", { script, out: "out/deck.pptx", shot: "changed", dpi: 70 }, cwd);
	assert.equal(result.details.ok, true);
	assert.equal(result.details.exit_code, 0);
	assert.equal(result.details.out, path.join(cwd, "out/deck.pptx"));
	assert.equal(result.details.deck.slides, 1);
	assert.ok(result.details.deck.shapes >= 1);
	assert.deepEqual(result.details.changed_slides, [1]);
	assert.match(result.content[0].text, /^OK ppt_run_deck build\.py/);
	assert.match(result.content[0].text, /deck: 1 slides/);
	// The auto-shot is part of the same call: one image, no second round trip.
	const images = result.content.filter((block: any) => block.type === "image");
	assert.equal(images.length, 1);
	assert.equal(images[0].mimeType, "image/png");
	assert.ok(images[0].data.length > 1000);
});

test("ppt_run_deck returns a structured failure and keeps the previous output", async () => {
	const cwd = await tempDir();
	await buildScript(
		cwd,
		"ok.py",
		`
D = Deck.create(OUT, overwrite=True)
D.add_text(slide=1, text="上一版", x=1, y=1, w=6, h=1, size=24)
D.save()
`,
	);
	const first = await call("ppt_run_deck", { script: "ok.py", out: "deck.pptx", shot: false }, cwd);
	assert.equal(first.details.ok, true);
	assert.ok(first.details.deck, "deck summary falls back to inspect when @@DECK@@ is missing");

	const broken = await buildScript(
		cwd,
		"broken.py",
		`
print("第一页开始画")
raise ValueError("unknown shape '内容占位符 6'")
`,
	);
	const failed = await call("ppt_run_deck", { script: broken, out: "deck.pptx", shot: "changed" }, cwd);
	assert.equal(failed.details.ok, false);
	assert.equal(failed.details.stale, true);
	assert.equal(failed.details.error.type, "ValueError");
	assert.equal(failed.details.error.message, "unknown shape '内容占位符 6'");
	assert.ok(String(failed.details.error.file).endsWith("broken.py"));
	assert.ok(failed.details.error.line >= 10, "traceback line points into the body");
	assert.equal(failed.details.shot.skipped, "build failed");
	assert.match(failed.content[0].text, /^FAILED ppt_run_deck broken\.py/);
	assert.match(failed.content[0].text, /ERROR ValueError/);
	// The previous build is still on disk and still shot-able.
	const shot = await call("ppt_shot", { path: "deck.pptx", slides: "1", dpi: 60 }, cwd);
	assert.equal(shot.details.slides.length, 1);
});

test("ppt_run_deck shot=flagged renders only the pages from @@WARN@@", async () => {
	const cwd = await tempDir();
	const script = await buildScript(
		cwd,
		"three.py",
		`
D = Deck.create(OUT, empty=True)
D.add_slide(0)
D.add_slide(0)
D.add_slide(0)
D.add_text(slide=2, text="第二页", x=1, y=1, w=4, h=0.6, size=20)
print("@@WARN@@ slide 2: bottom_pt 12.0", flush=True)
D.save()
D.report()
`,
	);
	const flagged = await call(
		"ppt_run_deck",
		{ script, out: "deck.pptx", shot: "flagged", dpi: 60 },
		cwd,
	);
	assert.equal(flagged.details.ok, true);
	assert.deepEqual(flagged.details.deck.slides, 3);
	assert.deepEqual(flagged.details.shot.shots.map((shot: any) => shot.slide), [2]);

	// No @@WARN@@ means no render at all — the fast iteration path.
	const quiet = await buildScript(
		cwd,
		"quiet.py",
		`
D = Deck.create(OUT, overwrite=True)
D.add_text(slide=1, text="干净的一页", x=1, y=1, w=6, h=1, size=24)
D.save()
D.report()
`,
	);
	const skipped = await call("ppt_run_deck", { script: quiet, out: "deck.pptx", shot: "flagged" }, cwd);
	assert.equal(skipped.details.ok, true);
	assert.match(String(skipped.details.shot.skipped), /no @@WARN@@/);
});

test("ppt_inspect reports text_fit, font substitutions and a text mode", async () => {
	const cwd = await tempDir();
	const script = await buildScript(
		cwd,
		"build.py",
		`
D = Deck.create(OUT, overwrite=True)
D.add_text(slide=1, text="这一行放不下" * 4, x=1, y=1, w=1.5, h=0.3, size=18)
D.save()
`,
	);
	await call("ppt_run_deck", { script, out: "deck.pptx", shot: false }, cwd);

	const inspected = await call("ppt_inspect", { path: "deck.pptx", slides: "1" }, cwd);
	assert.equal(inspected.details.slide_count, 1);
	const shape = inspected.details.slides[0].shapes.find((item: any) => item.text?.startsWith("这一行放不下"));
	assert.ok(shape, "text shape is inspectable");
	assert.equal(shape.text_fit.estimated, true);
	assert.equal(shape.text_fit.fits, false);
	assert.ok(shape.text_fit.est_lines > shape.text_fit.capacity_lines);
	assert.ok(inspected.details.font_substitutions !== undefined);

	const textMode = await call("ppt_inspect", { path: "deck.pptx", slides: "1", mode: "text" }, cwd);
	assert.equal(textMode.details.mode, "text");
	assert.equal(textMode.details.layouts, undefined);
	const texts = textMode.details.slides[0].shapes.map((row: any) => row.text);
	assert.ok(texts.some((text: string) => text.startsWith("这一行放不下")));
});

test("ppt_shot exposes measured overflow, per-shape density and compareTo", async () => {
	const cwd = await tempDir();
	const script = await buildScript(
		cwd,
		"build.py",
		`
D = Deck.create(OUT, overwrite=True)
D.add_text(slide=1, text="确定性的溢出文字" * 12, x=1, y=1, w=3, h=0.4, size=18)
D.save()
`,
	);
	await call("ppt_run_deck", { script, out: "deck.pptx", shot: false }, cwd);

	const shot = await call(
		"ppt_shot",
		{ path: "deck.pptx", slides: "1", dpi: 60, density: true, expect: "clean", compareTo: "deck.pptx" },
		cwd,
	);
	assert.match(shot.content[0].text, /WARNING text overflow/);
	assert.match(shot.content[0].text, /EXPECTATION FAILED/);
	const image = shot.content.find((block: any) => block.type === "image");
	assert.ok(image, "shot returns an image block");
	assert.equal(image.mimeType, "image/png");

	const report = shot.details;
	const overflow = shot.content[0].text;
	assert.ok(overflow.includes("rendered_lines"));
	assert.deepEqual(report.flagged_slides, [1]);
	const dense = report.density[0].shapes.find((row: any) => row.rendered_lines > 1);
	assert.ok(dense, "density names the shape that actually overflows");
	assert.ok(dense.slack_in < 0);
	assert.deepEqual(report.compared.map((row: any) => row.changed_px_ratio), [0]);
});

test("ppt_doctor resolves versions, soffice and font substitution", async () => {
	const cwd = await tempDir();
	const doctor = await call("ppt_doctor", {}, cwd);
	assert.ok(doctor.details.python_pptx);
	assert.ok(doctor.details.fonts);
	assert.match(doctor.content[0].text, /python-pptx/);
});

test("keepVersions + compareTo diff the previous build", async () => {
	const cwd = await tempDir();
	await mkdir(path.join(cwd, "tpl"), { recursive: true });
	// A template deck, built with the backend itself.
	await writeFile(path.join(cwd, "tpl/make.py"), `
import os
from pathlib import Path
from pi_pptx import Deck
tpl = Path(os.environ["DECK_ROOT"]) / "tpl/template.pptx"
d = Deck.create(tpl, overwrite=True)
d.add_text(slide=1, text="模板页", x=1, y=1, w=6, h=1, size=24)
d.save()
`);
	await writeFile(path.join(cwd, "build.py"), `
import os, sys
from pathlib import Path
from pi_pptx import Deck
ROOT = Path(os.environ["DECK_ROOT"])
OUT = Path(os.environ["DECK_OUT"])
word = sys.argv[1] if len(sys.argv) > 1 else "v1"
D = Deck.create(OUT, template=ROOT / "tpl/template.pptx", overwrite=True)
D.add_text(slide=1, text=word, x=1, y=2.5, w=6, h=1, size=24)
D.save()
D.report()
`);
	await call("ppt_run_deck", { script: "tpl/make.py", shot: false }, cwd);

	const first = await call("ppt_run_deck", { script: "build.py", out: "deck.pptx", shot: false, expect: "clean" }, cwd);
	assert.equal(first.details.ok, true, JSON.stringify(first.content[0].text).slice(0, 500));

	const second = await call(
		"ppt_run_deck",
		{ script: "build.py", out: "deck.pptx", args: ["v2"], keepVersions: 1, shot: "changed", dpi: 60 },
		cwd,
	);
	assert.ok(second.details.ok);
	const text = second.content[0].text as string;
	const version = /previous version kept at (\S+?\.pptx)/.exec(text)?.[1];
	assert.ok(version, "stashed version path is reported");
	const diffed = await call(
		"ppt_run_deck",
		{ script: "build.py", out: "deck.pptx", args: ["v2"], shot: "changed", dpi: 60, compareTo: version },
		cwd,
	);
	// The stashed version is the v1 build, the current one is v2: one page differs.
	assert.ok(diffed.details.shot.compared[0].changed_px_ratio > 0);
	assert.deepEqual(diffed.details.shot.changed_vs_previous, [1]);
	// Comparing a deck with itself must be exactly zero on every page.
	const same = await call(
		"ppt_run_deck",
		{ script: "build.py", out: "deck.pptx", args: ["v2"], shot: "changed", dpi: 60, compareTo: "deck.pptx" },
		cwd,
	);
	assert.deepEqual(same.details.shot.compared.map((r: any) => r.changed_px_ratio), [0]);
	assert.deepEqual(same.details.shot.changed_vs_previous, []);
});

test("advertises the bundled skills once at startup instead of rewriting the prompt", async () => {
	const handler = events.find((event) => event.name === "resources_discover")!.handler;
	const result = await handler({ type: "resources_discover", cwd: process.cwd(), reason: "startup" });

	assert.equal(result?.skillPaths?.length, 1);
	const skillsDir = result.skillPaths[0];
	// Path comes from the extension's own location, so it works from a checkout, a
	// local install and an npm package.
	assert.ok(path.isAbsolute(skillsDir), `${skillsDir} must be absolute`);
	assert.equal(path.dirname(skillsDir), path.resolve(path.dirname(fileURLToPath(import.meta.url)), ".."));

	const dirs = (await readdir(skillsDir, { withFileTypes: true })).filter((entry) => entry.isDirectory());
	assert.ok(dirs.length >= 3, "deck-build / layout-design / visual-review skills ship with the plugin");

	const build = await readFile(path.join(skillsDir, "pptx-deck-build", "SKILL.md"), "utf8");
	assert.match(build, /^---\nname: pptx-deck-build\ndescription: .+\n---\n/);
	assert.match(build, /ppt_run_deck/);
	assert.match(build, /build_deck\.py/);
	// The skill points at the runnable example, which must exist next to the skills.
	await stat(path.join(skillsDir, "..", "examples", "build_deck.py"));

	for (const name of ["pptx-layout-design", "pptx-visual-review"]) {
		const text = await readFile(path.join(skillsDir, name, "SKILL.md"), "utf8");
		assert.match(text, new RegExp(`^---\\nname: ${name}\\ndescription: .+\\n---\\n`));
	}
});

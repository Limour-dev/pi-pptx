/**
 * Extension-level smoke test: load the real extension factory with a stub
 * ExtensionAPI, then call tool.execute() directly (no model involved).
 *
 * Run: node --test tests/extension.test.ts   (Node 22.18+/26 strips types)
 */
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { after, test } from "node:test";

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

after(async () => {
	await Promise.all(dirs.map((dir) => rm(dir, { recursive: true, force: true })));
});

test("registers the pptx toolset and command", () => {
	for (const name of [
		"ppt_inspect",
		"ppt_shot",
		"ppt_doctor",
		"ppt_create",
		"ppt_clone",
		"ppt_add_slide",
		"ppt_set_text",
		"ppt_set_cell",
		"ppt_add_text",
		"ppt_add_image",
		"ppt_add_table",
		"ppt_delete",
	]) {
		assert.ok(tools.has(name), `${name} registered`);
	}
	assert.deepEqual(commands, ["pptx"]);
	assert.ok(events.some((event) => event.name === "before_agent_start"));
	assert.equal(tools.get("ppt_shot")!.annotations?.readOnlyHint, true);
});

test("create -> edit -> shot returns images and overflow warnings", async () => {
	const cwd = await tempDir();
	const deck = "deck.pptx";

	const created = await call("ppt_create", { path: deck }, cwd);
	assert.match(created.content[0].text, /1 slides/);

	await call(
		"ppt_add_text",
		{ path: deck, slide: 1, text: "正常标题", x: 1, y: 1, w: 6, h: 1, size: 24 },
		cwd,
	);
	const longText = "这是一段特意写得很长很长的说明文字，用来制造确定性的文本溢出。".repeat(3);
	const bad = await call(
		"ppt_add_text",
		{ path: deck, slide: 1, text: longText, x: 1, y: 2.5, w: 4, h: 0.3, size: 18 },
		cwd,
	);
	const badName = bad.details.name as string;

	const inspected = await call("ppt_inspect", { path: deck, slides: "1" }, cwd);
	assert.equal(inspected.details.slide_count, 1);
	const shapePaths = inspected.details.slides[0].shapes.map((shape: any) => shape.path);
	assert.ok(shapePaths.includes(badName));

	const shot = await call("ppt_shot", { path: deck, slides: "1", dpi: 80 }, cwd);
	assert.equal(shot.content[0].type, "text");
	assert.match(shot.content[0].text, /WARNING text overflow/);
	assert.match(shot.content[0].text, new RegExp(badName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
	const image = shot.content.find((block: any) => block.type === "image");
	assert.ok(image, "shot returns an image block");
	assert.equal(image.mimeType, "image/png");
	assert.ok(image.data.length > 1000);
	assert.ok(shot.details.slides[0].overflow.includes(badName));
});

test("clone + set_text preserve style and update the copy only", async () => {
	const cwd = await tempDir();
	const deck = "deck.pptx";
	await call("ppt_create", { path: deck }, cwd);

	const added = await call(
		"ppt_add_text",
		{
			path: deck,
			slide: 1,
			text: "原始标题",
			x: 1,
			y: 1,
			w: 8,
			h: 1,
			size: 30,
			font: "微软雅黑",
			color: "#1F4E79",
		},
		cwd,
	);
	const name = added.details.name as string;

	await call("ppt_clone", { path: deck, source: 1, count: 1 }, cwd);
	const edited = await call(
		"ppt_set_text",
		{ path: deck, slide: 2, shape: name, text: "克隆页标题\n第二段" },
		cwd,
	);
	assert.equal(edited.details.paragraphs, 2);

	const inspected = await call("ppt_inspect", { path: deck }, cwd);
	assert.equal(inspected.details.slide_count, 2);
	const copyShape = inspected.details.slides[1].shapes.find((shape: any) => shape.name === name);
	assert.equal(copyShape.text, "克隆页标题\n第二段");
	assert.equal(copyShape.style.size_pt, 30);
	assert.equal(copyShape.style.font, "微软雅黑");
	const sourceShape = inspected.details.slides[0].shapes.find((shape: any) => shape.name === name);
	assert.equal(sourceShape.text, "原始标题");
});

test("create refuses overwrite and delete refuses emptying the deck", async () => {
	const cwd = await tempDir();
	const deck = "deck.pptx";
	await call("ppt_create", { path: deck }, cwd);
	await assert.rejects(() => call("ppt_create", { path: deck }, cwd), /已存在/);
	await assert.rejects(() => call("ppt_delete", { path: deck, slides: "1" }, cwd), /至少保留 1 页/);

	const doctor = await call("ppt_doctor", {}, cwd);
	assert.ok(doctor.details.python_pptx);
});

test("injects the pptx workflow only for PowerPoint prompts", async () => {
	const handler = events.find((event) => event.name === "before_agent_start")!.handler;

	const pptEvent = {
		type: "before_agent_start",
		messageType: "user",
		prompt: "用这个模板做一份 PPT",
		systemPromptOptions: { sections: {} as Record<string, string> },
	};
	await handler(pptEvent);
	assert.ok(pptEvent.systemPromptOptions.sections.pptx_workflow.includes("ppt_clone"));

	const chatEvent = {
		type: "before_agent_start",
		messageType: "user",
		prompt: "explain this TypeScript function",
		systemPromptOptions: { sections: {} as Record<string, string> },
	};
	await handler(chatEvent);
	assert.equal(chatEvent.systemPromptOptions.sections.pptx_workflow, undefined);
});

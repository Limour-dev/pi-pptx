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
		"ppt_add_shape",
		"ppt_add_image",
		"ppt_add_table",
		"ppt_set_shape",
		"ppt_delete_shape",
		"ppt_set_notes",
		"ppt_notes_check",
		"ppt_table_style",
		"ppt_apply",
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

test("apply batches edits atomically, supports dry run and reports failures", async () => {
	const cwd = await tempDir();
	const deck = "batch.pptx";
	await call("ppt_create", { path: deck }, cwd);
	const added = await call(
		"ppt_add_text",
		{ path: deck, slide: 1, text: "初始", x: 1, y: 1, w: 5, h: 1, size: 20 },
		cwd,
	);
	const id = added.details.id as number;
	const ops = [
		{ op: "set_text", slide: 1, shape: `#${id}`, text: "批量改完" },
		{ op: "add_text", slide: 1, x: 1, y: 3, w: 5, text: "新增一段", size: 16 },
		{ op: "set_notes", slide: 1, text: "讲者备注" },
	];

	const dry = await call("ppt_apply", { path: deck, ops, dryRun: true, shot: false }, cwd);
	assert.match(dry.content[0].text, /validated \(dry run\): 3\/3 ops/);
	let inspected = await call("ppt_inspect", { path: deck, slides: "1" }, cwd);
	assert.equal(
		inspected.details.slides[0].shapes.find((shape: any) => shape.id === id).text,
		"初始",
	);

	const applied = await call("ppt_apply", { path: deck, ops, shot: false }, cwd);
	assert.match(applied.content[0].text, /batch saved: 3\/3 ops/);
	inspected = await call("ppt_inspect", { path: deck, slides: "1" }, cwd);
	assert.equal(
		inspected.details.slides[0].shapes.find((shape: any) => shape.id === id).text,
		"批量改完",
	);
	assert.equal(inspected.details.slides[0].notes, "讲者备注");
	assert.ok(inspected.details.slides[0].shapes.some((shape: any) => shape.text === "新增一段"));

	const bad = await call(
		"ppt_apply",
		{ path: deck, ops: [{ op: "set_text", slide: 1, shape: "#999", text: "x" }], shot: false },
		cwd,
	);
	assert.match(bad.content[0].text, /NOT saved/);
	assert.match(bad.content[0].text, /1 op error/);
	inspected = await call("ppt_inspect", { path: deck, slides: "1" }, cwd);
	assert.equal(
		inspected.details.slides[0].shapes.find((shape: any) => shape.id === id).text,
		"批量改完",
	);
});

test("delete_shape, set_notes and table_style tools expose ids and results", async () => {
	const cwd = await tempDir();
	const deck = "tools.pptx";
	await call("ppt_create", { path: deck }, cwd);
	const added = await call(
		"ppt_add_text",
		{ path: deck, slide: 1, text: "待删除", x: 1, y: 1, w: 3, h: 1 },
		cwd,
	);
	const notes = await call("ppt_set_notes", { path: deck, slide: 1, text: "第一段\n第二段" }, cwd);
	assert.equal(notes.details.paragraphs, 2);
	const deleted = await call("ppt_delete_shape", { path: deck, slide: 1, shape: `#${added.details.id}` }, cwd);
	assert.equal(deleted.details.deleted, true);

	const table = await call(
		"ppt_add_table",
		{
			path: deck,
			slide: 1,
			rows: 2,
			cols: 2,
			x: 1,
			y: 3,
			w: 5,
			h: 1.5,
			data: [
				["指标", "值"],
				["EF", "52"],
			],
			headerFill: "#1F4E79",
			headerColor: "#FFFFFF",
			zebra: "#EAF1F8",
		},
		cwd,
	);
	assert.match(table.content[0].text, /模板里没有任何表格样例/);
	const styled = await call(
		"ppt_table_style",
		{ path: deck, slide: 1, shape: `#${table.details.id}`, border: "none", colWidths: [2, 3] },
		cwd,
	);
	assert.equal(styled.details.rows, 2);
	assert.equal(styled.details.cols, 2);

	const inspected = await call("ppt_inspect", { path: deck, slides: "1" }, cwd);
	const texts = inspected.details.slides[0].shapes.map((shape: any) => shape.text);
	assert.ok(!texts.includes("待删除"));
	assert.ok(inspected.details.slides[0].shapes.some((shape: any) => shape.id === table.details.id));
	assert.equal(inspected.details.slides[0].notes, "第一段\n第二段");
});

test("add_shape draws cards and chevrons with stable names", async () => {
	const cwd = await tempDir();
	const deck = "shapes.pptx";
	await call("ppt_create", { path: deck }, cwd);

	const card = await call(
		"ppt_add_shape",
		{
			path: deck,
			slide: 1,
			kind: "roundrect",
			x: 1,
			y: 1.5,
			w: 4,
			h: 1.4,
			radius: 0.06,
			fill: "#F4F7FB",
			line: "#CAD9E8",
			text: "卡片标题",
			size: 16,
			color: "#1F4E79",
			name: "Card 1",
		},
		cwd,
	);
	assert.equal(card.details.kind, "roundrect");
	assert.equal(card.details.name, "Card 1");
	assert.match(card.content[0].text, /added roundrect "Card 1"/);

	const chevron = await call(
		"ppt_add_shape",
		{ path: deck, slide: 1, kind: "chevron", x: 5.5, y: 1.5, w: 1.4, h: 0.9, fill: "#2E75B6" },
		cwd,
	);
	assert.equal(chevron.details.kind, "chevron");

	const bar = await call("ppt_add_shape", { path: deck, slide: 1, kind: "bar", x: 1, y: 1.4, w: 4 }, cwd);
	assert.ok(Math.abs(bar.details.box_in[3] - 0.06) < 0.01);

	const inspected = await call("ppt_inspect", { path: deck, slides: "1" }, cwd);
	const names = inspected.details.slides[0].shapes.map((shape: any) => shape.name);
	assert.ok(names.includes("Card 1"));

	// Shapes are addressable by id afterwards (e.g. for ignore= or deletion).
	const deleted = await call("ppt_delete_shape", { path: deck, slide: 1, shape: `#${chevron.details.id}` }, cwd);
	assert.equal(deleted.details.deleted, true);
});

test("shot supports onlyChanged, density and id/regex ignore", async () => {
	const cwd = await tempDir();
	const deck = "incremental.pptx";
	await call("ppt_create", { path: deck }, cwd);
	await call("ppt_clone", { path: deck, source: 1, count: 2 }, cwd);

	const first = await call("ppt_shot", { path: deck, slides: "1-3", dpi: 60, onlyChanged: true }, cwd);
	assert.deepEqual(first.details.slides.map((slide: any) => slide.slide), [1, 2, 3]);

	const second = await call("ppt_shot", { path: deck, dpi: 60, onlyChanged: true }, cwd);
	assert.deepEqual(second.details.slides, []);
	assert.match(second.content[0].text, /没有(任何)?页面变化|no_changes/);

	const longText = "这一段文字足够长以制造确定的溢出效果。".repeat(6);
	const flagged = await call(
		"ppt_add_text",
		{ path: deck, slide: 2, text: longText, x: 1, y: 1, w: 3, h: 0.4, size: 20, name: "Dense 1" },
		cwd,
	);
	const third = await call("ppt_shot", { path: deck, dpi: 60, onlyChanged: true, density: true }, cwd);
	assert.deepEqual(third.details.slides.map((slide: any) => slide.slide), [2]);
	assert.ok(third.details.density[0].capacity >= 1);

	const ignored = await call(
		"ppt_shot",
		{ path: deck, slides: "2", dpi: 60, overflow: "strict", ignore: [`#${flagged.details.id}`], density: true },
		cwd,
	);
	const overflowNames = (ignored.details.slides[0].overflow ?? []).map((item: any) => item.shape);
	assert.ok(!overflowNames.includes("Dense 1"));
	const stillFlagged = await call("ppt_shot", { path: deck, slides: "2", dpi: 60, overflow: "strict" }, cwd);
	assert.ok((stillFlagged.details.slides[0].overflow ?? []).length > 0);
});

test("set_shape moves a placeholder and changes z-order", async () => {
	const cwd = await tempDir();
	const deck = "layout.pptx";
	await call("ppt_create", { path: deck }, cwd);
	const added = await call(
		"ppt_add_text",
		{ path: deck, slide: 1, text: "底层图形", x: 1, y: 1, w: 3, h: 1, name: "Base 1" },
		cwd,
	);
	const moved = await call(
		"ppt_set_shape",
		{ path: deck, slide: 1, shape: `#${added.details.id}`, x: 2, y: 2, w: 5, h: 1.2, z: "front", name: "Moved 1" },
		cwd,
	);
	assert.deepEqual(moved.details.box_in, [2, 2, 5, 1.2]);
	assert.equal(moved.details.name, "Moved 1");
	assert.match(moved.content[0].text, /box \[/);
});

test("notes_check reports per-slide fingerprints", async () => {
	const cwd = await tempDir();
	const deck = "notes.pptx";
	await call("ppt_create", { path: deck }, cwd);
	await call("ppt_clone", { path: deck, source: 1, count: 1 }, cwd);
	await call("ppt_set_notes", { path: deck, slide: 1, text: "第一页备注：开场" }, cwd);
	await call("ppt_set_notes", { path: deck, slide: 2, text: "第二页备注：方法" }, cwd);

	const checked = await call("ppt_notes_check", { path: deck }, cwd);
	assert.equal(checked.details.notes.length, 2);
	assert.equal(checked.details.notes[0].first_20, "第一页备注：开场");
	assert.ok(checked.details.notes.every((row: any) => row.notes_len > 0));
});

test("shot montage returns one contact sheet", async () => {
	const cwd = await tempDir();
	const deck = "montage.pptx";
	await call("ppt_create", { path: deck }, cwd);
	await call("ppt_clone", { path: deck, source: 1, count: 2 }, cwd);
	const shot = await call("ppt_shot", { path: deck, slides: "1-3", dpi: 70, montage: 2 }, cwd);
	const images = shot.content.filter((block: any) => block.type === "image");
	assert.equal(images.length, 1);
	assert.equal(shot.details.montage.cols, 2);
	assert.equal(shot.details.slides.length, 3);
	assert.ok(shot.details.montage.px[0] > 600);
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

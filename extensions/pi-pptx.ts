/**
 * pi-pptx — author PowerPoint decks from existing templates.
 *
 * Tools wrap the Python backend in ../python/pptkit.py:
 *   python-pptx          read / clone / edit
 *   LibreOffice+PyMuPDF  render slides to PNG (+ heuristic overflow report)
 *
 * The rendered PNGs are returned as image content, so the model can actually
 * see the deck it is editing.
 */

import { execFile } from "node:child_process";
import { readFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Type } from "@earendil-works/pi-ai";
import {
	defineTool,
	withFileMutationQueue,
	type ExtensionAPI,
	type ExtensionToolContext,
} from "@earendil-works/pi-coding-agent";

const PYTHON = process.env.PI_PPTKIT_PYTHON || "python3";
const KIT =
	process.env.PI_PPTKIT_PY ||
	fileURLToPath(new URL("../python/pptkit.py", import.meta.url));

const MAX_SHOT_PAGES = 4;

type KitResult = Record<string, any>;

const PPT_PATH = Type.String({
	description: "PPTX file path. Relative paths resolve from the working directory.",
});

const SLIDE = Type.Number({
	description: "1-based slide number.",
	minimum: 1,
});

const OUT = Type.Optional(
	Type.String({
		description: "Write the result to this path instead of editing the input file in place.",
	}),
);

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

async function mutate(
	ctx: ExtensionToolContext,
	signal: AbortSignal | undefined,
	file: string,
	args: string[],
): Promise<KitResult> {
	const absolute = resolveFromCwd(ctx.cwd, file);
	return withFileMutationQueue(absolute, () => execKit(args, ctx.cwd, signal));
}

function jsonResult(data: KitResult, summary?: string) {
	const text = JSON.stringify(data);
	return {
		content: [{ type: "text" as const, text: summary ? `${summary}\n${text}` : text }],
		details: data,
	};
}

function overflowSummary(data: KitResult): string {
	const flagged = (data.shots ?? []).filter((shot: any) => shot.overflow?.length);
	if (!flagged.length) return "";
	const parts = flagged.map((shot: any) => {
		const shapes = shot.overflow.map((item: any) => item.shape).join(", ");
		return `slide ${shot.slide}: ${shapes}`;
	});
	return `WARNING text overflow (heuristic): ${parts.join("; ")}. Fix by resizing the shape, reducing text, or lowering the font size, then ppt_shot again.`;
}

const FONT_DESCRIPTION =
	"Font name, e.g. '微软雅黑'. Omit to inherit the presentation theme font.";

// ---------------------------------------------------------------- read tools

const inspectTool = defineTool({
	name: "ppt_inspect",
	label: "PPT Inspect",
	promptSnippet: "Inspect PPTX layouts, shapes (names, paths, inch boxes, text), tables and notes.",
	description: `Read a PPTX: slide size/count, every layout (index, name, placeholders), and for the requested slides the full shape tree with names, recursive paths ("组合 7/矩形 9"), inch boxes, text, tables, pictures and notes.

Call this before editing: use the returned shape paths verbatim and 1-based slide numbers. Pass \`slides\` (e.g. "1,3" or "2-4") to inspect only some pages; otherwise all pages are returned.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slides: Type.Optional(
			Type.String({ description: "Slide selection, e.g. '1' or '1,3' or '2-4'. Default: all." }),
		),
	}),
	annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["inspect", params.path];
		if (params.slides) args.push("--slides", params.slides);
		const data = await execKit(args, ctx.cwd, signal);
		const summary = `${data.slide_count} slides, ${data.slides?.length ?? 0} inspected, ${data.layouts?.length ?? 0} layouts, canvas ${data.slide_size_in?.[0]}x${data.slide_size_in?.[1]} in`;
		return jsonResult(data, summary);
	},
});

const shotTool = defineTool({
	name: "ppt_shot",
	label: "PPT Screenshot",
	promptSnippet: "Render PPTX slides to PNG images (with overflow warnings) so you can see the result.",
	description: `Render slides of a PPTX to PNG and return the images plus a JSON report. This is how you SEE a deck: use it to pick a style parent from a template's sample pages and to self-check every edit (overflow, wrapping, overlap, font size).

- Returns at most ${MAX_SHOT_PAGES} images per call. If \`slides\` is omitted it renders from slide 1 and reports the rest in \`omitted_slides\`; call again with explicit slide numbers for more.
- \`dpi\`: 80 for an overview, 110 (default) for normal checking, 150 for fine detail.
- Report fields: \`shots[].px\` size, \`shots[].overflow\` (heuristic: rendered text bbox exceeds the shape box; fix and re-shoot), \`pdf\` cached PDF path.
- Rendering is cached by file (path, mtime, size): re-shooting an unchanged deck is fast.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slides: Type.Optional(
			Type.String({ description: "Slides to render, e.g. '2' or '1,3'. Default: from slide 1." }),
		),
		dpi: Type.Optional(Type.Number({ description: "Raster DPI. Default 110.", minimum: 40, maximum: 300 })),
	}),
	annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["shot", params.path, "--dpi", String(params.dpi ?? 110), "--max-pages", String(MAX_SHOT_PAGES)];
		if (params.slides) args.push("--slides", params.slides);
		const data = await execKit(args, ctx.cwd, signal);

		const content: any[] = [];
		const warning = overflowSummary(data);
		const report = {
			pdf: data.pdf,
			dpi: data.dpi,
			shots: data.shots,
			...(data.omitted_slides ? { omitted_slides: data.omitted_slides, hint: data.hint } : {}),
		};
		content.push({
			type: "text" as const,
			text: warning ? `${warning}\n${JSON.stringify(report)}` : JSON.stringify(report),
		});
		for (const shot of data.shots ?? []) {
			content.push({
				type: "image" as const,
				data: (await readFile(shot.file)).toString("base64"),
				mimeType: "image/png",
			});
		}
		return {
			content,
			details: {
				pdf: data.pdf,
				dpi: data.dpi,
				slides: (data.shots ?? []).map((shot: any) => ({
					slide: shot.slide,
					file: shot.file,
					px: shot.px,
					overflow: shot.overflow?.map((item: any) => item.shape) ?? [],
				})),
				omitted_slides: data.omitted_slides,
			},
		};
	},
});

const doctorTool = defineTool({
	name: "ppt_doctor",
	label: "PPT Doctor",
	promptSnippet: "Check the PPTX backend: python-pptx, PyMuPDF, LibreOffice and CJK font substitution.",
	description:
		"Check the PPTX backend environment: Python, python-pptx, PyMuPDF, LibreOffice and how CJK template fonts resolve. Use it when ppt_shot/ppt_inspect fail or when text metrics look off.",
	parameters: Type.Object({}),
	annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
	async execute(_id, _params, signal, _onUpdate, ctx) {
		const data = await execKit(["doctor"], ctx.cwd, signal);
		return jsonResult(data, `python-pptx ${data.python_pptx}, PyMuPDF ${data.pymupdf}, ${data.soffice_version || "no soffice"}`);
	},
});

// ------------------------------------------------------------ authoring tools

const createTool = defineTool({
	name: "ppt_create",
	label: "PPT Create",
	promptSnippet: "Create a working PPTX, optionally as a copy of a template deck.",
	description: `Create a working .pptx.

- With \`template\`: copies that deck byte-for-byte (masters, layouts, sample pages, media all preserved). This is the recommended start: create a working copy and never edit the source template in place.
- Without \`template\`: blank 16:9 deck (13.333 x 7.5 in) with a title slide unless \`empty\` is true.
- Refuses to overwrite an existing file unless \`overwrite\` is true.
- \`width\`/\`height\` (inches) resize the canvas and may distort a template.`,
	parameters: Type.Object({
		path: Type.String({ description: "Output PPTX path to create." }),
		template: Type.Optional(Type.String({ description: "Source template deck to copy." })),
		width: Type.Optional(Type.Number({ description: "Canvas width in inches (default 13.333)." })),
		height: Type.Optional(Type.Number({ description: "Canvas height in inches (default 7.5)." })),
		empty: Type.Optional(Type.Boolean({ description: "For a blank deck, skip the generated title slide." })),
		overwrite: Type.Optional(Type.Boolean({ description: "Overwrite an existing file. Default false." })),
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["create", params.path];
		if (params.template) args.push("--template", params.template);
		if (params.width != null) args.push("--width", String(params.width));
		if (params.height != null) args.push("--height", String(params.height));
		if (params.empty) args.push("--empty");
		if (params.overwrite) args.push("--overwrite");
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `created ${data.file} (${data.slide_count} slides, ${data.slide_size_in?.[0]}x${data.slide_size_in?.[1]} in)`);
	},
});

const cloneTool = defineTool({
	name: "ppt_clone",
	label: "PPT Clone Slide",
	promptSnippet: "Duplicate an existing slide, preserving groups, images, charts, background and notes.",
	description: `Duplicate an existing slide with everything intact: shapes, groups, pictures, charts, hyperlinks, background and notes (relationships are re-mapped, which python-pptx's add_slide does not do).

This is the main authoring path. Real templates keep their design in the master and grouped shapes, so clone a sample page and replace its text instead of building pages from layouts. Clone the same source page once per new page, then edit each copy.

- \`source\`: 1-based source slide number.
- \`at\`: 1-based position for the first copy (default: append).
- \`count\`: number of copies (default 1).
- \`out\`: write to another file instead of editing in place.`,
	parameters: Type.Object({
		path: PPT_PATH,
		source: SLIDE,
		at: Type.Optional(Type.Number({ description: "1-based insert position. Default: append.", minimum: 1 })),
		count: Type.Optional(Type.Number({ description: "Number of copies. Default 1.", minimum: 1, maximum: 50 })),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["clone", params.path, String(params.source)];
		if (params.at != null) args.push("--at", String(params.at));
		if (params.count != null) args.push("--count", String(params.count));
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `cloned slide ${params.source} ${params.count ?? 1}x -> ${data.slide_count} slides total`);
	},
});

const addSlideTool = defineTool({
	name: "ppt_add_slide",
	label: "PPT Add Slide",
	promptSnippet: "Add an empty slide from a layout (prefer ppt_clone for designed templates).",
	description: `Add an empty slide from a layout. \`layout\` accepts the layout index or its exact name from ppt_inspect (e.g. "3_版本内页"). The layout's placeholders are created empty; fill them with ppt_set_text.

For designed templates prefer ppt_clone of a sample page: layouts often contain only a skeleton.`,
	parameters: Type.Object({
		path: PPT_PATH,
		layout: Type.Union([Type.Number(), Type.String()], {
			description: "Layout index (0-based) or exact layout name.",
		}),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["add-slide", params.path, "--layout", String(params.layout)];
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `added slide ${data.added_slide} from layout "${data.layout}" (placeholders: ${(data.placeholders ?? []).join(", ") || "none"})`);
	},
});

const setTextTool = defineTool({
	name: "ppt_set_text",
	label: "PPT Set Text",
	promptSnippet: "Replace a shape's text while keeping its font, size, color and position.",
	description: `Replace the text of one shape while preserving the first run's font/size/color and the shape's position. Do not rebuild boxes: edit the template's own shapes.

- \`shape\`: exact name ("标题 1"), recursive path ("组合 7/矩形 9"), or top-level index ("3") from ppt_inspect.
- \`text\`: use real newlines to create paragraphs; each new paragraph inherits the first paragraph's style.
- Table cells: use ppt_set_cell instead.
- Check the capacity of the original text: a title that fit 1 line may overflow at 2 lines. ppt_shot after editing.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		shape: Type.String({ description: "Shape name, recursive path, or top-level index." }),
		text: Type.String({ description: "New text; \\n separates paragraphs." }),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["set-text", params.path, "--slide", String(params.slide), "--shape", params.shape, "--text", params.text];
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `slide ${params.slide} "${data.path}" now has ${data.paragraphs} paragraph(s)`);
	},
});

const setCellTool = defineTool({
	name: "ppt_set_cell",
	label: "PPT Set Table Cell",
	promptSnippet: "Set one table cell's text (row/col are 0-based).",
	description: `Set one table cell's text, preserving the cell's run formatting. \`row\` and \`col\` are 0-based and match the \`table\` array returned by ppt_inspect. \`shape\` is the table shape's name or path.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		shape: Type.String({ description: "Table shape name or path." }),
		row: Type.Number({ description: "0-based row index.", minimum: 0 }),
		col: Type.Number({ description: "0-based column index.", minimum: 0 }),
		text: Type.String(),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = [
			"set-cell",
			params.path,
			"--slide",
			String(params.slide),
			"--shape",
			params.shape,
			"--row",
			String(params.row),
			"--col",
			String(params.col),
			"--text",
			params.text,
		];
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `slide ${params.slide} ${params.shape}[${params.row}][${params.col}] updated`);
	},
});

const addTextTool = defineTool({
	name: "ppt_add_text",
	label: "PPT Add Text",
	promptSnippet: "Add a text box at inch coordinates on the slide canvas.",
	description: `Add a new text box at inch coordinates (origin = top-left; default canvas 13.333 x 7.5 in). Use this for labels that do not exist in the template.

New text does not inherit the template's text styling. Pass \`font\`, \`size\` and \`color\` to match nearby text, or better, clone a sample page that already contains a similar box and edit it. \`text\` newlines create paragraphs. Word wrap is on and auto-size is off, so text can overflow — verify with ppt_shot.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		text: Type.String({ description: "Text; \\n separates paragraphs." }),
		x: Type.Number({ description: "Left, inches." }),
		y: Type.Number({ description: "Top, inches." }),
		w: Type.Optional(Type.Number({ description: "Width, inches. Default 6." })),
		h: Type.Optional(Type.Number({ description: "Height, inches. Default 0.8." })),
		size: Type.Optional(Type.Number({ description: "Font size in points." })),
		font: Type.Optional(Type.String({ description: FONT_DESCRIPTION })),
		color: Type.Optional(Type.String({ description: "Hex color, e.g. #1F4E79." })),
		bold: Type.Optional(Type.Boolean()),
		align: Type.Optional(
			Type.Union([Type.Literal("left"), Type.Literal("center"), Type.Literal("right"), Type.Literal("justify")]),
		),
		valign: Type.Optional(Type.Union([Type.Literal("top"), Type.Literal("middle"), Type.Literal("bottom")])),
		lineSpacing: Type.Optional(Type.Number({ description: "Line spacing multiple, e.g. 1.2." })),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = [
			"add-text",
			params.path,
			"--slide",
			String(params.slide),
			"--text",
			params.text,
			"--x",
			String(params.x),
			"--y",
			String(params.y),
			"--w",
			String(params.w ?? 6),
			"--h",
			String(params.h ?? 0.8),
		];
		if (params.size != null) args.push("--size", String(params.size));
		if (params.font) args.push("--font", params.font);
		if (params.color) args.push("--color", params.color);
		if (params.bold != null) args.push(params.bold ? "--bold" : "--no-bold");
		if (params.align) args.push("--align", params.align);
		if (params.valign) args.push("--valign", params.valign);
		if (params.lineSpacing != null) args.push("--line-spacing", String(params.lineSpacing));
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `added ${data.name} on slide ${params.slide} at [${data.box_in?.join(", ")}] in`);
	},
});

const addImageTool = defineTool({
	name: "ppt_add_image",
	label: "PPT Add Image",
	promptSnippet: "Insert a picture at inch coordinates.",
	description: `Insert a picture. Coordinates and sizes are inches. Give \`w\` and/or \`h\` to scale (aspect ratio is preserved when only one is given); omit both for the image's native size. Without x/y the image is placed at 0,0 — usually pass both.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		image: Type.String({ description: "Image file (PNG, JPEG, GIF, SVG where supported)." }),
		x: Type.Optional(Type.Number({ description: "Left, inches. Default 0." })),
		y: Type.Optional(Type.Number({ description: "Top, inches. Default 0." })),
		w: Type.Optional(Type.Number({ description: "Width, inches." })),
		h: Type.Optional(Type.Number({ description: "Height, inches." })),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["add-image", params.path, "--slide", String(params.slide), "--path", params.image];
		if (params.x != null) args.push("--x", String(params.x));
		if (params.y != null) args.push("--y", String(params.y));
		if (params.w != null) args.push("--w", String(params.w));
		if (params.h != null) args.push("--h", String(params.h));
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `added ${data.name} on slide ${params.slide} at [${data.box_in?.join(", ")}] in`);
	},
});

const addTableTool = defineTool({
	name: "ppt_add_table",
	label: "PPT Add Table",
	promptSnippet: "Insert a table (prefer cloning a slide with a table to keep template styling).",
	description: `Insert a table at inch coordinates with optional 2D \`data\` (array of rows, cells are strings).

New tables use python-pptx's default table style and will usually clash with a designed template. Prefer cloning a slide that already contains a table, then use ppt_set_cell. The header row is bold by default.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		rows: Type.Number({ minimum: 1 }),
		cols: Type.Number({ minimum: 1 }),
		x: Type.Number({ description: "Left, inches." }),
		y: Type.Number({ description: "Top, inches." }),
		w: Type.Number({ description: "Width, inches." }),
		h: Type.Number({ description: "Height, inches." }),
		data: Type.Optional(
			Type.Array(Type.Array(Type.Union([Type.String(), Type.Number(), Type.Null()])), {
				description: "Cell values, row-major. Missing cells stay empty.",
			}),
		),
		size: Type.Optional(Type.Number({ description: "Cell font size in points." })),
		header: Type.Optional(Type.Boolean({ description: "Bold the header row. Default true." })),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = [
			"add-table",
			params.path,
			"--slide",
			String(params.slide),
			"--rows",
			String(params.rows),
			"--cols",
			String(params.cols),
			"--x",
			String(params.x),
			"--y",
			String(params.y),
			"--w",
			String(params.w),
			"--h",
			String(params.h),
		];
		if (params.data) args.push("--data", JSON.stringify(params.data));
		if (params.size != null) args.push("--size", String(params.size));
		if (params.header === false) args.push("--no-header");
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `added ${data.name} (${params.rows}x${params.cols}) on slide ${params.slide}`);
	},
});

const deleteTool = defineTool({
	name: "ppt_delete",
	label: "PPT Delete Slides",
	promptSnippet: "Delete slides by 1-based numbers, e.g. '2,5-7'.",
	description: `Delete slides by 1-based numbers or ranges ("2", "1,3", "5-7"). At least one slide must remain. Use it to trim a template's sample pages.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slides: Type.String({ description: "Slides to delete, e.g. '2,5-7'." }),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["delete", params.path, "--slides", params.slides];
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `deleted slides ${(data.deleted ?? []).join(", ")}; ${data.slide_count} remain`);
	},
});

// ------------------------------------------------------------- prompt + cmd

const PPT_INTENT =
	/(pptx?|powerpoint|slide|deck|presentation|幻灯片|演示文稿|汇报|简报|课件|画册)/i;

const WORKFLOW = `The pi-pptx tools (ppt_*) author PowerPoint decks from existing templates. Follow this loop:
1. ppt_inspect the source deck: layouts, sample slides, exact shape paths and inch boxes.
2. ppt_shot 1-4 sample pages and actually look at them; pick the page whose style matches each needed layout.
3. ppt_create a working copy with template=<source deck>. Never edit the source template in place.
4. ppt_clone that sample page once per new page (this preserves groups, pictures, charts and background). Then replace content with ppt_set_text / ppt_set_cell / ppt_add_text / ppt_add_image / ppt_add_table.
5. ppt_shot every edited slide after each edit. Check overflow[], wrapping, overlaps and font sizes, then fix and re-shoot. Do not finish while overflow warnings are unresolved.
Numbers: slides and insert positions are 1-based; table row/col are 0-based. The canvas is 13.33x7.5 in unless ppt_inspect says otherwise. Clone a slide that already has a table/chart instead of adding a default-styled one. If ppt_shot fails, run ppt_doctor.`;

export default function (pi: ExtensionAPI) {
	for (const tool of [
		inspectTool,
		shotTool,
		doctorTool,
		createTool,
		cloneTool,
		addSlideTool,
		setTextTool,
		setCellTool,
		addTextTool,
		addImageTool,
		addTableTool,
		deleteTool,
	]) {
		pi.registerTool(tool);
	}

	// Only pay the prompt tokens when the run looks like PowerPoint work.
	pi.on("before_agent_start", async (event) => {
		if (event.messageType !== "user") return;
		if (!PPT_INTENT.test(event.prompt)) return;
		event.systemPromptOptions.sections = {
			...event.systemPromptOptions.sections,
			pptx_workflow: WORKFLOW,
		};
	});

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

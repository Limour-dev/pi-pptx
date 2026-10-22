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
const MAX_MONTAGE_PAGES = 12;

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

const PARAS = Type.Optional(
	Type.Array(Type.Record(Type.String(), Type.Any()), {
		description:
			"Rich paragraphs: [{runs:[{t,size,bold,italic,color,font}], align, bullet, bullet_color, indent, hanging, space_after, space_before, line_spacing, level}]. Unspecified run fields inherit the shape's first run.",
	}),
);

const APPLY_OP = Type.Record(Type.String(), Type.Any(), {
	description:
		"One op object. Required: op. Common: slide (1-based), shape (name/path/index/#id). See the tool description for per-op fields.",
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

function shotWarnings(data: KitResult): string {
	const warnings: string[] = [];
	const overflow = (data.shots ?? []).filter((shot: any) => shot.overflow?.length);
	if (overflow.length) {
		const parts = overflow.map((shot: any) => {
			const shapes = shot.overflow.map((item: any) => item.shape).join(", ");
			return `slide ${shot.slide}: ${shapes}`;
		});
		warnings.push(
			`WARNING text overflow (heuristic): ${parts.join("; ")}. Fix by resizing the shape, reducing text, or lowering the font size, then ppt_shot again.`,
		);
	}
	const overlaps = (data.shots ?? []).filter((shot: any) => shot.overlaps?.length);
	if (overlaps.length) {
		const parts = overlaps.map((shot: any) => {
			const pairs = shot.overlaps.map((item: any) => `${item.a} x ${item.b}`).join(", ");
			return `slide ${shot.slide}: ${pairs}`;
		});
		warnings.push(
			`WARNING shape overlap (heuristic): ${parts.join("; ")}. Check these boxes are intentional; move or resize if not.`,
		);
	}
	return warnings.join("\n");
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
		mode: Type.Optional(
			Type.Union([Type.Literal("full"), Type.Literal("summary")], {
				description: "full (default) = shape tree; summary = per-slide title/count/table/picture/notes overview.",
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
	promptSnippet: "Render PPTX slides to PNG images (with overflow warnings) so you can see the result.",
	description: `Render slides of a PPTX to PNG and return the images plus a JSON report. This is how you SEE a deck: use it to pick a style parent from a template's sample pages and to self-check every edit (overflow, wrapping, overlap, font size).

- Returns at most ${MAX_SHOT_PAGES} images per call (${MAX_MONTAGE_PAGES} when \`montage\` is set). If \`slides\` is omitted it renders from slide 1 and reports the rest in \`omitted_slides\`; call again with explicit slide numbers for more.
- \`dpi\`: 80 for an overview, 110 (default) for normal checking, 150 for fine detail.
- \`montage\`: one contact sheet (e.g. 3 columns) instead of N separate images — best for reviewing a whole deck's rhythm.
- \`region\`: crop [x,y,w,h] in inches (e.g. a small table) without re-rendering the whole slide.
- \`overflow\`: "loose" (default; skips left/right checks on centred text), "strict" (report every side), "off". \`ignore\` skips named shapes.
- Report fields: \`shots[].overflow\` (rendered text exceeds the shape box), \`shots[].overlaps\` (text collides with another text/picture/table), \`cache_hit\`, \`pdf\`.
- Rendering is cached by content hash: re-shooting an unchanged deck is fast. CJK font substitution can differ ~5% from PowerPoint — leave a little margin.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slides: Type.Optional(
			Type.String({ description: "Slides to render, e.g. '2' or '1,3'. Default: from slide 1." }),
		),
		dpi: Type.Optional(Type.Number({ description: "Raster DPI. Default 110.", minimum: 40, maximum: 300 })),
		montage: Type.Optional(
			Type.Number({ description: "Output a contact sheet with this many columns. Default: off.", minimum: 1, maximum: 6 }),
		),
		region: Type.Optional(
			Type.Array(Type.Number(), {
				description: "Crop each rendered slide to [x,y,w,h] in inches.",
				minItems: 4,
				maxItems: 4,
			}),
		),
		overflow: Type.Optional(
			Type.Union([Type.Literal("loose"), Type.Literal("strict"), Type.Literal("off")], {
				description: "Overflow heuristic mode. Default loose.",
			}),
		),
		ignore: Type.Optional(
			Type.Array(Type.String(), { description: "Shape names to skip in overflow/overlap checks." }),
		),
	}),
	annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const limit = params.montage ? MAX_MONTAGE_PAGES : MAX_SHOT_PAGES;
		const args = ["shot", params.path, "--dpi", String(params.dpi ?? 110), "--max-pages", String(limit)];
		if (params.slides) args.push("--slides", params.slides);
		if (params.montage != null) args.push("--montage", String(params.montage));
		if (params.region) args.push("--region", params.region.join(","));
		if (params.overflow) args.push("--overflow", params.overflow);
		if (params.ignore?.length) args.push("--ignore", params.ignore.join(","));
		const data = await execKit(args, ctx.cwd, signal);

		const content: any[] = [];
		const warning = shotWarnings(data);
		const report = {
			pdf: data.pdf,
			dpi: data.dpi,
			cache_hit: data.cache_hit,
			overflow_mode: data.overflow_mode,
			shots: data.shots,
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
				montage: data.montage,
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
	description: `Add an empty slide from a layout. \`layout\` accepts the layout index or its exact name from ppt_inspect (e.g. "3_版本内页"). The layout's placeholders are created empty; fill them with ppt_set_text, or delete the ones you do not need with ppt_delete_shape.

For designed templates prefer ppt_clone of a sample page: layouts often contain only a skeleton (and new placeholder names may differ from the sample page's names, e.g. "Title 1" vs "标题 1"). The result lists every new shape with its id, name, type and box so you can address them by #id.`,
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
		const shapes = (data.shapes ?? [])
			.map((shape: any) => `#${shape.id} ${shape.name} (${shape.type})`)
			.join(", ");
		return jsonResult(
			data,
			`added slide ${data.added_slide} from layout "${data.layout}" (placeholders: ${(data.placeholders ?? []).join(", ") || "none"}; shapes: ${shapes || "none"})`,
		);
	},
});

const setTextTool = defineTool({
	name: "ppt_set_text",
	label: "PPT Set Text",
	promptSnippet: "Replace a shape's text (or rich runs) while keeping its font, size, color and position.",
	description: `Replace the text of one shape while preserving the first run's font/size/color and the shape's position. Do not rebuild boxes: edit the template's own shapes.

- \`shape\`: \`#<id>\` is the most reliable form (ppt_inspect returns ids). Names are accepted only when unique; duplicate names are refused with a candidate table. Recursive paths ("组合 7/矩形 9") and 0-based top-level indexes also work.
- \`text\`: real newlines create paragraphs; each inherits the first paragraph's style. \`text: ""\` clears the text but KEEPS the shape (empty placeholders still draw their bullet) — use ppt_delete_shape to remove it.
- \`paras\`: rich alternative for mixed formatting, e.g. [{"runs":[{"t":"一句实话：","bold":true,"color":"C00000"},{"t":"没有事前效能计算。"}]}]. Unspecified run fields inherit the shape's first run.
- Table cells: use ppt_set_cell instead.
- Check the capacity of the original text: a title that fit 1 line may overflow at 2 lines. ppt_shot after editing.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		shape: Type.String({ description: "Shape #id (preferred), unique name, recursive path, or 0-based top-level index." }),
		text: Type.Optional(Type.String({ description: "New text; \\n separates paragraphs. Empty string keeps the shape." })),
		paras: PARAS,
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["set-text", params.path, "--slide", String(params.slide), "--shape", params.shape];
		if (params.paras) args.push("--paras", JSON.stringify(params.paras));
		else if (params.text != null) args.push("--text", params.text);
		else throw new Error("ppt_set_text needs either text or paras");
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
	promptSnippet: "Add a text box at inch coordinates with optional bullets, spacing and autofit.",
	description: `Add a new text box at inch coordinates (origin = top-left; default canvas 13.333 x 7.5 in). Use this for labels that do not exist in the template.

New text does not inherit the template's text styling. Pass \`font\`, \`size\` and \`color\` to match nearby text, or better, clone a sample page that already contains a similar box and edit it.
- \`h\` omitted => the height is estimated from the text and \`fit: "grow"\` (spAutoFit) is enabled, so the box grows instead of clipping. Pass \`fit: "shrink"\` to shrink text into a fixed box, or \`"none"\` for a fixed box.
- Layout: \`spaceAfter\`/\`spaceBefore\` (pt), \`lineSpacing\` (multiple), \`indent\` (inches; negative = hanging), \`hanging\`, \`bullet\` ("•" or true), \`bulletColor\`.
- \`paras\` supports run-level mixed formatting like ppt_set_text. \`font\` writes latin+ea+cs, so CJK actually uses the requested face.
- Always ppt_shot after adding: the overflow heuristic is a signal, not a guarantee.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		text: Type.Optional(Type.String({ description: "Text; \\n separates paragraphs. Omit when using paras." })),
		paras: PARAS,
		x: Type.Number({ description: "Left, inches." }),
		y: Type.Number({ description: "Top, inches." }),
		w: Type.Optional(Type.Number({ description: "Width, inches. Default 6." })),
		h: Type.Optional(Type.Number({ description: "Height, inches. Omit to auto-grow from the text." })),
		size: Type.Optional(Type.Number({ description: "Font size in points." })),
		font: Type.Optional(Type.String({ description: FONT_DESCRIPTION })),
		color: Type.Optional(Type.String({ description: "Hex color, e.g. #1F4E79." })),
		bold: Type.Optional(Type.Boolean()),
		align: Type.Optional(
			Type.Union([Type.Literal("left"), Type.Literal("center"), Type.Literal("right"), Type.Literal("justify")]),
		),
		valign: Type.Optional(Type.Union([Type.Literal("top"), Type.Literal("middle"), Type.Literal("bottom")])),
		lineSpacing: Type.Optional(Type.Number({ description: "Line spacing multiple, e.g. 1.2." })),
		spaceAfter: Type.Optional(Type.Number({ description: "Paragraph space after, points." })),
		spaceBefore: Type.Optional(Type.Number({ description: "Paragraph space before, points." })),
		indent: Type.Optional(Type.Number({ description: "Left indent in inches; negative means hanging." })),
		hanging: Type.Optional(Type.Number({ description: "First-line hanging indent in inches." })),
		bullet: Type.Optional(
			Type.Union([Type.String(), Type.Boolean()], { description: "Bullet glyph ('•'), true for '•', false/none to disable." }),
		),
		bulletColor: Type.Optional(Type.String({ description: "Bullet color, e.g. #2E75B6." })),
		fit: Type.Optional(
			Type.Union([Type.Literal("grow"), Type.Literal("shrink"), Type.Literal("none")], {
				description: "Autofit mode. Default: grow when h is omitted, otherwise none.",
			}),
		),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = [
			"add-text",
			params.path,
			"--slide",
			String(params.slide),
			"--x",
			String(params.x),
			"--y",
			String(params.y),
			"--w",
			String(params.w ?? 6),
		];
		if (params.paras) args.push("--paras", JSON.stringify(params.paras));
		else if (params.text != null) args.push("--text", params.text);
		else throw new Error("ppt_add_text needs either text or paras");
		if (params.h != null) args.push("--h", String(params.h));
		if (params.size != null) args.push("--size", String(params.size));
		if (params.font) args.push("--font", params.font);
		if (params.color) args.push("--color", params.color);
		if (params.bold != null) args.push(params.bold ? "--bold" : "--no-bold");
		if (params.align) args.push("--align", params.align);
		if (params.valign) args.push("--valign", params.valign);
		if (params.lineSpacing != null) args.push("--line-spacing", String(params.lineSpacing));
		if (params.spaceAfter != null) args.push("--space-after", String(params.spaceAfter));
		if (params.spaceBefore != null) args.push("--space-before", String(params.spaceBefore));
		if (params.indent != null) args.push("--indent", String(params.indent));
		if (params.hanging != null) args.push("--hanging", String(params.hanging));
		if (params.bullet != null) {
			args.push("--bullet", params.bullet === true ? "•" : params.bullet === false ? "none" : params.bullet);
		}
		if (params.bulletColor) args.push("--bullet-color", params.bulletColor);
		if (params.fit) args.push("--fit", params.fit);
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(
			data,
			`added ${data.name} (#${data.id}) on slide ${params.slide} at [${data.box_in?.join(", ")}] in, fit=${data.fit}`,
		);
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
	promptSnippet: "Insert a table (theme style off by default; prefer cloning a page that has a table).",
	description: `Insert a table at inch coordinates with optional 2D \`data\` (array of rows, cells are strings).

By default the theme's table style is DISABLED (no firstRow/bandRow, no tableStyleId) and a neutral grey border is applied, so a new table does not fight a designed template. \`themeStyle: true\` restores the theme look. Style it further with ppt_table_style.
- The header row is bold by default. \`headerFill\`/\`headerColor\`/\`zebra\`/\`bodyColor\` take hex colors.
- Prefer cloning a slide that already contains a table when the template has one; ppt_inspect with mode="summary" tells you (has_table). The tool returns a hint when the deck has no table sample.`,
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
		themeStyle: Type.Optional(Type.Boolean({ description: "Use the theme table style. Default false." })),
		headerFill: Type.Optional(Type.String({ description: "Header row fill, e.g. #1F4E79." })),
		headerColor: Type.Optional(Type.String({ description: "Header text color, e.g. #FFFFFF." })),
		bodyColor: Type.Optional(Type.String({ description: "Body text color." })),
		zebra: Type.Optional(Type.String({ description: "Even data-row fill, e.g. #EAF1F8." })),
		border: Type.Optional(Type.String({ description: "Border color, or 'none'. Default #D9D9D9." })),
		colWidths: Type.Optional(Type.Array(Type.Number(), { description: "Column widths in inches." })),
		rowHeight: Type.Optional(Type.Number({ description: "Row height in inches." })),
		aligns: Type.Optional(
			Type.Array(Type.String(), { description: "Per-column alignment, cycling: left|center|right|justify." }),
		),
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
		if (params.themeStyle) args.push("--theme-style");
		if (params.headerFill) args.push("--header-fill", params.headerFill);
		if (params.headerColor) args.push("--header-color", params.headerColor);
		if (params.bodyColor) args.push("--body-color", params.bodyColor);
		if (params.zebra) args.push("--zebra", params.zebra);
		if (params.border) args.push("--border", params.border);
		if (params.colWidths?.length) args.push("--col-widths", params.colWidths.join(","));
		if (params.rowHeight != null) args.push("--row-height", String(params.rowHeight));
		if (params.aligns?.length) args.push("--aligns", params.aligns.join(","));
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		const hint = data.hint ? ` ${data.hint}` : "";
		return jsonResult(
			data,
			`added ${data.name} (#${data.id}, ${params.rows}x${params.cols}) on slide ${params.slide}${hint}`,
		);
	},
});

const deleteShapeTool = defineTool({
	name: "ppt_delete_shape",
	label: "PPT Delete Shape",
	promptSnippet: "Delete a shape (e.g. an empty placeholder whose bullet still renders).",
	description: `Delete one shape from a slide. Use this instead of setting text to a space: an empty placeholder still renders its bullet glyph, and template sample text boxes cannot be removed by clearing their text.

\`shape\` accepts \`#<id>\` (preferred), a unique name, a recursive path, or a 0-based top-level index. Group children can be deleted with their path.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		shape: Type.String({ description: "Shape #id (preferred), unique name, path, or 0-based top-level index." }),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["delete-shape", params.path, "--slide", String(params.slide), "--shape", params.shape];
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `deleted "${data.path}" (#${data.id}) from slide ${params.slide}`);
	},
});

const setNotesTool = defineTool({
	name: "ppt_set_notes",
	label: "PPT Set Notes",
	promptSnippet: "Write (or clear) a slide's speaker notes.",
	description: `Set the speaker notes of one slide. Notes are part of the deck, so write them before the final review; ppt_clone already copies notes from the source page. Pass \`text: ""\` to clear.`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		text: Type.String({ description: "Notes text; newlines create paragraphs. Empty clears." }),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = ["set-notes", params.path, "--slide", String(params.slide), "--text", params.text];
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `slide ${params.slide} notes: ${data.paragraphs} paragraph(s), ${data.chars} chars`);
	},
});

const tableStyleTool = defineTool({
	name: "ppt_table_style",
	label: "PPT Table Style",
	promptSnippet: "Style a table: header fill/color, zebra rows, borders, widths, alignment.",
	description: `Style an existing table (usually created by ppt_add_table or cloned from a template). Colors are hex; all fields are optional.

- themeStyle: true restores the theme's firstRow/bandRow variants; false removes the theme table style for manual styling.
- headerFill / headerColor / bodyColor / zebra paint rows; border sets all cell borders ("none" removes them).
- colWidths (inches per column), rowHeight, aligns (per column, cycling), margin (cell padding), valign.
- This edits the table in place; it does not change cell text (use ppt_set_cell for that).`,
	parameters: Type.Object({
		path: PPT_PATH,
		slide: SLIDE,
		shape: Type.String({ description: "Table shape #id (preferred), name, path, or index." }),
		themeStyle: Type.Optional(Type.Boolean({ description: "Re-enable the theme table style. Default: leave unchanged." })),
		header: Type.Optional(Type.Boolean({ description: "Treat row 0 as a header. Default true." })),
		headerFill: Type.Optional(Type.String({ description: "Header fill, e.g. #1F4E79." })),
		headerColor: Type.Optional(Type.String({ description: "Header text color, e.g. #FFFFFF." })),
		bodyColor: Type.Optional(Type.String({ description: "Body text color." })),
		zebra: Type.Optional(Type.String({ description: "Even data-row fill, e.g. #EAF1F8." })),
		border: Type.Optional(Type.String({ description: "Border color, or 'none'." })),
		borderWidth: Type.Optional(Type.Number({ description: "Border width in points. Default 1." })),
		bodySize: Type.Optional(Type.Number({ description: "Body font size in points." })),
		headerSize: Type.Optional(Type.Number({ description: "Header font size in points." })),
		headerBold: Type.Optional(Type.Boolean({ description: "Bold header. Default true." })),
		bodyBold: Type.Optional(Type.Boolean({ description: "Bold body cells." })),
		colWidths: Type.Optional(Type.Array(Type.Number(), { description: "Column widths in inches." })),
		rowHeight: Type.Optional(Type.Number({ description: "Row height in inches." })),
		aligns: Type.Optional(Type.Array(Type.String(), { description: "Per-column alignment, cycling." })),
		margin: Type.Optional(Type.Number({ description: "Horizontal cell padding in inches." })),
		valign: Type.Optional(Type.Union([Type.Literal("top"), Type.Literal("middle"), Type.Literal("bottom")])),
		clearFill: Type.Optional(Type.Boolean({ description: "Clear fills on non-header rows." })),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		const args = [
			"table-style",
			params.path,
			"--slide",
			String(params.slide),
			"--shape",
			params.shape,
		];
		if (params.themeStyle != null) args.push(params.themeStyle ? "--theme-style" : "--no-theme-style");
		if (params.header === false) args.push("--no-header");
		if (params.headerFill) args.push("--header-fill", params.headerFill);
		if (params.headerColor) args.push("--header-color", params.headerColor);
		if (params.bodyColor) args.push("--body-color", params.bodyColor);
		if (params.zebra) args.push("--zebra", params.zebra);
		if (params.border) args.push("--border", params.border);
		if (params.borderWidth != null) args.push("--border-width", String(params.borderWidth));
		if (params.bodySize != null) args.push("--body-size", String(params.bodySize));
		if (params.headerSize != null) args.push("--header-size", String(params.headerSize));
		if (params.headerBold != null) args.push(params.headerBold ? "--header-bold" : "--no-header-bold");
		if (params.bodyBold != null) args.push(params.bodyBold ? "--body-bold" : "--no-body-bold");
		if (params.colWidths?.length) args.push("--col-widths", params.colWidths.join(","));
		if (params.rowHeight != null) args.push("--row-height", String(params.rowHeight));
		if (params.aligns?.length) args.push("--aligns", params.aligns.join(","));
		if (params.margin != null) args.push("--margin", String(params.margin));
		if (params.valign) args.push("--valign", params.valign);
		if (params.clearFill) args.push("--clear-fill");
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, `styled table "${data.path}" (#${data.id}): ${data.rows}x${data.cols}`);
	},
});

function batchSummary(data: KitResult): string {
	const state = data.dry_run ? "validated (dry run)" : data.saved ? "saved" : "NOT saved";
	const lines = [
		`batch ${state}: ${data.applied}/${data.op_count} ops, changed slides [${(data.changed_slides ?? []).join(", ")}]`,
	];
	if (data.errors?.length) {
		const details = data.errors
			.map((err: any) => `#${err.index} ${err.op ?? "?"}: ${String(err.error).split("\n")[0]}`)
			.join("; ");
		lines.push(`WARNING ${data.errors.length} op error(s): ${details}`);
	}
	for (const entry of data.overflow ?? []) {
		if (entry.overflow?.length) {
			const shapes = entry.overflow.map((item: any) => `${item.shape}${item.id ? `#${item.id}` : ""}`).join(", ");
			lines.push(`WARNING slide ${entry.slide} text overflow: ${shapes}`);
		}
		if (entry.overlaps?.length) {
			const pairs = entry.overlaps.map((item: any) => `${item.a} x ${item.b}`).join(", ");
			lines.push(`WARNING slide ${entry.slide} shape overlap: ${pairs}`);
		}
	}
	if (data.shot_error) lines.push(`WARNING auto-shot failed: ${data.shot_error}`);
	if (data.hint) lines.push(`HINT ${data.hint}`);
	return lines.join("\n");
}

const applyTool = defineTool({
	name: "ppt_apply",
	label: "PPT Batch Apply",
	promptSnippet: "Apply many edits in one atomic transaction (with dry-run and overflow summary).",
	description: `Apply a list of operations to one deck in a single transaction: the file is written once, after all ops, and a failing op aborts the batch without touching the file (unless stopOnError=false). Use this instead of dozens of single-edit calls for a multi-page build.

Supported ops (op + fields):
- set_text: slide, shape, text | paras
- set_cell: slide, shape, row, col, text
- add_text: slide, x, y, w, h?, text | paras, size, font, color, bold, align, valign, line_spacing, space_after, space_before, indent, hanging, bullet, bullet_color, fit
- add_image: slide, path, x, y, w, h
- add_table: slide, rows, cols, x, y, w, h, data, size, header, theme_style, header_fill, header_color, zebra, border, col_widths, row_height, aligns
- delete_shape: slide, shape
- set_notes: slide, text
- table_style: slide, shape + ppt_table_style fields (snake_case)
- add_slide: layout (index or name)
- clone: source (1-based), at?, count?
- delete_slides: slides ("2,5-7")

Behavior:
- dryRun: validate all ops in memory and report what would change; the file is not written.
- shot (default true): after saving, render the changed slides and report overflow/overlap warnings.
- Returns per-op results, errors with indexes, changed_slides and the final slide_count.`,
	parameters: Type.Object({
		path: PPT_PATH,
		ops: Type.Array(APPLY_OP, { description: "Ordered operations; later ops see the effects of earlier ones." }),
		stopOnError: Type.Optional(Type.Boolean({ description: "Abort and leave the file untouched on the first error. Default true." })),
		dryRun: Type.Optional(Type.Boolean({ description: "Validate without saving. Default false." })),
		shot: Type.Optional(Type.Boolean({ description: "Auto-render changed slides for overflow/overlap. Default true." })),
		overflow: Type.Optional(
			Type.Union([Type.Literal("loose"), Type.Literal("strict"), Type.Literal("off")], {
				description: "Overflow mode for the auto-shot. Default loose.",
			}),
		),
		ignore: Type.Optional(Type.Array(Type.String(), { description: "Shape names to skip in auto-shot checks." })),
		out: OUT,
	}),
	annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
	async execute(_id, params, signal, _onUpdate, ctx) {
		if (!params.ops?.length) throw new Error("ppt_apply needs at least one op");
		const args = ["apply", params.path, "--ops", JSON.stringify(params.ops)];
		if (params.stopOnError === false) args.push("--no-stop-on-error");
		if (params.dryRun) args.push("--dry-run");
		if (params.shot === false) args.push("--no-shot");
		if (params.overflow) args.push("--overflow", params.overflow);
		if (params.ignore?.length) args.push("--ignore", params.ignore.join(","));
		if (params.out) args.push("--out", params.out);
		const data = await mutate(ctx, signal, params.path, args);
		return jsonResult(data, batchSummary(data));
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
1. ppt_inspect the source deck: layouts and sample slides. Use mode="summary" for a large deck, then full inspect (with slides="...") only on pages you will edit. Shapes have an id: address them as shape="#<id>" when names are duplicated.
2. ppt_shot 1-4 sample pages (or montage=3 for an overview) and actually look at them; pick the page whose style matches each needed layout.
3. ppt_create a working copy with template=<source deck>. Never edit the source template in place.
4. Design that lives in the master/layout (logo, background) -> ppt_add_slide; design that lives in a single page (groups, pictures, charts, tables) -> ppt_clone that page. Delete unused placeholders with ppt_delete_shape (clearing text leaves a rendering bullet).
5. Text: ppt_set_text keeps the template's font/size/color. Use paras=[{runs:[{t,bold,color,size}]}] for mixed emphasis. New boxes: ppt_add_text (omit h to auto-grow; set spaceAfter/indent/bullet/fit as needed). Its font= applies to CJK too.
6. Tables: clone a page that has one when available (inspect summary has_table). Otherwise ppt_add_table (theme style off by default) + ppt_table_style. Notes: ppt_set_notes.
7. For a multi-page build or more than ~20 edits, use ppt_apply with an ops=[...] array (dryRun first). One transaction, one file write, per-op errors, and an automatic overflow/overlap summary of changed slides.
8. ppt_shot after each batch and check overflow[], overlaps[], wrapping and font sizes; fix and re-shoot. Do not finish while warnings are unresolved.
Numbers: slides and insert positions are 1-based; table row/col and shape index are 0-based. Canvas is 13.33x7.5 in unless ppt_inspect says otherwise. Undo = recreate the working copy from the template and re-run the batch. If ppt_shot fails, run ppt_doctor.`;

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
		deleteShapeTool,
		setNotesTool,
		tableStyleTool,
		applyTool,
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

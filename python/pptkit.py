#!/usr/bin/env python3
"""
pptkit — PPTX read/write + single-page rendering backend for the pi-pptx extension.

Layers (see RESEARCH.md):
  * python-pptx      : inspect / clone / edit
  * LibreOffice+PyMuPDF : render slides to PNG and compare text bboxes for overflow
  * JSON on stdout   : easy for the TypeScript tool layer to consume

Every mutating command writes atomically (temp file + fsync-by-replace) and
reopens the result to verify that the deck is still readable. Batch edits
(`apply`) mutate an in-memory deck and commit once, so a failed operation can
leave the file untouched.
"""
from __future__ import annotations

import argparse
import copy
import difflib
import hashlib
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400  # 1 inch
PT_PER_IN = 72.0
CACHE_ROOT = Path(tempfile.gettempdir()) / "pi-pptx"
MAX_CLONE_COUNT = 50


class KitError(Exception):
    """A user-facing failure (bad arguments, missing shape, ...)."""


# --------------------------------------------------------------------- utils
def fail(msg: str) -> None:
    raise KitError(msg)


def emit(obj) -> None:
    print(json.dumps(obj, ensure_ascii=False))


def to_in(emu) -> float:
    return round(emu / EMU, 3)


def to_pt(emu) -> float:
    return emu / EMU * PT_PER_IN


def box_in(shape) -> list[float] | None:
    if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
        return None
    return [to_in(shape.left), to_in(shape.top), to_in(shape.width), to_in(shape.height)]


def parse_rgb(value: str) -> RGBColor:
    v = value.strip().lstrip("#")
    if len(v) != 6:
        fail(f"颜色必须是 6 位十六进制，如 #1F4E79，收到: {value}")
    try:
        return RGBColor.from_string(v.upper())
    except Exception:
        fail(f"无法解析颜色: {value}")


def parse_enum(value: str | None, mapping: dict, what: str):
    if value is None:
        return None
    key = value.strip().lower()
    if key not in mapping:
        fail(f"未知的 {what}: {value} (可选: {', '.join(mapping)})")
    return mapping[key]


def parse_json_arg(value: str, what: str):
    """Accept a JSON string or '@/path/to/file.json' (and '-' for stdin)."""
    raw = value
    if value == "-":
        raw = sys.stdin.read()
    elif value.startswith("@"):
        p = Path(value[1:]).expanduser()
        if not p.exists():
            fail(f"{what} 文件不存在: {p}")
        raw = p.read_text(encoding="utf-8")
    elif not value.lstrip().startswith(("{", "[")):
        p = Path(value).expanduser()
        if p.exists() and p.is_file():
            raw = p.read_text(encoding="utf-8")
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        fail(f"{what} 不是合法 JSON: {exc}")


ALIGN = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}
VALIGN = {
    "top": MSO_ANCHOR.TOP,
    "middle": MSO_ANCHOR.MIDDLE,
    "bottom": MSO_ANCHOR.BOTTOM,
}
FIT = {
    "grow": MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT,
    "shrink": MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
    "none": MSO_AUTO_SIZE.NONE,
}
BULLET_NONE = {None, False, "", "none", "no", "off"}


def parse_pages(spec: str | None, total: int, what: str = "页") -> list[int]:
    """'2' | '1,3' | '2-4' -> [2] / [1, 3] / [2, 3, 4] (1-based, deduped)."""
    if not spec:
        return list(range(1, total + 1))
    pages: list[int] = []
    for raw in str(spec).split(","):
        part = raw.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                start, end = int(a), int(b)
            except ValueError:
                fail(f"无法解析{what}范围: {part}")
            if start > end:
                start, end = end, start
            if not (1 <= start and end <= total):
                fail(f"{what}范围 {part} 超出 1..{total}")
            pages.extend(range(start, end + 1))
        else:
            if not part.lstrip("-").isdigit():
                fail(f"无法解析{what}: {part}")
            p = int(part)
            if not 1 <= p <= total:
                fail(f"{what} {p} 超出 1..{total}")
            pages.append(p)
    seen: set[int] = set()
    out = []
    for p in pages:
        if p not in seen:
            seen.add(p)
            out.append(p)
    if not out:
        fail(f"没有有效的{what}")
    return out


def load(path: str):
    p = Path(path).expanduser()
    if not p.exists():
        fail(f"文件不存在: {p}")
    try:
        return Presentation(str(p)), p
    except Exception as exc:  # corrupt / not a pptx
        fail(f"无法打开 PPTX {p}: {exc}")


def save_deck(prs, src: Path, out: str | None) -> Path:
    """Atomic save + integrity check (reopen). Returns the written path."""
    target = Path(out).expanduser() if out else src
    if not target.is_absolute():
        target = (Path.cwd() / target).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.parent / f".{target.name}.{uuid.uuid4().hex}.tmp"
    try:
        prs.save(str(tmp))
        check = Presentation(str(tmp))
        if len(check.slides) < 1:
            raise RuntimeError("保存后页数为 0")
        os.replace(tmp, target)
    finally:
        if tmp.exists():
            tmp.unlink()
    return target


def save_and_reload(prs, src: Path, out: str | None):
    target = save_deck(prs, src, out)
    return Presentation(str(target)), target


def deck_meta(prs, path: Path) -> dict:
    return {
        "file": str(path.resolve()),
        "slide_count": len(prs.slides),
        "slide_size_in": [round(prs.slide_width / EMU, 3), round(prs.slide_height / EMU, 3)],
    }


# ------------------------------------------------------------------ shape tree
def iter_shapes(shapes, prefix: str = ""):
    """Yield (shape, path). Group children are addressed as 'Group 5/TextBox 2'."""
    for shape in shapes:
        path = f"{prefix}{shape.name}"
        yield shape, path
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes, path + "/")


def _candidate_line(shape, path: str, index: int | None, tf=None) -> str:
    box = _abs_box(shape, tf or IDENTITY_TF)
    box_txt = f"box={box}" if box else "box=?"
    text = ""
    if shape.has_text_frame:
        text = shape.text_frame.text.strip().replace("\n", " / ")
        if len(text) > 24:
            text = text[:24] + "…"
    path_txt = f" path={path!r}" if path != shape.name else ""
    return (
        f"  index={index if index is not None else '-'} id=#{shape.shape_id} "
        f"name={shape.name!r}{path_txt} type={str(shape.shape_type).split(' ')[0]} {box_txt}"
        + (f" text={text!r}" if text else "")
    )


def _candidate_table(slide, limit: int = 20) -> str:
    rows = ["可用形状路径/候选形状（用 shape=\"#<id>\" 精确定位）:"]
    top = list(slide.shapes)
    for i, shape in enumerate(top[:limit]):
        rows.append(_candidate_line(shape, shape.name, i))
    if len(top) > limit:
        rows.append(f"  … 还有 {len(top) - limit} 个顶层形状")
    return "\n".join(rows)


def resolve_shape(slide, key: str):
    """Resolve a shape by '#id', exact path, unique name, or top-level index.

    Indexes are 0-based. Duplicate names are refused instead of silently
    picking the first match: templates often have several shapes named
    '标题 3'. Use the '#<id>' form returned by ppt_inspect in that case.
    """
    key = str(key).strip()
    top = list(slide.shapes)
    all_shapes = list(iter_shapes(slide.shapes))

    if key.startswith("#"):
        id_part = key[1:].strip()
        if not id_part.isdigit():
            fail(f"形状 id 必须是数字，如 '#12'，收到: {key}")
        wanted = int(id_part)
        matches = [(s, p) for s, p in all_shapes if s.shape_id == wanted]
        if len(matches) == 1:
            return matches[0]
        if not matches:
            fail(f"找不到 id={wanted} 的形状。\n{_candidate_table(slide)}")
        fail(f"id={wanted} 在这页上不唯一，请检查 ppt_inspect 输出。")

    # exact recursive path
    exact = [(s, p) for s, p in all_shapes if p == key]
    if len(exact) == 1:
        return exact[0]

    # name match
    by_name = [(s, p) for s, p in all_shapes if s.name == key]
    if len(by_name) == 1:
        return by_name[0]
    if len(by_name) > 1:
        rows = [f"形状名 {key!r} 不唯一（{len(by_name)} 个），请改用 #<id>："]
        for shape, path in by_name:
            index = next(
                (i for i, s in enumerate(top) if s._element is shape._element),
                None,
            )
            rows.append(_candidate_line(shape, path, index))
        fail("\n".join(rows))

    # 0-based top-level index
    if key.lstrip("-").isdigit():
        idx = int(key)
        if -len(top) <= idx < len(top):
            shape = top[idx]
            return shape, shape.name
        fail(
            f"顶层序号 {idx} 超出范围 -{len(top)}..{len(top) - 1}（index 从 0 开始，本页共 {len(top)} 个顶层形状）。"
            f"提示：可改用 ppt_inspect 返回的 #id。\n{_candidate_table(slide)}"
        )

    # unique leaf name inside a group
    leaves = [(s, p) for s, p in all_shapes if s.name == key.rsplit("/", 1)[-1]]
    if len(leaves) == 1:
        return leaves[0]

    paths = [p for _, p in all_shapes]
    close = difflib.get_close_matches(key, paths, n=3, cutoff=0.35)
    hint = f"\n最相近的候选: {close}" if close else ""
    fail(f"找不到形状 {key!r}。{hint}\n{_candidate_table(slide)}")


def text_style(text_frame) -> dict | None:
    for para in text_frame.paragraphs:
        for run in para.runs:
            f = run.font
            style = {}
            if f.size is not None:
                style["size_pt"] = round(f.size.pt, 1)
            if f.name:
                style["font"] = f.name
            if f.bold is not None:
                style["bold"] = bool(f.bold)
            if f.italic is not None:
                style["italic"] = bool(f.italic)
            if para.alignment is not None:
                style["align"] = str(para.alignment).split(" ")[0].lower()
            return style or None
    return None


IDENTITY_TF = (1.0, 1.0, 0.0, 0.0)


def _abs_box(shape, tf) -> list[float] | None:
    sx, sy, ox, oy = tf
    if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
        return None
    return [
        round((shape.left * sx + ox) / EMU, 3),
        round((shape.top * sy + oy) / EMU, 3),
        round(shape.width * sx / EMU, 3),
        round(shape.height * sy / EMU, 3),
    ]


def _group_child_tf(group, tf):
    """Map a group's child coordinate space onto the slide canvas."""
    try:
        xfrm = group._element.grpSpPr.xfrm
        ch_off, ch_ext = xfrm.chOff, xfrm.chExt
        chex, chey = ch_ext.cx, ch_ext.cy
        if not chex or not chey:
            return tf
    except Exception:
        return tf
    sx, sy, ox, oy = tf
    gx, gy = group.left * sx + ox, group.top * sy + oy
    gw, gh = group.width * sx, group.height * sy
    csx, csy = gw / chex, gh / chey
    return (csx, csy, gx - ch_off.x * csx, gy - ch_off.y * csy)


def shape_info(shape, path: str, top_index: int | None = None, tf=IDENTITY_TF, depth: int = 99) -> dict:
    info = {
        "name": shape.name,
        "path": path,
        "type": str(shape.shape_type).split(" ")[0],
        "id": shape.shape_id,
    }
    if top_index is not None:
        # Usable as `shape` in the editing tools. Matters when a slide has
        # duplicate shape names (common when a template hand-places text boxes).
        info["index"] = top_index
    box = _abs_box(shape, tf)
    if box:
        info["box_in"] = box
    if shape.has_text_frame and shape.text_frame.text.strip():
        info["text"] = shape.text_frame.text
        style = text_style(shape.text_frame)
        if style:
            info["style"] = style
    if getattr(shape, "has_table", False) and shape.has_table:
        info["table"] = [[cell.text for cell in row.cells] for row in shape.table.rows]
    if getattr(shape, "has_chart", False) and shape.has_chart:
        info["chart"] = True
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        image = getattr(shape, "image", None)
        if image is not None:
            info["image"] = {"ext": image.ext, "px": [image.size[0], image.size[1]]}
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP and depth > 0:
        child_tf = _group_child_tf(shape, tf)
        info["children"] = [
            shape_info(child, f"{path}/{child.name}", None, child_tf, depth - 1) for child in shape.shapes
        ]
    return info


def shape_brief(shape, index: int | None = None) -> dict:
    info = {
        "id": shape.shape_id,
        "name": shape.name,
        "type": str(shape.shape_type).split(" ")[0],
        "box_in": box_in(shape),
    }
    if index is not None:
        info["index"] = index
    if shape.is_placeholder:
        info["placeholder_idx"] = shape.placeholder_format.idx
        info["placeholder_type"] = str(shape.placeholder_format.type).split(" ")[0]
    return info


def slide_title(slide) -> str:
    texts = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if not text:
            continue
        name = shape.name.lower()
        if "标题" in name or "title" in name:
            return text.replace("\n", " ")[:60]
        texts.append(text.replace("\n", " "))
    return (texts[0][:60] if texts else "")


def slide_summary(prs, idx: int) -> dict:
    slide = prs.slides[idx]
    types: dict[str, int] = {}
    has_table = has_picture = has_chart = False
    for shape in slide.shapes:
        kind = str(shape.shape_type).split(" ")[0]
        types[kind] = types.get(kind, 0) + 1
        has_table = has_table or bool(getattr(shape, "has_table", False) and shape.has_table)
        has_chart = has_chart or bool(getattr(shape, "has_chart", False) and shape.has_chart)
        has_picture = has_picture or shape.shape_type == MSO_SHAPE_TYPE.PICTURE
    info = {
        "index": idx + 1,
        "layout": slide.slide_layout.name,
        "title": slide_title(slide),
        "shape_count": len(slide.shapes),
        "types": types,
        "has_table": has_table,
        "has_picture": has_picture,
        "has_chart": has_chart,
    }
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
        if notes:
            info["notes"] = notes[:40] + ("…" if len(notes) > 40 else "")
    return info


def slide_info(prs, idx: int, depth: int = 99) -> dict:
    slide = prs.slides[idx]
    info = {
        "index": idx + 1,
        "layout": slide.slide_layout.name,
        "layout_index": list(prs.slide_layouts).index(slide.slide_layout),
        "shape_count": len(slide.shapes),
        "shapes": [
            shape_info(shape, shape.name, i, depth=depth) for i, shape in enumerate(slide.shapes)
        ],
    }
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text
        if notes.strip():
            info["notes"] = notes
    return info


# ---------------------------------------------------------------- inspect cmd
def cmd_inspect(args):
    prs, path = load(args.pptx)
    total = len(prs.slides)
    pages = parse_pages(args.slides, total) if args.slides else list(range(1, total + 1))
    depth = max(0, int(args.depth))
    out = deck_meta(prs, path)
    out["mode"] = args.mode
    out["layouts"] = [
        {
            "index": i,
            "name": layout.name,
            "placeholders": [
                {
                    "idx": ph.placeholder_format.idx,
                    "type": str(ph.placeholder_format.type).split(" ")[0],
                    "name": ph.name,
                }
                for ph in layout.placeholders
            ],
        }
        for i, layout in enumerate(prs.slide_layouts)
    ]
    if args.mode == "summary":
        out["slides"] = [slide_summary(prs, p - 1) for p in pages]
    else:
        out["slides"] = [slide_info(prs, p - 1, depth) for p in pages]
    out["inspected_slides"] = pages
    emit(out)


# ---------------------------------------------------------------- render / shot
def deck_pdf(pptx: Path) -> tuple[Path, str, bool]:
    """Convert the whole deck to PDF, cached by content hash + mtime + size."""
    pptx = pptx.resolve()
    st = pptx.stat()
    digest = hashlib.sha256()
    with open(pptx, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    key = hashlib.sha1(
        f"{pptx}|{st.st_mtime_ns}|{st.st_size}|{digest.hexdigest()}".encode()
    ).hexdigest()[:16]
    cache = CACHE_ROOT / key
    pdf = cache / (pptx.stem + ".pdf")
    if pdf.exists():
        return pdf, key, True

    cache.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="conv-", dir=cache))
    profile = work / "lo-profile"
    cmd = [
        "soffice",
        f"-env:UserInstallation=file://{profile}",
        "--headless",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        str(work),
        str(pptx),
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        produced = work / (pptx.stem + ".pdf")
        if not produced.exists():
            pdfs = list(work.glob("*.pdf"))
            produced = pdfs[0] if pdfs else None
        if produced is None:
            raise RuntimeError(
                "LibreOffice 转换失败（确认已安装 soffice）:\n"
                f"{result.stdout.strip()}\n{result.stderr.strip()}"
            )
        os.replace(produced, pdf)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return pdf, key, False


def _norm_text(value: str) -> str:
    return "".join(value.split())


def _text_match_score(block: str, shape_text: str) -> int:
    if not block or not shape_text:
        return 0
    if block == shape_text:
        return 4
    if len(shape_text) < 2 or len(block) < 2:
        return 0
    if block.startswith(shape_text) or shape_text.startswith(block):
        return 3
    n = min(8, len(shape_text), len(block))
    if n >= 2 and block[:n] == shape_text[:n]:
        return 2
    if block in shape_text or shape_text in block:
        return 1
    return 0


def detect_overflow(
    slide,
    page,
    slide_w_pt: float,
    slide_h_pt: float,
    mode: str = "loose",
    ignore: set[str] | None = None,
) -> list[dict]:
    """Compare rendered text-block bboxes with the shapes' declared boxes.

    Only top-level shapes with text are checked: group children store relative
    coordinates, which do not map onto the rendered page.

    Each rendered text block is assigned to at most one shape (best text match,
    then nearest top edge), which avoids a neighbouring label dragging the
    measured bbox of a short shape sideways.

    ``mode``:
      * loose (default): skip left/right checks for centred/justified text —
        that is where the old heuristic produced most false positives.
      * strict: report every side with tight tolerances.
      * off: skip text overflow checks entirely.
    """
    if mode == "off":
        return []
    ignore = ignore or set()
    sx = page.rect.width / slide_w_pt if slide_w_pt else 1.0
    sy = page.rect.height / slide_h_pt if slide_h_pt else 1.0
    shapes = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if getattr(shape, "rotation", 0):
            continue
        if not shape.has_text_frame:
            continue
        if shape.name in ignore:
            continue
        txt = _norm_text(shape.text_frame.text)
        if not txt or shape.left is None or shape.width is None or shape.height is None:
            continue
        aligns = {
            para.alignment
            for para in shape.text_frame.paragraphs
            if para.alignment is not None
        }
        box = (
            to_pt(shape.left) * sx,
            to_pt(shape.top) * sy,
            to_pt(shape.left + shape.width) * sx,
            to_pt(shape.top + shape.height) * sy,
        )
        shapes.append((shape.name, shape.shape_id, txt, box, box_in(shape), aligns))
    if not shapes:
        return []

    # Work per rendered *line*: two separate text boxes on the same baseline
    # can land in one PyMuPDF block, which used to make a short label inherit
    # its neighbour's width (and report a bogus 400pt horizontal overflow).
    blocks = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            text = _norm_text("".join(span.get("text", "") for span in line.get("spans", [])))
            if not text:
                continue
            x0, y0, x1, y1 = line["bbox"]
            blocks.append((x0, y0, x1, y1, text))

    assigned: dict[int, list[tuple[float, float, float, float]]] = {i: [] for i in range(len(shapes))}
    for bx0, by0, bx1, by1, btext in blocks:
        best = None
        for i, (_name, _sid, txt, (x0, y0, x1, y1), _box, _aligns) in enumerate(shapes):
            if bx1 < x0 - 6 or bx0 > x1 + 6:
                continue
            overlaps_y = not (by1 < y0 - 8 or by0 > y1 + 8)
            starts_near = y0 - 8 <= by0 <= y1 + 12
            if not (overlaps_y or starts_near):
                continue
            score = _text_match_score(btext, txt)
            if score <= 0:
                continue
            candidate = (score, -abs(by0 - y0), i)
            if best is None or candidate > best:
                best = candidate
        if best is not None:
            assigned[best[2]].append((bx0, by0, bx1, by1))

    found = []
    for i, (name, sid, txt, (x0, y0, x1, y1), box_inches, aligns) in enumerate(shapes):
        matched = assigned[i]
        if not matched:
            continue
        ux0 = min(b[0] for b in matched)
        uy0 = min(b[1] for b in matched)
        ux1 = max(b[2] for b in matched)
        uy1 = max(b[3] for b in matched)
        # Rendered line boxes include leading/descender padding, so a small
        # vertical spill is normal. A real overflow is roughly one extra line.
        if mode == "strict":
            vtol = 4.0
            htol = 2.0
        else:
            vtol = max(10.0, 0.35 * max(b[3] - b[1] for b in matched))
            htol = 6.0
        over = {}
        centered = any(a in (PP_ALIGN.CENTER, PP_ALIGN.JUSTIFY) for a in aligns)
        horizontal = mode == "strict" or not centered
        if horizontal and ux1 > x1 + htol:
            over["right_pt"] = round(ux1 - x1, 1)
        if horizontal and ux0 < x0 - htol:
            over["left_pt"] = round(x0 - ux0, 1)
        if uy1 > y1 + vtol:
            over["bottom_pt"] = round(uy1 - y1, 1)
        if uy0 < y0 - vtol:
            over["top_pt"] = round(y0 - uy0, 1)
        if over:
            found.append(
                {
                    "shape": name,
                    "id": sid,
                    "box_in": box_inches,
                    "overflow_pt": over,
                    "text": txt[:40],
                }
            )
    return found


def detect_overlaps(slide, ignore: set[str] | None = None) -> list[dict]:
    """Report text boxes that visibly collide with other text/pictures/tables."""
    ignore = ignore or set()
    items = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if getattr(shape, "rotation", 0):
            continue
        if shape.name in ignore:
            continue
        box = box_in(shape)
        if not box:
            continue
        has_text = bool(shape.has_text_frame and shape.text_frame.text.strip())
        is_table = bool(getattr(shape, "has_table", False) and shape.has_table)
        is_picture = shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        kind = "text" if has_text else ("table" if is_table else ("picture" if is_picture else None))
        if kind is None:
            continue
        x, y, w, h = box
        items.append({"name": shape.name, "id": shape.shape_id, "kind": kind, "box": box, "area": max(w * h, 0.0001)})

    found = []
    for i in range(len(items)):
        a = items[i]
        ax, ay, aw, ah = a["box"]
        for j in range(i + 1, len(items)):
            b = items[j]
            if a["kind"] != "text" and b["kind"] != "text":
                continue
            bx, by, bw, bh = b["box"]
            ix = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
            iy = max(0.0, min(ay + ah, by + bh) - max(ay, by))
            area = ix * iy
            if area < 0.05:
                continue
            ratio = area / min(a["area"], b["area"])
            minimal = 0.22 if a["kind"] == b["kind"] == "text" else 0.35
            if ratio < minimal:
                continue
            found.append(
                {
                    "a": a["name"],
                    "a_id": a["id"],
                    "b": b["name"],
                    "b_id": b["id"],
                    "kinds": [a["kind"], b["kind"]],
                    "area_in2": round(area, 2),
                    "ratio": round(ratio, 2),
                }
            )
    return found[:20]


def make_montage(pymupdf, files: list[Path], out_path: Path, cols: int | None = None, cell_w: int = 640) -> dict:
    """Compose slide PNGs into one contact sheet (no Pillow dependency)."""
    count = len(files)
    cols = cols or max(1, math.ceil(math.sqrt(count)))
    cols = max(1, min(cols, count))
    rows = math.ceil(count / cols)
    first = pymupdf.Pixmap(str(files[0]))
    aspect = first.height / first.width if first.width else 0.5625
    cell_h = max(1, round(cell_w * aspect))
    gap = 12
    margin = 12
    width = cols * cell_w + (cols - 1) * gap + 2 * margin
    height = rows * cell_h + (rows - 1) * gap + 2 * margin
    doc = pymupdf.open()
    try:
        page = doc.new_page(width=width, height=height)
        for i, file in enumerate(files):
            col = i % cols
            row = i // cols
            x = margin + col * (cell_w + gap)
            y = margin + row * (cell_h + gap)
            page.insert_image(pymupdf.Rect(x, y, x + cell_w, y + cell_h), filename=str(file))
        pix = page.get_pixmap(dpi=72)
        pix.save(str(out_path))
        return {
            "file": str(out_path.resolve()),
            "px": [pix.width, pix.height],
            "cols": cols,
            "rows": rows,
            "slides": count,
        }
    finally:
        doc.close()


def _region_rect(page, region: list[float], slide_w_pt: float, slide_h_pt: float):
    import pymupdf

    sx = page.rect.width / slide_w_pt if slide_w_pt else 1.0
    sy = page.rect.height / slide_h_pt if slide_h_pt else 1.0
    x, y, w, h = region
    return pymupdf.Rect(x * PT_PER_IN * sx, y * PT_PER_IN * sy, (x + w) * PT_PER_IN * sx, (y + h) * PT_PER_IN * sy)


def preview_overflow(prs, pages: list[int], mode: str = "loose", ignore: set[str] | None = None) -> list[dict]:
    """Render an in-memory deck to a throwaway file to project overflow (dry run)."""
    work = Path(tempfile.mkdtemp(prefix="preview-", dir=CACHE_ROOT if CACHE_ROOT.exists() else None))
    try:
        tmp = work / "preview.pptx"
        prs.save(str(tmp))
        pdf, _key, _cached = deck_pdf(tmp)
        return render_overflow(prs, pdf, pages, mode, ignore)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def render_overflow(prs, pdf: Path, pages: list[int], mode: str = "loose", ignore: set[str] | None = None) -> list[dict]:
    import pymupdf

    doc = pymupdf.open(pdf)
    slide_w_pt = prs.slide_width / EMU * PT_PER_IN
    slide_h_pt = prs.slide_height / EMU * PT_PER_IN
    out = []
    try:
        for page_no in pages:
            if not 1 <= page_no <= doc.page_count:
                continue
            page = doc[page_no - 1]
            overflow = detect_overflow(prs.slides[page_no - 1], page, slide_w_pt, slide_h_pt, mode, ignore)
            overlaps = detect_overlaps(prs.slides[page_no - 1], ignore)
            if overflow or overlaps:
                entry: dict = {"slide": page_no}
                if overflow:
                    entry["overflow"] = overflow
                if overlaps:
                    entry["overlaps"] = overlaps
                out.append(entry)
    finally:
        doc.close()
    return out


def cmd_shot(args):
    import pymupdf

    prs, path = load(args.pptx)
    total = len(prs.slides)
    pages = parse_pages(args.slides, total)
    pdf, key, cache_hit = deck_pdf(path)
    doc = pymupdf.open(pdf)
    if doc.page_count < total:
        fail(f"PDF 只有 {doc.page_count} 页，但 PPTX 有 {total} 页（转换异常）")

    outdir = Path(args.out).expanduser() if args.out else CACHE_ROOT / key / f"shots-{args.dpi}"
    outdir.mkdir(parents=True, exist_ok=True)

    slide_w_pt = prs.slide_width / EMU * PT_PER_IN
    slide_h_pt = prs.slide_height / EMU * PT_PER_IN
    region = None
    if args.region:
        try:
            vals = [float(v) for v in str(args.region).split(",")]
        except ValueError:
            fail(f"--region 需要 4 个数字 x,y,w,h（英寸），收到: {args.region}")
        if len(vals) != 4:
            fail(f"--region 需要 4 个数字 x,y,w,h（英寸），收到: {args.region}")
        region = vals
    ignore = {s.strip() for s in (args.ignore or "").split(",") if s.strip()}
    if args.overflow not in ("loose", "strict", "off"):
        fail(f"--overflow 只能是 loose|strict|off，收到: {args.overflow}")

    shots = []
    files: list[Path] = []
    for page_no in pages:
        page = doc[page_no - 1]
        suffix = "-region" if region else ""
        target = outdir / f"slide-{page_no}{suffix}.png"
        if not target.exists():
            if region:
                page.get_pixmap(dpi=args.dpi, clip=_region_rect(page, region, slide_w_pt, slide_h_pt)).save(str(target))
            else:
                page.get_pixmap(dpi=args.dpi).save(str(target))
        pix = pymupdf.Pixmap(str(target))
        entry = {
            "slide": page_no,
            "file": str(target.resolve()),
            "px": [pix.width, pix.height],
        }
        overflow = detect_overflow(prs.slides[page_no - 1], page, slide_w_pt, slide_h_pt, args.overflow, ignore)
        if overflow:
            entry["overflow"] = overflow
        overlaps = detect_overlaps(prs.slides[page_no - 1], ignore)
        if overlaps:
            entry["overlaps"] = overlaps
        shots.append(entry)
        files.append(target)
    doc.close()

    omitted = []
    if args.max_pages and len(shots) > args.max_pages:
        omitted = [s["slide"] for s in shots[args.max_pages :]]
        shots = shots[: args.max_pages]
        files = files[: args.max_pages]

    result = {
        "pdf": str(pdf.resolve()),
        "dpi": args.dpi,
        "cache_hit": cache_hit,
        "overflow_mode": args.overflow,
        "shots": shots,
    }
    if region:
        result["region_in"] = region
    if args.montage:
        cols = max(1, int(args.montage))
        montage_path = outdir / "montage.png"
        result["montage"] = make_montage(
            pymupdf, files, montage_path, cols=cols if cols > 1 else None
        )
    if omitted:
        result["omitted_slides"] = omitted
        limit = args.max_pages
        result["hint"] = f"一次最多渲染 {limit} 页；其余页请再调用一次。"
    emit(result)


# ---------------------------------------------------------------- clone slide
def remap_rels(dest_slide, src_slide) -> None:
    """Re-point every r:id / r:embed in copied XML at the destination rels.

    Without this step pictures, charts and links silently disappear.
    """
    dpart, spart = dest_slide.part, src_slide.part
    for element in dest_slide.element.iter():
        for attr, value in list(element.attrib.items()):
            if not attr.startswith("{%s}" % R_NS):
                continue
            rel = spart.rels.get(value)
            if rel is None:
                continue
            if rel.is_external:
                element.attrib[attr] = dpart.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
            else:
                element.attrib[attr] = dpart.relate_to(rel.target_part, rel.reltype)


def clone_slide(prs, src_index: int, at: int | None = None, count: int = 1) -> list:
    src = prs.slides[src_index]
    created = []
    for k in range(count):
        new = prs.slides.add_slide(src.slide_layout)
        for shape in list(new.shapes):  # drop placeholders inherited from the layout
            shape._element.getparent().remove(shape._element)
        for shape in src.shapes:
            new.shapes._spTree.append(copy.deepcopy(shape._element))

        # copy explicit slide background (if any) so themed pages stay themed
        src_bg = src.element.cSld.find(qn("p:bg"))
        dest_bg = new.element.cSld.find(qn("p:bg"))
        if src_bg is not None:
            if dest_bg is not None:
                new.element.cSld.replace(dest_bg, copy.deepcopy(src_bg))
            else:
                new.element.cSld.insert(0, copy.deepcopy(src_bg))
        elif dest_bg is not None:
            new.element.cSld.remove(dest_bg)

        remap_rels(new, src)

        if src.has_notes_slide:
            notes = src.notes_slide.notes_text_frame.text
            if notes.strip():
                new.notes_slide.notes_text_frame.text = notes

        created.append(new)
        if at is not None:
            sld_id_lst = prs.slides._sldIdLst
            ids = list(sld_id_lst)
            moved = ids[-1]
            sld_id_lst.remove(moved)
            pos = min(max(at - 1 + k, 0), len(ids) - 1)
            sld_id_lst.insert(pos, moved)
    return created


# ------------------------------------------------------------- text formatting
def _set_run_font(run, font: str) -> None:
    """Set latin *and* East-Asian/complex-script typefaces.

    python-pptx's ``run.font.name`` only writes ``a:latin``; CJK glyphs follow
    ``a:ea`` (or the theme), so ``font='微软雅黑'`` used to be a silent no-op for
    Chinese text.
    """
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    prev = rPr.find(qn("a:latin"))
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            if prev is not None:
                prev.addnext(el)
            else:
                rPr.append(el)
        el.set("typeface", font)
        prev = el


def _first_run_rpr(text_frame):
    for para in text_frame.paragraphs:
        for run in para.runs:
            rPr = run._r.rPr
            if rPr is not None:
                return copy.deepcopy(rPr)
            return None
    return None


def apply_run_spec(run, spec: dict) -> None:
    if spec.get("size") is not None:
        run.font.size = Pt(float(spec["size"]))
    if spec.get("font"):
        _set_run_font(run, str(spec["font"]))
    if spec.get("color"):
        run.font.color.rgb = parse_rgb(str(spec["color"]))
    if spec.get("bold") is not None:
        run.font.bold = bool(spec["bold"])
    if spec.get("italic") is not None:
        run.font.italic = bool(spec["italic"])
    if spec.get("underline") is not None:
        run.font.underline = bool(spec["underline"])


def _bullet_anchor(pPr):
    for tag in ("a:tabLst", "a:defRPr", "a:extLst"):
        el = pPr.find(qn(tag))
        if el is not None:
            return el
    return None


def set_paragraph_bullet(
    para,
    char=None,
    color: str | None = None,
    indent: float | None = None,
    hanging: float | None = None,
) -> None:
    """Write a bullet (or explicit buNone) plus hanging indent into a:pPr.

    ``char="__keep__"`` only updates the indents and leaves the existing bullet
    definition alone (used when a paragraph is given an indent but no bullet).
    """
    pPr = para._p.get_or_add_pPr()
    keep = char == "__keep__"
    if not keep:
        for tag in ("a:buClr", "a:buFont", "a:buChar", "a:buAutoNum", "a:buNone"):
            for el in pPr.findall(qn(tag)):
                pPr.remove(el)

    if indent is not None:
        mar = abs(float(indent))
        pPr.set("marL", str(int(Inches(mar))))
        if hanging is None:
            hanging = mar if char not in BULLET_NONE and not keep else 0.0
        pPr.set("indent", str(-int(Inches(max(0.0, float(hanging))))))
    elif hanging is not None:
        pPr.set("indent", str(-int(Inches(max(0.0, float(hanging))))))
    if keep:
        return

    anchor = _bullet_anchor(pPr)
    elements = []
    if char in BULLET_NONE:
        elements.append(pPr.makeelement(qn("a:buNone"), {}))
    else:
        glyph = "•" if char is True else str(char)
        if color:
            buClr = pPr.makeelement(qn("a:buClr"), {})
            buClr.append(buClr.makeelement(qn("a:srgbClr"), {"val": str(parse_rgb(str(color)))}))
            elements.append(buClr)
        buFont = pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
        elements.append(buFont)
        elements.append(pPr.makeelement(qn("a:buChar"), {"char": glyph}))
    for el in elements:
        if anchor is not None:
            anchor.addprevious(el)
        else:
            pPr.append(el)


def apply_para_spec(para, spec: dict) -> None:
    align = parse_enum(spec.get("align"), ALIGN, "对齐方式")
    if align is not None:
        para.alignment = align
    if spec.get("space_after") is not None:
        para.space_after = Pt(float(spec["space_after"]))
    if spec.get("space_before") is not None:
        para.space_before = Pt(float(spec["space_before"]))
    if spec.get("line_spacing") is not None:
        para.line_spacing = float(spec["line_spacing"])
    if spec.get("level") is not None:
        para.level = int(spec["level"])
    bullet_given = "bullet" in spec
    if bullet_given or spec.get("indent") is not None or spec.get("hanging") is not None:
        set_paragraph_bullet(
            para,
            char=spec.get("bullet") if bullet_given else "__keep__",
            color=spec.get("bullet_color"),
            indent=spec.get("indent"),
            hanging=spec.get("hanging"),
        )


def _write_runs(para, runs: list[dict], base_rpr) -> None:
    for run in list(para.runs):
        run._r.getparent().remove(run._r)
    for spec in runs:
        if not isinstance(spec, dict):
            spec = {"t": spec}
        run = para.add_run()
        run.text = "" if spec.get("t", spec.get("text", "")) is None else str(spec.get("t", spec.get("text", "")))
        if base_rpr is not None:
            rPr = run._r.get_or_add_rPr()
            run._r.remove(rPr)
            run._r.insert(0, copy.deepcopy(base_rpr))
        apply_run_spec(run, spec)


def write_paragraphs(
    text_frame,
    paras: list[dict],
    defaults: dict | None = None,
    run_defaults: dict | None = None,
) -> int:
    """Replace all paragraphs with run-level rich text (the `paras=` form)."""
    if not isinstance(paras, list):
        fail("paras 必须是数组，每项是一段")
    defaults = defaults or {}
    run_defaults = run_defaults or {}
    base_rpr = _first_run_rpr(text_frame)
    base_p = copy.deepcopy(text_frame.paragraphs[0]._p) if text_frame.paragraphs else None
    for para in list(text_frame.paragraphs):
        para._p.getparent().remove(para._p)
    if not paras:
        paras = [{"runs": [{"t": ""}]}]
    for spec in paras:
        if not isinstance(spec, dict):
            fail("paras 的每一项必须是对象")
        el = copy.deepcopy(base_p) if base_p is not None else text_frame._txBody.makeelement(qn("a:p"), {})
        text_frame._txBody.append(el)
        para = text_frame.paragraphs[-1]
        for run in list(para.runs):
            run._r.getparent().remove(run._r)
        merged = {**defaults, **spec}
        apply_para_spec(para, merged)
        runs = spec.get("runs")
        if runs is None:
            runs = [{"t": spec.get("t", spec.get("text", ""))}]
        if not isinstance(runs, list):
            fail("paras[].runs 必须是数组")
        if run_defaults:
            runs = [{**run_defaults, **run} if isinstance(run, dict) else run for run in runs]
        _write_runs(para, runs, base_rpr)
    return len(text_frame.paragraphs)


def _set_para_text(para, text: str) -> None:
    """Replace a paragraph's text while keeping the first run's formatting."""
    runs = list(para.runs)
    if runs:
        runs[0].text = text
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        para.add_run().text = text


def fill_paragraphs(text_frame, text: str) -> None:
    """Write text (\\n separated) inheriting the first paragraph's style."""
    lines = text.split("\n")
    first = text_frame.paragraphs[0]
    base = copy.deepcopy(first._p)
    for para in list(text_frame.paragraphs)[1:]:
        para._p.getparent().remove(para._p)
    _set_para_text(first, lines[0])
    prev = first._p
    for _ in lines[1:]:
        element = copy.deepcopy(base)
        prev.addnext(element)
        prev = element
    for para, line in zip(text_frame.paragraphs, lines):
        _set_para_text(para, line)


def apply_text_frame_defaults(text_frame, spec: dict) -> None:
    fit = spec.get("fit")
    if fit is not None:
        text_frame.auto_size = parse_enum(fit, FIT, "fit")
    if spec.get("valign"):
        text_frame.vertical_anchor = parse_enum(spec["valign"], VALIGN, "垂直对齐")
    if spec.get("wrap") is False:
        text_frame.word_wrap = False


# -------------------------------------------------------------- table styling
def _table_element(shape):
    return shape._element.graphic.graphicData.tbl


def _tbl_pr(tbl):
    tblPr = tbl.find(qn("a:tblPr"))
    if tblPr is None:
        tblPr = tbl.makeelement(qn("a:tblPr"), {})
        tbl.insert(0, tblPr)
    return tblPr


DEFAULT_TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"


def set_theme_table_style(shape, enabled: bool) -> None:
    """Turn the theme's firstRow/bandRow variants and tableStyleId on/off."""
    tbl = _table_element(shape)
    tblPr = _tbl_pr(tbl)
    if enabled:
        tblPr.set("firstRow", "1")
        tblPr.set("bandRow", "1")
        if tblPr.find(qn("a:tableStyleId")) is None:
            el = tblPr.makeelement(qn("a:tableStyleId"), {})
            el.text = DEFAULT_TABLE_STYLE
            tblPr.insert(0, el)
        return
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    for child in list(tblPr):
        if child.tag == qn("a:tableStyleId"):
            tblPr.remove(child)


def _set_cell_borders(cell, color: str | None, width_pt: float = 1.0) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    for idx, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        ln = tcPr.makeelement(
            qn(tag),
            {"w": str(int(width_pt * 12700)), "cap": "flat", "cmpd": "sng", "algn": "ctr"},
        )
        if color:
            fill = ln.makeelement(qn("a:solidFill"), {})
            fill.append(fill.makeelement(qn("a:srgbClr"), {"val": str(parse_rgb(color))}))
            ln.append(fill)
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "solid"}))
        else:
            ln.append(ln.makeelement(qn("a:noFill"), {}))
        tcPr.insert(idx, ln)


def style_table(shape, spec: dict) -> dict:
    table = shape.table
    nrows, ncols = len(table.rows), len(table.columns)

    if spec.get("theme_style") is not None:
        set_theme_table_style(shape, bool(spec["theme_style"]))

    col_widths = spec.get("col_widths")
    if isinstance(col_widths, str):
        col_widths = [v.strip() for v in col_widths.split(",") if v.strip()]
    if col_widths:
        try:
            widths = [float(v) for v in col_widths]
        except (TypeError, ValueError):
            fail(f"col_widths 必须是数字数组（英寸），收到: {col_widths}")
        if len(widths) != ncols:
            fail(f"col_widths 需要 {ncols} 个值，收到 {len(widths)}")
        for col, width in zip(table.columns, widths):
            col.width = Inches(width)
    if spec.get("row_height") is not None:
        for row in table.rows:
            row.height = Inches(float(spec["row_height"]))

    body_size = spec.get("body_size", spec.get("size"))
    header_size = spec.get("header_size", body_size)
    aligns = spec.get("aligns")
    if isinstance(aligns, str):
        aligns = [a.strip() for a in aligns.split(",")]
    if aligns:
        aligns = [parse_enum(a, ALIGN, "对齐方式") for a in aligns]
    margin = spec.get("margin")
    valign = parse_enum(spec.get("valign"), VALIGN, "垂直对齐")
    border = spec.get("border")
    bold_header = spec.get("header_bold", True)
    header_fill = spec.get("header_fill")
    header_color = spec.get("header_color")
    body_color = spec.get("body_color")
    zebra = spec.get("zebra")

    for r in range(nrows):
        for c in range(ncols):
            cell = table.cell(r, c)
            is_header = r == 0 and bool(spec.get("header", True))
            if is_header and header_fill:
                cell.fill.solid()
                cell.fill.fore_color.rgb = parse_rgb(str(header_fill))
            elif not is_header and zebra and r % 2 == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = parse_rgb(str(zebra))
            elif spec.get("clear_fill"):
                cell.fill.background()
            if border is not None:
                _set_cell_borders(
                    cell,
                    None if border in (False, "none", "false") else str(border),
                    float(spec.get("border_width", 1.0)),
                )
            if margin is not None:
                cell.margin_left = Inches(float(margin))
                cell.margin_right = Inches(float(margin))
                cell.margin_top = Inches(float(margin) / 2)
                cell.margin_bottom = Inches(float(margin) / 2)
            if valign is not None:
                cell.vertical_anchor = valign
            for para in cell.text_frame.paragraphs:
                if aligns:
                    para.alignment = aligns[c % len(aligns)]
                elif is_header:
                    para.alignment = PP_ALIGN.CENTER
                for run in para.runs:
                    size = header_size if is_header else body_size
                    if size is not None:
                        run.font.size = Pt(float(size))
                    color = header_color if is_header else body_color
                    if color:
                        run.font.color.rgb = parse_rgb(str(color))
                    if is_header and bold_header:
                        run.font.bold = True
                    elif spec.get("body_bold") is not None:
                        run.font.bold = bool(spec["body_bold"])
    return {"rows": nrows, "cols": ncols}


# -------------------------------------------------------------- slide helpers
def get_slide(prs, number) -> "object":
    total = len(prs.slides)
    try:
        number = int(number)
    except (TypeError, ValueError):
        fail(f"页号必须是整数，收到: {number!r}")
    if not 1 <= number <= total:
        fail(f"页号 {number} 超出 1..{total}")
    return prs.slides[number - 1]


def get_layout(prs, key):
    key = str(key).strip()
    if key.lstrip("-").isdigit():
        idx = int(key)
        if not 0 <= idx < len(prs.slide_layouts):
            fail(f"版式序号 {idx} 超出 0..{len(prs.slide_layouts) - 1}")
        return prs.slide_layouts[idx]
    for cand in prs.slide_layouts:
        if cand.name == key:
            return cand
    names = [f"{i}:{l.name}" for i, l in enumerate(prs.slide_layouts)]
    fail(f"找不到版式 {key!r}。可用版式: {names}")


def remove_shape(shape) -> None:
    shape._element.getparent().remove(shape._element)


# ------------------------------------------------------------------- op layer
# Every op mutates the in-memory presentation and returns a JSON-able result.
# The CLI commands and the batch `apply` command both use this layer.
def op_set_text(prs, spec: dict) -> dict:
    slide = get_slide(prs, spec.get("slide"))
    shape, path = resolve_shape(slide, spec.get("shape"))
    if not shape.has_text_frame:
        if getattr(shape, "has_table", False) and shape.has_table:
            fail(f"{path} 是表格，请用 set_cell 修改单元格")
        fail(f"{path} 没有文本框，无法设置文本")
    tf = shape.text_frame
    if spec.get("paras") is not None:
        count = write_paragraphs(tf, spec["paras"])
    else:
        if spec.get("text") is None:
            fail("set_text 需要 text 或 paras")
        fill_paragraphs(tf, str(spec["text"]))
        count = len(tf.paragraphs)
    # text="" clears the runs but keeps the shape; use delete_shape to remove it.
    return {
        "slide": int(spec["slide"]),
        "op": "set_text",
        "id": shape.shape_id,
        "shape": shape.name,
        "path": path,
        "text": tf.text,
        "paragraphs": count,
    }


def op_set_cell(prs, spec: dict) -> dict:
    slide = get_slide(prs, spec.get("slide"))
    shape, path = resolve_shape(slide, spec.get("shape"))
    if not (getattr(shape, "has_table", False) and shape.has_table):
        fail(f"{path} 不是表格")
    table = shape.table
    nrows, ncols = len(table.rows), len(table.columns)
    row, col = int(spec.get("row")), int(spec.get("col"))
    if not (0 <= row < nrows and 0 <= col < ncols):
        fail(f"单元格 ({row},{col}) 超出表格范围 {nrows}x{ncols}（0 起）")
    text = "" if spec.get("text") is None else str(spec["text"])
    cell = table.cell(row, col)
    old = cell.text_frame.paragraphs[0]
    if old.runs:
        old.runs[0].text = text
        for run in old.runs[1:]:
            run._r.getparent().remove(run._r)
        for para in list(cell.text_frame.paragraphs)[1:]:
            para._p.getparent().remove(para._p)
    else:
        cell.text = text
    return {
        "slide": int(spec["slide"]),
        "op": "set_cell",
        "id": shape.shape_id,
        "shape": shape.name,
        "path": path,
        "row": row,
        "col": col,
        "text": text,
    }


def op_add_text(prs, spec: dict) -> dict:
    slide = get_slide(prs, spec.get("slide"))
    x = float(spec["x"])
    y = float(spec["y"])
    w = float(spec.get("w", 6.0))
    fit = spec.get("fit")
    if spec.get("h") is not None:
        h = float(spec["h"])
    else:
        # No explicit height: grow with the text instead of a fixed 0.8" box
        # that clips three lines at 18pt.
        if fit is None:
            fit = "grow"
        size = float(spec.get("size") or 18)
        text = spec.get("text") or ""
        lines = max(1, len(str(text).split("\n")))
        if spec.get("paras"):
            lines = max(1, len(spec["paras"]))
        h = round(0.25 + lines * size * 1.35 / 72.0, 2)
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    apply_text_frame_defaults(tf, {**spec, "fit": fit or "none"})

    para_defaults = {}
    for key in (
        "align",
        "space_after",
        "space_before",
        "line_spacing",
        "indent",
        "hanging",
        "bullet",
        "bullet_color",
        "level",
    ):
        if spec.get(key) is not None:
            para_defaults[key] = spec[key]
    # A bullet with no explicit indent gets the usual 0.25" hanging layout
    # instead of drawing the glyph on top of the first character.
    if (
        para_defaults.get("bullet") not in BULLET_NONE
        and para_defaults.get("indent") is None
        and para_defaults.get("hanging") is None
    ):
        para_defaults["indent"] = 0.25
    # top-level size/font/color/bold apply to runs unless a run overrides them
    run_defaults = {
        key: spec[key]
        for key in ("size", "font", "color", "bold", "italic", "underline")
        if spec.get(key) is not None
    }
    if spec.get("paras") is not None:
        count = write_paragraphs(tf, spec["paras"], para_defaults, run_defaults)
    else:
        text = "" if spec.get("text") is None else str(spec["text"])
        lines = text.split("\n")
        for i, line in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            apply_para_spec(para, para_defaults)
            run = para.add_run()
            run.text = line
            apply_run_spec(run, spec)
        count = len(lines)

    if spec.get("name"):
        box.name = str(spec["name"])
    if spec.get("valign"):
        tf.vertical_anchor = parse_enum(spec["valign"], VALIGN, "垂直对齐")
    return {
        "slide": int(spec["slide"]),
        "op": "add_text",
        "added": "textbox",
        "id": box.shape_id,
        "name": box.name,
        "path": box.name,
        "box_in": box_in(box),
        "paragraphs": count,
        "fit": fit or "none",
    }


def op_add_image(prs, spec: dict) -> dict:
    slide = get_slide(prs, spec.get("slide"))
    image_path = Path(str(spec.get("path"))).expanduser()
    if not image_path.is_absolute():
        image_path = (Path.cwd() / image_path).resolve()
    if not image_path.exists():
        fail(f"图片不存在: {image_path}")
    kwargs = {}
    if spec.get("x") is not None and spec.get("y") is not None:
        kwargs["left"] = Inches(float(spec["x"]))
        kwargs["top"] = Inches(float(spec["y"]))
    if spec.get("w") is not None:
        kwargs["width"] = Inches(float(spec["w"]))
    if spec.get("h") is not None:
        kwargs["height"] = Inches(float(spec["h"]))
    picture = slide.shapes.add_picture(str(image_path), **kwargs)
    return {
        "slide": int(spec["slide"]),
        "op": "add_image",
        "added": "picture",
        "id": picture.shape_id,
        "name": picture.name,
        "path": picture.name,
        "box_in": box_in(picture),
    }


def op_add_table(prs, spec: dict) -> dict:
    slide = get_slide(prs, spec.get("slide"))
    rows, cols = int(spec.get("rows")), int(spec.get("cols"))
    if rows < 1 or cols < 1:
        fail("表格至少 1 行 1 列")
    data = spec.get("data") or []
    if not isinstance(data, list) or any(not isinstance(row, list) for row in data):
        fail("data 必须是二维数组")
    if len(data) > rows or any(len(row) > cols for row in data):
        fail(f"data 尺寸 {len(data)}x{max((len(r) for r in data), default=0)} 超过 {rows}x{cols}")

    had_table = any(
        getattr(shape, "has_table", False) and shape.has_table
        for slide_obj in prs.slides
        for shape in slide_obj.shapes
    )
    frame = slide.shapes.add_table(
        rows,
        cols,
        Inches(float(spec["x"])),
        Inches(float(spec["y"])),
        Inches(float(spec["w"])),
        Inches(float(spec["h"])),
    )
    table = frame.table
    theme_style = bool(spec.get("theme_style"))
    set_theme_table_style(frame, theme_style)
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            value = ""
            if r < len(data) and c < len(data[r]):
                value = "" if data[r][c] is None else str(data[r][c])
            cell.text = value
    border = spec.get("border", "#D9D9D9")
    if border is None:
        border = "#D9D9D9"
    style_spec = {
        "header": bool(spec.get("header", True)),
        "header_fill": spec.get("header_fill"),
        "header_color": spec.get("header_color"),
        "body_color": spec.get("body_color"),
        "zebra": spec.get("zebra"),
        "border": border,
        "body_size": spec.get("size"),
        "header_size": spec.get("size"),
        "col_widths": spec.get("col_widths"),
        "row_height": spec.get("row_height"),
        "aligns": spec.get("aligns"),
        "margin": spec.get("margin"),
        "valign": spec.get("valign"),
        "header_bold": spec.get("header", True),
    }
    style_table(frame, style_spec)
    result = {
        "slide": int(spec["slide"]),
        "op": "add_table",
        "added": "table",
        "id": frame.shape_id,
        "name": frame.name,
        "path": frame.name,
        "box_in": box_in(frame),
        "rows": rows,
        "cols": cols,
        "theme_style": theme_style,
    }
    if not had_table:
        result["hint"] = (
            "模板里没有任何表格样例：已关闭主题表样式并使用中性边框。"
            "可用 ppt_table_style 调整，或克隆带表格的页。"
        )
    return result


def op_delete_shape(prs, spec: dict) -> dict:
    slide = get_slide(prs, spec.get("slide"))
    shape, path = resolve_shape(slide, spec.get("shape"))
    info = {
        "slide": int(spec["slide"]),
        "op": "delete_shape",
        "id": shape.shape_id,
        "shape": shape.name,
        "path": path,
        "deleted": True,
    }
    remove_shape(shape)
    return info


def op_set_notes(prs, spec: dict) -> dict:
    slide = get_slide(prs, spec.get("slide"))
    text = "" if spec.get("text") is None else str(spec["text"])
    tf = slide.notes_slide.notes_text_frame
    tf.text = text
    return {
        "slide": int(spec["slide"]),
        "op": "set_notes",
        "chars": len(text),
        "paragraphs": len(tf.paragraphs),
    }


def op_table_style(prs, spec: dict) -> dict:
    slide = get_slide(prs, spec.get("slide"))
    shape, path = resolve_shape(slide, spec.get("shape"))
    if not (getattr(shape, "has_table", False) and shape.has_table):
        fail(f"{path} 不是表格")
    info = style_table(shape, spec)
    return {
        "slide": int(spec["slide"]),
        "op": "table_style",
        "id": shape.shape_id,
        "shape": shape.name,
        "path": path,
        **info,
    }


def op_add_slide(prs, spec: dict) -> dict:
    layout = get_layout(prs, spec.get("layout"))
    slide = prs.slides.add_slide(layout)
    return {
        "slide": len(prs.slides),
        "op": "add_slide",
        "layout": layout.name,
        "placeholders": [ph.name for ph in slide.placeholders],
        "shapes": [shape_brief(shape, i) for i, shape in enumerate(slide.shapes)],
    }


def op_clone(prs, spec: dict) -> dict:
    total = len(prs.slides)
    src = int(spec.get("source", spec.get("src", 0)))
    if not 1 <= src <= total:
        fail(f"源页号 {src} 超出 1..{total}")
    count = int(spec.get("count") or 1)
    if not 1 <= count <= MAX_CLONE_COUNT:
        fail(f"克隆数量应在 1..{MAX_CLONE_COUNT}，收到 {count}")
    at = spec.get("at")
    if at is not None and not 1 <= int(at) <= total + 1:
        fail(f"插入位置 {at} 超出 1..{total + 1}")
    clone_slide(prs, src - 1, int(at) if at is not None else None, count)
    return {
        "slide": int(at) if at is not None else len(prs.slides),
        "op": "clone",
        "created_from": src,
        "count": count,
        "at": at,
        "slide_count": len(prs.slides),
    }


def op_delete_slides(prs, spec: dict) -> dict:
    total = len(prs.slides)
    pages = parse_pages(str(spec.get("slides")), total)
    if len(pages) >= total:
        fail("不能删除全部页（至少保留 1 页）")
    sld_id_lst = prs.slides._sldIdLst
    ids = list(sld_id_lst)
    for page in sorted(pages, reverse=True):
        sld_id_lst.remove(ids[page - 1])
    return {
        "slide": pages[0],
        "op": "delete_slides",
        "deleted": pages,
        "slide_count": len(prs.slides),
    }


OP_HANDLERS = {
    "set_text": op_set_text,
    "set_cell": op_set_cell,
    "add_text": op_add_text,
    "add_image": op_add_image,
    "add_table": op_add_table,
    "delete_shape": op_delete_shape,
    "set_notes": op_set_notes,
    "table_style": op_table_style,
    "add_slide": op_add_slide,
    "clone": op_clone,
    "delete_slides": op_delete_slides,
}


def apply_one(prs, op) -> dict:
    if not isinstance(op, dict):
        fail("每个 op 必须是 JSON 对象")
    name = op.get("op") or op.get("type")
    handler = OP_HANDLERS.get(str(name))
    if handler is None:
        fail(f"未知 op {name!r}。支持: {', '.join(sorted(OP_HANDLERS))}")
    result = handler(prs, op)
    result.setdefault("op", name)
    return result


# ---------------------------------------------------------------- deck command
def cmd_create(args):
    target = Path(args.out).expanduser()
    if not target.is_absolute():
        target = (Path.cwd() / target).resolve()
    if target.exists() and not args.overwrite:
        fail(f"目标文件已存在: {target}（需要覆盖请传 overwrite=true）")
    target.parent.mkdir(parents=True, exist_ok=True)

    if args.template:
        src = Path(args.template).expanduser()
        if not src.exists():
            fail(f"模板不存在: {src}")
        if src.resolve() == target.resolve():
            fail("template 与目标文件相同：请指定新的输出路径，不要覆盖模板")
        Presentation(str(src))  # verify it parses
        if args.width is None and args.height is None:
            tmp = target.parent / f".{target.name}.{uuid.uuid4().hex}.tmp"
            try:
                shutil.copyfile(src, tmp)
                Presentation(str(tmp))  # integrity
                os.replace(tmp, target)
            finally:
                if tmp.exists():
                    tmp.unlink()
        else:
            prs = Presentation(str(src))
            if args.width:
                prs.slide_width = Inches(args.width)
            if args.height:
                prs.slide_height = Inches(args.height)
            save_deck(prs, target, str(target))
        Presentation(str(target))  # integrity
    else:
        prs = Presentation()
        prs.slide_width = Inches(args.width or 13.333)
        prs.slide_height = Inches(args.height or 7.5)
        if not args.empty:
            layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(layout)
            if len(slide.placeholders) > 0:
                body = slide.placeholders[0]
                if body.has_text_frame and not body.text_frame.text.strip():
                    body.text_frame.text = "标题"
            if len(slide.placeholders) > 1:
                sub = slide.placeholders[1]
                if sub.has_text_frame and not sub.text_frame.text.strip():
                    sub.text_frame.text = "副标题"
        prs.save(str(target))
        Presentation(str(target))

    prs = Presentation(str(target))
    out = deck_meta(prs, target)
    out["template"] = str(Path(args.template).expanduser().resolve()) if args.template else None
    emit(out)


def cmd_add_slide(args):
    prs, path = load(args.pptx)
    result = op_add_slide(prs, {"layout": args.layout})
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    out["added_slide"] = result["slide"]
    emit(out)


def cmd_delete(args):
    prs, path = load(args.pptx)
    result = op_delete_slides(prs, {"slides": args.slides})
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_clone(args):
    prs, path = load(args.pptx)
    result = op_clone(
        prs,
        {"source": args.src, "at": args.at, "count": args.count or 1},
    )
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_set_text(args):
    prs, path = load(args.pptx)
    spec = {"slide": args.slide, "shape": args.shape}
    if args.paras:
        spec["paras"] = parse_json_arg(args.paras, "--paras")
    else:
        spec["text"] = args.text
    result = op_set_text(prs, spec)
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_set_cell(args):
    prs, path = load(args.pptx)
    result = op_set_cell(
        prs,
        {"slide": args.slide, "shape": args.shape, "row": args.row, "col": args.col, "text": args.text},
    )
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_add_text(args):
    prs, path = load(args.pptx)
    spec = {
        "slide": args.slide,
        "text": args.text,
        "x": args.x,
        "y": args.y,
        "w": args.w,
        "h": args.h,
        "size": args.size,
        "font": args.font,
        "color": args.color,
        "bold": args.bold,
        "align": args.align,
        "valign": args.valign,
        "line_spacing": args.line_spacing,
        "space_after": args.space_after,
        "space_before": args.space_before,
        "indent": args.indent,
        "hanging": args.hanging,
        "bullet": args.bullet,
        "bullet_color": args.bullet_color,
        "fit": args.fit,
    }
    if args.paras:
        spec["paras"] = parse_json_arg(args.paras, "--paras")
        spec.pop("text", None)
    result = op_add_text(prs, spec)
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_add_image(args):
    prs, path = load(args.pptx)
    result = op_add_image(
        prs,
        {
            "slide": args.slide,
            "path": args.path,
            "x": args.x,
            "y": args.y,
            "w": args.w,
            "h": args.h,
        },
    )
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_add_table(args):
    prs, path = load(args.pptx)
    spec = {
        "slide": args.slide,
        "rows": args.rows,
        "cols": args.cols,
        "x": args.x,
        "y": args.y,
        "w": args.w,
        "h": args.h,
        "size": args.size,
        "header": args.header,
        "theme_style": args.theme_style,
        "header_fill": args.header_fill,
        "header_color": args.header_color,
        "zebra": args.zebra,
        "border": args.border,
        "col_widths": args.col_widths,
        "row_height": args.row_height,
        "aligns": args.aligns,
    }
    if args.data:
        data = parse_json_arg(args.data, "--data")
        spec["data"] = data
    result = op_add_table(prs, spec)
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_delete_shape(args):
    prs, path = load(args.pptx)
    result = op_delete_shape(prs, {"slide": args.slide, "shape": args.shape})
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_set_notes(args):
    prs, path = load(args.pptx)
    result = op_set_notes(prs, {"slide": args.slide, "text": args.text})
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_table_style(args):
    prs, path = load(args.pptx)
    spec = {
        "slide": args.slide,
        "shape": args.shape,
        "header": args.header,
        "header_fill": args.header_fill,
        "header_color": args.header_color,
        "body_color": args.body_color,
        "zebra": args.zebra,
        "border": args.border,
        "border_width": args.border_width,
        "body_size": args.body_size,
        "header_size": args.header_size,
        "header_bold": args.header_bold,
        "body_bold": args.body_bold,
        "col_widths": args.col_widths,
        "row_height": args.row_height,
        "aligns": args.aligns,
        "margin": args.margin,
        "valign": args.valign,
        "theme_style": args.theme_style,
        "clear_fill": args.clear_fill,
    }
    result = op_table_style(prs, spec)
    prs2, target = save_and_reload(prs, path, args.out)
    out = deck_meta(prs2, target)
    out.update(result)
    emit(out)


def cmd_apply(args):
    prs, path = load(args.pptx)
    ops = parse_json_arg(args.ops, "--ops")
    if not isinstance(ops, list):
        fail("--ops 必须是 JSON 数组")
    results = []
    errors = []
    for i, op in enumerate(ops):
        try:
            result = apply_one(prs, op)
        except KitError as exc:
            errors.append({"index": i, "op": (op or {}).get("op") if isinstance(op, dict) else None, "error": str(exc)})
            if args.stop_on_error:
                break
        except Exception as exc:  # unexpected backend failure: still report which op
            errors.append(
                {
                    "index": i,
                    "op": (op or {}).get("op") if isinstance(op, dict) else None,
                    "error": f"{type(exc).__name__}: {exc}",
                }
            )
            if args.stop_on_error:
                break
        else:
            results.append({"index": i, **result})

    touched = sorted({r["slide"] for r in results if isinstance(r.get("slide"), int)})
    out = {
        "dry_run": bool(args.dry_run),
        "stop_on_error": bool(args.stop_on_error),
        "op_count": len(ops),
        "applied": len(results),
        "changed_slides": touched,
        "ops": results,
        "errors": errors,
        "ok": not errors,
    }
    if errors and args.stop_on_error:
        out["saved"] = False
        out["hint"] = "批次已中止，文件未修改（原子提交）。修正 errors 后重试。"
        emit(out)
        return
    if args.dry_run:
        out["saved"] = False
        if args.shot and touched:
            try:
                ignore = {s.strip() for s in (args.ignore or "").split(",") if s.strip()}
                pages = [p for p in touched if 1 <= p <= len(prs.slides)]
                warnings = preview_overflow(prs, pages, args.overflow, ignore)
                if warnings:
                    out["overflow"] = warnings
                    out["overflow_preview"] = True
            except Exception as exc:  # rendering must never hide the validation result
                out["shot_error"] = str(exc)
        emit(out)
        return

    prs2, target = save_and_reload(prs, path, args.out)
    out.update(deck_meta(prs2, target))
    out["saved"] = True
    if args.shot and touched:
        try:
            pdf, _key, _cached = deck_pdf(target)
            ignore = {s.strip() for s in (args.ignore or "").split(",") if s.strip()}
            pages = [p for p in touched if 1 <= p <= len(prs2.slides)]
            warnings = render_overflow(prs2, pdf, pages, args.overflow, ignore)
            if warnings:
                out["overflow"] = warnings
        except Exception as exc:  # rendering must never lose the edit result
            out["shot_error"] = str(exc)
    emit(out)


# -------------------------------------------------------------------- doctor
def theme_fonts(prs) -> dict:
    """Read major/minor latin + East-Asian typefaces from the first theme."""
    try:
        theme_part = prs.slide_masters[0].part.part_related_by(RT.THEME)
        root = parse_xml(theme_part.blob)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = root.find(".//a:fontScheme", ns)
        if scheme is None:
            return {}
        out = {}
        for kind in ("major", "minor"):
            node = scheme.find(f"a:{kind}Font", ns)
            if node is None:
                continue
            entry = {}
            for tag in ("latin", "ea", "cs"):
                el = node.find(f"a:{tag}", ns)
                if el is not None and el.get("typeface"):
                    entry[tag] = el.get("typeface")
            scripts = {}
            for el in node.findall("a:font", ns):
                script = el.get("script")
                if script in ("Hans", "Hant", "Jpan", "Hang"):
                    scripts[script] = el.get("typeface")
            if scripts:
                entry["scripts"] = scripts
            out[kind] = entry
        return out
    except Exception:
        return {}


def cmd_doctor(args):
    info = {
        "python": sys.version.split()[0],
        "cache_dir": str(CACHE_ROOT),
    }
    try:
        import pptx

        info["python_pptx"] = pptx.__version__
    except Exception as exc:
        info["python_pptx"] = f"ERROR: {exc}"
    try:
        import pymupdf

        info["pymupdf"] = pymupdf.__version__
    except Exception as exc:
        info["pymupdf"] = f"ERROR: {exc}"
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    info["soffice"] = soffice
    if soffice:
        try:
            result = subprocess.run([soffice, "--version"], capture_output=True, text=True, timeout=30)
            info["soffice_version"] = result.stdout.strip().splitlines()[0] if result.stdout.strip() else ""
        except Exception as exc:
            info["soffice_version"] = f"ERROR: {exc}"
    fonts = {}
    for font in ["微软雅黑", "思源黑体 CN", "宋体", "Noto Sans CJK SC"]:
        try:
            result = subprocess.run(["fc-match", font], capture_output=True, text=True, timeout=10)
            fonts[font] = result.stdout.strip()
        except Exception:
            fonts[font] = "fc-match unavailable"
    info["fonts"] = fonts
    if getattr(args, "pptx", None):
        try:
            prs, _path = load(args.pptx)
            info["deck_theme_fonts"] = theme_fonts(prs)
        except KitError as exc:
            info["deck_theme_fonts"] = f"ERROR: {exc}"
    info["notes"] = [
        "ppt_add_text(font=) writes a:latin + a:ea + a:cs, so CJK uses the same face.",
        "LibreOffice substitutes missing fonts: rendered widths can differ ~5%; leave 5-10% margin.",
    ]
    emit(info)


# ----------------------------------------------------------------------- cli
def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="PPTX agent toolkit")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("inspect", help="版式/形状 JSON")
    p.add_argument("pptx")
    p.add_argument("--slides", help="如 1,3 或 2-4；默认全部")
    p.add_argument("--mode", choices=["full", "summary"], default="full", help="summary 只给每页概览")
    p.add_argument("--depth", type=int, default=99, help="组合递归深度（0 = 不展开子形状）")
    p.set_defaults(func=cmd_inspect)

    p = sub.add_parser("shot", help="渲染页面为 PNG")
    p.add_argument("pptx")
    p.add_argument("-o", "--out", help="输出目录（默认放在缓存里）")
    p.add_argument("--slides", help="如 1,3 或 2-4；默认全部")
    p.add_argument("--dpi", type=int, default=110)
    p.add_argument("--max-pages", type=int, default=0, help="最多返回多少页")
    p.add_argument("--overflow", choices=["loose", "strict", "off"], default="loose",
                   help="溢出启发式：loose 跳过居中文本的左右误报（默认），strict 全部报告，off 关闭")
    p.add_argument("--ignore", help="逗号分隔的形状名，跳过溢出/重叠检查")
    p.add_argument("--montage", type=int, metavar="COLS", help="额外输出拼图，如 3（3 列）")
    p.add_argument("--region", help="局部裁剪 x,y,w,h（英寸）")
    p.set_defaults(func=cmd_shot)

    p = sub.add_parser("create", help="新建/复制模板生成 pptx")
    p.add_argument("out")
    p.add_argument("--template")
    p.add_argument("--width", type=float, help="英寸")
    p.add_argument("--height", type=float, help="英寸")
    p.add_argument("--empty", action="store_true", help="不生成标题页")
    p.add_argument("--overwrite", action="store_true")
    p.set_defaults(func=cmd_create)

    p = sub.add_parser("clone", help="克隆已有页")
    p.add_argument("pptx")
    p.add_argument("src", type=int, help="源页号(1 起)")
    p.add_argument("--at", type=int, help="插入到第几页(1 起)")
    p.add_argument("--count", type=int, default=1)
    p.add_argument("--out")
    p.set_defaults(func=cmd_clone)

    p = sub.add_parser("add-slide", help="按版式新建页")
    p.add_argument("pptx")
    p.add_argument("--layout", required=True, help="版式序号或名称")
    p.add_argument("--out")
    p.set_defaults(func=cmd_add_slide)

    p = sub.add_parser("delete", help="删除页")
    p.add_argument("pptx")
    p.add_argument("--slides", required=True)
    p.add_argument("--out")
    p.set_defaults(func=cmd_delete)

    p = sub.add_parser("set-text", help="替换形状文本（保留原样式，可混排 runs）")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--shape", required=True, help="名称、路径、顶层序号或 #id")
    p.add_argument("--text", help="用 \\n 分段（与 --paras 二选一）")
    p.add_argument("--paras", help="段落/run 级样式 JSON 数组；@file 或 - 读文件/stdin")
    p.add_argument("--out")
    p.set_defaults(func=cmd_set_text)

    p = sub.add_parser("set-cell", help="修改表格单元格")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--shape", required=True, help="名称、路径、顶层序号或 #id")
    p.add_argument("--row", type=int, required=True, help="0 起")
    p.add_argument("--col", type=int, required=True, help="0 起")
    p.add_argument("--text", required=True)
    p.add_argument("--out")
    p.set_defaults(func=cmd_set_cell)

    p = sub.add_parser("add-text", help="新增文本框")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--text", help="文本；\\n 分段")
    p.add_argument("--paras", help="段落/run 级样式 JSON 数组")
    p.add_argument("--x", type=float, required=True)
    p.add_argument("--y", type=float, required=True)
    p.add_argument("--w", type=float, default=6.0)
    p.add_argument("--h", type=float, help="省略时按内容估算并启用 grow")
    p.add_argument("--size", type=float)
    p.add_argument("--font")
    p.add_argument("--color")
    p.add_argument("--bold", action=argparse.BooleanOptionalAction, default=None)
    p.add_argument("--align", help="left|center|right|justify")
    p.add_argument("--valign", help="top|middle|bottom")
    p.add_argument("--line-spacing", type=float)
    p.add_argument("--space-after", type=float, help="段后间距 pt")
    p.add_argument("--space-before", type=float, help="段前间距 pt")
    p.add_argument("--indent", type=float, help="左缩进英寸；负值表示悬挂缩进")
    p.add_argument("--hanging", type=float, help="首行悬挂英寸（默认等于 indent）")
    p.add_argument("--bullet", help="项目符号，如 '•'；传 none 关闭")
    p.add_argument("--bullet-color")
    p.add_argument("--fit", choices=["grow", "shrink", "none"], help="默认：省略 h 时 grow，否则 none")
    p.add_argument("--out")
    p.set_defaults(func=cmd_add_text)

    p = sub.add_parser("add-image", help="插入图片")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--path", required=True)
    p.add_argument("--x", type=float)
    p.add_argument("--y", type=float)
    p.add_argument("--w", type=float)
    p.add_argument("--h", type=float)
    p.add_argument("--out")
    p.set_defaults(func=cmd_add_image)

    p = sub.add_parser("add-table", help="插入表格")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--rows", type=int, required=True)
    p.add_argument("--cols", type=int, required=True)
    p.add_argument("--x", type=float, required=True)
    p.add_argument("--y", type=float, required=True)
    p.add_argument("--w", type=float, required=True)
    p.add_argument("--h", type=float, required=True)
    p.add_argument("--data", help="二维 JSON 数组；@file 亦可")
    p.add_argument("--size", type=float)
    p.add_argument("--header", action=argparse.BooleanOptionalAction, default=True)
    p.add_argument("--theme-style", action=argparse.BooleanOptionalAction, default=False,
                   help="沿用主题表格样式（默认关闭，避免与设计模板打架）")
    p.add_argument("--header-fill")
    p.add_argument("--header-color")
    p.add_argument("--zebra", help="偶数行填充色")
    p.add_argument("--border", help="边框颜色或 none")
    p.add_argument("--col-widths", help="逗号分隔的列宽（英寸）")
    p.add_argument("--row-height", type=float)
    p.add_argument("--aligns", help="逗号分隔的列对齐，如 left,center")
    p.add_argument("--out")
    p.set_defaults(func=cmd_add_table)

    p = sub.add_parser("delete-shape", help="删除形状")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--shape", required=True, help="名称、路径、顶层序号或 #id")
    p.add_argument("--out")
    p.set_defaults(func=cmd_delete_shape)

    p = sub.add_parser("set-notes", help="写讲者备注")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--text", required=True)
    p.add_argument("--out")
    p.set_defaults(func=cmd_set_notes)

    p = sub.add_parser("table-style", help="表格样式（填充/边框/列宽/对齐）")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--shape", required=True, help="表格形状名/路径/#id")
    p.add_argument("--header", action=argparse.BooleanOptionalAction, default=True)
    p.add_argument("--header-fill")
    p.add_argument("--header-color")
    p.add_argument("--body-color")
    p.add_argument("--zebra", help="偶数行填充色")
    p.add_argument("--border", help="边框颜色或 none")
    p.add_argument("--border-width", type=float, default=1.0)
    p.add_argument("--body-size", type=float)
    p.add_argument("--header-size", type=float)
    p.add_argument("--header-bold", action=argparse.BooleanOptionalAction, default=True)
    p.add_argument("--body-bold", action=argparse.BooleanOptionalAction, default=None)
    p.add_argument("--col-widths", help="逗号分隔的列宽（英寸）")
    p.add_argument("--row-height", type=float)
    p.add_argument("--aligns", help="逗号分隔的列对齐，如 left,center")
    p.add_argument("--margin", type=float, help="单元格横向内边距（英寸）")
    p.add_argument("--valign", help="top|middle|bottom")
    p.add_argument("--theme-style", action=argparse.BooleanOptionalAction, default=None,
                   help="是否沿用主题表格样式（默认不改）")
    p.add_argument("--clear-fill", action="store_true", help="清除非表头行的填充")
    p.add_argument("--out")
    p.set_defaults(func=cmd_table_style)

    p = sub.add_parser("apply", help="批量操作（一次事务提交 + 溢出汇总）")
    p.add_argument("pptx")
    p.add_argument("--ops", required=True, help="JSON 数组；@file 或 - 读文件/stdin")
    p.add_argument("--out")
    p.add_argument("--stop-on-error", action=argparse.BooleanOptionalAction, default=True)
    p.add_argument("--dry-run", action="store_true", help="只验证，不写文件")
    p.add_argument("--shot", action=argparse.BooleanOptionalAction, default=True,
                   help="保存后自动渲染改动页并汇总溢出/重叠")
    p.add_argument("--overflow", choices=["loose", "strict", "off"], default="loose")
    p.add_argument("--ignore", help="逗号分隔的形状名，跳过检查")
    p.set_defaults(func=cmd_apply)

    p = sub.add_parser("doctor", help="环境自检")
    p.add_argument("pptx", nargs="?", help="可选：同时读取该文件的主题字体")
    p.set_defaults(func=cmd_doctor)

    return ap


def main() -> None:
    ap = build_parser()
    args = ap.parse_args()
    try:
        args.func(args)
    except SystemExit:
        raise
    except KitError as exc:
        print(f"pptkit {args.cmd} 失败: {exc}", file=sys.stderr)
        raise SystemExit(2)
    except BrokenPipeError:
        pass
    except Exception as exc:
        if os.environ.get("PI_PPTKIT_DEBUG"):
            raise
        print(f"pptkit {args.cmd} 失败: {exc}", file=sys.stderr)
        raise SystemExit(2)


if __name__ == "__main__":
    main()

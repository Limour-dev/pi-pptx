#!/usr/bin/env python3
"""
pptkit — PPTX read/write + single-page rendering backend for the pi-pptx extension.

Layers (see RESEARCH.md):
  * python-pptx      : inspect / clone / edit
  * LibreOffice+PyMuPDF : render slides to PNG and compare text bboxes for overflow
  * JSON on stdout   : easy for the TypeScript tool layer to consume

Every mutating command writes atomically (temp file + fsync-by-replace) and
reopens the result to verify that the deck is still readable.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
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
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 914400  # 1 inch
PT_PER_IN = 72.0
CACHE_ROOT = Path(tempfile.gettempdir()) / "pi-pptx"
MAX_CLONE_COUNT = 50


# --------------------------------------------------------------------- utils
def fail(msg: str) -> None:
    print(msg, file=sys.stderr)
    raise SystemExit(2)


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


def resolve_shape(slide, key: str):
    key = str(key).strip()
    candidates = list(iter_shapes(slide.shapes))
    for shape, path in candidates:
        if path == key or shape.name == key:
            return shape, path
    if key.lstrip("-").isdigit():
        idx = int(key)
        top = list(slide.shapes)
        if -len(top) <= idx < len(top):
            shape = top[idx]
            return shape, shape.name
    # unique leaf-name match (inside a group)
    leaves = [(s, p) for s, p in candidates if s.name == key.rsplit("/", 1)[-1]]
    if len(leaves) == 1:
        return leaves[0]
    names = [p for _, p in candidates][:60]
    fail(f"找不到形状 {key!r}。可用形状路径: {names}")


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
    if shape.left is None or shape.top is None:
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


def shape_info(shape, path: str, top_index: int | None = None, tf=IDENTITY_TF) -> dict:
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
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        child_tf = _group_child_tf(shape, tf)
        info["children"] = [
            shape_info(child, f"{path}/{child.name}", None, child_tf) for child in shape.shapes
        ]
    return info


def slide_info(prs, idx: int) -> dict:
    slide = prs.slides[idx]
    info = {
        "index": idx + 1,
        "layout": slide.slide_layout.name,
        "layout_index": list(prs.slide_layouts).index(slide.slide_layout),
        "shapes": [shape_info(shape, shape.name, i) for i, shape in enumerate(slide.shapes)],
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
    out = deck_meta(prs, path)
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
    out["slides"] = [slide_info(prs, p - 1) for p in pages]
    out["inspected_slides"] = pages
    emit(out)


# ---------------------------------------------------------------- render / shot
def deck_pdf(pptx: Path) -> tuple[Path, str]:
    """Convert the whole deck to PDF, cached by (path, mtime, size)."""
    pptx = pptx.resolve()
    st = pptx.stat()
    key = hashlib.sha1(f"{pptx}|{st.st_mtime_ns}|{st.st_size}".encode()).hexdigest()[:16]
    cache = CACHE_ROOT / key
    pdf = cache / (pptx.stem + ".pdf")
    if pdf.exists():
        return pdf, key

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
    return pdf, key


def _norm_text(value: str) -> str:
    return "".join(value.split())


def detect_overflow(slide, page, slide_w_pt: float, slide_h_pt: float) -> list[dict]:
    """Compare rendered text-block bboxes with the shapes' declared boxes.

    Only top-level shapes with text are checked: group children store relative
    coordinates, which do not map onto the rendered page.
    """
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
        txt = shape.text_frame.text.strip()
        if not txt or shape.left is None:
            continue
        box = (
            to_pt(shape.left) * sx,
            to_pt(shape.top) * sy,
            to_pt(shape.left + shape.width) * sx,
            to_pt(shape.top + shape.height) * sy,
        )
        shapes.append((shape.name, _norm_text(txt), box, box_in(shape)))

    blocks = []
    for block in page.get_text("blocks"):
        if len(block) > 6 and block[6] != 0:
            continue
        text = _norm_text(block[4] if len(block) > 4 else "")
        if not text:
            continue
        blocks.append((block[0], block[1], block[2], block[3], text))

    found = []
    for name, txt, (x0, y0, x1, y1), box_inches in shapes:
        matched = []
        for bx0, by0, bx1, by1, btext in blocks:
            # Accept blocks that overlap the shape or that start at its top
            # edge (overflowing text grows downward, so its center can be far
            # outside a deliberately small box).
            overlaps = not (bx1 < x0 - 8 or bx0 > x1 + 8 or by1 < y0 - 8 or by0 > y1 + 8)
            starts_near = (
                y0 - 8 <= by0 <= y1 + 12
                and not (bx1 < x0 - 6 or bx0 > x1 + 6)
            )
            if not (overlaps or starts_near):
                continue
            if not (
                btext in txt
                or txt in btext
                or btext.startswith(txt[:8])
                or txt.startswith(btext[:8])
            ):
                continue
            matched.append((bx0, by0, bx1, by1))
        if not matched:
            continue
        ux0 = min(b[0] for b in matched)
        uy0 = min(b[1] for b in matched)
        ux1 = max(b[2] for b in matched)
        uy1 = max(b[3] for b in matched)
        # Rendered line boxes include leading/descender padding, so a small
        # vertical spill is normal. A real overflow is roughly one extra line.
        vtol = max(10.0, 0.35 * max(b[3] - b[1] for b in matched))
        htol = 6.0
        over = {}
        if ux1 > x1 + htol:
            over["right_pt"] = round(ux1 - x1, 1)
        if uy1 > y1 + vtol:
            over["bottom_pt"] = round(uy1 - y1, 1)
        if ux0 < x0 - htol:
            over["left_pt"] = round(x0 - ux0, 1)
        if uy0 < y0 - vtol:
            over["top_pt"] = round(y0 - uy0, 1)
        if over:
            found.append(
                {
                    "shape": name,
                    "box_in": box_inches,
                    "overflow_pt": over,
                    "text": txt[:40],
                }
            )
    return found


def cmd_shot(args):
    import pymupdf

    prs, path = load(args.pptx)
    total = len(prs.slides)
    pages = parse_pages(args.slides, total)
    pdf, key = deck_pdf(path)
    doc = pymupdf.open(pdf)
    if doc.page_count < total:
        fail(f"PDF 只有 {doc.page_count} 页，但 PPTX 有 {total} 页（转换异常）")

    outdir = Path(args.out).expanduser() if args.out else CACHE_ROOT / key / f"shots-{args.dpi}"
    outdir.mkdir(parents=True, exist_ok=True)

    slide_w_pt = prs.slide_width / EMU * PT_PER_IN
    slide_h_pt = prs.slide_height / EMU * PT_PER_IN
    shots = []
    for page_no in pages:
        page = doc[page_no - 1]
        target = outdir / f"slide-{page_no}.png"
        if not target.exists():
            page.get_pixmap(dpi=args.dpi).save(str(target))
        pix = page.get_pixmap(dpi=args.dpi)
        entry = {
            "slide": page_no,
            "file": str(target.resolve()),
            "px": [pix.width, pix.height],
        }
        overflow = detect_overflow(prs.slides[page_no - 1], page, slide_w_pt, slide_h_pt)
        if overflow:
            entry["overflow"] = overflow
        shots.append(entry)
    doc.close()

    omitted = []
    if args.max_pages and len(shots) > args.max_pages:
        omitted = [s["slide"] for s in shots[args.max_pages :]]
        shots = shots[: args.max_pages]

    result = {
        "pdf": str(pdf.resolve()),
        "dpi": args.dpi,
        "shots": shots,
    }
    if omitted:
        result["omitted_slides"] = omitted
        result["hint"] = f"一次最多渲染 {args.max_pages} 页；其余页请再调用一次。"
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


def cmd_clone(args):
    prs, path = load(args.pptx)
    total = len(prs.slides)
    if not 1 <= args.src <= total:
        fail(f"源页号 {args.src} 超出 1..{total}")
    count = args.count or 1
    if not 1 <= count <= MAX_CLONE_COUNT:
        fail(f"克隆数量应在 1..{MAX_CLONE_COUNT}，收到 {count}")
    if args.at is not None and not 1 <= args.at <= total + 1:
        fail(f"插入位置 {args.at} 超出 1..{total + 1}")
    clone_slide(prs, args.src - 1, args.at, count)
    target = save_deck(prs, path, args.out)
    prs2 = Presentation(str(target))
    out = deck_meta(prs2, target)
    out.update({"created_from": args.src, "count": count, "at": args.at})
    emit(out)


# ------------------------------------------------------------------ deck edits
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
    layout = None
    key = str(args.layout).strip()
    if key.lstrip("-").isdigit():
        idx = int(key)
        if not 0 <= idx < len(prs.slide_layouts):
            fail(f"版式序号 {idx} 超出 0..{len(prs.slide_layouts) - 1}")
        layout = prs.slide_layouts[idx]
    else:
        for cand in prs.slide_layouts:
            if cand.name == key:
                layout = cand
                break
        if layout is None:
            names = [f"{i}:{l.name}" for i, l in enumerate(prs.slide_layouts)]
            fail(f"找不到版式 {key!r}。可用版式: {names}")
    slide = prs.slides.add_slide(layout)
    target = save_deck(prs, path, args.out)
    prs2 = Presentation(str(target))
    out = deck_meta(prs2, target)
    out.update(
        {
            "added_slide": len(prs2.slides),
            "layout": layout.name,
            "placeholders": [ph.name for ph in slide.placeholders],
        }
    )
    emit(out)


def cmd_delete(args):
    prs, path = load(args.pptx)
    total = len(prs.slides)
    pages = parse_pages(args.slides, total)
    if len(pages) >= total:
        fail("不能删除全部页（至少保留 1 页）")
    sld_id_lst = prs.slides._sldIdLst
    ids = list(sld_id_lst)
    for page in sorted(pages, reverse=True):
        sld_id_lst.remove(ids[page - 1])
    target = save_deck(prs, path, args.out)
    prs2 = Presentation(str(target))
    out = deck_meta(prs2, target)
    out["deleted"] = pages
    emit(out)


# ---------------------------------------------------------------- text editing
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


def cmd_set_text(args):
    prs, path = load(args.pptx)
    total = len(prs.slides)
    if not 1 <= args.slide <= total:
        fail(f"页号 {args.slide} 超出 1..{total}")
    slide = prs.slides[args.slide - 1]
    shape, shape_path = resolve_shape(slide, args.shape)
    if not shape.has_text_frame:
        if getattr(shape, "has_table", False) and shape.has_table:
            fail(f"{shape_path} 是表格，请用 set-cell 修改单元格")
        fail(f"{shape_path} 没有文本框，无法设置文本")
    fill_paragraphs(shape.text_frame, args.text)
    target = save_deck(prs, path, args.out)
    prs2 = Presentation(str(target))
    slide2 = prs2.slides[args.slide - 1]
    shape2, path2 = resolve_shape(slide2, shape_path)
    out = deck_meta(prs2, target)
    out.update(
        {
            "slide": args.slide,
            "shape": shape2.name,
            "path": path2,
            "text": shape2.text_frame.text,
            "paragraphs": len(shape2.text_frame.paragraphs),
        }
    )
    emit(out)


def cmd_set_cell(args):
    prs, path = load(args.pptx)
    total = len(prs.slides)
    if not 1 <= args.slide <= total:
        fail(f"页号 {args.slide} 超出 1..{total}")
    slide = prs.slides[args.slide - 1]
    shape, shape_path = resolve_shape(slide, args.shape)
    if not (getattr(shape, "has_table", False) and shape.has_table):
        fail(f"{shape_path} 不是表格")
    table = shape.table
    nrows, ncols = len(table.rows), len(table.columns)
    if not (0 <= args.row < nrows and 0 <= args.col < ncols):
        fail(f"单元格 ({args.row},{args.col}) 超出表格范围 {nrows}x{ncols}（0 起）")
    cell = table.cell(args.row, args.col)
    old = cell.text_frame.paragraphs[0]
    if old.runs:
        old.runs[0].text = args.text
        for run in old.runs[1:]:
            run._r.getparent().remove(run._r)
        for para in list(cell.text_frame.paragraphs)[1:]:
            para._p.getparent().remove(para._p)
    else:
        cell.text = args.text
    target = save_deck(prs, path, args.out)
    prs2 = Presentation(str(target))
    out = deck_meta(prs2, target)
    out.update(
        {
            "slide": args.slide,
            "shape": shape.name,
            "row": args.row,
            "col": args.col,
            "text": args.text,
        }
    )
    emit(out)


# -------------------------------------------------------------- add elements
def _apply_run_style(run, args):
    if getattr(args, "size", None):
        run.font.size = Pt(args.size)
    if getattr(args, "font", None):
        run.font.name = getattr(args, "font")
    if getattr(args, "color", None):
        run.font.color.rgb = parse_rgb(args.color)
    if getattr(args, "bold", None) is not None:
        run.font.bold = args.bold


def cmd_add_text(args):
    prs, path = load(args.pptx)
    total = len(prs.slides)
    if not 1 <= args.slide <= total:
        fail(f"页号 {args.slide} 超出 1..{total}")
    slide = prs.slides[args.slide - 1]
    box = slide.shapes.add_textbox(Inches(args.x), Inches(args.y), Inches(args.w), Inches(args.h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    align = parse_enum(args.align, ALIGN, "对齐方式")
    valign = parse_enum(args.valign, VALIGN, "垂直对齐")
    if valign is not None:
        tf.vertical_anchor = valign

    lines = args.text.split("\n")
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if align is not None:
            para.alignment = align
        if args.line_spacing:
            para.line_spacing = args.line_spacing
        run = para.add_run()
        run.text = line
        _apply_run_style(run, args)

    target = save_deck(prs, path, args.out)
    prs2 = Presentation(str(target))
    box2 = prs2.slides[args.slide - 1].shapes[-1]
    out = deck_meta(prs2, target)
    out.update(
        {
            "slide": args.slide,
            "added": "textbox",
            "name": box2.name,
            "path": box2.name,
            "box_in": box_in(box2),
            "paragraphs": len(lines),
        }
    )
    emit(out)


def cmd_add_image(args):
    prs, path = load(args.pptx)
    total = len(prs.slides)
    if not 1 <= args.slide <= total:
        fail(f"页号 {args.slide} 超出 1..{total}")
    image_path = Path(args.path).expanduser()
    if not image_path.is_absolute():
        image_path = (Path.cwd() / image_path).resolve()
    if not image_path.exists():
        fail(f"图片不存在: {image_path}")
    slide = prs.slides[args.slide - 1]
    kwargs = {}
    if args.x is not None and args.y is not None:
        kwargs["left"] = Inches(args.x)
        kwargs["top"] = Inches(args.y)
    if args.w is not None:
        kwargs["width"] = Inches(args.w)
    if args.h is not None:
        kwargs["height"] = Inches(args.h)
    picture = slide.shapes.add_picture(str(image_path), **kwargs)
    target = save_deck(prs, path, args.out)
    prs2 = Presentation(str(target))
    pic2 = prs2.slides[args.slide - 1].shapes[-1]
    out = deck_meta(prs2, target)
    out.update(
        {
            "slide": args.slide,
            "added": "picture",
            "name": pic2.name,
            "path": pic2.name,
            "box_in": box_in(pic2),
        }
    )
    emit(out)


def cmd_add_table(args):
    prs, path = load(args.pptx)
    total = len(prs.slides)
    if not 1 <= args.slide <= total:
        fail(f"页号 {args.slide} 超出 1..{total}")
    if args.rows < 1 or args.cols < 1:
        fail("表格至少 1 行 1 列")
    data = []
    if args.data:
        try:
            data = json.loads(args.data)
        except json.JSONDecodeError as exc:
            fail(f"--data 不是合法 JSON: {exc}")
        if not isinstance(data, list) or any(not isinstance(row, list) for row in data):
            fail("--data 必须是二维数组")
        if len(data) > args.rows or any(len(row) > args.cols for row in data):
            fail(f"--data 尺寸 {len(data)}x{max((len(r) for r in data), default=0)} 超过 {args.rows}x{args.cols}")
    slide = prs.slides[args.slide - 1]
    frame = slide.shapes.add_table(
        args.rows, args.cols, Inches(args.x), Inches(args.y), Inches(args.w), Inches(args.h)
    )
    table = frame.table
    table.first_row = bool(args.header)
    for r in range(args.rows):
        for c in range(args.cols):
            cell = table.cell(r, c)
            value = ""
            if r < len(data) and c < len(data[r]):
                value = "" if data[r][c] is None else str(data[r][c])
            cell.text = value
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    if args.size:
                        run.font.size = Pt(args.size)
                    if args.header and r == 0:
                        run.font.bold = True
    target = save_deck(prs, path, args.out)
    prs2 = Presentation(str(target))
    shape2 = prs2.slides[args.slide - 1].shapes[-1]
    out = deck_meta(prs2, target)
    out.update(
        {
            "slide": args.slide,
            "added": "table",
            "name": shape2.name,
            "path": shape2.name,
            "box_in": box_in(shape2),
            "rows": args.rows,
            "cols": args.cols,
        }
    )
    emit(out)


# -------------------------------------------------------------------- doctor
def cmd_doctor(_args):
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
    emit(info)


# ----------------------------------------------------------------------- cli
def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="PPTX agent toolkit")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("inspect", help="版式/形状 JSON")
    p.add_argument("pptx")
    p.add_argument("--slides", help="如 1,3 或 2-4；默认全部")
    p.set_defaults(func=cmd_inspect)

    p = sub.add_parser("shot", help="渲染页面为 PNG")
    p.add_argument("pptx")
    p.add_argument("-o", "--out", help="输出目录（默认放在缓存里）")
    p.add_argument("--slides", help="如 1,3 或 2-4；默认全部")
    p.add_argument("--dpi", type=int, default=110)
    p.add_argument("--max-pages", type=int, default=0, help="最多返回多少页")
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

    p = sub.add_parser("set-text", help="替换形状文本（保留原样式）")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--shape", required=True, help="名称、路径或顶层序号")
    p.add_argument("--text", required=True, help="用 \\n 分段")
    p.add_argument("--out")
    p.set_defaults(func=cmd_set_text)

    p = sub.add_parser("set-cell", help="修改表格单元格")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--shape", required=True)
    p.add_argument("--row", type=int, required=True, help="0 起")
    p.add_argument("--col", type=int, required=True, help="0 起")
    p.add_argument("--text", required=True)
    p.add_argument("--out")
    p.set_defaults(func=cmd_set_cell)

    p = sub.add_parser("add-text", help="新增文本框")
    p.add_argument("pptx")
    p.add_argument("--slide", type=int, required=True)
    p.add_argument("--text", required=True)
    p.add_argument("--x", type=float, required=True)
    p.add_argument("--y", type=float, required=True)
    p.add_argument("--w", type=float, default=6.0)
    p.add_argument("--h", type=float, default=0.8)
    p.add_argument("--size", type=float)
    p.add_argument("--font")
    p.add_argument("--color")
    p.add_argument("--bold", action=argparse.BooleanOptionalAction, default=None)
    p.add_argument("--align", help="left|center|right|justify")
    p.add_argument("--valign", help="top|middle|bottom")
    p.add_argument("--line-spacing", type=float)
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
    p.add_argument("--data", help="二维 JSON 数组")
    p.add_argument("--size", type=float)
    p.add_argument("--header", action=argparse.BooleanOptionalAction, default=True)
    p.add_argument("--out")
    p.set_defaults(func=cmd_add_table)

    p = sub.add_parser("doctor", help="环境自检")
    p.set_defaults(func=cmd_doctor)

    return ap


def main() -> None:
    ap = build_parser()
    args = ap.parse_args()
    try:
        args.func(args)
    except SystemExit:
        raise
    except BrokenPipeError:
        pass
    except Exception as exc:
        if os.environ.get("PI_PPTKIT_DEBUG"):
            raise
        print(f"pptkit {args.cmd} 失败: {exc}", file=sys.stderr)
        raise SystemExit(2)


if __name__ == "__main__":
    main()

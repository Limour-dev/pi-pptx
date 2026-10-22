"""LibreOffice rendering, PNG/PDF caching, montage and layout heuristics.

Two heuristics run on the rendered PDF: text overflow (rendered text bbox vs.
declared shape box) and visible shape overlap. Both are advisory signals.
"""
from __future__ import annotations

import hashlib
import math
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN

from .utils import (
    CACHE_ROOT,
    EMU,
    PT_PER_IN,
    box_in,
    fail,
    parse_pages,
    read_only_ignore,
    to_pt,
)

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


def _vertical_anchor(shape) -> str:
    """Effective text anchor: 'top' | 'center' | 'bottom'.

    DrawingML defaults to top when ``a:bodyPr/@anchor`` is absent, and
    python-pptx reports that as ``None``.
    """
    try:
        anchor = shape.text_frame.vertical_anchor
    except Exception:
        return "top"
    if anchor is None:
        return "top"
    name = str(anchor).split(" ")[0].upper()
    if name == "TOP":
        return "top"
    if name == "BOTTOM":
        return "bottom"
    return "center"


def _text_shapes(slide, ignore, slide_w_pt: float, page):
    """Top-level text shapes with their declared box (pt) and text."""
    sx = page.rect.width / slide_w_pt if slide_w_pt else 1.0
    sy = page.rect.height / _slide_h_pt(slide_w_pt, page) if slide_w_pt else 1.0
    shapes = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if getattr(shape, "rotation", 0):
            continue
        if not shape.has_text_frame:
            continue
        if ignore.matches(shape, shape.name):
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
        shapes.append(
            {
                "shape": shape,
                "name": shape.name,
                "id": shape.shape_id,
                "text": txt,
                "box_pt": box,
                "box_in": box_in(shape),
                "aligns": aligns,
                "anchor": _vertical_anchor(shape),
            }
        )
    return shapes


def _slide_h_pt(slide_w_pt: float, page) -> float:
    """Slide height in pt inferred from the page geometry (aspect preserved)."""
    if not slide_w_pt or not page.rect.width:
        return page.rect.height
    return slide_w_pt * page.rect.height / page.rect.width


def _rendered_lines(page) -> list[tuple[float, float, float, float, str]]:
    """Rendered text per *line*.

    Two separate text boxes on the same baseline can land in one PyMuPDF block,
    which used to make a short label inherit its neighbour's width (and report a
    bogus 400pt horizontal overflow).
    """
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
    return blocks


def _assign_lines(shapes, blocks) -> dict[int, list[tuple[float, float, float, float]]]:
    """Assign each rendered line to at most one shape (best text match, then top edge)."""
    assigned: dict[int, list[tuple[float, float, float, float]]] = {i: [] for i in range(len(shapes))}
    for bx0, by0, bx1, by1, btext in blocks:
        best = None
        for i, entry in enumerate(shapes):
            x0, y0, x1, y1 = entry["box_pt"]
            if bx1 < x0 - 6 or bx0 > x1 + 6:
                continue
            overlaps_y = not (by1 < y0 - 8 or by0 > y1 + 8)
            starts_near = y0 - 8 <= by0 <= y1 + 12
            if not (overlaps_y or starts_near):
                continue
            score = _text_match_score(btext, entry["text"])
            if score <= 0:
                continue
            candidate = (score, -abs(by0 - y0), i)
            if best is None or candidate > best:
                best = candidate
        if best is not None:
            assigned[best[2]].append((bx0, by0, bx1, by1))
    return assigned


def rendered_text_boxes(slide, page, slide_w_pt: float, ignore=None) -> dict[int, list[float]]:
    """shape id -> bbox of the text actually rendered, in inches.

    A placeholder's declared box is often much larger than its text (a title
    placeholder spanning the whole page), so collision checks must use what is
    on screen, not the inherited geometry.
    """
    ignore = read_only_ignore(ignore)
    shapes = _text_shapes(slide, ignore, slide_w_pt, page)
    if not shapes:
        return {}
    assigned = _assign_lines(shapes, _rendered_lines(page))
    scale = EMU / PT_PER_IN
    boxes = {}
    for index, matched in assigned.items():
        if not matched:
            continue
        ux0 = min(b[0] for b in matched)
        uy0 = min(b[1] for b in matched)
        ux1 = max(b[2] for b in matched)
        uy1 = max(b[3] for b in matched)
        boxes[shapes[index]["id"]] = [
            round(ux0 * scale, 3),
            round(uy0 * scale, 3),
            round((ux1 - ux0) * scale, 3),
            round((uy1 - uy0) * scale, 3),
        ]
    return boxes


def detect_overflow(
    slide,
    page,
    slide_w_pt: float,
    slide_h_pt: float,
    mode: str = "loose",
    ignore=None,
) -> tuple[list[dict], list[dict]]:
    """Compare rendered text bboxes with the shapes' declared boxes.

    Returns ``(overflow, suppressed)``. ``suppressed`` lists the sides that
    *would* have been reported but are impossible by construction, so a caller
    can tell "clean" apart from "not checked".

    Only top-level shapes with text are checked: group children store relative
    coordinates, which do not map onto the rendered page.

    Each rendered text block is assigned to at most one shape (best text match,
    then nearest top edge), which avoids a neighbouring label dragging the
    measured bbox of a short shape sideways.

    A side is skipped when the text anchor makes it impossible:
      * anchor=top    -> the first line cannot spill above the box (``top``)
      * anchor=bottom -> the last line cannot spill below the box (``bottom``)
    Rendered line boxes carry ascender/leading padding, so those two sides used
    to fire on every large first line of a top-anchored title.

    ``mode``:
      * loose (default): skip left/right checks for centred/justified text —
        that is where the old heuristic produced most false positives.
      * strict: report every side with tight tolerances.
      * off: skip text overflow checks entirely.
    """
    if mode == "off":
        return [], []
    ignore = read_only_ignore(ignore)
    shapes = _text_shapes(slide, ignore, slide_w_pt, page)
    if not shapes:
        return [], []
    assigned = _assign_lines(shapes, _rendered_lines(page))

    found = []
    suppressed = []
    for i, entry in enumerate(shapes):
        matched = assigned[i]
        if not matched:
            continue
        name, sid, txt, anchor = entry["name"], entry["id"], entry["text"], entry["anchor"]
        box_inches, aligns = entry["box_in"], entry["aligns"]
        x0, y0, x1, y1 = entry["box_pt"]
        ux0 = min(b[0] for b in matched)
        uy0 = min(b[1] for b in matched)
        ux1 = max(b[2] for b in matched)
        uy1 = max(b[3] for b in matched)
        # Rendered line boxes include leading/descender padding, so a small
        # vertical spill is normal. A real overflow is roughly one extra line.
        line_height = max(b[3] - b[1] for b in matched)
        if mode == "strict":
            # Font substitution changes ascent/descent a little, so "strict"
            # still tolerates a fraction of one line — a real overflow is a
            # whole extra line, far above this.
            vtol = max(4.0, 0.22 * line_height)
            htol = 2.0
        else:
            vtol = max(10.0, 0.35 * line_height)
            htol = 6.0
        over = {}
        hidden = []
        centered = any(a in (PP_ALIGN.CENTER, PP_ALIGN.JUSTIFY) for a in aligns)
        horizontal = mode == "strict" or not centered
        if horizontal and ux1 > x1 + htol:
            over["right_pt"] = round(ux1 - x1, 1)
        if horizontal and ux0 < x0 - htol:
            over["left_pt"] = round(x0 - ux0, 1)
        if uy1 > y1 + vtol:
            if anchor == "bottom":
                hidden.append({"side": "bottom", "spill_pt": round(uy1 - y1, 1), "reason": "anchor=b"})
            else:
                over["bottom_pt"] = round(uy1 - y1, 1)
        elif anchor == "bottom" and uy1 > y1 + 0.5:
            hidden.append({"side": "bottom", "spill_pt": round(uy1 - y1, 1), "reason": "anchor=b"})
        if uy0 < y0 - vtol:
            if anchor == "top":
                hidden.append({"side": "top", "spill_pt": round(y0 - uy0, 1), "reason": "anchor=t"})
            else:
                over["top_pt"] = round(y0 - uy0, 1)
        elif anchor == "top" and uy0 < y0 - 0.5:
            hidden.append({"side": "top", "spill_pt": round(y0 - uy0, 1), "reason": "anchor=t"})
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
        if hidden:
            suppressed.append(
                {
                    "shape": name,
                    "id": sid,
                    "sides": hidden,
                    "text": txt[:40],
                }
            )
    return found, suppressed


CONTAINER_PLACEHOLDERS = {
    "BODY",
    "OBJECT",
    "CONTENT",
    "PICTURE",
    "TABLE",
    "CHART",
    "MEDIA",
}


def _covers(outer: list[float], inner: list[float]) -> float:
    """Fraction of ``inner`` (x,y,w,h in inches) that ``outer`` covers."""
    ox, oy, ow, oh = outer
    ix, iy, iw, ih = inner
    if iw <= 0 or ih <= 0:
        return 0.0
    cx = max(0.0, min(ox + ow, ix + iw) - max(ox, ix))
    cy = max(0.0, min(oy + oh, iy + ih) - max(oy, iy))
    return (cx * cy) / (iw * ih)


def _container_of(a: dict, b: dict) -> str | None:
    """Return a reason when one shape is a container that lays out the other.

    Templates place pictures/text boxes *inside* a body placeholder on purpose;
    judging container vs. content as a collision makes every such slide warn.
    Only shapes that can legitimately act as containers (placeholders and
    autoshapes) are exempt, and only when one box really encloses the other.
    """
    for outer, inner in ((a, b), (b, a)):
        if not outer["containerish"]:
            continue
        if _covers(outer["box"], inner["box"]) >= 0.92:
            return f"{outer['name']} 是容器"
    return None


def detect_overlaps(slide, ignore=None, text_boxes: dict[int, list[float]] | None = None) -> tuple[list[dict], list[dict]]:
    """Report text boxes that visibly collide with other text/pictures/tables.

    ``text_boxes`` (shape id -> rendered bbox) replaces the declared box of text
    shapes: a title placeholder's declared box can cover half the page while its
    text is one line, and comparing boxes flagged every card placed below it.
    Returns ``(overlaps, suppressed)``; ``suppressed`` holds container/content
    pairs that are layout by design.
    """
    ignore = read_only_ignore(ignore)
    text_boxes = text_boxes or {}
    items = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if getattr(shape, "rotation", 0):
            continue
        if ignore.matches(shape, shape.name):
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
        rendered = text_boxes.get(shape.shape_id) if kind == "text" else None
        if rendered and rendered[2] > 0.02 and rendered[3] > 0.02:
            box = rendered
        x, y, w, h = box
        # Only body/object-style placeholders and autoshapes may act as a
        # container. A title placeholder that happens to cover a text box is a
        # real collision, not a layout intent.
        containerish = shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
        if shape.is_placeholder:
            ph_type = str(shape.placeholder_format.type).split(" ")[0]
            containerish = containerish or ph_type in CONTAINER_PLACEHOLDERS
        items.append(
            {
                "name": shape.name,
                "id": shape.shape_id,
                "kind": kind,
                "box": box,
                "area": max(w * h, 0.0001),
                "containerish": containerish,
            }
        )

    found = []
    suppressed = []
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
            reason = _container_of(a, b)
            entry = {
                "a": a["name"],
                "a_id": a["id"],
                "b": b["name"],
                "b_id": b["id"],
                "kinds": [a["kind"], b["kind"]],
                "area_in2": round(area, 2),
                "ratio": round(ratio, 2),
            }
            if reason:
                suppressed.append({**entry, "reason": reason})
            else:
                found.append(entry)
    return found[:20], suppressed[:20]


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


def check_slide(slide, page, slide_w_pt: float, slide_h_pt: float, mode: str = "loose", ignore=None) -> dict:
    """Run both heuristics for one rendered slide.

    Returns ``{"overflow": [...], "overlaps": [...], "suppressed": [...]}``;
    empty lists are omitted by the callers that build the report.
    """
    overflow, hidden = detect_overflow(slide, page, slide_w_pt, slide_h_pt, mode, ignore)
    text_boxes = rendered_text_boxes(slide, page, slide_w_pt, ignore)
    overlaps, containers = detect_overlaps(slide, ignore, text_boxes)
    entry: dict = {}
    if overflow:
        entry["overflow"] = overflow
    if overlaps:
        entry["overlaps"] = overlaps
    suppressed = [{"kind": "overflow", **item} for item in hidden] + [
        {"kind": "overlap", **item} for item in containers
    ]
    if suppressed:
        entry["suppressed"] = suppressed
    return entry


def render_overflow(prs, pdf: Path, pages: list[int], mode: str = "loose", ignore=None) -> list[dict]:
    """Heuristics for ``pages`` of an already-rendered deck (no PNG writes)."""
    import pymupdf

    doc = pymupdf.open(pdf)
    slide_w_pt = prs.slide_width / EMU * PT_PER_IN
    slide_h_pt = prs.slide_height / EMU * PT_PER_IN
    out = []
    try:
        for page_no in pages:
            if not 1 <= page_no <= doc.page_count:
                continue
            entry = check_slide(
                prs.slides[page_no - 1], doc[page_no - 1], slide_w_pt, slide_h_pt, mode, ignore
            )
            if entry:
                out.append({"slide": page_no, **entry})
    finally:
        doc.close()
    return out


def preview_overflow(prs, pages: list[int], mode: str = "loose", ignore=None) -> list[dict]:
    """Render an in-memory deck to a throwaway file to project overflow (dry run)."""
    work = Path(tempfile.mkdtemp(prefix="preview-", dir=CACHE_ROOT if CACHE_ROOT.exists() else None))
    try:
        tmp = work / "preview.pptx"
        prs.save(str(tmp))
        pdf, _key, _cached = deck_pdf(tmp)
        return render_overflow(prs, pdf, pages, mode, ignore)
    finally:
        shutil.rmtree(work, ignore_errors=True)


# --------------------------------------------------------------- shot payload
def _region_tuple(value) -> list[float] | None:
    if value is None:
        return None
    if isinstance(value, str):
        value = [v for v in value.replace(" ", "").split(",") if v]
    try:
        vals = [float(v) for v in value]
    except (TypeError, ValueError):
        fail(f"region 需要 4 个数字 x,y,w,h（英寸），收到: {value}")
    if len(vals) != 4:
        fail(f"region 需要 4 个数字 x,y,w,h（英寸），收到: {value}")
    return vals


def _region_spec(spec):
    if isinstance(spec, str):
        text = spec.strip()
        if text.startswith("{"):
            import json

            try:
                return json.loads(text)
            except json.JSONDecodeError as exc:
                fail(f"region 不是合法 JSON: {exc}")
    return spec


def region_for_page(spec, page_no: int) -> list[float] | None:
    """Resolve a region spec that may be global or per page.

    Accepted: ``[x,y,w,h]`` (every page), ``"x,y,w,h"``, or a dict such as
    ``{"3": [x,y,w,h], "default": [x,y,w,h]}``. Per-page crops matter when only
    one slide needs a magnified table.
    """
    spec = _region_spec(spec)
    if spec is None:
        return None
    if isinstance(spec, dict):
        by_page = {str(k): v for k, v in spec.items()}
        value = by_page.get(str(page_no), by_page.get("default"))
        return _region_tuple(value)
    return _region_tuple(spec)


def estimate_capacity(slide) -> int:
    """Rough declared line capacity of one slide (screening number).

    Capacity comes from the shapes' boxes and their first run's font size
    (falling back to 18pt). The *rendered* line count is the measurement; this
    only answers "could this page ever hold that much text".
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE as _TYPE

    capacity = 0
    for shape in slide.shapes:
        if shape.shape_type in (_TYPE.GROUP, _TYPE.PICTURE):
            continue
        if not shape.has_text_frame or shape.height is None:
            continue
        text = shape.text_frame.text.strip()
        if not text:
            continue
        size = 18.0
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size is not None:
                    size = run.font.size.pt
                    break
            if size != 18.0:
                break
        line_h = max(6.0, size * 1.35)
        capacity += max(1, int((shape.height / EMU * PT_PER_IN) // line_h))
    return capacity


def slide_density(slide, page, slide_w_pt: float, slide_h_pt: float) -> dict:
    """First text line + rendered/expected line counts (a scanning table)."""
    rendered = 0
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            if "".join(span.get("text", "") for span in line.get("spans", [])).strip():
                rendered += 1
    capacity = estimate_capacity(slide)
    info: dict = {"slide": int(page.number) + 1}
    first = ""
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            first = shape.text_frame.text.strip().split("\n")[0][:40]
            break
    if first:
        info["first_line"] = first
    info["lines"] = rendered
    if capacity:
        info["capacity"] = capacity
        if rendered > capacity:
            info["dense"] = True
    return info


def changed_slides_since_last_shot(prs, path: Path) -> list[int]:
    """Slide numbers whose XML changed since the last shot of this file.

    Each slide is hashed separately, so an edit loop can render one page instead
    of the deck. The snapshot is kept per file (not per content hash) so it
    survives saves. First call on a deck returns every slide.
    """
    import json

    state_dir = CACHE_ROOT / ("state-" + hashlib.sha1(str(Path(path).resolve()).encode()).hexdigest()[:12])
    snapshot = state_dir / "slides.json"
    current = [
        hashlib.sha1(slide.element.xml.encode("utf-8")).hexdigest() for slide in prs.slides
    ]
    previous: list[str] = []
    if snapshot.exists():
        try:
            previous = json.loads(snapshot.read_text(encoding="utf-8"))
        except Exception:
            previous = []
    changed = [
        index + 1
        for index, digest in enumerate(current)
        if index >= len(previous) or previous[index] != digest
    ]
    try:
        state_dir.mkdir(parents=True, exist_ok=True)
        snapshot.write_text(json.dumps(current), encoding="utf-8")
    except Exception:
        pass
    return changed


def shot_deck(
    prs,
    path: Path,
    *,
    slides=None,
    dpi: int = 110,
    out: str | None = None,
    max_pages: int = 0,
    montage: int | None = None,
    region=None,
    overflow: str = "loose",
    ignore=None,
    only_changed: bool = False,
    changed_slides: list[int] | None = None,
    density: bool = False,
) -> dict:
    """Render pages to PNG and return the report payload (no printing).

    ``only_changed`` diffs against ``changed_slides`` (the pages an edit
    touched) so an iteration loop can render just those pages instead of
    re-shooting the whole deck.
    """
    import pymupdf

    total = len(prs.slides)
    detected: list[int] = []
    if only_changed:
        changed = sorted({int(p) for p in (changed_slides or []) if 1 <= int(p) <= total})
        if not changed:
            # No caller-provided list: diff slide XML against the last shot.
            detected = changed_slides_since_last_shot(prs, path)
            changed = detected
        if not changed:
            result = {
                "dpi": dpi,
                "only_changed": True,
                "cache_hit": True,
                "overflow_mode": overflow,
                "shot_slides": [],
                "shots": [],
                "no_changes": True,
                "hint": "与上次渲染相比没有任何页面变化，未重新渲染。用 only_changed=false 强制全量。",
            }
            return result
        pages = parse_pages(slides, total) if slides else changed
        pages = [p for p in pages if p in changed] or changed
    else:
        pages = parse_pages(slides, total)
    if overflow not in ("loose", "strict", "off"):
        fail(f"overflow 只能是 loose|strict|off，收到: {overflow}")
    ignore_filter = read_only_ignore(ignore)

    pdf, key, cache_hit = deck_pdf(path)
    doc = pymupdf.open(pdf)
    if doc.page_count < total:
        doc.close()
        fail(f"PDF 只有 {doc.page_count} 页，但 PPTX 有 {total} 页（转换异常）")

    outdir = Path(out).expanduser() if out else CACHE_ROOT / key / f"shots-{dpi}"
    outdir.mkdir(parents=True, exist_ok=True)

    slide_w_pt = prs.slide_width / EMU * PT_PER_IN
    slide_h_pt = prs.slide_height / EMU * PT_PER_IN

    shots = []
    files: list[Path] = []
    density_rows: list[dict] = []
    for page_no in pages:
        page = doc[page_no - 1]
        page_region = region_for_page(region, page_no)
        suffix = ""
        if page_region:
            tag = hashlib.sha1(",".join(f"{v:g}" for v in page_region).encode()).hexdigest()[:6]
            suffix = f"-region-{tag}"
        target = outdir / f"slide-{page_no}{suffix}.png"
        if not target.exists():
            if page_region:
                page.get_pixmap(dpi=dpi, clip=_region_rect(page, page_region, slide_w_pt, slide_h_pt)).save(str(target))
            else:
                page.get_pixmap(dpi=dpi).save(str(target))
        pix = pymupdf.Pixmap(str(target))
        entry: dict = {
            "slide": page_no,
            "file": str(target.resolve()),
            "px": [pix.width, pix.height],
        }
        entry.update(check_slide(prs.slides[page_no - 1], page, slide_w_pt, slide_h_pt, overflow, ignore_filter))
        shots.append(entry)
        files.append(target)
        if density:
            row = slide_density(prs.slides[page_no - 1], page, slide_w_pt, slide_h_pt)
            row["slide"] = page_no
            density_rows.append(row)
    doc.close()

    omitted = []
    if max_pages and len(shots) > max_pages:
        omitted = [s["slide"] for s in shots[max_pages:]]
        shots = shots[:max_pages]
        files = files[:max_pages]

    result: dict = {
        "pdf": str(pdf.resolve()),
        "dpi": dpi,
        "only_changed": bool(only_changed),
        "cache_hit": cache_hit,
        "overflow_mode": overflow,
        "shot_slides": [s["slide"] for s in shots],
        "shots": shots,
    }
    if region is not None:
        resolved = _region_spec(region)
        result["region_in"] = resolved if isinstance(resolved, dict) else _region_tuple(resolved)
    if density:
        result["density"] = density_rows
    if montage:
        cols = max(1, int(montage))
        montage_path = outdir / "montage.png"
        result["montage"] = make_montage(pymupdf, files, montage_path, cols=cols if cols > 1 else None)
    if omitted:
        result["omitted_slides"] = omitted
        result["hint"] = f"一次最多渲染 {max_pages} 页；其余页请再调用一次。"
    substitutions = _font_substitutions()
    if substitutions:
        result["font_substitutions"] = substitutions
        result["render_note"] = (
            "LibreOffice 渲染 ≠ PowerPoint：上面的字体被替换，字宽可能差约 5%，行宽请留 5-10% 余量。"
        )
    return result


def _font_substitutions() -> dict:
    """Only the requested faces that were actually substituted."""
    try:
        from .doctor import _font_substitutions as probe

        fonts = probe()
    except Exception:
        return {}
    return {name: matched.split(":")[0] for name, matched in fonts.items() if name not in matched}


def font_report(path: Path) -> dict:
    """Theme fonts + substitution chain, so renderer output can be judged."""
    try:
        from .doctor import doctor_payload

        info = doctor_payload(str(path))
    except Exception as exc:  # a font probe must never break a render
        return {"error": str(exc)}
    keep = ("fonts", "deck_theme_fonts")
    return {k: info[k] for k in keep if k in info}

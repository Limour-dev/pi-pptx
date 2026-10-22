"""Shared constants, errors, argument parsing and atomic (re)loading.

Everything here is dependency-free apart from python-pptx, so every other
module of the package can import it without cycles.
"""
from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import uuid
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN

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


def deck_report(prs, path, changed_slides=None) -> dict:
    """One-line deck summary for the ``@@DECK@@`` stdout protocol.

    A build script ends with ``deck.report()`` and ``ppt_run_deck`` learns the
    page/shape counts and which pages changed without a second inspect call.
    """
    shapes = tables = pictures = notes = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            shapes += 1
            if getattr(shape, "has_table", False) and shape.has_table:
                tables += 1
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pictures += 1
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            notes += 1
    payload = {
        "out": str(Path(path).resolve()) if path else None,
        "slides": len(prs.slides),
        "shapes": shapes,
        "tables": tables,
        "pictures": pictures,
        "notes": notes,
    }
    if changed_slides:
        payload["changed_slides"] = list(changed_slides)
    return payload


class ShapeFilter:
    """Match shapes by name, ``#id`` or ``/regex/`` for ``ignore=`` arguments.

    Shape ids drift when a deck is rebuilt and names repeat in templates, so
    "ignore this known-harmless warning" needs more than an exact name match.
    Entries are tried in this order: ``#id``, ``/regex/`` (matched against both
    name and path), then exact name.
    """

    def __init__(self, entries=None):
        if isinstance(entries, ShapeFilter):
            self.names = set(entries.names)
            self.ids = set(entries.ids)
            self.patterns = list(entries.patterns)
            self.raw = list(entries.raw)
            return
        if isinstance(entries, str):
            entries = [e.strip() for e in entries.split(",")]
        self.names: set[str] = set()
        self.ids: set[int] = set()
        self.patterns: list[re.Pattern] = []
        self.raw: list[str] = []
        for raw in entries or []:
            entry = str(raw).strip()
            if not entry:
                continue
            self.raw.append(entry)
            if entry.startswith("#") and entry[1:].isdigit():
                self.ids.add(int(entry[1:]))
            elif len(entry) > 2 and entry.startswith("/") and entry.endswith("/"):
                try:
                    self.patterns.append(re.compile(entry[1:-1]))
                except re.error as exc:
                    fail(f"ignore 里的正则 {entry} 无效: {exc}")
            else:
                self.names.add(entry)

    def __bool__(self) -> bool:
        return bool(self.raw)

    def matches(self, shape, path: str | None = None) -> bool:
        if not self.raw:
            return False
        if shape.shape_id in self.ids:
            return True
        if shape.name in self.names:
            return True
        if self.patterns:
            candidates = [shape.name]
            if path:
                candidates.append(path)
            for pattern in self.patterns:
                if any(pattern.search(value) for value in candidates):
                    return True
        return False


def read_only_ignore(spec) -> ShapeFilter:
    """Build a ShapeFilter from a CLI string / JSON list (``None`` = match none)."""
    return ShapeFilter(spec)

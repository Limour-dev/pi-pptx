"""Importable PPTX authoring library — the same primitives the MCP tools use.

A scripted deck's natural unit is one re-runnable Python file, not forty tool
calls, so the whole op layer is available without an agent in the loop:

    from pi_pptx import Deck

    deck = Deck.create("out.pptx", template="模板.pptx")
    deck.clone(source=3, count=2)
    deck.set_text(slide=4, shape="#12", text="标题")
    deck.add_shape(slide=4, kind="roundrect", x=0.9, y=1.6, w=3.6, h=1.5,
                   fill="#F4F7FB", line="#CAD9E8", radius=0.06, text="卡片")
    deck.shots(only_changed=True)      # renders only the pages the ops touched
    deck.save()

Every method mirrors an ``ops`` entry, so `build_deck.py` and `ppt_apply` share
one implementation instead of two half-copies.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from .doctor import doctor_payload
from .ops import OP_HANDLERS, apply_one
from .render import preview_overflow, render_overflow, shot_deck
from .shapes import inspect_deck, notes_check
from .utils import KitError, deck_meta, deck_report, emit, fail, load, save_and_reload

__all__ = [
    "Deck",
    "KitError",
    "OP_HANDLERS",
    "deck_meta",
    "deck_report",
    "doctor_payload",
    "emit",
    "inspect_deck",
    "load",
    "notes_check",
    "shot_deck",
]


@dataclass
class OpResult:
    """Result of one op: the raw payload plus the slides it touched."""

    data: dict
    slide: int | None = None

    def __getitem__(self, key):
        return self.data[key]

    def get(self, key, default=None):
        return self.data.get(key, default)


class Deck:
    """An in-memory deck with the CLI op layer attached.

    Mutations happen in memory; nothing touches the file until :meth:`save`.
    That makes "re-run the script from the template" the undo mechanism, and a
    failed op leaves the previous file intact.
    """

    def __init__(self, prs, path: Path | None = None, template: Path | None = None):
        self.prs = prs
        self.path = Path(path) if path else None
        self.template = Path(template) if template else None
        self.changed: set[int] = set()

    # ------------------------------------------------------------- lifecycle
    @classmethod
    def open(cls, path: str | Path) -> "Deck":
        prs, resolved = load(str(path))
        return cls(prs, resolved)

    @classmethod
    def create(
        cls,
        path: str | Path,
        *,
        template: str | Path | None = None,
        overwrite: bool = False,
        width: float | None = None,
        height: float | None = None,
        empty: bool = False,
    ) -> "Deck":
        """Byte-copy a template (the usual case) or start from a blank deck.

        Copying the file verbatim keeps every layout, theme font and sample page
        exactly as the designer left them.
        """
        import shutil
        import uuid

        target = Path(path).expanduser()
        if not target.is_absolute():
            target = (Path.cwd() / target).resolve()
        if target.exists() and not overwrite:
            fail(f"目标文件已存在: {target}（需要覆盖请传 overwrite=True）")
        target.parent.mkdir(parents=True, exist_ok=True)

        if template:
            src = Path(template).expanduser()
            if not src.exists():
                fail(f"模板不存在: {src}")
            if src.resolve() == target:
                fail("template 与目标文件相同：请指定新的输出路径，不要覆盖模板")
            if width is None and height is None:
                tmp = target.parent / f".{target.name}.{uuid.uuid4().hex}.tmp"
                try:
                    shutil.copyfile(src, tmp)
                    Presentation(str(tmp))
                    tmp.replace(target)
                finally:
                    if tmp.exists():
                        tmp.unlink()
            else:
                prs = Presentation(str(src))
                if width:
                    prs.slide_width = Inches(width)
                if height:
                    prs.slide_height = Inches(height)
                prs.save(str(target))
            return cls(Presentation(str(target)), target, src)

        prs = Presentation()
        prs.slide_width = Inches(width or 13.333)
        prs.slide_height = Inches(height or 7.5)
        if not empty:
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            placeholders = list(slide.placeholders)
            if placeholders and placeholders[0].has_text_frame:
                placeholders[0].text_frame.text = "标题"
            if len(placeholders) > 1 and placeholders[1].has_text_frame:
                placeholders[1].text_frame.text = "副标题"
        prs.save(str(target))
        return cls(Presentation(str(target)), target)

    @property
    def slide_count(self) -> int:
        return len(self.prs.slides)

    def meta(self) -> dict:
        return deck_meta(self.prs, self.path or Path("deck.pptx"))

    def report(self) -> dict:
        """Print the ``@@DECK@@`` line that ``ppt_run_deck`` parses.

        Put ``deck.report()`` right after ``save()``: the tool then knows the
        slide/shape counts and the pages this run touched, without a second
        ``ppt_inspect`` round trip.
        """
        import json

        payload = deck_report(self.prs, self.path, sorted(self.changed))
        print("@@DECK@@ " + json.dumps(payload, ensure_ascii=False))
        return payload

    def save(self, out: str | Path | None = None) -> Path:
        """Atomic save (temp file + integrity reopen + replace)."""
        if self.path is None:
            fail("这个 Deck 还没有目标路径，请用 Deck.open/create 或传 out=")
        prs, target = save_and_reload(self.prs, self.path, str(out) if out else None)
        self.prs = prs
        self.path = target
        self.changed.clear()
        return target

    # -------------------------------------------------------------- inspection
    def inspect(self, *, mode: str = "full", slides=None, depth: int = 99) -> dict:
        return inspect_deck(self.prs, self.path or Path("deck.pptx"), mode=mode, slides=slides, depth=depth)

    def notes_check(self, slides=None) -> dict:
        return notes_check(self.prs, slides)

    def shots(self, *, slides=None, dpi: int = 110, only_changed: bool = False, **kwargs) -> dict:
        """Render pages. ``only_changed=True`` uses the pages ops touched."""
        if self.path is None:
            fail("这个 Deck 还没有目标路径")
        return shot_deck(
            self.prs,
            self.path,
            slides=slides,
            dpi=dpi,
            only_changed=only_changed,
            changed_slides=sorted(self.changed),
            **kwargs,
        )

    def check(self, slides=None, mode: str = "loose", ignore=None) -> list[dict]:
        """Heuristics without writing PNGs (needs a saved file for the PDF)."""
        if self.path is None:
            fail("这个 Deck 还没有目标路径")
        return preview_overflow(self.prs, slides or list(range(1, self.slide_count + 1)), mode, ignore)

    def check_rendered(self, slides=None, mode: str = "loose", ignore=None) -> list[dict]:
        from .render import deck_pdf

        if self.path is None:
            fail("这个 Deck 还没有目标路径")
        pdf, _key, _cached = deck_pdf(self.path)
        pages = slides or list(range(1, self.slide_count + 1))
        return render_overflow(self.prs, pdf, pages, mode, ignore)

    # ---------------------------------------------------------------- op layer
    def op(self, operation: str, **spec) -> dict:
        """Run any op by name (see ``OP_HANDLERS``); ``None`` values are dropped."""
        clean = {k: v for k, v in spec.items() if v is not None}
        result = apply_one(self.prs, {"op": operation, **clean})
        slide = result.get("slide")
        if isinstance(slide, int):
            self.changed.add(slide)
        deleted = result.get("deleted")
        if isinstance(deleted, list):
            self.changed.update(int(page) for page in deleted)
        return result

    def apply(self, ops: list[dict], *, stop_on_error: bool = True) -> dict:
        """Batch: run every op in memory, commit nothing (call :meth:`save`)."""
        results, errors = [], []
        for index, spec in enumerate(ops):
            try:
                name = spec.get("op") or spec.get("type")
                fields = {k: v for k, v in spec.items() if k not in ("op", "type")}
                results.append({"index": index, **self.op(name, **fields)})
            except KitError as exc:
                errors.append({"index": index, "op": spec.get("op"), "error": str(exc)})
                if stop_on_error:
                    break
        return {
            "op_count": len(ops),
            "applied": len(results),
            "changed_slides": sorted(self.changed),
            "ops": results,
            "errors": errors,
            "ok": not errors,
        }

    # ------------------------------------------------------------ op wrappers
    def add_slide(self, layout, **kw) -> dict:
        return self.op("add_slide", layout=layout, **kw)

    def clone(self, source: int, at: int | None = None, count: int = 1) -> dict:
        return self.op("clone", source=source, at=at, count=count)

    def delete_slides(self, slides) -> dict:
        return self.op("delete_slides", slides=slides)

    def set_text(self, slide: int, shape, *, text=None, paras=None, **kw) -> dict:
        return self.op("set_text", slide=slide, shape=shape, text=text, paras=paras, **kw)

    def set_cell(self, slide: int, shape, *, row: int, col: int, text=None, cell_runs=None, **kw) -> dict:
        return self.op(
            "set_cell", slide=slide, shape=shape, row=row, col=col, text=text, cell_runs=cell_runs, **kw
        )

    def add_text(self, slide: int, **spec) -> dict:
        return self.op("add_text", slide=slide, **spec)

    def add_shape(self, slide: int, **spec) -> dict:
        return self.op("add_shape", slide=slide, **spec)

    def add_image(self, slide: int, **spec) -> dict:
        return self.op("add_image", slide=slide, **spec)

    def add_table(self, slide: int, **spec) -> dict:
        return self.op("add_table", slide=slide, **spec)

    def table_style(self, slide: int, shape, **spec) -> dict:
        return self.op("table_style", slide=slide, shape=shape, **spec)

    def delete_shape(self, slide: int, shape) -> dict:
        return self.op("delete_shape", slide=slide, shape=shape)

    def set_notes(self, slide: int, text: str) -> dict:
        return self.op("set_notes", slide=slide, text=text)

    # ----------------------------------------------------------------- drawing
    def card(self, slide: int, *, x, y, w, h, fill="#F4F7FB", line="#CAD9E8",
             radius=0.06, text=None, size=None, color=None, **spec) -> dict:
        """Rounded card with the deck-default look (no shadow, hairline border)."""
        return self.add_shape(
            slide=slide,
            kind="roundrect",
            x=x,
            y=y,
            w=w,
            h=h,
            fill=fill,
            line=line,
            radius=radius,
            text=text,
            size=size,
            color=color,
            **spec,
        )

    def bar(self, slide: int, *, x, y, w, h=0.055, fill="#1F4E79", **spec) -> dict:
        """Accent bar (a filled rectangle, no border)."""
        return self.add_shape(slide=slide, kind="bar", x=x, y=y, w=w, h=h, fill=fill, line="none", **spec)

    def rule(self, slide: int, *, x, y, w, color="#CAD9E8", width=1.0, **spec) -> dict:
        """Horizontal divider line."""
        return self.add_shape(
            slide=slide, kind="line", x=x, y=y, w=w, h=0, fill=color, line_width=width, **spec
        )

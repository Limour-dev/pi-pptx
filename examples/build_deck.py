#!/usr/bin/env python3
"""The recommended shape of a deck build script: re-runnable, idempotent, one page per block.

Run it through the `ppt_run_deck` tool (which injects PYTHONPATH, DECK_ROOT,
DECK_OUT, DECK_TEMPLATE) or by hand:

    DECK_ROOT=. DECK_TEMPLATE=模板.pptx DECK_OUT=/tmp/deck.pptx \
        PYTHONPATH=python python3 examples/build_deck.py

Re-running is the undo: `Deck.create(..., overwrite=True)` starts from the
template every time, and `Deck.save()` is atomic, so a failed run leaves the
previous deck intact.

What this file demonstrates (each item is one column of the layout):
  * design tokens and geometry constants at the top, used everywhere;
  * a handful of semantic wrappers (tx / card / bar / title / table) so the page
    code reads like the slide instead of like OOXML;
  * `table()` computes its own height and returns the bottom edge;
  * `if want(n):` around every page, so `--only 3,7` rebuilds two pages instead
    of the whole deck;
  * `@@PAGE@@` / `@@WARN@@` / `@@DECK@@` so `ppt_run_deck` can report progress,
    render only the flagged pages, and skip a second inspect call.
"""
from __future__ import annotations

import argparse
import os
from pathlib import Path

from pi_pptx import Deck  # PYTHONPATH is injected by ppt_run_deck

ROOT = Path(os.environ.get("DECK_ROOT") or Path.cwd())
OUT = Path(os.environ.get("DECK_OUT") or ROOT / "deck.pptx")
TPL = os.environ.get("DECK_TEMPLATE")

# ---------------------------------------------------------------- design tokens
INK, MUTED = "#1A1A1A", "#5B6B7C"
BLUE, ORANGE = "#1F4E79", "#ED7D31"
CARD_FILL, CARD_LINE = "#F4F7FB", "#DCE5EF"
CX, CW, BODY_Y = 0.68, 11.97, 1.84

PAGES = 4


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("template", nargs="?", default=TPL, help="template .pptx (or DECK_TEMPLATE)")
    parser.add_argument("--only", default="", help="rebuild only these pages, e.g. '3,7'")
    options = parser.parse_args()
    only = {int(item) for item in options.only.split(",") if item.strip().isdigit()}

    def want(page: int) -> bool:
        return not only or page in only

    if not options.template:
        raise SystemExit("需要一个模板：传参、DECK_TEMPLATE，或通过 ppt_run_deck(template=...)")
    template = Path(options.template).expanduser()
    if not template.exists():
        raise SystemExit(f"模板不存在: {template}")

    D = Deck.create(OUT, template=template, overwrite=True)
    # Normalise the canvas to PAGES: a template may carry one sample page or
    # twenty. Cloning the sample page copies its design (groups, pictures,
    # tables) and its rels; trimming removes the pages we will not use.
    while D.slide_count < PAGES:
        D.clone(source=1, at=D.slide_count + 1)
    if D.slide_count > PAGES:
        D.delete_slides(",".join(str(page) for page in range(PAGES + 1, D.slide_count + 1)))
    # The sample page's content is the design language, not the content: drop
    # those shapes and keep whatever the layout itself draws.
    sample_shapes = [
        shape["name"]
        for shape in D.inspect(slides="1", depth=0)["slides"][0]["shapes"]
        if shape["name"] in ("图片 4", "图片 5", "内容占位符 6")
    ]
    for page in range(1, D.slide_count + 1):
        for name in sample_shapes:
            D.op("delete_shape", slide=page, shape=name)

    # ------------------------------------------------------------------ helpers
    def tx(page, x, y, w, text=None, *, h=None, size=13, color=INK, paras=None, **kw):
        # "size x lines x 1.35 + a little" is the only height rule that holds:
        # a 40pt line does not fit a 0.4in box, and nothing auto-grows a box that
        # already has an explicit h.
        if h is None:
            lines = max(1, len(paras)) if paras else max(1, str(text or "").count("\n") + 1)
            h = round(0.1 + lines * size * 1.35 / 72.0, 2)
        spec = dict(slide=page, x=x, y=y, w=w, h=h, size=size, color=color, **kw)
        if paras is not None:
            spec["paras"] = paras
        else:
            spec["text"] = text
        return D.op("add_text", **{k: v for k, v in spec.items() if v is not None})

    def card(page, x, y, w, h, **kw):
        return D.op(
            "add_shape", slide=page, kind="roundrect", x=x, y=y, w=w, h=h,
            fill=CARD_FILL, line=CARD_LINE, line_width=0.75, radius=0.10, shadow=False, **kw,
        )

    def bar(page, x, y, w, h=0.055, fill=BLUE, **kw):
        return D.op("add_shape", slide=page, kind="bar", x=x, y=y, w=w, h=h, fill=fill, line="none", **kw)

    def title_shape():
        """The deck's own title placeholder name (templates differ: 标题 1 / Title 1)."""
        shapes = D.inspect(slides="1", depth=0)["slides"][0]["shapes"]
        for shape in shapes:
            if "TITLE" in (shape.get("placeholder_type") or ""):
                return shape["name"]
        for shape in shapes:
            name = shape["name"].lower()
            if "标题" in name or "title" in name:
                return shape["name"]
        return None

    TITLE = title_shape()

    def title(page, text, kicker=None):
        if TITLE:
            D.op("set_text", slide=page, shape=TITLE, text=text)
        else:
            tx(page, CX, 0.85, CW, text, size=40, bold=True)
        if kicker:
            tx(page, CX, 1.45, 9.0, kicker, size=11.5, color=BLUE, bold=True, letter_spacing=1.2)
        bar(page, CX, BODY_Y - 0.12, 0.9, h=0.05, fill=ORANGE)

    def table(page, x, y, w, data, widths, *, row_h=0.32, head_h=0.42, size=11):
        # Height is yours to compute: "auto fit" does not exist at this layer. A
        # wrong h is how the caption under a table ends up on top of it.
        h = round(head_h + row_h * (len(data) - 1), 2)
        D.op(
            "add_table", slide=page, rows=len(data), cols=len(data[0]), x=x, y=y, w=w, h=h,
            data=data, size=size, header_fill=BLUE, header_color="#FFFFFF",
            zebra="#EEF3F9", border="#C9D6E4", col_widths=widths,
            row_height=row_h, header_height=head_h, first_col_bold=True,
        )
        return y + h  # the next block starts below this line

    def notes(page, text):
        D.op("set_notes", slide=page, text=text)

    # -------------------------------------------------------------------- pages
    if want(1):
        print("@@PAGE@@ 1", flush=True)
        tx(1, CX, 2.4, 10.5, "从模板到成稿", size=40, bold=True)
        tx(1, CX, 3.3, 9.5, "一个能重跑的 build_deck.py，加上两个看图的工具", size=16, color=MUTED)
        bar(1, CX, 4.2, 1.6, h=0.06, fill=ORANGE)
        notes(1, "开场：这套流程解决的是'改脚本 → 跑 → 看'的闭环，不是逐形状编辑。")

    if want(2):
        print("@@PAGE@@ 2", flush=True)
        title(2, "三个数字", "WHY SCRIPTED DECKS")
        stats = [("536", "形状", "整本一次生成"), ("15", "轮返工", "全部是改脚本重跑"), ("0", "次单点编辑", "工具层不提供")]
        for index, (value, label, note) in enumerate(stats):
            x = CX + index * 4.0
            card(2, x, BODY_Y, 3.7, 2.1)
            tx(2, x + 0.3, BODY_Y + 0.35, 3.1, value, size=34, color=BLUE, bold=True)
            tx(2, x + 0.3, BODY_Y + 1.05, 3.1, label, size=14, bold=True)
            tx(2, x + 0.3, BODY_Y + 1.42, 3.1, note, size=11, color=MUTED)
        tx(2, CX, 4.6, CW, "一句话：编辑 API 留在 Python 包里，工具只负责跑和看。", size=15)
        notes(2, "把'工具'和'库'分开讲：脚本是事务，图是验收。")

    if want(3):
        print("@@PAGE@@ 3", flush=True)
        title(3, "表格高度自己算", "TABLE GEOMETRY")
        bottom = table(
            3, CX, BODY_Y, 7.2,
            [["阶段", "做法", "验收"],
             ["排版", "常量 + 语义包装", "text_fit 全绿"],
             ["运行", "ppt_run_deck", "0 告警"],
             ["复查", "shot flagged / compareTo", "只有预期页变化"]],
            [1.4, 3.4, 2.4],
        )
        tx(3, CX, bottom + 0.2, 7.2, "h = head_h + row_h × (rows-1)：返回底边，下一块从这里开始。", size=10.5, color=MUTED)
        card(3, 8.4, BODY_Y, 4.2, 2.4)
        tx(3, 8.7, BODY_Y + 0.25, 3.6, "别指望自动适应", size=14, bold=True, color=BLUE)
        tx(3, 8.7, BODY_Y + 0.7, 3.6, paras=[
            {"runs": [{"t": "表 4 行 + h=2.35", "bold": True}, {"t": " 是这次第一个 bug："}]},
            {"runs": [{"t": "说明文字压到了表格上。"}]},
        ], size=11.5)
        notes(3, "表格高度是它最容易出错的地方，因为工具不会帮你算。")

    if want(4):
        print("@@PAGE@@ 4", flush=True)
        title(4, "讨论", "DISCUSSION")
        for index, line in enumerate(["哪些页还需要 5–10% 行宽余量？", "下次要不要把文案常量单独抽成一个文件？"]):
            D.op("add_text", slide=4, x=CX, y=BODY_Y + index * 0.6, w=CW, h=0.4,
                 text=line, size=16, bullet="•", indent=0.32, hanging=0.32)
        notes(4, "留两个真问题给听众，不要用'谢谢'收尾。")

    D.save()
    D.report()  # prints @@DECK@@ {...}: counts + which pages this run touched


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Re-runnable card/chevron page built with the library only (no python-pptx, no OOXML).

This is the acceptance case for `ppt_add_shape`: five rounded cards, four
chevrons, one accent bar, a header rule and a footer label — every element is a
library call, and re-running the script rebuilds the deck from the template.

    python3 examples/cards_demo.py /home/limour/pi-test/PPT模板.pptx /tmp/cards.pptx
    python3 python/pptkit.py shot /tmp/cards.pptx --dpi 110
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "python"))

from pi_pptx import Deck, emit  # noqa: E402

TEMPLATE = "/home/limour/pi-test/PPT模板.pptx"
OUT = "/tmp/pi-pptx-cards.pptx"

CARD_FILL = "#F4F7FB"
CARD_LINE = "#CAD9E8"
ACCENT = "#1F4E79"
CHEVRON = "#2E75B6"
MUTED = "#5B6B7C"

STEPS = [
    ("选题", "先问清谁在什么场合看"),
    ("检索", "证据、指南、本中心数据"),
    ("复现", "脚本跑通再写进正文"),
    ("成稿", "图表先行，文字收敛"),
    ("复核", "渲染一遍，逐页看溢出"),
]


def build(template: str, out: str) -> Deck:
    deck = Deck.create(out, template=template, overwrite=True)
    layout = deck.prs.slides[0].slide_layout.name
    deck.add_slide(layout)
    slide = deck.slide_count

    # The reused layout places its title low and large; move it to the header
    # band and give it the deck's own header style (keeps the placeholder).
    deck.op("set_shape", slide=slide, shape="#2", x=0.75, y=0.62, w=11.8, h=0.8)
    deck.set_text(
        slide=slide,
        shape="#2",
        paras=[{"runs": [{"t": "从选题到成稿的五步闭环", "size": 28, "bold": True, "color": ACCENT}]}],
    )
    deck.delete_shape(slide=slide, shape="#3")  # unused subtitle placeholder

    # header rule + accent bar
    deck.rule(slide=slide, x=0.75, y=1.62, w=11.83, color=CARD_LINE, width=1.0)
    deck.bar(slide=slide, x=0.75, y=1.58, w=0.9, h=0.055, fill=ACCENT)

    # five cards + four chevrons
    card_w, card_h, gap = 2.06, 1.62, 0.46
    x, y = 0.75, 2.25
    for index, (title, body) in enumerate(STEPS, start=1):
        deck.card(
            slide=slide,
            x=x,
            y=y,
            w=card_w,
            h=card_h,
            fill=CARD_FILL,
            line=CARD_LINE,
            radius=0.06,
            name=f"Card {index}",
        )
        deck.add_text(
            slide=slide,
            x=x + 0.18,
            y=y + 0.2,
            w=card_w - 0.36,
            h=0.42,
            text=title,
            size=17,
            bold=True,
            color=ACCENT,
            letter_spacing=0.3,
            name=f"CardTitle {index}",
        )
        deck.add_text(
            slide=slide,
            x=x + 0.18,
            y=y + 0.68,
            w=card_w - 0.36,
            h=0.8,
            text=body,
            size=11.5,
            color=MUTED,
            line_spacing=1.15,
            name=f"CardBody {index}",
        )
        if index < len(STEPS):
            deck.add_shape(
                slide=slide,
                kind="chevron",
                x=x + card_w + 0.06,
                y=y + card_h / 2 - 0.19,
                w=0.34,
                h=0.38,
                fill=CHEVRON,
                line="none",
                name=f"Chevron {index}",
            )
        x += card_w + gap

    deck.add_text(
        slide=slide,
        x=0.75,
        y=6.55,
        w=11.83,
        h=0.4,
        text="图 1｜五步闭环：每一步都以「能不能复现」为交付标准。",
        size=12,
        color=MUTED,
        name="Footnote 1",
    )

    deck.set_notes(
        slide=slide,
        text="讲这页时先说结论：五步里只有第三步（复现）是可以被验证的，其余四步都是判断。",
    )
    deck.save()
    return deck


def main() -> None:
    template = sys.argv[1] if len(sys.argv) > 1 else TEMPLATE
    out = sys.argv[2] if len(sys.argv) > 2 else OUT
    if not Path(template).exists():
        print(f"模板不存在: {template}", file=sys.stderr)
        raise SystemExit(2)

    deck = build(template, out)
    report = deck.shots(only_changed=True, dpi=110, overflow="strict")
    flags = [
        (entry["slide"], entry.get("overflow"), entry.get("overlaps"))
        for entry in report["shots"]
        if entry.get("overflow") or entry.get("overlaps")
    ]
    emit(
        {
            "deck": str(deck.path),
            "slides": deck.slide_count,
            "rendered": report["shot_slides"],
            "images": [shot["file"] for shot in report["shots"]],
            "flags": flags,
            "suppressed": [entry.get("suppressed") for entry in report["shots"] if entry.get("suppressed")],
            "notes": deck.notes_check()["notes"],
        }
    )


if __name__ == "__main__":
    main()

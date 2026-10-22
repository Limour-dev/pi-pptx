#!/usr/bin/env python3
"""Backend smoke tests for python/pptkit.py.

Run: python3 tests/test_backend.py
Set PI_PPTX_TEST_TEMPLATE to also test against a real deck (optional).
"""
from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KIT = ROOT / "python" / "pptkit.py"

# 1x1 transparent PNG
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def kit(*args: str, expect_ok: bool = True) -> dict:
    result = subprocess.run(
        [sys.executable, str(KIT), *args],
        capture_output=True,
        text=True,
        timeout=300,
    )
    if expect_ok:
        if result.returncode != 0:
            raise AssertionError(f"pptkit {' '.join(args)} failed: {result.stderr}")
        return json.loads(result.stdout)
    if result.returncode == 0:
        raise AssertionError(f"pptkit {' '.join(args)} was expected to fail")
    return {"stderr": result.stderr}


class PptkitTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="pi-pptx-test-"))
        self.deck = self.tmp / "deck.pptx"
        self.png = self.tmp / "pixel.png"
        self.png.write_bytes(PNG_1PX)

    def create(self, name: str = "deck.pptx", **kw) -> Path:
        deck = self.tmp / name
        args = ["create", str(deck)]
        if kw.get("template"):
            args += ["--template", kw["template"]]
        kit(*args)
        return deck

    def inspect(self, deck: Path, slides: str | None = None) -> dict:
        args = ["inspect", str(deck)]
        if slides:
            args += ["--slides", slides]
        return kit(*args)

    def make_template(self) -> Path:
        from pptx import Presentation
        from pptx.util import Inches

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Sample title"
        slide.shapes.add_picture(str(self.png), Inches(1), Inches(1), Inches(2), Inches(2))
        table = slide.shapes.add_table(2, 2, Inches(1), Inches(3.5), Inches(4), Inches(1.2)).table
        table.cell(0, 0).text = "指标"
        table.cell(0, 1).text = "值"
        table.cell(1, 0).text = "EF"
        table.cell(1, 1).text = "55"
        template = self.tmp / "template.pptx"
        prs.save(template)
        return template

    # ------------------------------------------------------------------ tests
    def test_doctor(self):
        info = kit("doctor")
        self.assertIn("python_pptx", info)
        self.assertNotIn("ERROR", str(info.get("python_pptx")))
        self.assertIn("pymupdf", info)

    def test_create_blank_and_overwrite_guard(self):
        deck = self.create()
        info = self.inspect(deck)
        self.assertEqual(info["slide_count"], 1)
        self.assertEqual(info["slide_size_in"], [13.333, 7.5])
        failure = kit("create", str(deck), expect_ok=False)
        self.assertIn("已存在", failure["stderr"])
        kit("create", str(deck), "--overwrite")
        self.assertEqual(self.inspect(deck)["slide_count"], 1)

    def test_create_from_template_copies_slides(self):
        template = self.make_template()
        deck = self.create("copy.pptx", template=template)
        info = self.inspect(deck)
        self.assertEqual(info["slide_count"], 1)
        types = [s["type"] for s in info["slides"][0]["shapes"]]
        self.assertIn("PICTURE", types)
        self.assertIn("TABLE", types)

    def test_clone_preserves_pictures_table_and_notes(self):
        from pptx import Presentation

        template = self.make_template()
        deck = self.create("copy.pptx", template=template)
        # add notes to the source page
        prs = Presentation(str(deck))
        prs.slides[0].notes_slide.notes_text_frame.text = "speaker note"
        prs.save(str(deck))

        out = kit("clone", str(deck), "1", "--at", "1", "--count", "1")
        self.assertEqual(out["slide_count"], 2)
        info = self.inspect(deck)
        first = info["slides"][0]["shapes"]
        self.assertIn("PICTURE", [s["type"] for s in first])
        table = next(s for s in first if s.get("table"))
        self.assertEqual(table["table"][1][1], "55")
        self.assertEqual(info["slides"][0].get("notes"), "speaker note")

    def test_set_text_preserves_style_and_paragraphs(self):
        deck = self.create()
        kit(
            "add-text",
            str(deck),
            "--slide",
            "1",
            "--text",
            "原始标题",
            "--x",
            "1",
            "--y",
            "1",
            "--w",
            "8",
            "--h",
            "1",
            "--size",
            "28",
            "--font",
            "微软雅黑",
            "--color",
            "#1F4E79",
            "--bold",
        )
        info = self.inspect(deck, "1")
        shape = info["slides"][0]["shapes"][-1]
        name = shape["name"]
        self.assertEqual(shape["style"]["size_pt"], 28.0)
        self.assertEqual(shape["style"]["font"], "微软雅黑")

        kit("set-text", str(deck), "--slide", "1", "--shape", name, "--text", "新标题\n第二段")
        info = self.inspect(deck, "1")
        shape = info["slides"][0]["shapes"][-1]
        self.assertEqual(shape["text"], "新标题\n第二段")
        self.assertEqual(shape["style"]["size_pt"], 28.0)
        self.assertEqual(shape["style"]["font"], "微软雅黑")

    def test_table_add_and_set_cell(self):
        deck = self.create()
        added = kit(
            "add-table",
            str(deck),
            "--slide",
            "1",
            "--rows",
            "2",
            "--cols",
            "2",
            "--x",
            "1",
            "--y",
            "1",
            "--w",
            "4",
            "--h",
            "1.5",
            "--data",
            '[["指标","值"],["EF","52"]]',
        )
        name = added["name"]
        kit("set-cell", str(deck), "--slide", "1", "--shape", name, "--row", "1", "--col", "1", "--text", "58")
        info = self.inspect(deck, "1")
        table = next(s for s in info["slides"][0]["shapes"] if s.get("table"))
        self.assertEqual(table["table"], [["指标", "值"], ["EF", "58"]])
        kit("set-cell", str(deck), "--slide", "1", "--shape", name, "--row", "9", "--col", "0", "--text", "x", expect_ok=False)

    def test_add_image_and_group_path_error(self):
        deck = self.create()
        added = kit(
            "add-image",
            str(deck),
            "--slide",
            "1",
            "--path",
            str(self.png),
            "--x",
            "2",
            "--y",
            "2",
            "--w",
            "1",
        )
        self.assertEqual(added["added"], "picture")
        self.assertAlmostEqual(added["box_in"][2], 1.0, places=1)
        failure = kit(
            "set-text",
            str(deck),
            "--slide",
            "1",
            "--shape",
            "不存在的形状",
            "--text",
            "x",
            expect_ok=False,
        )
        self.assertIn("可用形状路径", failure["stderr"])

    def test_shot_render_overflow_and_cache(self):
        deck = self.create()
        kit("add-text", str(deck), "--slide", "1", "--text", "正常的一行", "--x", "1", "--y", "1", "--w", "4", "--h", "1", "--size", "20")
        bad = kit(
            "add-text",
            str(deck),
            "--slide",
            "1",
            "--text",
            "这是一段很长很长的文字，用来制造确定性的溢出，" * 3,
            "--x",
            "1",
            "--y",
            "2.5",
            "--w",
            "4",
            "--h",
            "0.3",
            "--size",
            "18",
        )
        first = kit(
            "shot",
            str(deck),
            "--slides",
            "1",
            "--dpi",
            "90",
        )
        self.assertEqual(len(first["shots"]), 1)
        shot = first["shots"][0]
        self.assertTrue(Path(shot["file"]).exists())
        self.assertGreater(shot["px"][0], 100)
        names = [item["shape"] for item in shot.get("overflow", [])]
        self.assertIn(bad["name"], names)
        texts = [item["text"] for item in shot.get("overflow", [])]
        self.assertFalse(any("正常的一行" in text for text in texts))

        # second shot hits the PDF cache but still works
        second = kit("shot", str(deck), "--slides", "1", "--dpi", "90")
        self.assertEqual(second["shots"][0]["file"], shot["file"])
        self.assertEqual(second["pdf"], first["pdf"])

    def test_shot_max_pages_and_delete(self):
        deck = self.create()
        for _ in range(3):
            kit("clone", str(deck), "1")
        result = kit("shot", str(deck), "--dpi", "80", "--max-pages", "2")
        self.assertEqual([s["slide"] for s in result["shots"]], [1, 2])
        self.assertEqual(result["omitted_slides"], [3, 4])
        deleted = kit("delete", str(deck), "--slides", "2-3")
        self.assertEqual(deleted["slide_count"], 2)
        kit("delete", str(deck), "--slides", "1-2", expect_ok=False)

    def test_add_slide_by_name(self):
        deck = self.create()
        info = self.inspect(deck)
        layout = info["layouts"][0]
        out = kit("add-slide", str(deck), "--layout", layout["name"])
        self.assertEqual(out["layout"], layout["name"])
        self.assertEqual(out["added_slide"], 2)

    def test_real_template_with_pictures(self):
        template = os.environ.get("PI_PPTX_TEST_TEMPLATE") or "/home/limour/pi-test/PPT模板.pptx"
        if not Path(template).exists():
            self.skipTest("real template not available")
        info = kit("inspect", template, "--slides", "1-3")
        self.assertEqual(len(info["slides"]), 3)
        deck = self.create("real.pptx", template=template)
        before = self.inspect(deck, "3")
        before_pics = [s for s in before["slides"][0]["shapes"] if s["type"] == "PICTURE"]
        kit("clone", str(deck), "3")
        after = self.inspect(deck, "4")
        after_pics = [s for s in after["slides"][0]["shapes"] if s["type"] == "PICTURE"]
        self.assertEqual(len(before_pics), len(after_pics))
        self.assertGreater(len(after_pics), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)

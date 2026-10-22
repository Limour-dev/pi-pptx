#!/usr/bin/env python3
"""Backend smoke tests for python/pptkit.py.

Run: python3 tests/test_backend.py
Set PI_PPTX_TEST_TEMPLATE to also test against a real deck (optional).
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
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

    def inspect(self, deck: Path, slides: str | None = None, **kw) -> dict:
        args = ["inspect", str(deck)]
        if slides:
            args += ["--slides", slides]
        for key, value in kw.items():
            args += ["--" + key.replace("_", "-"), str(value)]
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

    def slide_xml(self, deck: Path, slide: int = 1) -> str:
        with zipfile.ZipFile(deck) as archive:
            return archive.read(f"ppt/slides/slide{slide}.xml").decode("utf-8")

    def test_add_text_cjk_font_layout_and_autofit(self):
        deck = self.create()
        added = kit(
            "add-text",
            str(deck),
            "--slide",
            "1",
            "--text",
            "中文探测行\n第二段",
            "--x",
            "1",
            "--y",
            "1",
            "--w",
            "6",
            "--size",
            "20",
            "--font",
            "微软雅黑",
            "--space-after",
            "6",
            "--indent",
            "0.32",
            "--bullet",
            "•",
            "--bullet-color",
            "#2E75B6",
        )
        # h omitted -> grow + estimated height
        self.assertEqual(added["fit"], "grow")
        self.assertGreater(added["box_in"][3], 0.8)
        xml = self.slide_xml(deck)
        self.assertIn('<a:ea typeface="微软雅黑"/>', xml)
        self.assertIn('<a:cs typeface="微软雅黑"/>', xml)
        self.assertIn("<a:spAutoFit/>", xml)
        self.assertIn("<a:buChar", xml)
        self.assertIn("<a:buClr>", xml)
        self.assertIn('marL="292608"', xml)
        self.assertIn('indent="-292608"', xml)
        self.assertIn('<a:spcAft>', xml)

        # explicit h + fit=shrink writes normAutofit, no grow
        shrunk = kit(
            "add-text",
            str(deck),
            "--slide",
            "1",
            "--text",
            "固定框",
            "--x",
            "1",
            "--y",
            "3",
            "--w",
            "3",
            "--h",
            "1",
            "--fit",
            "shrink",
        )
        self.assertEqual(shrunk["fit"], "shrink")
        self.assertIn("<a:normAutofit", self.slide_xml(deck))

    def test_set_text_rich_runs_keep_base_style(self):
        deck = self.create()
        added = kit(
            "add-text",
            str(deck),
            "--slide",
            "1",
            "--text",
            "原始",
            "--x",
            "1",
            "--y",
            "1",
            "--w",
            "8",
            "--h",
            "1.5",
            "--size",
            "20",
            "--font",
            "微软雅黑",
            "--color",
            "#1F1F1F",
        )
        paras = json.dumps(
            [
                {
                    "runs": [
                        {"t": "一句实话：", "bold": True, "color": "C00000", "size": 22},
                        {"t": "没有事前效能计算。"},
                    ],
                    "space_after": 6,
                },
                {"runs": [{"t": "第二段"}], "align": "center"},
            ],
            ensure_ascii=False,
        )
        out = kit("set-text", str(deck), "--slide", "1", "--shape", "#" + str(added["id"]), "--paras", paras)
        self.assertEqual(out["paragraphs"], 2)

        from pptx import Presentation

        prs = Presentation(str(deck))
        shape = prs.slides[0].shapes[-1]
        first, second = shape.text_frame.paragraphs
        bold, normal = first.runs
        self.assertEqual(bold.text, "一句实话：")
        self.assertTrue(bold.font.bold)
        self.assertEqual(bold.font.size.pt, 22.0)
        self.assertEqual(str(bold.font.color.rgb), "C00000")
        # base run style inherited where not overridden
        self.assertEqual(normal.text, "没有事前效能计算。")
        self.assertEqual(normal.font.size.pt, 20.0)
        self.assertEqual(normal.font.name, "微软雅黑")
        self.assertAlmostEqual(first.space_after.pt, 6.0)
        self.assertEqual(str(second.alignment), "CENTER (2)")

    def test_shape_id_addressing_duplicate_names_and_delete(self):
        deck = self.create()
        first = kit("add-text", str(deck), "--slide", "1", "--text", "A", "--x", "1", "--y", "1", "--w", "3", "--h", "1")
        second = kit("add-text", str(deck), "--slide", "1", "--text", "B", "--x", "5", "--y", "1", "--w", "3", "--h", "1")

        from pptx import Presentation

        prs = Presentation(str(deck))
        for shape in list(prs.slides[0].shapes)[-2:]:
            shape.name = "标题 3"
        prs.save(str(deck))

        failure = kit(
            "set-text", str(deck), "--slide", "1", "--shape", "标题 3", "--text", "x", expect_ok=False
        )
        self.assertIn("不唯一", failure["stderr"])
        self.assertIn(f"id=#{first['id']}", failure["stderr"])
        self.assertIn(f"id=#{second['id']}", failure["stderr"])
        self.assertIn("index=", failure["stderr"])

        out = kit("set-text", str(deck), "--slide", "1", "--shape", f"#{second['id']}", "--text", "B2")
        self.assertEqual(out["text"], "B2")
        info = self.inspect(deck, "1")
        texts = {shape["id"]: shape.get("text") for shape in info["slides"][0]["shapes"]}
        self.assertEqual(texts[first["id"]], "A")
        self.assertEqual(texts[second["id"]], "B2")

        # out-of-range numeric index explains the 0-based convention
        failure = kit(
            "set-text", str(deck), "--slide", "1", "--shape", "99", "--text", "x", expect_ok=False
        )
        self.assertIn("index 从 0 开始", failure["stderr"])

        deleted = kit("delete-shape", str(deck), "--slide", "1", "--shape", f"#{first['id']}")
        self.assertTrue(deleted["deleted"])
        ids = [shape["id"] for shape in self.inspect(deck, "1")["slides"][0]["shapes"]]
        self.assertNotIn(first["id"], ids)

    def test_notes_summary_inspect_and_doctor_theme(self):
        deck = self.create()
        out = kit("set-notes", str(deck), "--slide", "1", "--text", "开场白\n第二段")
        self.assertEqual(out["paragraphs"], 2)
        summary = self.inspect(deck, "1", mode="summary")
        self.assertTrue(summary["slides"][0]["title"])
        self.assertIn("开场白", summary["slides"][0]["notes"])
        self.assertEqual(summary["slides"][0]["shape_count"], 2)

        full = kit("inspect", str(deck), "--slides", "1", "--depth", "0")
        self.assertNotIn("children", json.dumps(full["slides"], ensure_ascii=False))
        doctor = kit("doctor", str(deck))
        self.assertIn("deck_theme_fonts", doctor)
        self.assertIn("major", doctor["deck_theme_fonts"])

    def test_table_theme_off_style_and_borders(self):
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
            "3",
            "--w",
            "5",
            "--h",
            "1.5",
            "--data",
            '[["指标","值"],["EF","52"]]',
            "--size",
            "13",
            "--header-fill",
            "#1F4E79",
            "--header-color",
            "#FFFFFF",
            "--zebra",
            "#EAF1F8",
        )
        self.assertIn("模板里没有任何表格样例", added["hint"])

        from pptx import Presentation
        from pptx.oxml.ns import qn

        prs = Presentation(str(deck))
        shape = next(s for s in prs.slides[0].shapes if s.has_table)
        tbl = shape._element.graphic.graphicData.tbl
        tblPr = tbl.find(qn("a:tblPr"))
        self.assertEqual(tblPr.get("firstRow"), "0")
        self.assertEqual(tblPr.get("bandRow"), "0")
        self.assertIsNone(tblPr.find(qn("a:tableStyleId")))
        header = shape.table.cell(0, 0)
        self.assertEqual(str(header.fill.fore_color.rgb), "1F4E79")
        body = shape.table.cell(1, 0)
        self.assertEqual(str(body.fill.fore_color.rgb), "EAF1F8")
        self.assertEqual([el.tag.split("}")[1] for el in list(body._tc.tcPr)[:4]], ["lnL", "lnR", "lnT", "lnB"])

        kit(
            "table-style",
            str(deck),
            "--slide",
            "1",
            "--shape",
            f"#{added['id']}",
            "--header-fill",
            "#C00000",
            "--border",
            "none",
            "--col-widths",
            "2,3",
        )
        prs = Presentation(str(deck))
        shape = next(s for s in prs.slides[0].shapes if s.has_table)
        self.assertEqual(str(shape.table.cell(0, 1).fill.fore_color.rgb), "C00000")
        self.assertAlmostEqual(shape.table.columns[0].width.inches, 2.0, places=2)
        lnL = shape.table.cell(1, 1)._tc.tcPr.find(qn("a:lnL"))
        self.assertIsNotNone(lnL.find(qn("a:noFill")))

    def test_apply_dry_run_atomic_failure_and_commit(self):
        deck = self.create()
        added = kit("add-text", str(deck), "--slide", "1", "--text", "初始", "--x", "1", "--y", "1", "--w", "5", "--h", "1")
        before = hashlib.sha256(deck.read_bytes()).hexdigest()

        good_ops = [
            {"op": "set_text", "slide": 1, "shape": f"#{added['id']}", "text": "批量改完"},
            {"op": "add_text", "slide": 1, "x": 1, "y": 3, "w": 5, "text": "新增", "size": 16},
            {"op": "set_notes", "slide": 1, "text": "备注"},
        ]
        dry = kit("apply", str(deck), "--ops", json.dumps(good_ops, ensure_ascii=False), "--dry-run")
        self.assertTrue(dry["dry_run"])
        self.assertFalse(dry["saved"])
        self.assertTrue(dry["ok"])
        self.assertEqual(dry["applied"], 3)
        self.assertEqual(hashlib.sha256(deck.read_bytes()).hexdigest(), before)
        self.assertEqual([s.get("text") for s in self.inspect(deck, "1")["slides"][0]["shapes"]], ["标题", "副标题", "初始"])

        bad_ops = [
            {"op": "set_text", "slide": 1, "shape": f"#{added['id']}", "text": "不应写入"},
            {"op": "set_text", "slide": 1, "shape": "#999", "text": "x"},
        ]
        bad = kit("apply", str(deck), "--ops", json.dumps(bad_ops), "--no-shot")
        self.assertFalse(bad["ok"])
        self.assertFalse(bad["saved"])
        self.assertEqual(bad["errors"][0]["index"], 1)
        self.assertEqual(hashlib.sha256(deck.read_bytes()).hexdigest(), before)

        applied = kit("apply", str(deck), "--ops", json.dumps(good_ops, ensure_ascii=False), "--no-shot")
        self.assertTrue(applied["ok"])
        self.assertTrue(applied["saved"])
        self.assertEqual(applied["changed_slides"], [1])
        info = self.inspect(deck, "1")
        texts = [s.get("text") for s in info["slides"][0]["shapes"]]
        self.assertIn("批量改完", texts)
        self.assertIn("新增", texts)
        self.assertEqual(info["slides"][0]["notes"], "备注")

    def test_apply_auto_shot_reports_overflow(self):
        deck = self.create()
        added = kit("add-text", str(deck), "--slide", "1", "--text", "短", "--x", "1", "--y", "1", "--w", "2", "--h", "0.3", "--size", "18")
        ops = [
            {
                "op": "set_text",
                "slide": 1,
                "shape": f"#{added['id']}",
                "text": "这是一段很长很长的文字，" * 4,
            }
        ]
        out = kit("apply", str(deck), "--ops", json.dumps(ops, ensure_ascii=False))
        self.assertTrue(out["saved"])
        self.assertIn("overflow", out)
        names = [item["shape"] for entry in out["overflow"] for item in entry.get("overflow", [])]
        self.assertIn(added["name"], names)

    def test_shot_montage_region_and_overlap(self):
        deck = self.create()
        kit("add-text", str(deck), "--slide", "1", "--text", "重叠一", "--x", "1", "--y", "2", "--w", "4", "--h", "1", "--size", "20")
        kit("add-text", str(deck), "--slide", "1", "--text", "重叠二", "--x", "1.2", "--y", "2.1", "--w", "4", "--h", "1", "--size", "20")
        kit("clone", str(deck), "1", "--count", "2")
        shot = kit(
            "shot",
            str(deck),
            "--slides",
            "1-3",
            "--dpi",
            "70",
            "--montage",
            "2",
            "--overflow",
            "off",
        )
        self.assertEqual(shot["overflow_mode"], "off")
        self.assertEqual(len(shot["shots"]), 3)
        montage = shot["montage"]
        self.assertEqual(montage["cols"], 2)
        self.assertEqual(montage["rows"], 2)
        self.assertTrue(Path(montage["file"]).exists())
        self.assertGreater(montage["px"][0], 600)
        self.assertTrue(shot["shots"][0]["overlaps"])

        region = kit(
            "shot",
            str(deck),
            "--slides",
            "1",
            "--dpi",
            "90",
            "--region",
            "0.5,0.5,3,2",
            "--overflow",
            "off",
        )
        self.assertEqual(region["region_in"], [0.5, 0.5, 3.0, 2.0])
        self.assertLess(region["shots"][0]["px"][0], 400)

    def test_loose_overflow_skips_centered_labels(self):
        deck = self.create()
        added = kit(
            "add-text",
            str(deck),
            "--slide",
            "1",
            "--text",
            "LVEF，%",
            "--x",
            "2",
            "--y",
            "2",
            "--w",
            "2",
            "--h",
            "0.6",
            "--size",
            "18",
            "--align",
            "center",
        )
        shot = kit("shot", str(deck), "--slides", "1", "--dpi", "90", "--overflow", "loose")
        flagged = [item["shape"] for item in shot["shots"][0].get("overflow", [])]
        self.assertNotIn(added["name"], flagged)

    # ---------------------------------------------------------- P0-1 anchors
    def test_template_page_is_clean_under_strict(self):
        """A never-edited template must report zero flags (acceptance for P0-1)."""
        template = os.environ.get("PI_PPTX_TEST_TEMPLATE", "")
        if not template or not Path(template).exists():
            self.skipTest("set PI_PPTX_TEST_TEMPLATE to a real deck to run this")
        shot = kit("shot", template, "--slides", "1-3", "--dpi", "70", "--overflow", "strict")
        for entry in shot["shots"]:
            self.assertFalse(entry.get("overflow"), f"slide {entry['slide']} overflow: {entry.get('overflow')}")
            self.assertFalse(entry.get("overlaps"), f"slide {entry['slide']} overlaps: {entry.get('overlaps')}")

    def test_template_title_leading_is_not_reported(self):
        """A top-anchored spill under 5pt is template leading, on every page."""
        deck = self.create()
        added = kit(
            "add-text", str(deck), "--slide", "1", "--text", "大字号标题", "--x", "1", "--y", "1",
            "--w", "6", "--h", "1.2", "--size", "44",
        )
        shot = kit("shot", str(deck), "--slides", "1", "--dpi", "90", "--overflow", "strict")
        entry = shot["shots"][0]
        flagged = [item["shape"] for item in entry.get("overflow", [])]
        self.assertNotIn(added["name"], flagged)
        # Not merely "suppressed": the noise is dropped, not moved to another list.
        noisy = [
            item for item in entry.get("suppressed", [])
            if item["shape"] == added["name"]
        ]
        self.assertFalse(noisy)

    def test_ignores_accept_id_and_regex(self):
        deck = self.create()
        first = kit("add-text", str(deck), "--slide", "1", "--text", "很长的文字" * 8,
                    "--x", "1", "--y", "1", "--w", "3", "--h", "0.4", "--size", "20")
        second = kit("add-text", str(deck), "--slide", "1", "--text", "另一段很长的文字" * 8,
                     "--x", "1", "--y", "4", "--w", "3", "--h", "0.4", "--size", "20")
        by_id = kit("shot", str(deck), "--slides", "1", "--dpi", "80", "--ignore", f"#{first['id']}")
        names = [item["shape"] for entry in by_id["shots"] for item in entry.get("overflow", [])]
        self.assertNotIn(first["name"], names)
        self.assertIn(second["name"], names)

        by_regex = kit("shot", str(deck), "--slides", "1", "--dpi", "80", "--ignore", "/^TextBox/")
        self.assertFalse([e for e in by_regex["shots"] if e.get("overflow")])

    # ---------------------------------------------------------- P0-3 shapes
    def test_add_shape_draws_cards_bars_and_chevrons(self):
        deck = self.create()
        card = kit(
            "add-shape", str(deck), "--slide", "1", "--kind", "roundrect",
            "--x", "1", "--y", "1.5", "--w", "4", "--h", "1.5",
            "--fill", "#F4F7FB", "--line", "#CAD9E8", "--radius", "0.06",
            "--text", "卡片", "--size", "16", "--color", "#1F4E79", "--name", "Card 1",
        )
        self.assertEqual(card["name"], "Card 1")
        self.assertEqual(card["paragraphs"], 1)
        self.assertAlmostEqual(card["box_in"][3], 1.5, places=2)

        bar = kit("add-shape", str(deck), "--slide", "1", "--kind", "bar",
                  "--x", "1", "--y", "3.4", "--w", "4", "--fill", "#1F4E79", "--line", "none")
        self.assertAlmostEqual(bar["box_in"][3], 0.06, places=2)

        chevron = kit("add-shape", str(deck), "--slide", "1", "--kind", "chevron",
                      "--x", "5.5", "--y", "1.5", "--w", "1.6", "--h", "0.9", "--fill", "#2E75B6")
        self.assertEqual(chevron["kind"], "chevron")

        rule = kit("add-shape", str(deck), "--slide", "1", "--kind", "line",
                   "--x", "1", "--y", "5", "--w", "5", "--fill", "#CAD9E8")
        self.assertEqual(rule["kind"], "line")

        info = self.inspect(deck, "1")
        shapes = {s["name"]: s for s in info["slides"][0]["shapes"]}
        self.assertIn("Card 1", shapes)
        # A drawn card has no inherited theme shadow.
        from pptx import Presentation

        prs = Presentation(str(deck))
        card_shape = next(s for s in prs.slides[0].shapes if s.name == "Card 1")
        self.assertFalse(card_shape.shadow.inherit)
        self.assertEqual(card_shape.adjustments[0], round(0.06 / 1.5, 3))

    def test_apply_supports_add_shape(self):
        deck = self.create()
        ops = [
            {"op": "add_shape", "slide": 1, "kind": "roundrect", "x": 1, "y": 1, "w": 3, "h": 1,
             "fill": "#F4F7FB", "name": "Card A", "text": "A"},
            {"op": "add_shape", "slide": 1, "kind": "chevron", "x": 4.2, "y": 1, "w": 1, "h": 1},
        ]
        out = kit("apply", str(deck), "--ops", json.dumps(ops), "--no-shot")
        self.assertTrue(out["ok"])
        self.assertEqual([op["kind"] for op in out["ops"]], ["roundrect", "chevron"])
        self.assertIn("Card A", [s["name"] for s in self.inspect(deck, "1")["slides"][0]["shapes"]])

    def test_set_shape_moves_restyles_and_reorders(self):
        deck = self.create()
        kit("add-text", str(deck), "--slide", "1", "--text", "叠在上层", "--x", "1", "--y", "1", "--w", "3", "--h", "1")
        card = kit("add-shape", str(deck), "--slide", "1", "--kind", "rect",
                   "--x", "1", "--y", "1", "--w", "4", "--h", "2", "--fill", "#F4F7FB")
        out = kit(
            "set-shape", str(deck), "--slide", "1", "--shape", f"#{card['id']}",
            "--x", "2", "--y", "2.5", "--w", "5", "--h", "1.5",
            "--fill", "#EAF1F8", "--line", "#CAD9E8", "--z", "front", "--name", "Panel 1",
        )
        self.assertEqual(out["name"], "Panel 1")
        self.assertEqual(out["box_before"], [1.0, 1.0, 4.0, 2.0])
        self.assertEqual(out["box_in"], [2.0, 2.5, 5.0, 1.5])

        from pptx import Presentation
        from pptx.oxml.ns import qn

        prs = Presentation(str(deck))
        shapes = list(prs.slides[0].shapes)
        self.assertEqual(shapes[-1].name, "Panel 1")  # z=front
        panel = shapes[-1]
        self.assertIsNotNone(panel._element.find(qn("p:spPr")).find(qn("a:solidFill")))

        back = kit("set-shape", str(deck), "--slide", "1", "--shape", "Panel 1", "--z", "back")
        self.assertEqual(back["z"], "back")
        prs = Presentation(str(deck))
        drawing = [s.name for s in prs.slides[0].shapes]
        self.assertEqual(drawing[0], "Panel 1")

    # ---------------------------------------------------------- P1-5 tables
    def test_table_header_height_first_column_and_rich_cells(self):
        deck = self.create()
        kit(
            "add-table", str(deck), "--slide", "1", "--rows", "3", "--cols", "2",
            "--x", "1", "--y", "1", "--w", "8", "--h", "2.5",
            "--data", json.dumps([["指标", "数值"], ["LVEF", "58"], ["BNP", "120"]], ensure_ascii=False),
            "--header-fill", "#1F4E79", "--header-color", "FFFFFF",
            "--row-height", "0.7", "--header-height", "0.45",
            "--first-col-bold", "--first-col-color", "#1F4E79",
            "--border", "none",
        )
        from pptx import Presentation

        prs = Presentation(str(deck))
        table = next(s for s in prs.slides[0].shapes if s.has_table).table
        self.assertAlmostEqual(table.rows[0].height.inches, 0.45, places=2)
        self.assertAlmostEqual(table.rows[1].height.inches, 0.7, places=2)
        self.assertTrue(table.cell(1, 0).text_frame.paragraphs[0].runs[0].font.bold)
        self.assertFalse(table.cell(1, 1).text_frame.paragraphs[0].runs[0].font.bold)

        rich = [[[{"t": "指标"}], [{"t": "数值"}]],
                [[{"t": "LVEF", "bold": True}, {"t": "（%）", "color": "888888"}], [{"t": "58"}]],
                [[{"t": "BNP"}], [{"t": "120"}]]]
        styled = kit("table-style", str(deck), "--slide", "1", "--shape",
                     f"#{next(s.shape_id for s in prs.slides[0].shapes if s.has_table)}",
                     "--cell-runs", json.dumps(rich, ensure_ascii=False), "--header-height", "0.5")
        self.assertTrue(styled["rich_cells"])
        prs = Presentation(str(deck))
        table = next(s for s in prs.slides[0].shapes if s.has_table).table
        runs = table.cell(1, 0).text_frame.paragraphs[0].runs
        self.assertEqual([run.text for run in runs], ["LVEF", "（%）"])
        self.assertTrue(runs[0].font.bold)
        self.assertAlmostEqual(table.rows[0].height.inches, 0.5, places=2)

    # ---------------------------------------------------------- P1-6 layouts
    def test_inspect_lists_layout_owned_shapes(self):
        template = os.environ.get("PI_PPTX_TEST_TEMPLATE", "")
        if not template or not Path(template).exists():
            self.skipTest("set PI_PPTX_TEST_TEMPLATE to a real deck to run this")
        info = kit("inspect", template, "--mode", "summary")
        with_shapes = [layout for layout in info["layouts"] if layout.get("shapes")]
        self.assertTrue(with_shapes, "at least one layout should expose inherited shapes")
        first = with_shapes[0]
        self.assertTrue(first.get("shapes_hint"))
        for shape in first["shapes"]:
            self.assertTrue(shape["inherited"])
            self.assertIn("box_in", shape)

    # ---------------------------------------------------------- P2-8 extras
    def test_letter_spacing_and_names(self):
        deck = self.create()
        added = kit(
            "add-text", str(deck), "--slide", "1", "--text", "字距测试", "--x", "1", "--y", "1",
            "--size", "20", "--letter-spacing", "0.75", "--name", "Label 1",
        )
        self.assertEqual(added["name"], "Label 1")
        from pptx import Presentation
        from pptx.oxml.ns import qn

        prs = Presentation(str(deck))
        shape = next(s for s in prs.slides[0].shapes if s.name == "Label 1")
        rPr = shape.text_frame.paragraphs[0].runs[0]._r.find(qn("a:rPr"))
        self.assertEqual(rPr.get("spc"), "75")

    def test_notes_check_and_delete_reports_note_shift(self):
        deck = self.create()
        kit("add-text", str(deck), "--slide", "1", "--text", "第一页", "--x", "1", "--y", "1")
        kit("clone", str(deck), "1", "--count", "1")
        kit("set-notes", str(deck), "--slide", "1", "--text", "第一页的备注：讲开场")
        kit("set-notes", str(deck), "--slide", "2", "--text", "第二页的备注：讲方法")
        first = kit("notes-check", str(deck))
        self.assertEqual([row["notes_len"] > 0 for row in first["notes"]], [True, True])
        self.assertEqual(first["notes"][0]["first_20"], "第一页的备注：讲开场")

        out = kit("delete", str(deck), "--slides", "1")
        self.assertEqual(out["slide_count"], 1)
        self.assertEqual(len(out["notes_after"]), 1)
        self.assertEqual(out["notes_after"][0]["first_20"], "第二页的备注：讲方法")
        self.assertIn("notes_hint", out)

    def test_only_changed_renders_only_touched_pages(self):
        deck = self.create()
        kit("clone", str(deck), "1", "--count", "2")
        # First call on an unknown deck renders everything it can.
        first = kit("shot", str(deck), "--only-changed", "--dpi", "60")
        self.assertEqual(first["shot_slides"], [1, 2, 3])
        second = kit("shot", str(deck), "--only-changed", "--dpi", "60")
        self.assertTrue(second["no_changes"])
        self.assertEqual(second["shots"], [])
        kit("add-text", str(deck), "--slide", "2", "--text", "改动", "--x", "1", "--y", "1")
        third = kit("shot", str(deck), "--only-changed", "--dpi", "60")
        self.assertEqual(third["shot_slides"], [2])

    def test_shot_density_table(self):
        deck = self.create()
        added = kit("add-text", str(deck), "--slide", "1", "--text", "第一行\n第二行", "--x", "1", "--y", "1",
                    "--w", "5", "--h", "1.2", "--size", "18")
        shot = kit("shot", str(deck), "--slides", "1", "--dpi", "70", "--density")
        row = shot["density"][0]
        self.assertEqual(row["slide"], 1)
        self.assertIn("first_line", row)
        self.assertIn("lines", row)
        # Density is per shape: text that does not fit is one box that is too
        # short, so the row names it instead of averaging the whole page.
        self.assertTrue(row["shapes"])
        entry = next(item for item in row["shapes"] if item["shape"] == added["name"])
        for key in ("shape", "rendered_lines", "needed_in", "box_h_in", "slack_in"):
            self.assertIn(key, entry)
        self.assertEqual(entry["rendered_lines"], 2)

    def test_overflow_reports_measured_lines_and_height(self):
        deck = self.create()
        bad = kit(
            "add-text", str(deck), "--slide", "1", "--text", "溢出文字" * 30,
            "--x", "1", "--y", "1", "--w", "3", "--h", "0.4", "--size", "18",
        )
        shot = kit("shot", str(deck), "--slides", "1", "--dpi", "80", "--overflow", "strict")
        entry = next(item for item in shot["shots"][0]["overflow"] if item["shape"] == bad["name"])
        self.assertGreater(entry["rendered_lines"], 1)
        self.assertGreater(entry["needed_in"], 0.4)
        self.assertLess(entry["slack_in"], 0)

    def test_overlap_area_is_square_inches(self):
        """area_in2 used to mix points with inches (1e11 on a 100 in² canvas)."""
        deck = self.create()
        kit("add-text", str(deck), "--slide", "1", "--text", "重叠一", "--x", "1", "--y", "2", "--w", "4", "--h", "1", "--size", "20")
        kit("add-text", str(deck), "--slide", "1", "--text", "重叠二", "--x", "1.2", "--y", "2.1", "--w", "4", "--h", "1", "--size", "20")
        shot = kit("shot", str(deck), "--slides", "1", "--dpi", "70", "--overflow", "off")
        overlaps = shot["shots"][0]["overlaps"]
        self.assertTrue(overlaps)
        for item in overlaps:
            self.assertLessEqual(item["area_in2"], 4 * 1.0 + 0.01)
            self.assertLessEqual(item["ratio"], 1.0)
            self.assertGreater(item["ratio"], 0.2)

    def test_inspect_text_fit_and_text_mode(self):
        deck = self.create()
        tight = kit(
            "add-text", str(deck), "--slide", "1", "--text", "两行文字需要的空间\n比这个盒子高" * 2,
            "--x", "1", "--y", "1", "--w", "2", "--h", "0.3", "--size", "18",
        )
        info = kit("inspect", str(deck), "--slides", "1")
        shapes = {shape["name"]: shape for shape in info["slides"][0]["shapes"]}
        fit = shapes[tight["name"]]["text_fit"]
        self.assertTrue(fit["estimated"])
        self.assertFalse(fit["fits"])
        self.assertGreater(fit["est_lines"], fit["capacity_lines"])
        self.assertIn("font_substitutions", info)

        text_mode = kit("inspect", str(deck), "--slides", "1", "--mode", "text")
        self.assertEqual(text_mode["mode"], "text")
        texts = [row["text"] for row in text_mode["slides"][0]["shapes"]]
        self.assertTrue(any("两行文字" in text for text in texts))
        self.assertNotIn("layouts", text_mode)

    def test_shot_expect_clean_and_flagged_slides(self):
        deck = self.create()
        kit("add-text", str(deck), "--slide", "1", "--text", "正常", "--x", "1", "--y", "1", "--w", "4", "--h", "1", "--size", "20")
        quiet = kit("shot", str(deck), "--slides", "1", "--dpi", "70", "--expect", "clean")
        self.assertTrue(quiet["clean"])
        self.assertEqual(quiet["flagged_slides"], [])
        self.assertNotIn("expectation_failed", quiet)

        kit("add-text", str(deck), "--slide", "1", "--text", "溢出" * 40, "--x", "1", "--y", "3",
            "--w", "3", "--h", "0.4", "--size", "18")
        flagged = kit("shot", str(deck), "--slides", "1", "--dpi", "70", "--expect", "clean")
        self.assertFalse(flagged["clean"])
        self.assertTrue(flagged["expectation_failed"])
        self.assertEqual(flagged["flagged_slides"], [1])

    def test_shot_compare_to_reports_changed_pages(self):
        deck = self.create()
        kit("add-text", str(deck), "--slide", "1", "--text", "版本一", "--x", "1", "--y", "1", "--w", "5", "--h", "1", "--size", "24")
        kit("clone", str(deck), "1", "--count", "1")
        previous = self.tmp / "previous.pptx"
        edited = kit(
            "add-text", str(deck), "--slide", "2", "--text", "版本一", "--x", "1", "--y", "2",
            "--w", "5", "--h", "1", "--size", "24",
        )
        previous.write_bytes(deck.read_bytes())
        kit("set-text", str(deck), "--slide", "2", "--shape", f"#{edited['id']}", "--text", "版本二")
        shot = kit("shot", str(deck), "--slides", "1,2", "--dpi", "60", "--compare-to", str(previous))
        compared = {row["slide"]: row["changed_px_ratio"] for row in shot["compared"]}
        self.assertEqual(compared[1], 0.0)
        self.assertGreater(compared[2], 0.0)
        self.assertEqual(shot["changed_vs_previous"], [2])

    def test_deck_report_protocol(self):
        sys.path.insert(0, str(ROOT / "python"))
        try:
            from pi_pptx import Deck

            deck = Deck.create(str(self.tmp / "report.pptx"))
            deck.add_text(slide=1, text="标题", x=1, y=1, w=5, size=24)
            deck.save()
            result = deck.report()
        finally:
            sys.path.pop(0)
        self.assertEqual(result["slides"], 1)
        self.assertEqual(result["shapes"], 3)
        self.assertTrue(Path(result["out"]).exists())

    def test_per_page_region(self):
        deck = self.create()
        kit("clone", str(deck), "1", "--count", "1")
        spec = json.dumps({"2": [0.5, 0.5, 3, 2], "default": [0, 0, 13.333, 7.5]})
        shot = kit("shot", str(deck), "--slides", "1,2", "--dpi", "80", "--region", spec, "--overflow", "off")
        sizes = {entry["slide"]: entry["px"][0] for entry in shot["shots"]}
        self.assertLess(sizes[2], sizes[1])

    def test_importable_deck_library(self):
        sys.path.insert(0, str(ROOT / "python"))
        try:
            from pi_pptx import Deck
        finally:
            sys.path.pop(0)

        deck = Deck.create(str(self.tmp / "lib.pptx"))
        deck.add_text(slide=1, text="库生成的标题", x=1, y=1, w=8, size=28)
        deck.card(slide=1, x=1, y=2.4, w=4, h=1.5, text="卡片", size=16)
        deck.bar(slide=1, x=1, y=2.3, w=4)
        result = deck.apply([{"op": "add_shape", "slide": 1, "kind": "chevron", "x": 6, "y": 2.4, "w": 1.2, "h": 1.5}])
        self.assertTrue(result["ok"])
        deck.save()
        info = deck.inspect(slides="1")
        names = [s["name"] for s in info["slides"][0]["shapes"]]
        self.assertIn("TextBox 3", names)
        self.assertIn("Rounded Rectangle 4", names)
        # The library and the CLI operate on the same file.
        cli_info = self.inspect(self.tmp / "lib.pptx", "1")
        self.assertIn("卡" + "片", [s.get("text") for s in cli_info["slides"][0]["shapes"]])

    def test_real_template_with_pictures(self):
        template = os.environ.get("PI_PPTX_TEST_TEMPLATE", "")
        if not template or not Path(template).exists():
            self.skipTest("set PI_PPTX_TEST_TEMPLATE to a real deck to run this")
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

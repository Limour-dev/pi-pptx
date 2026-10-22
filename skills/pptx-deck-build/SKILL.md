---
name: pptx-deck-build
description: 用一个可重跑的 Python 脚本 build_deck.py 从既有 .pptx 模板生成或修复整本 PPTX（pi-pptx）。任何 PPT/幻灯片/演示文稿/汇报/课件的制作、重建、改版都先读这个技能：脚本骨架、ppt_run_deck 的注入环境与 stdout 协议、模板克隆、页面与形状的坐标/下标约定。
---

# 脚本即 deck：build_deck.py

pi-pptx 的核心约定只有一句：**deck 的真源是 `build_deck.py`，不是 `.pptx`**。
编辑 = 改脚本再跑；`Deck.create(..., overwrite=True)` 每次从模板重建，所以"重跑"同时就是 undo。
不要手动编辑 pptx，不要逐个形状调工具，不要拿渲染出的 PNG 当交付物。

## 循环

0. `ppt_doctor`：环境与字体替换。字体替换是排版参数（替换字面宽约 +5%），行宽一开始就留 5–10% 余量。
1. `ppt_inspect(<模板>, mode="summary")`：几个版式、样例页有什么、有没有表格样例。
   `ppt_inspect(<模板>, mode="full", slides="1-3")`：要克隆的页里，形状名、`box_in`、占位符类型。
2. 写 `build_deck.py`（骨架见下；可直接跑通的完整示例在 `../../examples/build_deck.py`）。
3. `ppt_run_deck({script, out, template, shot: "flagged"})` 边写边跑：只渲染本次报 `@@WARN@@` 的页。
4. 干净后 `shot: "changed"`，或对整本 `ppt_shot(montage=3)` 看节奏。
5. 有溢出/错位 → 回脚本改，回到第 3 步。整本轮次通常比逐形状编辑少一个数量级。

## 骨架

```python
import argparse, os
from pathlib import Path
from pi_pptx import Deck                      # PYTHONPATH 由 ppt_run_deck 注入

ROOT = Path(os.environ.get("DECK_ROOT") or Path.cwd())
OUT  = Path(os.environ.get("DECK_OUT") or ROOT / "deck.pptx")
TPL  = os.environ.get("DECK_TEMPLATE")        # 相对路径由工具转成绝对路径

CX, CW, BODY_Y = 0.68, 11.97, 1.84            # 从 ppt_inspect 的样例页量出来，不要猜
PAGES = 4

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", default="", help="只重建这些页，如 '3,7'")
    only = {int(x) for x in ap.parse_args().only.split(",") if x.strip().isdigit()}
    want = lambda n: not only or n in only

    D = Deck.create(OUT, template=TPL, overwrite=True)
    while D.slide_count < PAGES:               # 模板页数不定：克隆样例页补齐
        D.clone(source=1, at=D.slide_count + 1)
    if D.slide_count > PAGES:                  # 页码串 1 起，如 "5-8"
        D.delete_slides(",".join(map(str, range(PAGES + 1, D.slide_count + 1))))

    if want(1):
        print("@@PAGE@@ 1", flush=True)        # 这是 shot="changed" 的依据
        D.add_text(slide=1, x=CX, y=2.4, w=10.5, h=1.0, text="标题", size=40, bold=True)
        D.set_notes(slide=1, text="开场备注")

    D.save()
    D.report()                                 # 打印 @@DECK@@ 计数行

if __name__ == "__main__":
    main()
```

要点：

- 每页一个 `if want(n):` 块，块首 `print("@@PAGE@@ n", flush=True)`；`--only 3,7` 时 22 页能降到 2 页重建。
- `--only` 依赖脚本自己给 `want()`；`ppt_run_deck` 只负责把 `args: ["--only", "3,7"]` 传进去。
- `@@WARN@@` 行不是协议必需，但脚本里做了自检（例如表格底边越界）时用 `print("@@WARN@@ slide 7: bottom_pt 17.7")` 报出来，`shot:"flagged"` 就只渲这些页。
- 末尾 `D.save()` 是原子写（临时文件 + 回读校验 + replace）；失败时上一版 `out` 还在，工具会标 `stale: true`。
- 脚本超过 ~300 行就把文案抽成常量表：22 页的 deck 有 60% 是文案，不是几何。

## 工具注入的环境（不要硬编码）

| 变量 | 含义 |
|---|---|
| `PYTHONPATH` | 已含插件 `python/` → `from pi_pptx import Deck` 直接可用，不要写 `sys.path` |
| `DECK_ROOT` | 最近 git 根（无则 cwd）；脚本 cwd 就是它，`ROOT / "data/x.csv"` 可直接用 |
| `DECK_OUT` / `DECK_TEMPLATE` | `out=` / `template=` 的绝对路径 |
| `PI_PPTKIT_PY` | 后端 CLI，脚本要 shell 出去时用 |
| `PYTHONDONTWRITEBYTECODE=1` / `PI_PPTX_RUN_ID` | 无需关心 |

## 模板与克隆

- `Deck.create(OUT, template=TPL, overwrite=True)` 是对模板做**字节复制**，版式、主题字体、样例页全部保留。
- 样例页既提供设计语言，也带着示例文案：先 `delete_shape` 掉样例内容（图片、占位符文本），保留版式自带的东西。
- 版式自带的 logo / 分隔线 / 底色在 `ppt_inspect` 里是 `inherited: true`，不必再画一遍；也不要改母版/版式。
- `D.clone(source=n, at=k, count=m)` 深拷贝整页（含图片/表格 rels、背景、备注）；`D.add_slide(layout="…")` 从空版式新建，但新页占位符名可能与样例页不同（`Title 1` vs `标题 1`），按 `placeholder_type` 找。
- 由 `add_slide` 新建页留下的空内容占位符，用 `delete_shape` 删掉；只 `set_text(" ")` 会留下项目符号，渲染时是个孤零零的圆点。

## 坐标与下标

| 对象 | 基准 |
|---|---|
| 页码、插入位置、`slides` 选择串 | **1 起**（`"1,3"`、`"2-4"`） |
| 表格 `row`/`col`、形状顶层 `index` | **0 起** |
| 形状 `shape=` | 名称、`"#<id>"`（inspect 返回的 id）或 `index` 字符串；重名时用 `#id` |
| 长度 | 一律英寸（in），画布默认 13.33×7.5；以 `ppt_inspect` 的 `slide_size_in` 为准 |

组合内子形状的 `box_in` 已换算到画布坐标；`ppt_inspect(depth=0)` 可把组合摊平成一项。给 `name=` 的形状重跑后仍然是同一句柄，可以当锚点用。

## 接着读

- `references/deck-api.md`：每个 op 的参数（文字/形状/表格/图片/备注/克隆）、富文本段落、表格样式、错误与寻址。
- `../pptx-layout-design/SKILL.md`：字号、行宽余量、表格高度、卡片与节奏等版面规则。
- `../pptx-visual-review/SKILL.md`：`ppt_shot` / `ppt_inspect` 怎么用来验收，溢出怎么判、怎么修。

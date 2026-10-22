# pi-pptx

让 pi agent **跑一个 build_deck.py 把整本 PPT 画出来，再看图验收**：一个工具负责跑脚本，两个工具负责看结果。

| 层 | 方案 | 作用 |
|---|---|---|
| 解析/写入 | [python-pptx](https://python-pptx.readthedocs.io/) 1.0+ | 读/写版式、形状、表格、备注；克隆整页；重映射 rels |
| 渲染 | LibreOffice headless → PDF → PyMuPDF | 任意页出 PNG/拼图/裁剪；文本溢出、形状重叠、逐形状密度、版本像素对比；PDF/PNG 缓存 |
| 插件 | pi extension `defineTool` | **4 个工具**：`ppt_run_deck` / `ppt_inspect` / `ppt_shot` / `ppt_doctor`；截图以 `ImageContent` 直接回给模型 |
| 库 | `python/pi_pptx/` | 与工具层同一套原语：`from pi_pptx import Deck`，脚本即事务 |

这套工具面来自一次真实产出（22 页、536 个形状、15 轮返工、1054 行 `build_deck.py`）：全程只用了 `ppt_inspect`（3 次）和 `ppt_shot`（14 次），编辑类工具**一次都没用**——因为整本是从零排的版，"改"+ 动作的循环是"改脚本 → 重跑 → 看图"。于是编辑 API 留成库，工具层只留"跑"和"看"。

## 安装

```bash
# 1. 系统依赖
sudo apt install libreoffice            # 提供 soffice
python3 -m pip install python-pptx pymupdf

# 2. 安装本插件（在插件目录内；发布后可用 pi install pi-pptx）
pi install ./pi-pptx                     # 写入 pi 配置
pi -e ./pi-pptx                          # 本次运行临时加载

# 3. 自检
pi -e ./pi-pptx -p "/pptx"               # python-pptx / PyMuPDF / soffice / 字体替换链

# 4. 可选：把后端当普通 Python 库用（脚本化「幂等重建」）
python3 -m pip install -e ./pi-pptx      # 之后 from pi_pptx import Deck
```

中文模板通常使用「微软雅黑 / 思源黑体 / 宋体」。本机若只有 Noto Sans CJK，字形能显示但字宽和行距会有差异（详见「已知边界」）。可以只装字体，或写 fontconfig 别名把它们统一指向 `Noto Sans CJK SC`。

## 工具

| 工具 | 作用 |
|---|---|
| `ppt_run_deck` | **跑 build_deck.py**：注入解释器与后端路径（`PYTHONPATH`/`DECK_ROOT`/`DECK_OUT`/`DECK_TEMPLATE`），解析 `@@PAGE@@`/`@@DECK@@`/`@@WARN@@`，失败时给结构化 traceback，成功后自动只渲"本次改动的页"或"报红的那几页" |
| `ppt_inspect` | 地图：版式目录 + 形状树（`box_in`/文本/表格/图片/备注 + 每个文本框的 `text_fit`）；`mode="summary"` 每页概览、`mode="text"` 纯文字（校对成品）、`depth` 控制组合递归 |
| `ppt_shot` | 眼睛：渲染 PNG（`montage` 拼图 / `region` 裁剪 / `density` 逐形状密度 / `compareTo` 版本对比 / `expect="clean"` 门禁），返回 `overflow[]`/`overlaps[]`/`suppressed[]` |
| `ppt_doctor` | 环境自检：版本、soffice、`font_substitutions`（字体替换 = 排版参数，字宽差约 5%）、模板主题字体 |

### 为什么只有 4 个

一本 22 页的 deck 有 536 个形状：文本框 334、卡片/色条/分隔线 187、表格 12、图片 3。其中 **187 + 334 个形状在调用发生前并不存在**，没有任何 `shape=` 参数能指到它们；逐次 `set_text` 要点 536 次；`apply(ops=[...])` 写 536 条 JSON 比写 1054 行 Python 更难维护，还没有循环、函数和共享常量。

而真正的迭代单元是**整本重建**：`Deck.create(overwrite=True)` + 从头画一遍 = 秒级、幂等、天然就是 undo（15 轮返工里一次都没需要过快照）。所以位置从"编辑 PPTX"移到"运行生成 PPTX 的脚本"。

**只删了 TypeScript 的 `defineTool` 包装，`python/pi_pptx/` 一个文件都没删**：ops 层是脚本的 API，`render.py`/`render_overflow` 是 `ppt_shot` 的实现。下面所有 op 在脚本里照样能用。

## `ppt_run_deck`

```ts
ppt_run_deck({
  script: "report/20261023/build_deck.py",  // 必填，相对 cwd
  out: "report/20261023/deck.pptx",         // 可选；注入 DECK_OUT，并用于失败时判定 stale
  template: "report/PPT模板.pptx",           // 可选；注入 DECK_TEMPLATE
  args: ["--only", "3,7"],                  // 透传给脚本（argparse）
  env: { EXTRA: "1" },                      // 额外环境变量
  python: "python3",                        // 默认 PI_PPTKIT_PYTHON 或 python3
  timeout: 600,
  shot: "changed",                          // false | "changed" | "flagged" | "all"
  dpi: 110,
  overflow: "loose",
  expect: "clean",                          // 有任何 flag 就让 ok=false
  keepVersions: 2,                          // 旧版挪到 <tmp>/pi-pptx/versions/
  compareTo: "<version>.pptx",              // 与上一版逐页像素对比
})
```

### 注入的环境变量

| 变量 | 值 | 脚本侧用法 |
|---|---|---|
| `PYTHONPATH` | 追加 `<插件>/python`（由 `import.meta.url` 解析，与安装位置无关） | `from pi_pptx import Deck`，**脚本里永不出现 `sys.path`** |
| `PI_PPTKIT_PY` | `pptkit.py` 绝对路径 | 脚本需要用 CLI 兜底时 |
| `DECK_ROOT` | 最近的 git root（找不到则 cwd） | `ROOT = Path(os.environ["DECK_ROOT"])` |
| `DECK_OUT` | `out` 绝对化后的路径 | `OUT = Path(os.environ["DECK_OUT"])` |
| `DECK_TEMPLATE` | `template` 绝对化后的路径 | `TPL = Path(os.environ["DECK_TEMPLATE"])` |
| `PYTHONDONTWRITEBYTECODE` | `1` | 别在仓库里留 `__pycache__` |
| `PI_PPTX_RUN_ID` | 本次运行 id | 可写进 `@@DECK@@` 便于对账 |

进程 `cwd` 固定为 `DECK_ROOT`（不是脚本所在目录）——脚本里通常写 `ROOT / "data/x.csv"` 这种仓库相对路径。

### stdout 协议（可选，强烈建议）

脚本按行打印，插件只认三种前缀，其余原样进 `stdout_tail`：

```
@@PAGE@@ 3                                  ← 开始画第 3 页；崩溃时能报"死在第 3 页"
@@DECK@@ {"out":"…","slides":22,"shapes":536,"tables":12,"pictures":3,"notes":22}
@@WARN@@ slide 7: TextBox 12 bottom_pt 17.7  ← 脚本自检已知问题；shot="flagged" 只渲这些页
```

`Deck.report()` 直接输出 `@@DECK@@` 这一行（含本次改动页 `changed_slides`），脚本零成本。没有 `@@DECK@@` 时，插件对 `out` 做一次 `inspect --mode summary` 兜底填计数。

### 返回值

```json
{
  "ok": true, "exit_code": 0, "duration_s": 3.4, "python": "python3",
  "script": "…/build_deck.py", "root": "…", "out": "…/deck.pptx", "stale": false,
  "deck": { "slides": 22, "shapes": 536, "tables": 12, "pictures": 3, "notes": 22 },
  "changed_slides": [1, 2, 3],
  "warnings": ["slide 7: TextBox 12 bottom_pt 17.7"],
  "stdout_tail": "…最后 40 行…",
  "shot": { "shots": [ … 同 ppt_shot 的 overflow/overlaps … ], "flagged_slides": [7] }
}
```

失败时：`ok:false` + `error {type,message,file,line}` + `traceback_tail`（最后 30 行 / 8 KB）+ `stale:true`（`out` 仍是上一版：`Deck.save()` 原子写，不会留半成品）。`stdout_tail` 硬截断 40 行 / 16 KB，防止脚本把整个 `DataFrame` deparse 进上下文。**失败不删 `out`**：上一版能看、能 shot、能对比。

### shot 模式

| 值 | 渲哪些页 | 用途 |
|---|---|---|
| `"changed"`（默认） | `@@DECK@@.changed_slides`（退化为 `@@PAGE@@`） | 日常迭代 |
| `"flagged"` | `@@WARN@@` 里提到的页；没有就不渲 | **最省 token**：多数轮次只看上一轮报红的 1–2 页 |
| `"all"` | 前 6 页（`--max-pages`），其余见 `omitted_slides` | 阶段性全量看一遍 |
| `false` | 不渲 | 还没到看图那一步 |

## 坐标与下标约定

| 对象 | 基准 |
|---|---|
| 页码、插入位置 | **1 起**（`slides` 选择串如 `"1,3"`、`"2-4"`） |
| 表格 `row` / `col` | **0 起**（与 `inspect` 的 `table` 数组一致） |
| 形状顶层 `index` | **0 起** |
| 形状 `#<id>` | `ppt_inspect` 返回的 `id`，重名/组合/下标歧义时用它 |
| `name=` | 每个 op 都接受；重建后仍是同一句柄 |
| `ignore=` | 名称、`#<id>`、`/regex/`（同时匹配 name 与 path） |

`box_in` 一律是英寸，原点在画布左上；默认画布 13.33 × 7.5 in（16:9）。组合内子形状的 `box_in` 已换算到画布坐标。**所有 `h` 都要自己按"字号 × 行数 × 1.35 + 0.25"估**——卡片不会自动长高，`add_table` 也不会自动适应。

## Agent 工作流

插件在检测到 PPT 类请求时（prompt 含 ppt/pptx/slides/幻灯片/演示文稿/模板等关键词）注入 `pptx_workflow`：

```
deck 的唯一真源是 report/<date>/build_deck.py。不要手动改 pptx，改脚本再 ppt_run_deck。
0. ppt_doctor / ppt_inspect 先拿字体替换与画布尺寸，行宽留 5–10% 余量。
1. ppt_inspect mode="summary" 看模板：几个 layout、样例页有什么、有没有表格样例。
2. 写 build_deck.py：常量（CX/CW/每格高）+ 9 个语义包装（tx/card/bar/title/kicker/table/pic/bullets/notes），
   裸 D.op() 只留给边角；每页一个小节，编号与页号一致；text_fit 是"能不能放下"的预检。
3. 表格高度自己算：h = header_h + row_h × (rows-1)，包装器返回底边给下一块用。
4. ppt_run_deck(shot="flagged") 迭代；干净了再 shot="changed" 或 all 全量看一遍。
5. 数字：slides 1 起；row/col 与 shape index 0 起；画布 13.33×7.5 in。
   撤销 = 重跑脚本。"改了 A 有没有挤坏 B" → ppt_shot compareTo=<上一版>。
6. 复用模板已有设计（版式自带的 logo/分隔线/底纹不必重画）；20+ 卡片用数据列表批量生成。
7. 脚本超过 ~300 行就把文案抽成常量表（22 页的 deck 有 60% 是文案）。
```

## build_deck.py 骨架

完整可跑示例见 [`examples/build_deck.py`](examples/build_deck.py)：4 页、语义包装器、表格返回底边、`--only` 单页重建、`@@PAGE@@`/`@@DECK@@`。核心结构：

```python
from pi_pptx import Deck                       # PYTHONPATH 由 ppt_run_deck 注入
ROOT = Path(os.environ.get("DECK_ROOT", Path.cwd()))
OUT  = Path(os.environ.get("DECK_OUT", ROOT / "deck.pptx"))

D = Deck.create(OUT, template=TPL, overwrite=True)   # 重跑即撤销
while D.slide_count < PAGES:                          # 模板页数不定：克隆样本页补齐
    D.clone(source=1, at=D.slide_count + 1)

def tx(s, x, y, w, text=None, *, h=None, size=13, **kw):   # 语义包装，覆盖 95% 的形状
    h = h or round(0.1 + size * 1.35 / 72, 2)
    return D.op("add_text", slide=s, x=x, y=y, w=w, h=h, text=text, size=size, **kw)

def table(s, x, y, w, data, widths, *, row_h=0.32, head_h=0.42):
    h = round(head_h + row_h * (len(data) - 1), 2)          # ← 自己算，别指望自动适应
    D.op("add_table", slide=s, rows=len(data), cols=len(data[0]), x=x, y=y, w=w, h=h,
         data=data, col_widths=widths, row_height=row_h, header_height=head_h)
    return y + h                                            # ← 返回底边，下游接着排

if want(3):                                     # --only 靠它把 22 页重建降到 1 页
    print("@@PAGE@@ 3", flush=True)
    bottom = table(3, CX, BODY_Y, 7.2, DATA, [1.4, 3.4, 2.4])
    tx(3, CX, bottom + 0.2, 7.2, "说明文字", size=10.5)

D.save()
D.report()                                      # 打 @@DECK@@：计数 + 本次改动页
```

## 直接用 Python 后端

工具层删掉的只是 TS 包装；CLI 与库都还在（`python3 -m pi_pptx.cli` 或旧的 `python pptkit.py`）。所有命令输出单行 JSON，失败时 stderr 给原因并以非 0 退出；默认原地修改，`--out` 写副本。

```bash
python3 python/pptkit.py doctor 模板.pptx
python3 python/pptkit.py inspect 模板.pptx --slides 1-3
python3 python/pptkit.py inspect deck.pptx --mode summary
python3 python/pptkit.py inspect deck.pptx --mode text          # 校对成品文字
python3 python/pptkit.py create deck.pptx --template 模板.pptx
python3 python/pptkit.py clone deck.pptx 3 --at 2 --count 2
python3 python/pptkit.py set-text deck.pptx --slide 2 --shape '#12' --text $'第一行\n第二行'
python3 python/pptkit.py add-shape deck.pptx --slide 2 --kind roundrect --x 1 --y 1.5 --w 4 --h 1.4 \
    --fill '#F4F7FB' --line '#CAD9E8' --radius 0.06 --text '卡片' --size 16 --name 'Card 1'
python3 python/pptkit.py add-table deck.pptx --slide 2 --rows 3 --cols 2 --x 1 --y 3 --w 8 --h 2 \
    --header-height 0.48 --row-height 0.7 --first-col-bold --border none
python3 python/pptkit.py table-style deck.pptx --slide 2 --shape '#9' --header-fill 1F4E79 --border none
python3 python/pptkit.py set-notes deck.pptx --slide 4 --text '讲者备注…'
python3 python/pptkit.py shot deck.pptx --slides 1-6 --montage 3
python3 python/pptkit.py shot deck.pptx --slides 6 --region 0.5,4,6,2.5 --dpi 150
python3 python/pptkit.py shot deck.pptx --only-changed --density --expect clean
python3 python/pptkit.py shot deck.pptx --slides 1-6 --compare-to 上一版.pptx
python3 python/pptkit.py notes-check deck.pptx                  # 备注是否错位一页
python3 python/pptkit.py apply deck.pptx --ops ops.json          # 批量事务（脚本化的旧入口）
python3 python/pptkit.py delete deck.pptx --slides 4-6
```

## 当库用（幂等重建）

```python
from pi_pptx import Deck

deck = Deck.create("报告.pptx", template="模板.pptx", overwrite=True)
deck.clone(source=1, count=2)                       # 或 deck.add_slide("3_版本内页")
deck.set_text(slide=2, shape="#12", text="标题")
deck.card(slide=2, x=0.9, y=1.8, w=4.2, h=1.5, text="卡片", size=16)
deck.bar(slide=2, x=0.845, y=1.8, w=0.055)
deck.add_table(slide=2, rows=3, cols=2, x=8.2, y=1.8, w=4.2, h=2,
               data=[["指标", "值"], ["LVEF", "58"]], header_height=0.48, first_col_bold=True)
deck.rule(slide=2, x=0.9, y=6.6, w=11.5)
deck.apply([{"op": "add_shape", "slide": 3, "kind": "chevron", "x": 2, "y": 3, "w": 1.4, "h": 0.9}])
deck.save()
deck.report()                                       # @@DECK@@ 行
print(deck.inspect(slides="2")["slides"][0]["shape_count"])
print(deck.shots(only_changed=True)["shot_slides"])
```

- 每个方法就是一个 op：`deck.op("table_style", slide=2, shape="#9", border="none")` 能调任何 op。
- 修改都在内存里；`save()` 才落盘（临时文件 + 完整性回读 + 原子替换）。重跑脚本即撤销。
- `deck.changed` 记录被改过的页，`deck.shots(only_changed=True)` 只渲这几页。
- 没装库时等价写法：`PYTHONPATH=python python3 build_deck.py`。

## 实现要点

- **原位路径解析**：插件用 `fileURLToPath(new URL("../python", import.meta.url))` 找到自家后端，`PYTHONPATH` 与 `PI_PPTKIT_PY` 都由它推导。开发目录、安装目录、npm 包布局都不需要写死路径。
- **形状寻址**：`#<id>` 优先；重名时**拒绝执行**并列出 id/name/type/box/文本候选表；找不到时给最相近的 3 个路径。顶层 `index` 0 起，越界报错会说明该约定。
- **中文字体**：所有写 run 的路径同时写 `a:latin` / `a:ea` / `a:cs`。python-pptx 的 `run.font.name` 只写 `a:latin`，对中文是静默无效的。
- **新文本框版式**：省略 `h` 时按文本估算高度并写 `<a:spAutoFit/>`（grow）；`fit='shrink'` 写 `<a:normAutofit/>`；支持 `spcAft/spcBef`、`marL/indent` 悬挂缩进、`buClr/buFont/buChar/buNone`。
- **混排文字**：`paras=[{runs:[…]}]` 保留原段落的 pPr 和首个 run 的 rPr，未指定的字段继续继承模板样式。
- **表格**：`add_table` 默认去掉主题 `firstRow/bandRow` 和 `a:tableStyleId`，使用中性边框；**必须删掉 `a:tableStyleId`**，否则主题表样式会盖住手工填色。另有 `header_height`（表头行高 ≠ 正文行高）、`first_col_bold/first_col_color`、`header_border`、`cell_runs[row][col]=[{t,bold,color}]`、`border="none"`。
- **克隆页必须重映射 rels**：深拷贝 `spTree` 后遍历所有 `r:id`/`r:embed`，把源页关系重新绑到目标页，再复制显式背景和备注。不做这一步图片/图表会静默丢失（并且 PowerPoint 不报错）。
- **改文字只动 run**：保留第一个 run 的字体/字号/颜色，删除多余 run；多段落通过深拷贝首段 XML 实现，继承项目符号、行距等段落属性。
- **原子保存 + 完整性回读**：写临时文件 → 重新 `Presentation()` 读取校验 → `os.replace`，避免半写坏文件。
- **渲染缓存**：PDF 以 `(绝对路径, mtime_ns, size, sha256)` 的 sha1 为 key（内容相同才命中）；PNG 以 `key + dpi` 缓存。返回里带 `cache_hit`，连续 `shot` 未改动的文件基本只花 PyMuPDF 出图时间（约 0.1 s/页）。
- **并发安全**：每次 soffice 调用使用独立 `-env:UserInstallation` profile，并在独立临时目录转换后原子移动 PDF；`ppt_run_deck` 与所有写文件的工具都经过 pi 的 `withFileMutationQueue(path)`。
- **`text_fit`（预检，不用渲染）**：`est_lines` 按 CJK=1em / Latin≈0.52em 估宽，CJK 逐字可断行、Latin 只在空格断行（所以 `01` 这种宽 token 算"横向溢出"而不是换行）；`needed_in = lines × size × 1.35 + 内边距`；`fits` 比较**行数**而不是英寸——13pt 单行放进 0.28in 的标签框几何上"超"几分，渲染却完全正常。
- **溢出检测（启发式）**：渲染后按**行**取 PDF 文本 bbox（避免同一基线的两个文本框被合并成一个 block），每行只分配给文本最匹配、离上沿最近的一个形状，再与形状 bbox 比较。**按锚点分侧**：`anchor="t"` 不查 top、`anchor="b"` 不查 bottom；被跳过的判定进 `suppressed[]` 并带 `reason`。模板固有几何（顶锚点、溢漏 < 5pt）**直接不报**——它在 22 页里页页出现，会淹没真问题。每条 overflow 都带 `rendered_lines` / `needed_in` / `slack_in`，所以"改框还是改字号"一眼可判。`overflow=loose`（默认）跳过居中/两端对齐的左右检查；`strict` 报告每一侧；`off` 关闭。
- **重叠检测**：只比较有内容的顶层形状，文本×文本重叠比 > 22%、文本×图片/表格 > 35%（交叠面积 ≥ 0.05 in²）时报警。面积一律**平方英寸且裁剪到两框交集/较小框**（曾把 pt 与 in 混算，报出 1e11 in²）。容器与内容不算重叠（覆盖 ≥ 92% 视为版式设计，记入 `suppressed[]`）。
- **逐形状密度**：`--density` 给每个文本框的 `rendered_lines` / `needed_in` / `box_h_in` / `slack_in` / `capacity_lines` / `tight`。"字比卡片高"永远是**某一个框**太矮，页级 `lines vs capacity` 平均只会制造假阳性（旧实现把 48 行表格算成 20 行容量）。
- **版本对比**：`--compare-to 上一版.pptx` 把两本 deck 渲染成灰度位图后异或计差异字节，给出每页 `changed_px_ratio`；`changed_vs_previous` 列出非零页。这是"改了 A 结果挤坏 B"唯一可靠的检测。
- **差分渲染**：`shot --only-changed` 把每页 XML 的 sha1 存到 `/tmp/pi-pptx/state-<path hash>/slides.json`，只渲与上次不同的页；没有改动时返回 `no_changes: true`。
- **拼图与局部裁剪**：`--montage N` 用 PyMuPDF 直接拼 N 列 contact sheet（不依赖 Pillow），一次看完整本节奏；`--region x,y,w,h` 按英寸裁剪，也可按页给 `{"3":[x,y,w,h],"default":[…]}`。
- **按页限量**：`ppt_shot` 普通模式单次最多 6 页（约 1.6k tokens/页 @110dpi）；`montage` 只回一张拼图，最多 60 页。超过部分返回 `omitted_slides`。拼图每格约 640px 宽——**12pt 脚注读不出来**，文字校对请 100–110 dpi 单页。
- **字体回显**：`ppt_inspect`/`ppt_shot`/`ppt_doctor` 都带 `font_substitutions`（渲染 ≠ PowerPoint，行宽留 5–10%）。fc-match 结果按 1 小时缓存在 `/tmp/pi-pptx/fonts.json`。
- **备注错位 lint**：`notes_check()` 给每页 `notes_len` + 前 20 字 + 该页标题——"先写备注再删模板页"导致整体错位一页就是靠它发现的。

## 配置

| 环境变量 | 默认值 | 作用 |
|---|---|---|
| `PI_PPTKIT_PYTHON` | `python3` | 解释器路径（如虚拟环境里的 python） |
| `PI_PPTKIT_PY` | 包内 `python/pptkit.py` | 后端脚本路径（默认由插件位置推导） |
| `PI_PPTX_TEST_TEMPLATE` | 未设置 | 后端测试用的真实模板；设置后额外验证模板 `strict` 0 告警 |
| `PI_PPTKIT_DEBUG` | 未设置 | 设置后显示 Python traceback |

`ppt_run_deck` 另外注入 `DECK_ROOT` / `DECK_OUT` / `DECK_TEMPLATE` / `PYTHONPATH` / `PYTHONDONTWRITEBYTECODE` / `PI_PPTX_RUN_ID`（见上表）。渲染缓存、临时 profile、旧版本（`keepVersions`）都在 `/tmp/pi-pptx/`，可随时删除。

## 已知边界

- 不支持 SmartArt、动画/切换、节（section）；选中窗格层级等细节也拿不到。
- 不能新建母版/版式；不要改模板版式（会影响所有引用页）。
- **工具层没有单点编辑**：`ppt_set_text`/`ppt_add_shape`/`ppt_apply` 这类工具已删除，等价写法在脚本里（`D.set_text(...)` / `D.op("add_shape", ...)` / `D.apply([...])`）。CLI 子命令都保留。
- 字体缺失时 LibreOffice 会替换，跨行位置可能与 PowerPoint 差 1–2 个字（约 5% 字宽）。溢出结果请留 5–10% 余量；要严格对齐请安装原字体。`ppt_doctor <deck>` 会输出主题 major/minor 字体与替换链。
- 溢出/重叠/密度都是启发式信号，不能替代看渲染图；文字超出页面被裁剪时不会报告。
- 同一 layout 新建页的占位符名可能与模板样例页不同（如 `Title 1` vs `标题 1`）；`ppt_inspect` 里按 `placeholder_type` 找，或直接用返回的 `#id`。
- 没有内建撤销：撤销 = 从模板重建（`Deck.create(..., overwrite=True)` 的重跑），或 `keepVersions` + `compareTo` 做版本对照。
- 备注整页错位不会自己报错：删页后跑一次 `notes_check()`。
- `add_shape` 的文字不会自动长高（形状是设计元素，不是文本框）：文字长了请自己给够 `h`，或用 `add_text` 覆盖在卡片上。
- 中文文件名正常支持；soffice 的输出目录必须存在（插件已处理）。

## 本次改进（2026-10-23，22 页实战复盘）

| 建议 | 落地 | 验收（自动测试） |
|---|---|---|
| 工具面收紧到 4 个 | 删掉 16 个编辑类 `defineTool`（`create/clone/add_slide/set_text/set_cell/add_text/add_shape/add_image/add_table/set_shape/delete_shape/set_notes/notes_check/table_style/apply/delete`），保留 Python op 层 | 扩展测试 `registers only the run/inspect/shot/doctor toolset` |
| 新增 `ppt_run_deck` | 注入 `PYTHONPATH`/`DECK_ROOT`/`DECK_OUT`/`DECK_TEMPLATE`；`@@PAGE@@`/`@@DECK@@`/`@@WARN@@` 协议；结构化 traceback；`stdout_tail` 截断；`stale` 判定；`keepVersions`；`shot=changed/flagged/all/false` | `ppt_run_deck … shoots changed pages`、`… structured failure`、`shot=flagged` |
| `area_in2` 量级错误 | `rendered_text_boxes` 把 pt 当 in（差 12700×）；改为 `/72`，面积裁剪到较小框、`ratio ≤ 1` | `test_overlap_area_is_square_inches` |
| 模板标题页页报 `top spill 2.6pt` | 顶锚点 + 溢漏 < 5pt 直接不报（也不进 `suppressed[]`） | `test_template_title_leading_is_not_reported`、`test_template_page_is_clean_under_strict` |
| 单行标签假阳性难判定 | overflow 带 `rendered_lines`/`needed_in`/`slack_in` | `test_overflow_reports_measured_lines_and_height` |
| `density` 细化到形状 | 每框 `rendered_lines`/`needed_in`/`box_h_in`/`slack_in`/`capacity_lines`/`tight`；`tight` 按行数判定 | `test_shot_density_table`、扩展测试 `ppt_shot exposes measured overflow…` |
| `compare_to` 版本对比 | `--compare-to` → 每页 `changed_px_ratio` + `changed_vs_previous`；`ppt_shot(compareTo=)` / `ppt_run_deck(compareTo=)` | `test_shot_compare_to_reports_changed_pages` |
| `expect="clean"` 构建门禁 | `clean` / `expectation_failed` / `flagged_slides`；`ppt_run_deck(expect="clean")` 让 `ok=false` | `test_shot_expect_clean_and_flagged_slides`、扩展测试 |
| `inspect` 补 `text_lines`/`box_in` 一致性 | 每个文本形状 `text_fit`（`est_lines`/`capacity_lines`/`needed_in`/`slack_in`/`fits`），排版前就能查 | `test_inspect_text_fit_and_text_mode` |
| `inspect mode="text"` | 纯 `{shape,id,text}` 扁平表，用于校对成品术语/错别字 | 同上 |
| `ppt_doctor` 字体信息前置 | `ppt_inspect` 也返回 `font_substitutions`；`ppt_doctor(path=)` 读主题字体 | 同上 |
| 文档纪律替换 | `pptx_workflow` 改成"脚本是唯一真源"的 8 条 | 扩展测试 `injects the pptx workflow…` |

## 开发与测试

```bash
npm install
PI_PPTX_TEST_TEMPLATE=/path/to/模板.pptx ./tests/smoke.sh   # py_compile + pyflakes + 后端 + 扩展 + tsc
npm run test:backend      # 仅 Python 后端
npm run test:extension    # 仅扩展（stub pi API，无需模型）
```

`tests/test_backend.py` 在临时目录自建合成模板（含图片/表格/备注），覆盖克隆保真、样式保留、单元格修改、中文字体 XML、批量事务/回滚、表格样式、备注、拼图/裁剪、溢出与重叠检测、缓存；以及本轮验收点：模板 `strict` 0 告警、顶锚点模板 leading 不报、overflow 的实测行高、`area_in2` 单位、`text_fit`/`mode=text`、逐形状密度、`expect=clean`、`compare_to`、`Deck.report()`。设置 `PI_PPTX_TEST_TEMPLATE` 后会额外验证真实模板。

`tests/extension.test.ts` 用 stub 的 `ExtensionAPI` 加载真实扩展，直接调用 `tool.execute()`，验证：只注册 4 个工具、`ppt_run_deck` 注入环境并跑通真实脚本（脚本里没有 `sys.path`）、`@@DECK@@` 与 `inspect` 兜底、失败时的结构化 traceback + `stale`、`shot=flagged` 只渲报红页、`text_fit`/`mode=text`、`density`+`expect`+`compareTo`，以及工作流的条件注入。

## 目录

```
pi-pptx/
├── extensions/pi-pptx.ts    # pi 工具层（4 个工具 + 工作流提示）
├── python/pptkit.py         # CLI 入口（薄壳）
├── python/pi_pptx/
│   ├── __init__.py          # 可 import 的 Deck API（含 report()）
│   ├── utils.py             # 常量/错误/参数解析/原子读写/ShapeFilter/deck_report
│   ├── shapes.py            # 形状树、寻址、inspect（full/summary/text）、备注 lint
│   ├── text.py              # run/paragraph/字距/项目符号/text_fit 预估
│   ├── tables.py            # 表格样式（填充/边框/表头行高/富文本单元格）
│   ├── slides.py            # 克隆（rels 重映射）/版式查找/删页
│   ├── render.py            # 渲染、缓存、拼图、溢出/重叠/密度/版本对比
│   ├── ops.py               # op 层（CLI、apply、Deck 共用）
│   ├── doctor.py            # 环境自检 + 字体替换（带缓存）
│   └── cli.py               # argparse 子命令 + main()
├── examples/build_deck.py   # 推荐的 build_deck.py 骨架（4 页、语义包装、表格返回底边）
├── examples/cards_demo.py   # 只用库原语重建「5 卡片 + 4 chevron」页
├── tests/                   # 后端 + 扩展测试
├── package.json             # pi 包清单（pi.extensions）
├── pyproject.toml           # pip install -e .（from pi_pptx import Deck）
└── tsconfig.json
```

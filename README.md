# pi-pptx

让 pi agent 真正“画” PPT：读模板、克隆样本页、批量改文字/表格/图片/备注，并把改完的每一页**渲染成图片看一遍**再继续。

实现依据调研笔记 `~/pi-test/pptkit/RESEARCH.md` 的三层结论：

| 层 | 方案 | 作用 |
|---|---|---|
| 解析/写入 | [python-pptx](https://python-pptx.readthedocs.io/) 1.0+ | 读版式/形状/表格/备注；克隆整页；改文字；加元素 |
| 渲染 | LibreOffice headless → PDF → PyMuPDF | 任意页出 PNG/拼图/局部裁剪；文本溢出与形状重叠检测；PDF/PNG 缓存 |
| 插件 | pi extension `defineTool` | 16 个 `ppt_*` 工具 + 批量事务入口；截图以 `ImageContent` 直接回给模型 |

## 安装

```bash
# 1. 系统依赖
sudo apt install libreoffice            # 提供 soffice
python3 -m pip install python-pptx pymupdf

# 2. 安装本插件（二选一）
pi install /home/limour/pi-pptx          # 写入 pi 配置
pi -e /home/limour/pi-pptx               # 本次运行临时加载

# 3. 自检
pi -e /home/limour/pi-pptx -p "/pptx"    # 查看 python-pptx / PyMuPDF / soffice / 字体
```

中文模板通常使用「微软雅黑 / 思源黑体 / 宋体」。本机若只有 Noto Sans CJK，字形能显示但字宽和行距会略有差异（详见下文「已知边界」）。可以只装字体，或写 fontconfig 别名把它们统一指向 `Noto Sans CJK SC`。

## 工具

| 工具 | 作用 |
|---|---|
| `ppt_inspect` | 版式目录 + 形状树 JSON；`mode="summary"` 给每页概览，`depth` 控制组合递归 |
| `ppt_shot` | 渲染页面为 PNG（可 `montage` 拼图、`region` 局部裁剪），返回溢出/重叠警告 |
| `ppt_create` | 新建空白 16:9，或**逐字节复制模板**作为工作副本 |
| `ppt_clone` | 克隆已有页（图片/图表/组合/背景/备注全部保留，重映射 rels） |
| `ppt_add_slide` | 按版式新建页（返回每个新形状的 id/name/type/box） |
| `ppt_set_text` | 改形状文字，保留原样式；支持 `paras` 段内混排（加粗小标题 + 正文） |
| `ppt_set_cell` | 改表格单元格（行列从 0 开始，与 `inspect` 的数组一致） |
| `ppt_add_text` | 按英寸坐标加文本框；自动 grow、段间距、悬挂缩进、项目符号、中英文字体同时生效 |
| `ppt_add_image` | 按英寸坐标插图，可只给宽或高保持比例 |
| `ppt_add_table` | 加表格（默认关闭主题表样式，避免与设计模板打架） |
| `ppt_table_style` | 表格样式：表头填充/字色、斑马行、边框、列宽、行高、对齐、内边距 |
| `ppt_delete_shape` | 删除形状（清掉空占位符残留的项目符号、模板自带文本框） |
| `ppt_set_notes` | 写/清空讲者备注 |
| `ppt_apply` | **批量事务**：几十个 op 一次提交，失败整体回滚；`dryRun` 预检 + 自动溢出汇总 |
| `ppt_delete` | 按页码/范围删页，至少保留 1 页 |
| `ppt_doctor` | 环境自检：版本、字体替换链；传 `path` 可读出模板主题字体 |

### 坐标与下标约定

| 对象 | 基准 |
|---|---|
| 页码、插入位置 | **1 起** |
| 表格 `row` / `col` | **0 起**（与 `inspect` 的 `table` 数组一致） |
| 形状顶层 `index` | **0 起** |
| 形状 `#<id>` | `ppt_inspect` 返回的 `id`，**优先使用**：重名、组合、下标歧义都不会出错 |

`box_in` 一律是英寸，原点在画布左上；默认画布 13.33 × 7.5 in（16:9）。组合内子形状的 `box_in` 已换算到画布坐标。

`ppt_set_text(text="")` 明确表示“清空文字但保留形状”（空占位符仍会画出项目符号）；要移除形状请用 `ppt_delete_shape`。

## Agent 工作流

插件在**检测到 PPT 类请求**时（prompt 含 ppt/pptx/slides/幻灯片/演示文稿/模板等关键词）向系统提示注入一段 `pptx_workflow`，要求 agent 按下面闭环工作：

```
inspect（大稿先用 summary）
  → shot（样本页 / montage 总览）挑“样式母本”
  → create（从模板复制工作副本，绝不直接改模板）
  → add_slide（设计在母版/版式里）或 clone（设计在这一页里）
  → set_text / set_cell / add_* / delete_shape / set_notes
  → shot / apply 的自动溢出汇总，看 overflow / overlaps，再改
```

**何时用 `ppt_apply`**：页数 > 5 或形状级操作 > 20 次时，写一个 `ops` 数组一次提交，而不是逐次调用。先 `dryRun=true` 预检，再正式执行；失败时文件完全不动（原子提交），所以“从模板重建”仍是最终的撤销手段：

```ts
ppt_apply({
  path: "deck.pptx",
  ops: [
    { op: "clone", source: 3, at: 4 },
    { op: "set_text", slide: 4, shape: "#12", paras: [
      { runs: [
        { t: "一句实话：", bold: true, color: "C00000" },
        { t: "没有事前效能计算。" },
      ]},
    ]},
    { op: "add_text", slide: 4, x: 0.9, y: 1.5, w: 11.5,
      text: "第二段", size: 18, space_after: 6, bullet: "•" },
    { op: "delete_shape", slide: 4, shape: "#7" },
    { op: "set_notes", slide: 4, text: "这一页讲三件事：…" },
  ],
  shot: true,
})
```

`ops` 支持 `set_text`、`set_cell`、`add_text`、`add_image`、`add_table`、`delete_shape`、`set_notes`、`table_style`、`add_slide`、`clone`、`delete_slides`；后续 op 能看到前面 op 的结果。返回值包含每个 op 的结果、带下标的 `errors`、`changed_slides`，以及改动页的 `overflow` / `overlaps` 汇总。

Python 后端同样可批量调用（`--ops` 支持 `@file`）：

```bash
python3 python/pptkit.py apply deck.pptx --ops ops.json --dry-run
python3 python/pptkit.py apply deck.pptx --ops ops.json            # 默认自动 shot
```

## 直接使用 Python 后端

```bash
cd /home/limour/pi-pptx

python3 python/pptkit.py doctor 模板.pptx
python3 python/pptkit.py inspect 模板.pptx --slides 1-3
python3 python/pptkit.py inspect deck.pptx --mode summary
python3 python/pptkit.py create deck.pptx --template 模板.pptx
python3 python/pptkit.py clone deck.pptx 3 --at 2 --count 2
python3 python/pptkit.py set-text deck.pptx --slide 2 --shape '#12' --text $'第一行\n第二行'
python3 python/pptkit.py set-text deck.pptx --slide 2 --shape '#12' --paras paragraphs.json
python3 python/pptkit.py set-cell deck.pptx --slide 2 --shape '#9' --row 1 --col 1 --text 58
python3 python/pptkit.py add-text deck.pptx --slide 2 --text "补充说明" --x 0.8 --y 6.4 --w 8 --size 16 --bullet '•' --space-after 6
python3 python/pptkit.py add-text deck.pptx --slide 2 --text "自动长高" --x 1 --y 1 --size 18 --fit grow
python3 python/pptkit.py table-style deck.pptx --slide 2 --shape '#9' --header-fill 1F4E79 --header-color FFFFFF --zebra EAF1F8 --border none
python3 python/pptkit.py delete-shape deck.pptx --slide 4 --shape '#7'
python3 python/pptkit.py set-notes deck.pptx --slide 4 --text '讲者备注…'
python3 python/pptkit.py shot deck.pptx --slides 1-6 --montage 3
python3 python/pptkit.py shot deck.pptx --slides 6 --region 0.5,4,6,2.5 --dpi 150
python3 python/pptkit.py apply deck.pptx --ops ops.json
python3 python/pptkit.py delete deck.pptx --slides 4-6
```

所有命令输出单行 JSON；失败时 stderr 给出中文原因并以非 0 退出。默认原地修改，`--out` 可写副本。

## 实现要点

- **批量事务**：`ppt_apply` 在内存里依次执行 op，全部成功才 `save_deck` 一次；`stop_on_error=true`（默认）下任一 op 失败即中止且不落盘，错误带 op 下标。`dry_run` 不写目标文件；开启 `shot`（默认）时会把内存结果存到临时副本渲染一次，提前报出“预计溢出/重叠”。
- **形状寻址**：`#<id>` 优先；重名时**拒绝执行**并列出 id/name/type/box/文本候选表；找不到时给最相近的 3 个路径。顶层 `index` 0 起，越界报错会说明该约定。
- **中文字体**：`ppt_add_text(font=)` 同时写 `a:latin` / `a:ea` / `a:cs`。python-pptx 的 `run.font.name` 只写 `a:latin`，对中文是静默无效的。
- **新文本框版式**：省略 `h` 时按文本估算高度并写 `<a:spAutoFit/>`（grow）；`fit='shrink'` 写 `<a:normAutofit/>`；支持 `spcAft/spcBef`、`marL/indent` 悬挂缩进、`buClr/buFont/buChar/buNone`。
- **混排文字**：`paras=[{runs:[…]}]` 保留原段落的 pPr 和首个 run 的 rPr，未指定的字段继续继承模板样式。
- **表格**：`ppt_add_table` 默认去掉主题 `firstRow/bandRow` 和 `a:tableStyleId`，使用中性边框；模板无表格样例时返回明确提示。`ppt_table_style` 可逐格刷填充/边框/字号/对齐，并支持列宽、行高、内边距。
- **克隆页必须重映射 rels**：`ppt_clone` 深拷贝 `spTree` 后，遍历所有 `r:id`/`r:embed`，把源页关系重新绑到目标页，再复制显式背景和备注。不做这一步图片/图表会静默丢失（并且 PowerPoint 不报错）。
- **改文字只动 run**：`ppt_set_text` 保留第一个 run 的字体/字号/颜色，删除多余 run；多段落通过深拷贝首段 XML 实现，继承项目符号、行距等段落属性。
- **原子保存 + 完整性回读**：写临时文件 → 重新 `Presentation()` 读取校验 → `os.replace`，避免半写坏文件。
- **渲染缓存**：PDF 以 `(绝对路径, mtime_ns, size, sha256)` 的 sha1 作为 key 缓存（内容相同才命中）；PNG 以 `key + dpi` 缓存。返回里带 `cache_hit`，连续 `shot` 未改动的文件基本只花 PyMuPDF 出图时间（约 0.1 s/页）。
- **并发安全**：每次 soffice 调用使用独立 `-env:UserInstallation` profile，并在独立临时目录转换后原子移动 PDF，两个并发转换不会互相打架。
- **溢出检测（启发式）**：渲染后按**行**取 PDF 文本 bbox（避免同一基线的两个文本框被合并成一个 block），每行只分配给文本最匹配、离上沿最近的一个形状，再与 `inspect` 的形状 bbox（EMU→pt，按 PDF 页面比例缩放）比较。`overflow=loose`（默认）跳过居中/两端对齐文本的左右检查；`strict` 全部报告；`off` 关闭。`ignore=["TextBox 9"]` 可跳过指定形状。
- **重叠检测**：只比较有内容的顶层形状，文本×文本重叠比 > 22%、文本×图片/表格 > 35%（且交叠面积 ≥ 0.05 in²）时报警。比溢出更容易发现“看不见的错”。
- **拼图与局部裁剪**：`--montage N` 用 PyMuPDF 直接拼 N 列 contact sheet（不依赖 Pillow），一次看完整本节奏；`--region x,y,w,h` 按英寸裁剪，省掉调 dpi 重渲。
- **按页限量**：`ppt_shot` 普通模式单次最多 4 页（约 1.6k tokens/页 @110dpi）；`montage` 模式最多 12 页但只回一张拼图。超过部分返回 `omitted_slides`。
- **文件写入串行化**：所有改文件的工具都经过 pi 的 `withFileMutationQueue(path)`，同一文件的并发工具调用不会互相覆盖。

## 配置

| 环境变量 | 默认值 | 作用 |
|---|---|---|
| `PI_PPTKIT_PYTHON` | `python3` | 解释器路径（如虚拟环境里的 python） |
| `PI_PPTKIT_PY` | 包内 `python/pptkit.py` | 后端脚本路径 |
| `PI_PPTKIT_DEBUG` | 未设置 | 设置后显示 Python traceback |

渲染缓存与临时 profile 位于 `/tmp/pi-pptx/`，可随时删除。

## 已知边界

- 不支持 SmartArt、动画/切换、节（section）；选中窗格层级等细节也拿不到。
- 不能新建母版/版式；不要改模板版式（会影响所有引用页）。
- 新表格默认不走主题样式；要贴合模板，优先克隆带表格/图表的样本页（`inspect --mode summary` 的 `has_table` 可直接判断），再用 `ppt_set_cell` / `ppt_table_style`。
- 字体缺失时 LibreOffice 会替换，跨行位置可能与 PowerPoint 差 1–2 个字（约 5% 字宽）。`ppt_shot` 的溢出结果请留 5–10% 余量；要严格对齐请安装原字体。`ppt_doctor <deck>` 会输出主题的 major/minor 字体与字体替换链。
- 溢出/重叠都是启发式信号，不能替代看渲染图；文字超出页面被裁剪时不会报告。
- 同一 layout 新建页的占位符名可能与模板样例页不同（如 `Title 1` vs `标题 1`）。`ppt_add_slide` 会返回新页所有形状的 id/name/type/box，按 `#id` 操作即可。
- 没有内建撤销：推荐把“从模板重建工作副本 + 重跑 ops”当作撤销（`ppt_create(template=…, overwrite=true)`），或分阶段 `out=` 落盘。批量 `ppt_apply` 的原子提交已消除“改到一半”的状态。
- 中文文件名正常支持；soffice 的输出目录必须存在（插件已处理）。

## 开发与测试

```bash
npm install
./tests/smoke.sh          # py_compile + 后端测试 + 扩展工具测试 + tsc
npm run test:backend      # 仅 Python 后端
npm run test:extension    # 仅扩展（stub pi API，无需模型）
```

`tests/test_backend.py` 会在临时目录自建合成模板（含图片/表格/备注）验证克隆保真、样式保留、单元格修改、中文字体 XML、批量事务/回滚、表格样式、备注、拼图/裁剪、溢出与重叠检测、缓存；若本机存在 `/home/limour/pi-test/PPT模板.pptx` 或设置了 `PI_PPTX_TEST_TEMPLATE`，还会验证真实模板的图片保留。

`tests/extension.test.ts` 用 stub 的 `ExtensionAPI` 加载真实扩展，直接调用 `tool.execute()`，验证工具注册、创建→编辑→截图（含 image content 与 overflow 警告）、批量 apply（dry-run/失败回滚）、删除形状/备注/表格样式、拼图、克隆后只改副本、拒绝覆盖/清空，以及提示词条件注入。

## 目录

```
pi-pptx/
├── extensions/pi-pptx.ts    # pi 工具层（16 个 ppt_* 工具 + 工作流提示）
├── python/pptkit.py         # inspect/shot/clone/编辑/批量 后端
├── tests/                   # 后端 + 扩展测试
├── package.json             # pi 包清单（pi.extensions）
└── tsconfig.json
```

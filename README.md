# pi-pptx

让 pi agent 真正“画” PPT：读模板、克隆样本页、改文字/表格/图片，并把改完的每一页**渲染成图片看一遍**再继续。

实现依据调研笔记 `~/pi-test/pptkit/RESEARCH.md` 的三层结论：

| 层 | 方案 | 作用 |
|---|---|---|
| 解析/写入 | [python-pptx](https://python-pptx.readthedocs.io/) 1.0+ | 读版式/形状/表格/备注；克隆整页；改文字；加元素 |
| 渲染 | LibreOffice headless → PDF → PyMuPDF | 任意页出 PNG；文本 bbox 溢出检测；PDF/PNG 缓存 |
| 插件 | pi extension `defineTool` | 12 个 `ppt_*` 工具；截图以 `ImageContent` 直接回给模型 |

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
| `ppt_inspect` | 版式目录 + 形状树（名称、递归路径、英寸坐标、文本、表格、图片、备注）JSON |
| `ppt_shot` | 渲染 1–4 页为 PNG（连同溢出警告和缓存 PDF 路径）返回给模型 |
| `ppt_create` | 新建空白 16:9，或**逐字节复制模板**作为工作副本 |
| `ppt_clone` | 克隆已有页（图片/图表/组合/背景/备注全部保留，重映射 rels） |
| `ppt_add_slide` | 按版式新建空白页（仅适合骨架型模板） |
| `ppt_set_text` | 改形状文字，保留原字体/字号/颜色/位置；`\n` 分段继承样式 |
| `ppt_set_cell` | 改表格单元格（行列从 0 开始，与 `inspect` 的数组一致） |
| `ppt_add_text` | 按英寸坐标加文本框（可设字号/字体/颜色/对齐/行距） |
| `ppt_add_image` | 按英寸坐标插图，可只给宽或高保持比例 |
| `ppt_add_table` | 加表格 + 二维数据（默认样式，通常不如克隆带表格的样本页） |
| `ppt_delete` | 按页码/范围删页，至少保留 1 页 |
| `ppt_doctor` | 环境自检：python-pptx、PyMuPDF、LibreOffice 版本与字体替换结果 |

坐标单位是英寸，默认画布 13.33 × 7.5 in（16:9）。组合内子形状的 `box_in` 已换算到画布坐标；页码从 1 开始，表格行列从 0 开始。`inspect` 里每个顶层形状还带 `index`，当同一页有重名形状（模板里手放的文本框很常见）时用它精确定位。

## Agent 工作流

插件在**检测到 PPT 类请求**时（prompt 含 ppt/pptx/slides/幻灯片/演示文稿/模板等关键词）向系统提示注入一段 `pptx_workflow`，要求 agent 按下面闭环工作：

```
inspect  →  看版式和样本页
shot     →  用图片确认哪一页是要复用的“样式母本”
create   →  从模板复制出工作副本（绝不直接改模板）
clone    →  每个新页克隆一次样本页（保住母版/组合/图片/图表）
set_text / set_cell / add_*  →  换内容
shot     →  看渲染结果，处理 overflow / 换行 / 遮挡 / 字号，再改
```

工具 description 里也写明了各自的坑（例如 `ppt_add_table` 的默认样式、`ppt_add_text` 不继承模板样式、`ppt_clone` 才是主路径）。`promptSnippet` 让工具在默认系统提示的 Available tools 区有简短说明。

## 直接使用 Python 后端

```bash
cd /home/limour/pi-pptx

python3 python/pptkit.py doctor
python3 python/pptkit.py inspect 模板.pptx --slides 1-3
python3 python/pptkit.py create deck.pptx --template 模板.pptx
python3 python/pptkit.py clone deck.pptx 3 --at 2 --count 2
python3 python/pptkit.py set-text deck.pptx --slide 2 --shape "组合 7/矩形 9" --text $'第一行\n第二行'
python3 python/pptkit.py set-cell deck.pptx --slide 2 --shape "Table 7" --row 1 --col 1 --text 58
python3 python/pptkit.py add-text deck.pptx --slide 2 --text "补充说明" --x 0.8 --y 6.4 --w 8 --size 16
python3 python/pptkit.py shot deck.pptx --slides 1,3 --dpi 110
python3 python/pptkit.py delete deck.pptx --slides 4-6
```

所有命令输出单行 JSON；失败时 stderr 给出中文原因并以非 0 退出。默认原地修改，`--out` 可写副本。

## 实现要点

- **克隆页必须重映射 rels**：`ppt_clone` 深拷贝 `spTree` 后，遍历所有 `r:id`/`r:embed`，把源页关系重新绑到目标页，再复制显式背景和备注。不做这一步图片/图表会静默丢失（并且 PowerPoint 不报错）。
- **改文字只动 run**：`ppt_set_text` 保留第一个 run 的字体/字号/颜色，删除多余 run；多段落通过深拷贝首段 XML 实现，继承项目符号、行距等段落属性。
- **原子保存 + 完整性回读**：写临时文件 → 重新 `Presentation()` 读取校验 → `os.replace`，避免半写坏文件。
- **渲染缓存**：PDF 以 `(绝对路径, mtime_ns, size)` 的 sha1 作为 key 缓存；PNG 以 `key + dpi` 缓存。连续 `shot` 未改动的文件基本只花 PyMuPDF 出图时间（约 0.1 s/页）。
- **并发安全**：每次 soffice 调用使用独立 `-env:UserInstallation` profile，并在独立临时目录转换后原子移动 PDF，两个并发转换不会互相打架。
- **溢出检测（启发式）**：渲染后取 PDF 文本块 bbox，与 `inspect` 的形状 bbox（EMU→pt，按 PDF 页面比例缩放）比较。匹配时接受与形状相交或从形状上沿开始的文本块，纵向允许约 0.35 行高/10pt 的行距误差；仅检查顶层、无旋转的文本形状（组合子形状坐标是相对的）。结果放在 `shots[].overflow`，并在给模型的文本里变成 `WARNING text overflow ...`。超出页面被裁掉的文字无法从 PDF 取到，仍以看图为准。
- **按页限量**：`ppt_shot` 单次最多 4 页（约 1.6k tokens/页 @110dpi），超过部分返回 `omitted_slides`，让模型分批查看。
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
- 新表格/图表使用 python-pptx 默认样式，通常与模板不搭——优先克隆带表格/图表的样本页，再用 `ppt_set_cell` 改内容。
- 字体缺失时 LibreOffice 会替换，跨行位置可能与 PowerPoint 差 1–2 个字。做版式自检够用；要严格对齐请安装原字体。
- 溢出检测是启发式信号，不能替代看渲染图；文字超出页面被裁剪时不会报告。
- 中文文件名正常支持；soffice 的输出目录必须存在（插件已处理）。

## 开发与测试

```bash
npm install
./tests/smoke.sh          # py_compile + 后端测试 + 扩展工具测试 + tsc
npm run test:backend      # 仅 Python 后端
npm run test:extension    # 仅扩展（stub pi API，无需模型）
```

`tests/test_backend.py` 会在临时目录自建合成模板（含图片/表格/备注）验证克隆保真、样式保留、单元格修改、溢出检测和缓存；若本机存在 `/home/limour/pi-test/PPT模板.pptx` 或设置了 `PI_PPTX_TEST_TEMPLATE`，还会验证真实模板的图片保留。

`tests/extension.test.ts` 用 stub 的 `ExtensionAPI` 加载真实扩展，直接调用 `tool.execute()`，验证工具注册、创建→编辑→截图（含 image content 与 overflow 警告）、克隆后只改副本、拒绝覆盖/清空，以及提示词条件注入。

## 目录

```
pi-pptx/
├── extensions/pi-pptx.ts    # pi 工具层
├── python/pptkit.py         # inspect/shot/clone/编辑 后端
├── tests/                   # 后端 + 扩展测试
├── package.json             # pi 包清单（pi.extensions）
└── tsconfig.json
```

---
name: pptx-visual-review
description: 用 ppt_inspect / ppt_shot 验收已生成的 PPTX：预检 text_fit、看渲染图、判溢出/重叠、逐形状密度、版本对比、拼图与裁剪、expect=clean 门禁。写完一页或一整本需要"看结果"时读它。
---

# 看图验收：先几何，再像素

`ppt_inspect` 回答"能不能放下"，`ppt_shot` 回答"长成什么样"。顺序是先 inspect 预检，再 shot 确认；
shot 报红时用 `density` 定位到具体那个框，回脚本改，再跑。

## 预检（不渲染）

```text
ppt_inspect({path, mode: "summary"})               # 每页标题/形状数/表格/图片/备注
ppt_inspect({path, mode: "full", slides: "1-3"})   # 形状树：name/path/box_in/text/table/notes
ppt_inspect({path, mode: "text"})                  # 扁平 {shape, id, text}，通读校对错别字
ppt_inspect({path: <模板>, ...})                    # 版式目录 + inherited 形状
```

每个文本形状带 `text_fit`：`est_lines` / `capacity_lines` / `needed_in` / `box_h_in` / `slack_in` / `fits`。
`fits: false` 或 `slack_in < 0.1` 就是"这框会换行/会压"，先改脚本，别等渲染。

## 渲染（像素）

```text
ppt_shot({path, slides: "1-6"})                    # 最多 6 张单页图，默认 110 dpi
ppt_shot({path, montage: 3})                       # 一张 3 列 contact sheet，最多 60 页：看整本节奏
ppt_shot({path, slides: "6", region: [0.5, 4, 6, 2.5], dpi: 150})   # 局部放大：量字距/对齐
ppt_shot({path, slides: "3", density: true})       # 逐形状 rendered_lines/needed_in/box_h_in/slack_in/tight
ppt_shot({path, slides: "2-8", compareTo: "上一版.pptx"})           # 每页 changed_px_ratio：改 A 有没有挤坏 B
ppt_shot({path, montage: 4, overflow: "strict", expect: "clean"})   # 构建门禁：任何一页报红即失败
```

- `dpi`：80 看节奏 / 110 默认检查 / 150 看细节。拼图每格约 640px 宽，**12pt 脚注在拼图里读不出来**——校对文字用 100+ dpi 单页。
- `overflow`：`loose`（默认，居中/两端对齐的左右不查）/ `strict`（每侧都报）/ `off`。
- `ignore`：按名称、`#<id>` 或 `/regex/` 跳过已知噪音形状（重建后依然稳定）。
- 超过 6 页普通模式会进 `omitted_slides`；要一次看全本就 `montage`。

## 报红怎么判

`shots[].overflow[]` 每条都带 `rendered_lines` / `needed_in` / `box_h_in` / `slack_in`，按它三选一：

1. `needed_in` 只比 `box_h_in` 大一点点 → 把框加高 0.1–0.2。
2. 行数明显超过 `capacity_lines` → 减字号或删字（先删字，字号是设计层级）。
3. `rendered_lines` 是 1 但横向溢出 → 框太窄或 `wrap` 关掉了，加宽或删字。

`shots[].overlaps[]` 是顶层有内容形状的重叠（文本×文本、文本×图片/表格）；容器包内容不算。
`suppressed[]` 列出被规则跳过、**并未修复**的侧与配对——它是解释，不是豁免清单。

别为了"变绿"去 `ignore` 真问题：模板固有的顶锚点 leading（< 5pt）本来就不报，剩下的红基本都是脚本要改。

## 验收节奏

1. 迭代中：`ppt_run_deck({script, out, template, shot: "flagged"})` → 只渲脚本 `@@WARN@@` 的页，一轮一改。
2. 页级干净后：`ppt_run_deck(... shot: "changed")` 或 `ppt_shot(montage=3)` 看整本。
3. 担心回归：`compareTo=<上一版 pptx>`（先让 `ppt_run_deck` 带 `keepVersions: 1` 把上一版留到 `/tmp/pi-pptx/versions/`）。
4. 定稿：`ppt_shot(montage=4, dpi=110)` 通看 + `ppt_inspect(mode="text")` 通读 + `notes_check`。
5. 自动化门禁：`ppt_run_deck(... expect="clean")` 或 `ppt_shot(... expect="clean")`，有红就是 `ok: false`。

## 字段速查

| 字段 | 含义 |
|---|---|
| `flagged_slides` | 本次有溢出/重叠的页 |
| `clean` / `expectation_failed` | `expect="clean"` 时的结论 |
| `density` | 逐形状的行数/所需高度/框高/余量/`tight` |
| `compared` / `changed_vs_previous` | 与 `compareTo` 的逐页像素差异 |
| `cache_hit` | 相同文件重复 shot 时命中 PDF/PNG 缓存（不是没渲染） |
| `font_substitutions` / `render_note` | 渲染字体与 PowerPoint 的差异（约 5% 字宽） |
| `omitted_slides` / `montage` | 本次没渲的页 / 拼图文件 |

`ppt_shot` 也是看"风格是否与模板一致"的工具：先 shot 模板的样例页，再 shot 成品对应页，两张图并排看
版式语言（字号层级、色条、卡片样式、页边距）对不对得上。

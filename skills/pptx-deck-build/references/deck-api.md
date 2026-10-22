# Deck API 速查（build_deck.py 里能直接用的）

两种写法等价：语义方法 `D.add_text(...)`，或通用 `D.op("add_text", ...)`。
每个 op 只读取自己认识的键（大小写/别名以实现为准）；`#<id>` / 名称 / `"<index>"` 都可以当 `shape=`。

## 生命周期

```python
D = Deck.create(OUT, template=TPL, overwrite=True)  # 字节复制模板；overwrite=True 才能重跑
D = Deck.create(OUT, overwrite=True)                 # 无模板：空白 16:9 画布
D = Deck.open("deck.pptx")                           # 在既有 deck 上继续
D.save()                                             # 原子写：临时文件 + 回读 + replace
D.save("copy.pptx")                                  # 另存
D.report()                                           # print("@@DECK@@ {...}")：页数/形状/表格/图片/备注
D.inspect(slides="2", mode="full")                   # 同 ppt_inspect，脚本里自检用
D.notes_check()                                      # 备注是否整体错位一页
D.changed                                            # 本次改过的页（save 后清空）
D.shots(only_changed=True, dpi=110)                  # 直接从库里渲染，一般不用（用 ppt_shot 更好）
```

## 页

```python
D.clone(source=1, at=D.slide_count + 1, count=2)     # 深拷贝整页：图片/表格 rels、背景、备注
D.add_slide(layout="3_版本内页")                      # 从版式新建（空占位符记得 delete_shape）
D.delete_slides("4-6")                               # 只收页码串（1 起）："3,5" / "4-6"
D.slide_count                                        # 当前页数
```

## 文字

```python
D.add_text(slide=2, x=0.68, y=1.84, w=6.0, h=0.9, text="正文", size=13, color="#1A1A1A")

D.add_text(
    slide=2, x=0.68, y=1.84, w=6.0,
    paras=[
        {"runs": [{"t": "小标题", "bold": True, "color": "#1F4E79"},
                  {"t": "  正体补充"}]},
        {"runs": [{"t": "第二段"}], "space_before": 6, "bullet": "•", "indent": 0.32, "hanging": 0.32},
    ],
    size=12, line_spacing=1.25,
)

D.set_text(slide=2, shape="#12", text="替换整段\n第二行")      # 保留首个 run 字体；\n = 段落
D.set_text(slide=2, shape="标题 1", paras=[...])              # 保留原段落属性 + 首个 run 样式
D.delete_shape(slide=2, shape="Content Placeholder 2")
```

常用键：`size`、`bold`、`italic`、`color`（`#RRGGBB` 或 `RRGGBB`）、`font`、`align`（`left/center/right/justify`）、
`valign`（`top/middle/bottom`）、`line_spacing`、`space_before`/`space_after`（pt）、`letter_spacing`、
`bullet`（字符、`True` 用 `•`、`"none"` 去掉）、`indent`/`hanging`（in）、`fit`（`grow/shrink/none`）、`wrap=False`、`name`。

- 不传 `h` 时按行数估高并写 `spAutoFit`（grow）；传了 `h` 就不会自动长高。
- 中文/日文/韩文写 run 时后端同时写 `a:latin`/`a:ea`/`a:cs`，`font=` 对中文有效。
- `paras[].runs[]` 支持每段混排；`{"runs":[{"t":...}]}` 是最小单元。

## 形状 / 图片 / 表格

```python
D.card(slide=2, x=0.68, y=1.84, w=3.7, h=2.1, text="卡片标题", size=16)   # 圆角卡片，无阴影细边
D.bar(slide=2, x=0.68, y=1.72, w=0.9, h=0.05, fill="#ED7D31")             # 强调色条
D.rule(slide=2, x=0.68, y=6.6, w=11.97)                                   # 分隔线

D.add_shape(slide=2, kind="chevron", x=2, y=3, w=1.4, h=0.9,
            fill="#1F4E79", line="none", radius=0.10, shadow=False, text="步骤")
# kind: rect / roundrect / oval / chevron / arrow / leftarrow / uparrow / downarrow /
#       triangle / parallelogram / diamond / pentagon / hexagon / bar / line
# 形状里的文字不会自动长高：文字多了自己给够 h，或另用 add_text 覆盖

D.add_image(slide=2, path="assets/fig1.png", x=7.6, y=1.84, w=5.0)        # 只给 w 等比缩放

D.add_table(
    slide=3, rows=4, cols=3, x=0.68, y=1.84, w=7.2, h=1.44,               # h 必须自己算！
    data=[["阶段", "做法", "验收"], ["排版", "常量+语义包装", "text_fit 全绿"], ...],
    col_widths=[1.4, 3.4, 2.4],               # 和 = w
    row_height=0.32, header_height=0.42,      # 表头行高可以 ≠ 正文行高
    size=11, header=True, header_fill="#1F4E79", header_color="#FFFFFF",
    zebra="#EEF3F9", border="#C9D6E4", first_col_bold=True,
    aligns=["left", "left", "center"],
)
D.table_style(slide=3, shape="#9", header_fill="1F4E79", border="none")
D.set_cell(slide=3, shape="#9", row=1, col=2, text="57",
           cell_runs=[{"t": "57", "bold": True}, {"t": " %"}])
```

`add_table` 默认**关掉主题表样式**（删 `a:tableStyleId` + firstRow/bandRow），否则主题配色会盖住手工填色。
返回值里带 `hint` 时说明模板没有任何表格样例。

## 备注

```python
D.set_notes(slide=4, text="讲者备注：这一页只讲一个数字。")
```

删过模板页之后跑一次 `D.notes_check()`（或 `ppt_inspect` 里看 `notes_len`/标题），备注会整体错位一页。

## 批量事务

```python
D.apply([
    {"op": "set_text", "slide": 2, "shape": "#12", "text": "标题"},
    {"op": "add_shape", "slide": 2, "kind": "chevron", "x": 2, "y": 3, "w": 1.4, "h": 0.9},
], stop_on_error=True)     # 失败时整体回滚，返回 {"ok", "ops", "errors"}
```

脚本里通常不需要 `apply`：直接顺序调用方法即可，`save()` 之前全部在内存里。
`apply` 适合"把一批 JSON 操作喂进来"的旧入口或数据驱动场景。

## 错误与寻址

- 找不到形状时，报错会列出候选（name/id/type/box/文本前 20 字）；重名时用 `#<id>`，不要赌名称。
- 顶层 `index` 是 0 起，越界报错会提示该约定；页码是 1 起。
- 一切失败都在 `D.save()` 之前，落盘的上一版仍可用于 `ppt_shot(compareTo=)`。

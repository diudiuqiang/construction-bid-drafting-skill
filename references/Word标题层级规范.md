# Word 标题层级规范

## 0. 为什么单独立这一篇

技术标交付物是 .docx，评审用 Word「大纲」视图 / 导航窗格翻结构。
标题层级不对会出现两种症状：

1. 导航窗格空或全平铺——评审找不到章节；
2. 章节层级错乱——「3.1.1」挂到了「3.2」下面。

根因几乎总是同一件事：**Markdown 的编号深度不能当层级用**。

## 1. 唯一可靠的层级信号：`#` 级数

本 skill 约定：**逻辑层级只看 Markdown `#` 级数，编号深度一律不作为依据。**

| Markdown | Word 样式 ID | Word 内置名 | outlineLvl | 大纲显示 |
|---|---|---|---|---|
| `#` | Title | Title | —（封面居中） | 封面 |
| `##` | Heading1 | heading 1 | 0 | 1 |
| `###` | Heading2 | heading 2 | 1 | 1.1 |
| `####` | Heading3 | heading 3 | 2 | 1.1.1 |
| `#####` | Heading4 | heading 4 | 3 | 1.1.1.1 |

`heading_style_map(lv)` 的实现：

```python
def heading_style_map(lv):
    if lv == 1:
        return 'Title'
    return 'Heading%d' % min(lv - 1, 4)
```

## 2. 编号深度为什么不可靠

同一份标书里，不同章节的编号深度可能不一致：

- 第三章用 `3.1 / 3.1.1 / 3.1.1.1`（三层编号）；
- 第四到第八章用 `4 / 4.1 / 4.1.1 / 4.1.1.1`（首层是单数字）。

若按「编号里点号的个数」推断层级，`4.1`（1 个点）和 `3.1.1`（2 个点）会被判成不同层级，
而实际上二者属于哪一级只由 `#` 数决定。**只有 `#` 级数是全局一致的。**

推论：分册写作时，`##` 级标题在哪个章节都一样是 H2；不要因为某册正文从 `##` 开始
就以为它「缺一级」。

## 3. Word 内置样式名必须小写

OOXML 里 `<w:name w:val="heading 1"/>` 必须是小写 `heading 1`。
写成 `Heading 1`、`标题 1`、`Heading1` 都会导致 Word 不认，大纲失效。

样式 ID（`w:styleId`）用驼峰 `Heading1`，样式名（`w:name`）用小写 `heading 1`——两者不同，别混。

## 4. outlineLvl 必须与样式层级匹配

| 样式 ID | 内置名 | outlineLvl |
|---|---|---|
| Heading1 | heading 1 | 0 |
| Heading2 | heading 2 | 1 |
| Heading3 | heading 3 | 2 |
| Heading4 | heading 4 | 3 |

Title 不设 outlineLvl（它是封面，不是章节）。

## 5. styles.xml 关键片段

```xml
<w:style w:type="paragraph" w:styleId="Heading1">
  <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
  <w:pPr><w:keepNext/><w:spacing w:before="360" w:after="180"/><w:outlineLvl w:val="0"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/>
         <w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
</w:style>
```

字号阶梯（`w:sz` 单位为半磅）：

| 样式 | w:sz | 折合 pt |
|---|---|---|
| Title | 36 | 18 |
| Heading1 | 32 | 16 |
| Heading2 | 28 | 14 |
| Heading3 | 26 | 13 |
| Heading4 | 24 | 12 |
| Normal | 24 | 12 |
| TableText | 21 | 10.5 |

## 6. `#` 级标题的双重用途与 `--title-keywords`

`#` 约定只用于封面（项目名 / 文档名），映射 Title。

若正文里也用了 `#`（比如分册文件首行写 `# 施工方案总述`），默认会被全映射成 Title，
导致大纲混入封面。此时用 `--title-keywords` 控制：

```bash
python build_docx.py --in x.md --out x.docx --title-keywords "横沙乡,施工方案"
```

命中关键词的 `#` 标题才映射 Title，其余映射 Heading1。
不传该参数时全部 `#` 映射 Title（封面约定）。

> **关键词必须写真实项目名**，不要照抄「项目名」这类占位符——占位符永远命中不了，
> 结果一个 Title 都没有，封面直接丢失（`--verify` 只提示、不判失败）。

## 7. 校验方法

```bash
python build_docx.py --verify ./out/施工方案.docx
```

校验四项：

| 项 | 判据 |
|---|---|
| Application | 必须为 `Microsoft Office Word` 或 `Microsoft Word` |
| creator | 不得含 workbuddy/agent/ai/claude/gpt/deepseek/assistant/智能体 |
| outlineLvl | 每个内置 heading 样式必须有正确的 `w:outlineLvl` |
| 标题段落 | 至少有一个 heading/Title 段落，否则大纲视图为空 |

对应函数：`C.docx_outline_check(path)`、`verify_docx(path)`。

## 8. 常见错误清单

| 错误 | 后果 | 修正 |
|---|---|---|
| 用编号深度推断层级 | 层级错乱 | 一律按 `#` 级数 |
| 样式名写成 `Heading 1` | 大纲失效 | 改 `heading 1` |
| 缺 `w:outlineLvl` | 大纲视图空 | 补 outlineLvl |
| outlineLvl 与层级不符 | 层级错乱 | heading N → outlineLvl N-1 |
| `#` 级标题混入正文 | 大纲混入封面 | 改用 `##`，或传 `--title-keywords` |
| Application 写工具名 | 暴露生成痕迹 | 改 `Microsoft Office Word` |
| creator 写 AI 名 | 暴露 AI 痕迹 | 改真人/组名 |
| 每个段落都套 Heading 样式 | 大纲爆炸 | 正文用 Normal |

## 9. 生成时的检查顺序

1. Markdown 里 `#` 级数是否全局一致（`check_heading_structure`）；
2. 生成 docx；
3. `--verify` 校验大纲 + 元数据；
4. Word 打开看导航窗格，确认 1 / 1.1 / 1.1.1 / 1.1.1.1 四级可见。

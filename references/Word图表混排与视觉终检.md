# Word 图表混排与视觉终检

## 0. 为什么单独立这一篇

前四阶段的校验全是**文本级**：禁用词、硬指标、编号、标题层级、AI 痕迹。
它们看不见下面这些问题——而这些问题在评审翻开 Word 的第一眼就会暴露：

- 图片在 docx 里「存在」但打开是空白；
- 题注跑到图片上面，或左对齐缩在角落；
- 表格列宽等分，中文长表头被挤成一列竖排字；
- 从 Excel 拆出来的表字体、字号、边框与正文两张皮。

所以合卷交付要加**阶段 5：视觉终检**。本文讲怎么生成、怎么验证。

---

## 1. 内嵌图片的约定

`build_docx.py` 支持 Markdown 行：

```markdown
[[IMG:pieces/gantt_W1-W11.png]]
```

- 路径相对 `--img-base`（默认取输入 md 所在目录）；也接受绝对路径。
- 图片按版心宽度铺满（`(page_width - 2*margin) * 635` EMU），高度按原图比例计算。
- 尺寸由纯标准库解析 PNG / JPEG / GIF / BMP 的头，**不需要 Pillow**。
- 同名文件自动加序号，避免 `word/media/` 覆盖。

```bash
python scripts/build_docx.py --in ./out/终稿.md --out ./out/终稿.docx \
    --title "技术标终稿" --creator "投标技术组" \
    --title-keywords "项目名关键词,施工方案" \
    --img-base ./out --footer-text "技术标  第 {page} 页"
python scripts/build_docx.py --verify ./out/终稿.docx
```

### 图在上、题注在下

Markdown 里必须写：

```markdown
[[IMG:pieces/图A.png]]

图 10-8 某某横道图（第 1~77 日历天，W1~W11）
```

不要写反。`--img-base` 生成时会检查**每张图片之后是否紧跟题注**，缺了会打印警告
（不阻断）。注意别反过来判「题注后面是不是图片」——一张图的题注后面本来就可能紧跟下一张图，
那个判据会把所有正常排版误报一遍。

### 题注判据的坑

题注要居中，且**不能误伤以「表 X-Y」开头的正文长句**。判据用：

```python
CAP_PAT = re.compile(r'^(图|表)\s*\d+(?:-\d+)?\s+\S')
is_caption = (CAP_PAT.match(s)
              and (s.endswith('表') or s.endswith('）') or s.endswith(')'))
              and len(s) <= 60)
```

**不要**用「不含句读标点」当判据——真实题注常含全角逗号，例如
`图 10-8 施工总进度横道图（第 1~77 日历天，W1~W11）`，一排除就全漏判。
以「表」或「）」收尾 + 短行，能把正文长句（以「。」收尾、长度上百）干净排除。

---

## 2. 最阴的坑：wp 命名空间 URI

`w:drawing` 依赖 `wp:` 前缀，**只有一种正确写法**：

```
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing
```

常见错写（多了一层 `wordprocessingml/`）：

```
http://schemas.openxmlformats.org/drawingml/2006/wordprocessingml/wordprocessingDrawing
```

**错写的后果是静默的**：docx 里 `<w:drawing>`、`word/media/*.png`、`document.xml.rels` 的 image 关系、
`[Content_Types].xml` 全都在，XML 也能解析，`--verify` 若不专门查也会过；但 LibreOffice 转 PDF 后
`get_image_info()` 返回 **0 张**。Word 里多半也是空白框。文本级校验完全发现不了。

`build_docx.py --verify` 现在会检查这一项，并报：

```
wp 命名空间 URI='...' —— 应为 '...'；写错时 LibreOffice 会静默丢弃图片
```

排查这类问题的可靠方法：**二分法**。拿 LibreOffice 自己导出的、确认能渲染的 docx 作参照，
逐块替换成自己的部件，定位到是 `document.xml` 还是包结构的问题，再比根元素声明。
不要靠猜「是不是 rId 编号 / 图片太大 / 缺 bwMode」——实测这些都无关。

### 图片部件四件套，缺一不可

| 部件 | 内容 |
|---|---|
| `word/document.xml` | `<w:drawing>` + 根上声明 `xmlns:wp` / `xmlns:a` / `xmlns:pic` |
| `word/_rels/document.xml.rels` | 每条 `r:embed` 对应一条 `.../relationships/image` |
| `word/media/<name>` | 实际图片字节 |
| `[Content_Types].xml` | `<Default Extension="png" ContentType="image/png"/>` 等 |

`--verify` 会比对 `drawing 数 / r:embed 数 / image 关系数 / media 文件数` 四者一致。

---

## 3. 表格列宽：按 CJK 权重自适应

等分列宽（`9000 / ncol`）对中文技术标是灾难：表头「施工区段/流水段」六个汉字比 `No.` 宽得多，
固定布局下会被挤成竖排，或者内容列被压到一两字一行。

```python
def _cjk_weight(s):
    return sum(2 if ord(c) > 0x2E80 else 1 for c in s)

# 每列取该列所有单元格的最大权重，clamp 到 4~40，再按总和分配
weights = [min(max([_cjk_weight(str(r[ci])) for r in rows if ci < len(r)] + [4]), 40)
           for ci in range(ncol)]
widths = [max(1, int(9000 * w / sum(weights))) for w in weights]
```

`<w:tcW w:w="<width>" w:type="dxa"/>` 与 `<w:tblLayout w:type="fixed"/>` 要配套，否则 Word 会忽略列宽。

### 从 Excel 拆工作表进 Word

用户常要求「把进度计划 Excel 的每个工作表放进正文相关位置」。两条路：

| 方式 | 适用 | 要点 |
|---|---|---|
| **转成 Markdown 表格** | 字段少、行数可控的计划表、里程碑表、劳动力表、机械进退场表 | 用 openpyxl 读值 → 生成 `\|` 表格 → 交给 `build_docx.py`，字体/字号/边框自动与正文一致 |
| **转成 PNG 图片** | 横道图、网络图这类单元格色块图 | 用绘图库画好后走 `[[IMG:]]`；图片宽高比要与版心匹配 |

无论哪种，**表题注居中、正文说明段落在表前**，并在表下写「注：本表与《XXX.xlsx》「YYY」工作表口径一致」，
让评委知道数据来源。数值口径必须与正文、Excel、待修正清单三处一致（见 `二次修正流程.md` 第 5 节）。

> 拆表时注意 Excel 里「注：…」这类合并单元格行，`r[0]` 往往不是字符串 `'注：'`；
> 判断用 `str(r[0]).startswith('注')`，别用等值比较，否则会把注释行当成数据行插进表体。

---

## 4. 阶段 5：视觉终检

### 4.1 转 PDF

```bash
"/c/Program Files/LibreOffice/program/soffice.exe" --headless \
    --convert-to pdf --outdir ./out ./out/终稿.docx
```

**环境坑**：某些沙箱 / 精简 shell 会剥离 `ProgramFiles` / `ProgramFiles(x86)` / `LOCALAPPDATA`
环境变量，依赖这些变量拼 LibreOffice 路径的脚本（如 `recalc.py`）会报
`LibreOffice executable not found`。运行时补上即可，不需要安装任何东西：

```bash
ProgramFiles="C:\Program Files" python <script.py> <args>
```

### 4.2 用 pymupdf 程序化核对

`pymupdf`（`import pymupdf`，旧名 `fitz`）是**只读**校验，不修改交付件：

```python
import pymupdf
doc = pymupdf.open('终稿.pdf')
print('pages:', doc.page_count)

# ① 图片数量与位置：必须与 md 里 [[IMG:]] 数量一致
imgs = [(i + 1, im['width'], im['height'], im['bbox'])
        for i in range(doc.page_count) for im in doc[i].get_image_info()]
print('images:', len(imgs))

# ② 题注居中：块中心应约等于页宽 / 2
for i in range(doc.page_count):
    for b in doc[i].get_text('blocks'):
        t = b[4].strip()
        if t.startswith('图 ') or t.startswith('表 '):
            print('p%d center=%.1f page=%.1f' %
                  (i + 1, (b[0] + b[2]) / 2, doc[i].rect.width / 2))

# ③ 字体 / 字号 / 越界
from collections import Counter
fonts, sizes, over = Counter(), Counter(), []
for pg in doc:
    for blk in pg.get_text('dict')['blocks']:
        for line in blk.get('lines', []):
            for sp in line['spans']:
                fonts[sp['font']] += len(sp['text'])
                sizes[round(sp['size'], 1)] += len(sp['text'])
                x0, y0, x1, y1 = sp['bbox']
                if x1 > pg.rect.width - 20 or x0 < 20 or y1 > pg.rect.height - 15:
                    over.append((pg.number + 1, sp['text'][:20]))
print('fonts:', fonts.most_common(5), 'sizes:', sorted(sizes))
print('overflow:', len(over))
```

判据：

- 图片数 = md 中 `[[IMG:]]` 数；尺寸与源图一致（LibreOffice 会重新编码 PNG，md5 不同是正常的，看像素尺寸）；
- 题注块中心 ≈ 页宽 / 2（本项目实测 297.8 vs 297.65）；
- 字体只有宋体 / 黑体 / Times New Roman 及少量替代字体；出现大段 `NotoSansSC-Thin` 之类要查是否缺字；
- 越界 0（页脚落在页面底部属正常，可单独排除）。

> 如果手头没有视觉模型可读图，程序化审计就是可靠替代：非白像素占比可以判断「整页横道图是否真的画出来了」。

---

## 5. 检查清单

- [ ] md 中图片顺序为「`[[IMG:]]` → 题注」，生成时无顺序警告
- [ ] `--verify` 通过：`wp` 命名空间正常、图片部件四件套数量一致、元数据无工具名
- [ ] 题注居中（PDF 块中心 ≈ 页心）
- [ ] 表格列宽不等分、中文表头不竖排
- [ ] Excel 拆出的表字体字号与正文一致，表下注明数据同源
- [ ] PDF 图片数、页数、字体、越界四项核对通过
- [ ] 页脚页码正常、无跨页空白

---

## 6. 一句话判据

> 文本级体检说「全绿」不等于 Word 能看。只要稿件里有图，就必须转一次 PDF，
> 数一遍图片、量一遍题注位置——`wp` 命名空间的坑是静默的，只能靠这一步兜住。

# -*- coding: utf-8 -*-
"""Markdown -> Word(.docx) 生成器（纯标准库，无第三方依赖）。

设计目标：
1. 生成的 docx 在 Word「大纲」视图 / 导航窗格中呈现 1 / 1.1 / 1.1.1 / 1.1.1.1 四级结构；
2. 标题层级按 Markdown `#` 级数判定——编号深度不可靠（见 references/Word标题层级规范.md）；
3. 表格列宽按内容 CJK 权重自适应；支持 `[[IMG:相对路径]]` 内嵌图片与居中图表题注
   （见 references/Word图表混排与视觉终检.md）；
4. 不留下生成工具痕迹：docProps/app.xml 的 Application 写「Microsoft Office Word」，
   creator 由 --creator 指定（默认「投标技术组」），不出现 AI / Agent / 脚本名。

用法：
  # 生成
  python build_docx.py --in ./out/施工方案.md --out ./out/施工方案.docx \
      --title "施工方案（技术标第十章）初稿" --creator "投标技术组" \
      --title-keywords "横沙乡,施工方案" --img-base ./out

  # 校验已有 docx 的大纲可用性、元数据、图片部件
  python build_docx.py --verify ./out/施工方案.docx
"""
import argparse
import os
import re
import struct
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cbd_common as C


# ---------------------------------------------------------------- 内联解析

def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


def clean_inline(s):
    """去掉 Markdown 行内标记：反引号、〔〕转半角括号。保留 ** 供 runs 解析。"""
    s = s.replace('`', '')
    s = s.replace('〔', '(').replace('〕', ')')
    return s.strip()


def runs_xml(text, style):
    """把 `**加粗**` 解析为多个 run；其余按普通文本。"""
    text = clean_inline(text)
    sizes = {'Normal': 24, 'Title': 36, 'Heading1': 32, 'Heading2': 29,
             'Heading3': 27, 'Heading4': 25, 'TableText': 21}
    sz = sizes.get(style, 24)
    ea = '黑体' if style.startswith('Heading') or style == 'Title' else '宋体'
    base = ('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
            'w:eastAsia="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (ea, sz, sz))
    out = []
    for part in re.split(r'(\*\*.+?\*\*)', text):
        if not part:
            continue
        bold = part.startswith('**') and part.endswith('**') and len(part) >= 4
        t = esc(part[2:-2] if bold else part)
        if not t:
            continue
        rpr = '<w:rPr>%s%s</w:rPr>' % (base, '<w:b/>' if bold else '')
        out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, t))
    return ''.join(out) if out else '<w:r><w:rPr>%s</w:rPr><w:t></w:t></w:r>' % base


def p_xml(text, style='Normal', indent=False, center=False):
    ppr = '<w:pPr><w:pStyle w:val="%s"/>' % style
    if indent:
        ppr += '<w:ind w:firstLineChars="200" w:firstLine="480"/>'
    ppr += '<w:spacing w:before="60" w:after="60" w:line="360" w:lineRule="auto"/>'
    if center:
        ppr += '<w:jc w:val="center"/>'
    if style.startswith('Heading') or style == 'Title':
        ppr += '<w:keepNext/>'
    ppr += '</w:pPr>'
    return '<w:p>%s%s</w:p>' % (ppr, runs_xml(text, style))


def _cjk_weight(s):
    """列宽权重：CJK 记 2、ASCII 记 1。中日韩表头比英文表头占宽，等分列宽会把中文挤成竖排。"""
    return sum(2 if ord(c) > 0x2E80 else 1 for c in s)


def table_xml(rows, total=9000):
    """按每列内容最大权重分配列宽，单列上下限 4~40，避免某个超长单元格独占整表。"""
    if not rows:
        return ''
    ncol = max(len(r) for r in rows)
    weights = []
    for ci in range(ncol):
        w = 4
        for r in rows:
            if ci < len(r):
                w = max(w, _cjk_weight(str(r[ci])))
        weights.append(min(w, 40))
    sw = sum(weights)
    widths = [max(1, int(total * w / sw)) for w in weights]
    grid = ''.join('<w:gridCol w:w="%d"/>' % w for w in widths)
    out = ['<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>'
           '<w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>'
           '<w:tblGrid>%s</w:tblGrid>' % grid]
    for ri, row in enumerate(rows):
        cells = []
        for ci in range(ncol):
            c = row[ci] if ci < len(row) else ''
            cells.append('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>'
                         % (widths[ci], p_xml(c, 'TableText')))
        trpr = '<w:trPr><w:tblHeader/></w:trPr>' if ri == 0 else ''
        out.append('<w:tr>%s%s</w:tr>' % (trpr, ''.join(cells)))
    out.append('</w:tbl>')
    return ''.join(out)


# ---------------------------------------------------------------- 内嵌图片

IMG_PAT = re.compile(r'^\[\[IMG:(.+?)\]\]$')
# 图表题注：以「表」或「）」结尾且较短。不能写「不含标点」判据——
# 真实题注常含全角逗号，如「图 10-8 施工总进度横道图（第 1~77 日历天，W1~W11）」。
CAP_PAT = re.compile(r'^(图|表)\s*\d+(?:-\d+)?\s+\S')
IMAGE_TYPES = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
               'gif': 'image/gif', 'bmp': 'image/bmp'}
# 关键：wp 命名空间必须是 wordprocessingDrawing（不含 wordprocessingml/）。
# 写错时 LibreOffice 会「静默」丢弃图片：docx 内 drawing / media / rels 全在，转 PDF 后 0 图。
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'


def is_caption(line):
    s = line.strip()
    return (bool(CAP_PAT.match(s))
            and (s.endswith('表') or s.endswith('）') or s.endswith(')'))
            and len(s) <= 60)


def image_size(path):
    """纯标准库读取位图尺寸，返回 (w, h)；识别不了返回 (4, 3)。"""
    try:
        with open(path, 'rb') as f:
            head = f.read(32)
            if head[:8] == b'\x89PNG\r\n\x1a\n' and head[12:16] == b'IHDR':
                return struct.unpack('>II', head[16:24])
            if head[:6] in (b'GIF87a', b'GIF89a'):
                return struct.unpack('<HH', head[6:10])
            if head[:2] == b'BM':
                w, h = struct.unpack('<ii', head[18:26])
                return (abs(w), abs(h))
            if head[:2] == b'\xff\xd8':
                f.seek(2)
                while True:
                    b = f.read(1)
                    if not b:
                        break
                    if b != b'\xff':
                        continue
                    marker = f.read(1)
                    while marker == b'\xff':
                        marker = f.read(1)
                    if not marker:
                        break
                    m = marker[0]
                    if m in (0xD8, 0xD9, 0x01) or 0xD0 <= m <= 0xD7:
                        continue
                    ln = f.read(2)
                    if len(ln) < 2:
                        break
                    seg = f.read(struct.unpack('>H', ln)[0] - 2)
                    if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                             0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                        h, w = struct.unpack('>HH', seg[1:5])
                        return (w, h)
    except (OSError, struct.error, IndexError):
        pass
    return (4, 3)


def img_xml(name, rid, docpr_id, cx, cy):
    """居中内嵌图片。a / pic / wp 命名空间在 document 根上声明（见 build_docx）。"""
    safe = esc(name)
    return (
        '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="120"/></w:pPr>'
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="%d" name="%s"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
        % (cx, cy, docpr_id, safe, docpr_id, safe, rid, cx, cy))


def figure_order_warnings(md):
    """规范顺序是「图片 → 题注」。返回「图片之后没有紧跟题注」的图片行号。

    不要反过来判断题注的下一行是不是图片——一张图的题注后面本来就可能紧跟下一张图，
    那样会把所有正常排版都误报一遍。
    """
    lines = md.split('\n')
    bad = []
    for i, line in enumerate(lines):
        if IMG_PAT.match(line.strip()):
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j >= len(lines) or not is_caption(lines[j]):
                bad.append(i + 1)
    return bad


# ---------------------------------------------------------------- Markdown -> OOXML

HR_PAT = re.compile(r'^\s*(---+|\*\*\*+|___+)\s*$')
UL_PAT = re.compile(r'^\s*[-*+]\s+(.*)$')
OL_PAT = re.compile(r'^\s*(\d+)[.、]\s*(.*)$')
BQ_PAT = re.compile(r'^\s*>\s?(.*)$')
TBL_SEP = re.compile(r'^\s*\|[\s:\-\|]+\|\s*$')


def md_to_body(md, title_keywords=None, img_base=None, img_cx=9026 * 635):
    """Markdown 转 Word 正文 XML，返回 (body_xml, images)。

    title_keywords: `#` 级标题中命中任一关键词的映射为 Title，其余映射为 Heading1。
        为空时全部 `#` 映射为 Title（封面约定）。
    img_base: `[[IMG:相对路径]]` 的解析基准目录，默认是输入 md 所在目录。
    img_cx: 图片显示宽度（EMU，1 twip=635），默认按 A4 版心 9026 twips 铺满。
    images: [(rId, 本地绝对路径, 包内文件名), ...]，由 build_docx 写入 media 与 rels。
    """
    kws = [k.strip() for k in (title_keywords or []) if k.strip()]
    body, images, used_names = [], [], set()
    lines = md.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        # 图片：必须排在表格/标题判断之前，否则会被当成普通段落写进正文
        m = IMG_PAT.match(line.strip())
        if m:
            rel = m.group(1).strip()
            path = rel if os.path.isabs(rel) else os.path.join(img_base or '.', rel)
            path = os.path.abspath(path)
            if not os.path.isfile(path):
                raise SystemExit('图片不存在: %s' % path)
            name = os.path.basename(path)
            stem, ext = os.path.splitext(name)
            ext = ext.lower().lstrip('.')
            k = 1
            while name in used_names:
                name = '%s_%d.%s' % (stem, k, ext or 'png')
                k += 1
            used_names.add(name)
            rid = 'rId%d' % (10 + len(images))
            images.append((rid, path, name))
            w, h = image_size(path)
            cx = img_cx
            cy = max(1, int(cx * h / w))
            body.append(img_xml(name, rid, 100 + len(images), cx, cy))
            i += 1
            continue
        # 表格块
        if line.lstrip().startswith('|') and i + 1 < len(lines) and TBL_SEP.match(lines[i + 1]):
            rows = []
            while i < len(lines) and lines[i].lstrip().startswith('|'):
                raw = lines[i].strip()
                if TBL_SEP.match(raw):
                    i += 1
                    continue
                cells = [c.strip() for c in raw.strip('|').split('|')]
                rows.append(cells)
                i += 1
            body.append(table_xml(rows))
            continue
        # 分隔线
        if HR_PAT.match(line):
            i += 1
            continue
        # 标题
        m = re.match(r'^(#{1,6})\s+(.*)$', line)
        if m:
            lv = len(m.group(1))
            txt = m.group(2).strip()
            if lv == 1:
                hit = any(k in txt for k in kws) if kws else True
                style = 'Title' if hit else 'Heading1'
            else:
                style = 'Heading%d' % min(lv - 1, 4)
            body.append(p_xml(txt, style))
            i += 1
            continue
        # 引用块
        m = BQ_PAT.match(line)
        if m:
            body.append(p_xml(m.group(1), 'Normal', indent=True))
            i += 1
            continue
        # 无序列表
        m = UL_PAT.match(line)
        if m:
            body.append(p_xml('· ' + m.group(1), 'Normal'))
            i += 1
            continue
        # 有序列表
        m = OL_PAT.match(line)
        if m:
            body.append(p_xml('%s. %s' % (m.group(1), m.group(2)), 'Normal'))
            i += 1
            continue
        # 图表题注居中（顺序要求：Markdown 里写「图片 → 题注」，图在上题注在下）
        if is_caption(line):
            body.append(p_xml(line.strip(), 'Normal', center=True))
            i += 1
            continue
        body.append(p_xml(line, 'Normal', indent=True))
        i += 1
    return ''.join(body), images


# ---------------------------------------------------------------- 部件模板

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="360"/><w:keepNext/></w:pPr>
<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:keepNext/><w:spacing w:before="360" w:after="180"/><w:outlineLvl w:val="0"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/><w:outlineLvl w:val="1"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="2"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
<w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="3"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="TableText"/><w:basedOn w:val="Normal"/>
<w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders>
<w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/>
<w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/>
<w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/>
</w:tblBorders></w:tblPr></w:style>
</w:styles>"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""


def content_types(images):
    """按实际图片扩展名补 <Default>，避免 png/jpeg 部件没有 Content-Type。"""
    exts = set()
    for _, _, name in images:
        ext = os.path.splitext(name)[1].lower().lstrip('.')
        if ext in IMAGE_TYPES:
            exts.add(ext)
    extra = ''.join('<Default Extension="%s" ContentType="%s"/>' % (e, IMAGE_TYPES[e])
                    for e in sorted(exts))
    return CONTENT_TYPES.replace('</Types>', extra + '</Types>')


def doc_rels(images):
    """rId1/rId2 已被页脚与样式占用，图片一律从 rId10 起编号。"""
    extra = ''.join(
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/image" Target="media/%s"/>' % (rid, name)
        for rid, _, name in images)
    return DOC_RELS.replace('</Relationships>', extra + '</Relationships>')


def footer_xml(text):
    """页脚：{text} 中的 {page} 替换为 PAGE 域。"""
    if not text:
        return ''
    parts = text.split('{page}')
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>']
    for idx, part in enumerate(parts):
        if idx:
            out.append('<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>')
        if part:
            out.append('<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % esc(part))
    out.append('</w:p></w:ftr>')
    return ''.join(out)


def core_xml(title, creator, subject='', keywords=''):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/">'
            '<dc:title>%s</dc:title><dc:creator>%s</dc:creator>'
            '<dc:subject>%s</dc:subject><cp:keywords>%s</cp:keywords>'
            '</cp:coreProperties>' % (esc(title), esc(creator), esc(subject), esc(keywords)))


def app_xml(application='Microsoft Office Word'):
    """应用程序属性。默认写 Word 官方名，避免暴露生成工具。"""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
            '<Application>%s</Application></Properties>' % esc(application))


def build_docx(md, out_path, title='施工方案', creator='投标技术组',
               subject='', keywords='', footer_text='', application='Microsoft Office Word',
               title_keywords=None, page_width=11906, page_height=16838, margin=1440,
               img_base=None):
    img_cx = max(1000, (page_width - 2 * margin) * 635)  # twip -> EMU
    body, images = md_to_body(md, title_keywords, img_base=img_base, img_cx=img_cx)
    sect = ('<w:sectPr>'
            + ('<w:footerReference w:type="default" r:id="rId1"/>' if footer_text else '')
            + '<w:pgSz w:w="%d" w:h="%d"/>' % (page_width, page_height)
            + '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
              'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
              % (margin, margin, margin, margin))
    # a / pic / wp 必须在根上声明；wp 的 URI 只有 wordprocessingDrawing 一种正确写法。
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
           'xmlns:wp="%s" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
           '<w:body>%s%s</w:body></w:document>' % (WP_NS, body, sect))
    d = os.path.dirname(os.path.abspath(out_path))
    if d and not os.path.isdir(d):
        os.makedirs(d)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types(images))
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/_rels/document.xml.rels', doc_rels(images))
        if footer_text:
            z.writestr('word/footer1.xml', footer_xml(footer_text))
        for _, path, name in images:
            with open(path, 'rb') as f:
                z.writestr('word/media/%s' % name, f.read())
        z.writestr('docProps/core.xml', core_xml(title, creator, subject, keywords))
        z.writestr('docProps/app.xml', app_xml(application))
    return out_path


# ---------------------------------------------------------------- 校验

def verify_docx(path):
    """校验 docx：大纲可用性、样式名、元数据。返回 (ok, 详情列表)。"""
    ok = True
    detail = []
    if not os.path.exists(path):
        return False, ['文件不存在: %s' % path]
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if 'word/document.xml' not in names:
            return False, ['缺少 word/document.xml（不是有效的 docx）']
        app = z.read('docProps/app.xml').decode('utf-8', errors='ignore') \
            if 'docProps/app.xml' in names else ''
        core = z.read('docProps/core.xml').decode('utf-8', errors='ignore') \
            if 'docProps/core.xml' in names else ''
        sty = z.read('word/styles.xml').decode('utf-8', errors='ignore') \
            if 'word/styles.xml' in names else ''
        docxml = z.read('word/document.xml').decode('utf-8', errors='ignore')
        relsxml = z.read('word/_rels/document.xml.rels').decode('utf-8', errors='ignore') \
            if 'word/_rels/document.xml.rels' in names else ''
        media = [n for n in names if n.startswith('word/media/')]
    # 图片部件一致性（文本级体检查不到「图片静默丢失」）
    ndraw = docxml.count('<w:drawing>')
    if ndraw:
        mns = re.search(r'xmlns:wp="([^"]+)"', docxml)
        if not mns:
            ok = False
            detail.append('含 %d 个 drawing 但根上未声明 xmlns:wp —— 图片不会渲染' % ndraw)
        elif mns.group(1) != WP_NS:
            ok = False
            detail.append('wp 命名空间 URI=%r —— 应为 %r；写错时 LibreOffice 会静默丢弃图片'
                          % (mns.group(1), WP_NS))
        else:
            detail.append('wp 命名空间 —— 正常')
        embeds = set(re.findall(r'<a:blip r:embed="([^"]+)"', docxml))
        img_tags = ' '.join(re.findall(r'<Relationship\b[^>]*?relationships/image"[^>]*?/>', relsxml))
        img_ids = set(re.findall(r'Id="([^"]+)"', img_tags))
        targets = set(re.findall(r'Target="media/([^"]+)"', img_tags))
        media_names = set(os.path.basename(n) for n in media)
        if embeds != img_ids or targets != media_names or len(media) != ndraw:
            ok = False
            detail.append('图片部件不一致: drawing=%d / r:embed=%d / image关系=%d / media文件=%d'
                          % (ndraw, len(embeds), len(img_ids), len(media)))
        else:
            detail.append('图片部件一致: %d 张（drawing / rels / media 三者匹配）' % ndraw)
    # 元数据
    m = re.search(r'<Application>([^<]*)</Application>', app)
    app_val = m.group(1) if m else '(无)'
    if app_val and app_val not in ('Microsoft Office Word', 'Microsoft Word'):
        ok = False
        detail.append('Application=%r —— 疑似生成工具痕迹' % app_val)
    else:
        detail.append('Application=%r —— 正常' % app_val)
    m = re.search(r'<dc:creator>([^<]*)</dc:creator>', core)
    creator = m.group(1) if m else '(无)'
    if re.search(r'(workbuddy|agent|ai|claude|gpt|deepseek|assistant|智能体)', creator, re.I):
        ok = False
        detail.append('creator=%r —— 疑似 AI 痕迹' % creator)
    else:
        detail.append('creator=%r —— 正常' % creator)
    # 大纲
    ol_ok, ol_detail = C.docx_outline_check(path)
    detail.extend(ol_detail)
    if not ol_ok:
        ok = False
    # 标题样式计数
    counts = C.docx_pstyle_counts(path)
    detail.append('段落样式计数: %s' % json_like(counts))
    heads = {k: v for k, v in counts.items()
             if k.lower().startswith('heading') or k == 'Title'}
    if not heads:
        ok = False
        detail.append('未发现任何标题样式段落 —— 大纲视图将为空')
    elif 'Title' not in counts:
        # 不判失败：有的稿子确实不用封面标题，但多数是 --title-keywords 写了占位符
        detail.append('未发现 Title 段落 —— 封面标题缺失；'
                      '若用了 --title-keywords，请确认关键词是真实项目名而非占位符')
    return ok, detail


def json_like(d):
    import json
    return json.dumps(d, ensure_ascii=False, sort_keys=True)


# ---------------------------------------------------------------- CLI

def main():
    ap = argparse.ArgumentParser(description='Markdown -> Word(.docx) 生成与校验')
    ap.add_argument('--in', dest='inp', help='输入 Markdown 文件')
    ap.add_argument('--out', help='输出 docx 路径')
    ap.add_argument('--title', default='施工方案', help='文档标题（core.xml）')
    ap.add_argument('--creator', default='投标技术组', help='作者（core.xml，不要写 AI 名）')
    ap.add_argument('--subject', default='', help='主题')
    ap.add_argument('--keywords', default='', help='关键词')
    ap.add_argument('--footer-text', default='',
                    help='页脚文字，{page} 处插入页码；留空则无页脚')
    ap.add_argument('--application', default='Microsoft Office Word',
                    help='Application 元数据（默认 Microsoft Office Word）')
    ap.add_argument('--title-keywords', default='',
                    help='逗号分隔；# 级标题命中任一关键词才映射为 Title')
    ap.add_argument('--img-base', default='',
                    help='[[IMG:相对路径]] 的基准目录，默认取输入 md 所在目录')
    ap.add_argument('--verify', help='校验已有 docx（不生成）')
    args = ap.parse_args()

    if args.verify:
        ok, detail = verify_docx(args.verify)
        for x in detail:
            print(('  [OK] ' if not x.startswith(('Application=', 'creator=')) else '  ') + x)
        print('校验结果: %s' % ('通过' if ok else '不通过'))
        raise SystemExit(0 if ok else 1)

    if not args.inp or not args.out:
        raise SystemExit('生成模式需要 --in 与 --out')
    md = C.read_text(args.inp)
    kws = [k for k in args.title_keywords.split(',') if k.strip()]
    img_base = args.img_base or os.path.dirname(os.path.abspath(args.inp))
    bad = figure_order_warnings(md)
    if bad:
        print('警告: %d 张图片之后未紧跟题注（图片所在行 %s%s）；规范顺序是「图片 → 题注」'
              % (len(bad), ','.join(str(x) for x in bad[:8]),
                 '...' if len(bad) > 8 else ''))
    build_docx(md, args.out, title=args.title, creator=args.creator,
               subject=args.subject, keywords=args.keywords,
               footer_text=args.footer_text, application=args.application,
               title_keywords=kws, img_base=img_base)
    n_img = sum(1 for l in md.split('\n') if IMG_PAT.match(l.strip()))
    print('已生成: %s（%d 字，%d 张图片）' % (os.path.abspath(args.out), len(md), n_img))
    ok, detail = verify_docx(args.out)
    print('自检: %s' % ('通过' if ok else '不通过'))
    for x in detail:
        print('  -', x)


if __name__ == '__main__':
    main()

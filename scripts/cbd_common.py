# -*- coding: utf-8 -*-
"""施工标编制共享工具库。纯标准库，无第三方依赖。

被 merge_validate.py / audit_structure.py / compliance_check.py / build_docx.py 共用。
设计原则：项目专属内容（禁用词、硬指标、章节骨架）一律走外部文件，本库只放通用方法。
"""
import json
import os
import re
import zipfile

# ---------------------------------------------------------------- 正则常量

# 表/图编号：形如「表 3-1」「图 3-11」，允许全角破折号与空白
NUM_PAT = re.compile(r'(表|图)\s*(\d+)\s*[-—]\s*(\d+)')

# Markdown 标题
HEAD_PAT = re.compile(r'^(#{1,6})\s+(.*)$')

# 标题行首的编号：形如「3.1.1」「4」
NUM_HEAD_PAT = re.compile(r'^(\d+(?:\.\d+)*)\s')

# 初稿注占位
NOTE_PAT = re.compile(r'〔初稿注[^〕]*〕')

# 标准规范版本号：GB 50204-2002、JGJ 46-2005、DL/T 5009.3-1997 等
STD_PAT = re.compile(
    r'(GB|GB/T|GB/Z|GBJ|JGJ|JGJ/T|DL|DL/T|NB/T|CJJ|JTJ|JGJ/T)'
    r'\s*([0-9.]+)\s*[-—]\s*((?:19|20\d)\d)'
)

# 标准规范白名单：确实现行有效的旧版号（例外需写明理由）
STD_WHITELIST = {
    ('JGJ', '46', '2005'),   # 《施工现场临时用电安全技术规范》JGJ 46-2005 现行有效
}

# 表格分隔行（Markdown）
TBL_SEP_PAT = re.compile(r'^\s*\|[\s:\-\|]+\|\s*$')

# ---------------------------------------------------------------- 基础 IO

def read_lines(path, comment='#'):
    """读规则文件：每行一条，`comment` 开头为注释，空行忽略。"""
    if not path:
        return []
    if not os.path.exists(path):
        raise SystemExit('规则文件不存在: %s' % path)
    out = []
    with open(path, encoding='utf-8') as f:
        for line in f:
            s = line.strip()
            if s and not s.startswith(comment):
                out.append(s)
    return out


def read_pairs(path, seps=None):
    """读「键 分隔符 值」规则文件，返回 [(键, 值), ...]。

    用于硬指标表（指标名: 关键字符串）。分隔符依次尝试，默认 ':' '：' '=' ' '。
    """
    if seps is None:
        seps = [':', '：', '=', ' ']
    out = []
    for line in read_lines(path):
        for sep in seps:
            if sep in line:
                k, v = line.split(sep, 1)
                out.append((k.strip(), v.strip()))
                break
    return out


def parse_hard_rules(pairs):
    """把硬指标规则解析为 [(指标名, 关键字符串, 最少出现次数), ...]。

    支持 `指标名: 关键字符串:N` 声明最少出现次数（N 省略时默认 1）。
    「双重出现」类硬指标（如关键线路、尾水）必须写明次数，否则只判有无。
    """
    out = []
    for k, v in pairs:
        m = re.match(r'^(.*):(\d+)$', v)
        if m:
            out.append((k, m.group(1).strip(), int(m.group(2))))
        else:
            out.append((k, v, 1))
    return out


def read_replace_rules(path, sep='==>'):
    """读措辞归一规则文件：每行 `原文 ==> 替换`，返回 [(原文, 替换), ...]。"""
    out = []
    for line in read_lines(path):
        if sep in line:
            a, b = line.split(sep, 1)
            out.append((a.strip(), b.strip()))
    return out


def load_json(path, default=None):
    if not path or not os.path.exists(path):
        return default
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def write_text(path, text):
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)


def read_text(path):
    with open(path, encoding='utf-8') as f:
        return f.read()


def list_segments(base):
    """列出分册文件，按文件名中的数字序排序（p01-2.md → (1,2)）。"""
    pat = os.path.join(base, 'p[0-9][0-9]-[0-9].md')
    import glob
    files = glob.glob(pat)
    return sorted(files, key=lambda x: tuple(int(v) for v in re.findall(r'\d+', os.path.basename(x))))


# ---------------------------------------------------------------- 编号

def plan_numbers(files, cross_ref=None):
    """两遍法编号第一遍：按分册顺序分配全篇唯一连续号，并解析跨分册引用。

    cross_ref 结构：{"分册名|表|章|册内号": ["目标分册", "表", 章, 号], ...}
    各分册内部用流水号，合并时统一重排为「章-连续号」。
    """
    cross_ref = cross_ref or {}
    defs = {}
    counters = {}
    pending = []
    for f in files:
        name = os.path.basename(f)
        text = read_text(f)
        seen = set()
        for m in NUM_PAT.finditer(text):
            kind, ch, seq = m.group(1), int(m.group(2)), int(m.group(3))
            if (kind, ch, seq) in seen:
                continue
            seen.add((kind, ch, seq))
            key = '%s|%s|%d|%d' % (name, kind, ch, seq)
            tgt = cross_ref.get(key)
            if tgt:
                pending.append((name, kind, ch, seq, tgt))
                continue
            ck = (kind, ch)
            counters[ck] = counters.get(ck, 0) + 1
            defs[key] = counters[ck]
    # 第二遍：跨册引用指向被引用册的定义号
    for name, kind, ch, seq, tgt in pending:
        tname, tkind, tch, tseq = tgt[0], tgt[1], int(tgt[2]), int(tgt[3])
        tkey = '%s|%s|%d|%d' % (tname, tkind, tch, tseq)
        if tkey in defs:
            defs['%s|%s|%d|%d' % (name, kind, ch, seq)] = defs[tkey]
    return defs


def apply_numbers(text, name, defs):
    """两遍法编号第二遍：把册内流水号回写为全篇统一号。"""
    def repl(m):
        kind, ch, seq = m.group(1), int(m.group(2)), int(m.group(3))
        new = defs.get('%s|%s|%d|%d' % (name, kind, ch, seq))
        return '%s %d-%d' % (kind, ch, new) if new else m.group(0)
    return NUM_PAT.sub(repl, text)


def check_numbering(full):
    """校验合并后 表/图 编号：各章应为 1..N 连续，无缺号。"""
    first = {}
    for m in NUM_PAT.finditer(full):
        kind, ch, seq = m.group(1), int(m.group(2)), int(m.group(3))
        first.setdefault((kind, ch), set()).add(seq)
    problems = []
    for (kind, ch), seqs in sorted(first.items()):
        n = max(seqs)
        missing = sorted(set(range(1, n + 1)) - seqs)
        if missing:
            problems.append('%s %d-* 缺号: %s' % (kind, ch, missing))
        if len(seqs) != n:
            problems.append('%s %d-* 编号不连续（出现 %d 个不同号，最大 %d）'
                            % (kind, ch, len(seqs), n))
    return problems


# 表/图「定义」的两种形态：
#   ① 独立题注行：行首（可含 **）是「表 3-1 名称…」；
#   ② 正文内嵌题注：「图 3-1 施工区段划分图（另册附图）」——另册附图的占位
#      写法普遍嵌在句子里，只认行首会漏掉这批，进而把正常引用误判为悬空。
DEF_LINE_PAT = re.compile(r'^\*{0,2}\s*(表|图)\s*(\d+)\s*[-—]\s*(\d+)\s*\*{0,2}\s*\S')
DEF_INLINE_PAT = re.compile(r'[「『]\s*(表|图)\s*(\d+)\s*[-—]\s*(\d+)\s*\S')


def find_dangling_refs(full):
    """找出「被引用但从未定义」的表/图号（B 类提示，人工复核）。

    定义判定见 DEF_LINE_PAT / DEF_INLINE_PAT。
    注意：不能简单统计「出现次数 > 1」——跨册引用解析后同一编号出现多次
    属正常，真正要抓的是「只有引用、没有题注」的悬空号。
    """
    defined = set()
    for line in full.split('\n'):
        s = line.strip()
        m = DEF_LINE_PAT.match(s) or DEF_INLINE_PAT.search(s)
        if m:
            defined.add((m.group(1), int(m.group(2)), int(m.group(3))))
    dangling = set()
    for m in NUM_PAT.finditer(full):
        k = (m.group(1), int(m.group(2)), int(m.group(3)))
        if k not in defined:
            dangling.add(k)
    return sorted(dangling)


# ---------------------------------------------------------------- 标题结构

def parse_headings(text):
    """解析 Markdown 标题，返回 [(行号, #级数, 标题文本, 编号或None), ...]。"""
    out = []
    for i, line in enumerate(text.split('\n'), 1):
        m = HEAD_PAT.match(line.strip())
        if not m:
            continue
        txt = m.group(2).strip()
        nm = NUM_HEAD_PAT.match(txt)
        out.append((i, len(m.group(1)), txt, nm.group(1) if nm else None))
    return out


def check_heading_structure(text, allow_start_any=False):
    """标题结构检查：层级跳跃、编号重复（A 类）；编号分支不衔接（B 类）。

    逻辑层级只看 Markdown `#` 级数——编号深度不可靠：
    有的章节用 3.1/3.1.1/3.1.1.1，有的用 4/4.1/4.1.1，同一深度分属不同层级。

    allow_start_any：允许首个标题出现在任意层级。分册片段与不含文首横幅的
    合并稿首个标题可能是 H2，此时不应判为跳跃。

    返回 (heads, issues_a, issues_b)：
      issues_a —— 层级跳跃、编号重复，属结构性硬错误；
      issues_b —— 编号与上层分支不衔接，属书写口径问题，不判红线。

    编号衔接的判定向上追溯所有祖先，而不是只看直接父节点：
    兄弟节点（如 3.3.4 跟在 3.3.3 之后）编号本就不以父节点为前缀，
    只看直接父节点会把这类正常书写误判成结构断裂。
    """
    heads = parse_headings(text)
    issues_a = []
    issues_b = []
    prev_lv = 0
    stack = []          # [(级数, 编号)] 逻辑层级栈
    seen_num = {}       # 编号 -> 行号
    for idx, (line, lv, txt, num) in enumerate(heads):
        if lv > prev_lv + 1 and not (allow_start_any and idx == 0):
            issues_a.append('L%d 标题层级跳跃 H%d->H%d: %s'
                            % (line, prev_lv, lv, txt[:60]))
        # 同层出现同编号 → 重复节点（分册撞车）
        if num:
            if num in seen_num:
                issues_a.append('L%d 编号重复: %s（首见 L%d）'
                                % (line, num, seen_num[num]))
            seen_num[num] = line
        # 维护层级栈：先弹出同级与更深层，栈顶即直接父节点
        while stack and stack[-1][0] >= lv:
            stack.pop()
        # 编号分支检查：向上追溯任一祖先编号能覆盖当前编号即视为衔接。
        # 祖先全无编号时无从校验，跳过——否则「## 3.1」挂在不带编号的 H1 下会被误判。
        numbered = [p for _, p in stack if p]
        if num and numbered:
            if not any(num == p or num.startswith(p + '.')
                       for p in numbered):
                issues_b.append('L%d 编号 %s 与上层分支不衔接（上层 %s）'
                                % (line, num, ' > '.join(numbered)))
        stack.append((lv, num))
        prev_lv = max(prev_lv, lv)
    return heads, issues_a, issues_b


def heading_style_map(lv):
    """Markdown `#` 级数 → Word 样式 ID。

    `#` 约定只用于封面（项目名/文档名），映射 Title；
    `##`→Heading1、`###`→Heading2、`####`→Heading3、`#####`→Heading4。
    这样 Word 导航窗格得到 1 / 1.1 / 1.1.1 / 1.1.1.1 四级大纲。
    """
    if lv == 1:
        return 'Title'
    return 'Heading%d' % min(lv - 1, 4)


# ---------------------------------------------------------------- 表格

def parse_tables(text):
    """解析 Markdown 表格块，返回 [(起始行, 表头列表, 数据行数, 是否紧接表名), ...]。"""
    lines = text.split('\n')
    out = []
    i = 0
    while i < len(lines):
        if lines[i].lstrip().startswith('|') and i + 1 < len(lines) and TBL_SEP_PAT.match(lines[i + 1]):
            start = i + 1
            hdr = [c.strip() for c in lines[i].strip().strip('|').split('|')]
            j = i + 2
            rows = 0
            while j < len(lines) and lines[j].lstrip().startswith('|'):
                cells = [c.strip() for c in lines[j].strip().strip('|').split('|')]
                if len(cells) != len(hdr):
                    out.append((j + 1, None, 0, '列数不符 %d vs %d' % (len(cells), len(hdr))))
                rows += 1
                j += 1
            # 表名判定：回看上文，命中「表 X-Y …」题注即视为已命名。
            # 题注可能加粗（**表 3-19 …**），与表格之间可能隔空行与「注：…」行，
            # 故按「非空行」计数回看，遇到标题或表格行即停止（避免跨表误判）。
            named = False
            name = None
            k = i - 1
            seen = 0
            while k >= 0 and seen < 6:
                s = lines[k].strip()
                if not s:
                    k -= 1
                    continue
                if s.startswith('#') or s.lstrip().startswith('|'):
                    break
                m = re.match(r'^\*{0,2}\s*表\s*(\d+)\s*[-—]\s*(\d+)\s*(.*?)\s*\*{0,2}$', s)
                if m:
                    named = True
                    name = m.group(3).strip().strip('*').strip()
                    break
                seen += 1
                k -= 1
            out.append((start, hdr, rows, ('named', name) if named else 'unnamed'))
            i = j
            continue
        i += 1
    return out


def find_unnamed_tables(text):
    """找出缺表名的表格块：表格前一行不是「表 X-Y …」。"""
    bad = []
    for start, hdr, rows, flag in parse_tables(text):
        if isinstance(flag, tuple):
            continue
        if flag == 'unnamed':
            first = ' | '.join(hdr[:3]) if hdr else ''
            bad.append((start, first[:60], rows))
        else:
            bad.append((start, flag, rows))
    return bad


# ---------------------------------------------------------------- 格式残留

def check_format(text):
    """扫描格式残留：空单元格、行尾空白、句号后多空格、NUL、表格块缩进。"""
    issues = []
    for i, line in enumerate(text.split('\n'), 1):
        if re.search(r'\|\s+\|', line):
            issues.append('L%d 空单元格: %s' % (i, line[:70]))
        if line != line.rstrip():
            issues.append('L%d 行尾空白' % i)
        if re.search(r'。\s{2,}\S', line):
            issues.append('L%d 句号后多空格: %s' % (i, line[:70]))
        if '\x00' in line:
            issues.append('L%d 含 NUL 字符' % i)
        if re.match(r'^[ \t]+\|', line):
            issues.append('L%d 表格块整体缩进: %s' % (i, line[:70]))
    return issues


# ---------------------------------------------------------------- docx 读取

def docx_paragraphs(path):
    """读 docx 段落，返回 [(样式ID, 样式名, 文本), ...]。"""
    with zipfile.ZipFile(path) as z:
        doc = z.read('word/document.xml').decode('utf-8', errors='ignore')
        try:
            sty = z.read('word/styles.xml').decode('utf-8', errors='ignore')
        except KeyError:
            sty = ''
    names = {}
    for m in re.finditer(r'<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>', sty, re.S):
        nm = re.search(r'<w:name w:val="([^"]+)"', m.group(2))
        names[m.group(1)] = nm.group(1) if nm else '?'
    out = []
    for m in re.finditer(r'<w:p[ >].*?</w:p>', doc, re.S):
        s = m.group(0)
        sid = re.search(r'w:pStyle w:val="([^"]+)"', s)
        txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', s)).strip()
        if not txt:
            continue
        key = sid.group(1) if sid else 'Normal'
        out.append((key, names.get(key, '?'), txt))
    return out


def docx_pstyle_counts(path):
    """统计 docx 各段落样式出现次数。"""
    import collections
    c = collections.Counter()
    for key, name, txt in docx_paragraphs(path):
        c[name] += 1
    return dict(c)


def docx_outline_check(path):
    """检查 docx 大纲可用性：标题样式名必须是 Word 内置名，且带 outlineLvl。

    返回 (可用?, 详情列表)。
    """
    with zipfile.ZipFile(path) as z:
        try:
            sty = z.read('word/styles.xml').decode('utf-8', errors='ignore')
        except KeyError:
            return False, ['缺少 word/styles.xml']
    builtins = {'heading 1': 0, 'heading 2': 1, 'heading 3': 2, 'heading 4': 3}
    detail = []
    ok = True
    for m in re.finditer(r'<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>', sty, re.S):
        sid, body = m.group(1), m.group(2)
        nm = re.search(r'<w:name w:val="([^"]+)"', body)
        if not nm or nm.group(1) not in builtins:
            continue
        ol = re.search(r'<w:outlineLvl w:val="([^"]+)"', body)
        want = builtins[nm.group(1)]
        if not ol:
            ok = False
            detail.append('%s: 缺 w:outlineLvl（大纲不可用）' % sid)
        elif int(ol.group(1)) != want:
            ok = False
            detail.append('%s: outlineLvl=%s 应为 %d' % (sid, ol.group(1), want))
        else:
            detail.append('%s: 正确（name=%s, outlineLvl=%d）' % (sid, nm.group(1), want))
    if not detail:
        ok = False
        detail.append('未发现任何内置标题样式')
    return ok, detail


def docx_full_text(path):
    """提取 docx 全文（段落文本拼接）。"""
    return '\n'.join(t for _, _, t in docx_paragraphs(path))


# ---------------------------------------------------------------- 旧规范扫描

def scan_old_standards(text, whitelist=None):
    """扫描 2008 年以前的标准版本号（视为疑似废止），白名单除外。"""
    wl = whitelist if whitelist is not None else STD_WHITELIST
    hits = []
    for m in STD_PAT.finditer(text):
        s, n, y = m.group(1), m.group(2), m.group(3)
        if int(y) < 2008 and (s, n, y) not in wl:
            hits.append((s, n, y, text[max(0, m.start() - 20):m.end() + 20].replace('\n', ' ')))
    return hits


def scan_notes(text):
    """列出全部〔初稿注〕及其行号。"""
    out = []
    for m in NOTE_PAT.finditer(text):
        line = text[:m.start()].count('\n') + 1
        tail = text[m.end():m.end() + 1]
        out.append((line, m.group(0), tail))
    return out

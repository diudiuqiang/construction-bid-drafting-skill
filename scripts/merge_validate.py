# -*- coding: utf-8 -*-
"""分册合并 + 两遍编号 + 合规校验，输出 Markdown 合并稿与构建报告。

用法示例：
  python merge_validate.py --base ./draft --out ./out/施工方案.md \
      --banner ./draft/_banner.md --banned ./conf/禁用词.txt \
      --hard ./conf/硬指标.txt --replace ./conf/措辞归一.txt \
      --cross-ref ./conf/cross_ref.json --report ./out/build_report.json

所有规则文件均可省略；省略时对应检查跳过（不臆造项目规则）。
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cbd_common as C


def build_banner(template, meta):
    """按 meta 替换 banner 模板中的 {占位符}。缺省占位符保留原样并告警。"""
    if not template:
        return ''
    out = template
    for k, v in (meta or {}).items():
        out = out.replace('{%s}' % k, str(v))
    left = [t for t in out.split('\n') if '{' in t and '}' in t]
    return out, left


def do_merge(files, banner, replace_rules, cross_ref):
    """合并分册：两遍编号 → 拼接 → 措辞归一。"""
    defs = C.plan_numbers(files, cross_ref)
    parts = []
    info = []
    if banner:
        parts.append(banner.strip())
        parts.append('\n\n')
    for f in files:
        name = os.path.basename(f)
        s = C.read_text(f).strip()
        s = C.apply_numbers(s, name, defs)
        parts.append(s)
        parts.append('\n\n')
        info.append([name, len(s)])
    full = ''.join(parts).strip() + '\n'
    for a, b in replace_rules:
        full = full.replace(a, b)
    return full, info, defs


def do_validate(full, banned, hard, notes_limit, seg_files):
    """合规校验：红线 / 硬指标 / 编号 / 结构 / 格式 / 表格名 / 初稿注。"""
    report = {}

    # 1. 禁用词（允许「不配置储能」这类正确声明，先剔除再匹配）
    check_text = full
    for prefix in ('不配置', '无', '不设', '不涉及', '不含', '不作', '不采用', '不再'):
        check_text = check_text.replace(prefix + '储能', '  ')
    hits = {}
    for pat in banned:
        try:
            ms = list(__import__('re').finditer(pat, check_text))
        except __import__('re').error:
            ms = []
        if ms:
            hits[pat] = len(ms)
    report['banned_hits'] = hits

    # 2. 旧规范版本号
    report['old_standards'] = C.scan_old_standards(full)

    # 3. 硬指标（支持「:N」声明最少出现次数，如「关键线路: 关键线路:2」）
    missing = []
    for k, needle, min_n in C.parse_hard_rules(hard):
        n = full.count(needle)
        if n < min_n:
            missing.append(k if min_n <= 1
                           else '%s（需 ≥%d 处，实际 %d 处）' % (k, min_n, n))
    report['missing_hard_items'] = missing

    # 4. 编号连续性
    report['numbering_problems'] = C.check_numbering(full)
    report['dangling_refs'] = ['%s %d-%d' % x for x in C.find_dangling_refs(full)]

    # 5. 标题结构（A 类：层级跳跃 / 编号重复；B 类：编号分支不衔接）
    heads, struct_a, struct_b = C.check_heading_structure(full)
    report['structure_issues'] = struct_a
    report['number_branch_issues'] = struct_b
    report['heading_count'] = len(heads)

    # 6. 格式残留
    report['format_issues'] = C.check_format(full)

    # 7. 缺表名表格
    report['unnamed_tables'] = C.find_unnamed_tables(full)

    # 8. 初稿注
    notes = C.scan_notes(full)
    report['notes_count'] = len(notes)
    report['notes_over_limit'] = len(notes) > notes_limit
    report['notes_detail'] = [{'line': l, 'text': t} for l, t, _ in notes]

    # 9. 红线汇总
    red = []
    if hits:
        red.append('禁用词命中 %d 种' % len(hits))
    if report['old_standards']:
        red.append('疑似废止规范 %d 处' % len(report['old_standards']))
    if missing:
        red.append('硬指标缺失 %d 项' % len(missing))
    report['red_lines'] = red
    report['status'] = 'fail' if red else 'pass'
    return report


def main():
    ap = argparse.ArgumentParser(description='分册合并与合规校验')
    ap.add_argument('--base', default='.', help='分册目录（含 pNN-M.md）')
    ap.add_argument('--out', required=True, help='输出 Markdown 合并稿路径')
    ap.add_argument('--banner', help='封面模板文件（含 {占位符}）')
    ap.add_argument('--meta', help='封面占位元信息 JSON 文件')
    ap.add_argument('--banned', help='禁用词文件（每行一条正则）')
    ap.add_argument('--hard', help='硬指标文件（每行「指标名: 关键字符串」）')
    ap.add_argument('--replace', help='措辞归一规则文件（每行「原文 ==> 替换」）')
    ap.add_argument('--cross-ref', help='跨分册引用登记表 JSON')
    ap.add_argument('--report', help='构建报告 JSON 输出路径')
    ap.add_argument('--notes-limit', type=int, default=10, help='初稿注上限（默认 10）')
    ap.add_argument('--quiet', action='store_true', help='只输出报告 JSON，不打印明细')
    args = ap.parse_args()

    files = C.list_segments(args.base)
    if not files:
        raise SystemExit('在 %s 未找到分册文件（pNN-M.md）' % args.base)

    banner = C.read_text(args.banner) if args.banner and os.path.exists(args.banner) else ''
    meta = C.load_json(args.meta, {}) or {}
    if banner:
        banner, unfilled = build_banner(banner, meta)
        if unfilled and not args.quiet:
            print('警告：banner 存在未替换占位符：', unfilled)

    banned = C.read_lines(args.banned)
    hard = C.read_pairs(args.hard)
    replace_rules = C.read_replace_rules(args.replace)
    cross_ref = C.load_json(args.cross_ref, {}) or {}

    full, info, _ = do_merge(files, banner, replace_rules, cross_ref)
    report = do_validate(full, banned, hard, args.notes_limit, files)

    C.write_text(args.out, full)
    report['segments'] = info
    report['total_files'] = len(files)
    report['total_chars'] = len(full)
    report['md'] = os.path.abspath(args.out)

    if args.report:
        C.write_text(args.report, json.dumps(report, ensure_ascii=False, indent=2))

    if not args.quiet:
        print('分册 %d 个，合并 %d 字' % (report['total_files'], report['total_chars']))
        print('状态: %s' % report['status'])
        if report['red_lines']:
            for x in report['red_lines']:
                print('  [红线]', x)
        if report['numbering_problems']:
            print('编号问题:')
            for x in report['numbering_problems']:
                print('  -', x)
        if report['structure_issues']:
            print('结构问题:')
            for x in report['structure_issues']:
                print('  -', x)
        if report['unnamed_tables']:
            print('缺表名表格 %d 处（记录待修正，不擅自改号）' % len(report['unnamed_tables']))
        if report['format_issues']:
            print('格式残留 %d 处' % len(report['format_issues']))
        print('初稿注 %d 处（上限 %d）' % (report['notes_count'], args.notes_limit))
    else:
        print(json.dumps({k: report[k] for k in
                          ('total_files', 'total_chars', 'status', 'red_lines')},
                         ensure_ascii=False))


if __name__ == '__main__':
    main()

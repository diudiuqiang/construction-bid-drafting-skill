# -*- coding: utf-8 -*-
"""结构审计（只读）：标题层级、编号连续性、表格完整性、docx 大纲可用性。

支持三种输入：
  1. 合并后的 Markdown 文件        --in ./out/施工方案.md
  2. 分册目录（pNN-M.md）          --in ./draft --segments
  3. 生成的 Word 文件              --in ./out/施工方案.docx

用法：
  python audit_structure.py --in ./out/施工方案.md --report ./out/audit.json
  python audit_structure.py --in ./draft --segments
  python audit_structure.py --in ./out/施工方案.docx
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cbd_common as C


def audit_md(text, label='', allow_start_any=False, check_num=True):
    """审计 Markdown 结构。

    check_num=False 用于分册片段：分册内表号按本地编号书写，跨册撞号属正常，
    编号连续性须在合并后由合并稿校验。
    """
    out = {'label': label, 'type': 'md'}
    heads, issues_a, issues_b = C.check_heading_structure(text, allow_start_any=allow_start_any)
    out['heading_count'] = len(heads)
    out['heading_levels'] = {}
    for _, lv, _, _ in heads:
        out['heading_levels']['H%d' % lv] = out['heading_levels'].get('H%d' % lv, 0) + 1
    out['structure_issues'] = issues_a
    out['number_branch_issues'] = issues_b
    out['numbering_problems'] = C.check_numbering(text) if check_num else []
    out['unnamed_tables'] = C.find_unnamed_tables(text)
    out['table_count'] = len(C.parse_tables(text))
    out['format_issues'] = C.check_format(text)
    out['notes'] = len(C.scan_notes(text))
    out['total_chars'] = len(text)
    out['status'] = 'fail' if (issues_a or out['numbering_problems']) else 'pass'
    # 说明：分册片段传 check_num=False 时，status 只反映标题层级结构，
    # 不覆盖册内编号连续性（跨册撞号属正常，须在合并稿上校验）。
    out['status_scope'] = ('标题层级 + 编号连续性' if check_num
                           else '仅标题层级（分册片段不校验册内编号）')
    return out


def audit_docx(path):
    """审计 docx 结构与元数据。"""
    out = {'label': os.path.basename(path), 'type': 'docx'}
    ok, detail = C.docx_outline_check(path)
    out['outline_ok'] = ok
    out['outline_detail'] = detail
    out['style_counts'] = C.docx_pstyle_counts(path)
    heads = {k: v for k, v in out['style_counts'].items()
             if k.lower().startswith('heading') or k == 'Title'}
    out['heading_counts'] = heads
    out['total_chars'] = len(C.docx_full_text(path))
    # 元数据
    import re as _re
    import zipfile
    with zipfile.ZipFile(path) as z:
        app = z.read('docProps/app.xml').decode('utf-8', errors='ignore') \
            if 'docProps/app.xml' in z.namelist() else ''
        core = z.read('docProps/core.xml').decode('utf-8', errors='ignore') \
            if 'docProps/core.xml' in z.namelist() else ''
    m = _re.search(r'<Application>([^<]*)</Application>', app)
    out['application'] = m.group(1) if m else '(无)'
    m = _re.search(r'<dc:creator>([^<]*)</dc:creator>', core)
    out['creator'] = m.group(1) if m else '(无)'
    m = _re.search(r'<dc:title>([^<]*)</dc:title>', core)
    out['title'] = m.group(1) if m else '(无)'
    out['meta_issues'] = []
    if out['application'] and out['application'] not in ('Microsoft Office Word', 'Microsoft Word'):
        out['meta_issues'].append('Application=%r 疑似工具痕迹' % out['application'])
    if _re.search(r'(workbuddy|agent|claude|gpt|deepseek|assistant|智能体)', out['creator'], _re.I):
        out['meta_issues'].append('creator=%r 疑似 AI 痕迹' % out['creator'])
    out['status'] = 'pass' if (ok and not out['meta_issues'] and heads) else 'fail'
    return out


def audit_segments(base):
    """审计分册目录：逐册结构 + 合并后结构。"""
    files = C.list_segments(base)
    if not files:
        raise SystemExit('在 %s 未找到分册文件（pNN-M.md）' % base)
    out = {'label': base, 'type': 'segments', 'segments': []}
    merged = []
    for f in files:
        text = C.read_text(f)
        r = audit_md(text, os.path.basename(f), allow_start_any=True, check_num=False)
        out['segments'].append(r)
        merged.append(text)
    defs = C.plan_numbers(files)
    parts = []
    for f, text in zip(files, merged):
        parts.append(C.apply_numbers(text, os.path.basename(f), defs))
    full = '\n\n'.join(parts)
    # 分册合并稿不含文首横幅（横幅由构建脚本注入），故允许首标题为任意层级
    out['merged'] = audit_md(full, '合并稿（不含横幅）', allow_start_any=True)
    out['status'] = 'fail' if out['merged']['status'] == 'fail' else 'pass'
    return out


def print_report(rep):
    t = rep['type']
    print('=' * 60)
    print('结构审计: %s' % rep['label'])
    print('类型: %s   状态: %s' % (t, rep.get('status', '-')))
    print('字符数: %s' % rep.get('total_chars', '-'))
    if t == 'segments':
        print('分册 %d 个:' % len(rep['segments']))
        for s in rep['segments']:
            print('  %s  标题 %d  表格 %d  缺名 %d  状态 %s'
                  % (s['label'], s['heading_count'], s['table_count'],
                     len(s['unnamed_tables']), s['status']))
        m = rep['merged']
        print('合并稿: 标题 %d  表格 %d  缺名 %d  状态 %s'
              % (m['heading_count'], m['table_count'], len(m['unnamed_tables']), m['status']))
        for x in m['structure_issues']:
            print('  [结构]', x)
        for x in m['numbering_problems']:
            print('  [编号]', x)
        if m['unnamed_tables']:
            print('  缺表名表格（前 5）:')
            for row in m['unnamed_tables'][:5]:
                print('    L%s %s' % (row[0], row[1]))
    elif t == 'md':
        print('标题数: %d  层级分布: %s'
              % (rep['heading_count'], json.dumps(rep['heading_levels'], ensure_ascii=False)))
        print('表格数: %d  缺表名: %d  初稿注: %d'
              % (rep['table_count'], len(rep['unnamed_tables']), rep['notes']))
        for x in rep['structure_issues']:
            print('  [结构]', x)
        for x in rep['numbering_problems']:
            print('  [编号]', x)
        if rep['unnamed_tables']:
            print('  缺表名表格（前 5）:')
            for row in rep['unnamed_tables'][:5]:
                print('    L%s %s' % (row[0], row[1]))
        if rep['format_issues']:
            print('  格式残留 %d 处（前 5）:' % len(rep['format_issues']))
            for x in rep['format_issues'][:5]:
                print('    ', x)
    elif t == 'docx':
        print('大纲可用: %s' % ('是' if rep['outline_ok'] else '否'))
        for x in rep['outline_detail']:
            print('  -', x)
        print('标题样式计数: %s' % json.dumps(rep['heading_counts'], ensure_ascii=False))
        print('元数据: Application=%s  creator=%s  title=%s'
              % (rep['application'], rep['creator'], rep['title']))
        for x in rep['meta_issues']:
            print('  [元数据]', x)
    print('=' * 60)


def main():
    ap = argparse.ArgumentParser(description='结构审计（只读）')
    ap.add_argument('--in', dest='inp', required=True, help='输入文件/目录')
    ap.add_argument('--segments', action='store_true', help='输入为分册目录')
    ap.add_argument('--report', help='报告 JSON 输出路径')
    ap.add_argument('--quiet', action='store_true', help='只输出 JSON 摘要')
    args = ap.parse_args()

    if args.segments:
        rep = audit_segments(args.inp)
    elif args.inp.lower().endswith('.docx'):
        rep = audit_docx(args.inp)
    else:
        rep = audit_md(C.read_text(args.inp), os.path.basename(args.inp))

    if args.report:
        C.write_text(args.report, json.dumps(rep, ensure_ascii=False, indent=2))
    if not args.quiet:
        print_report(rep)
    else:
        print(json.dumps({'type': rep['type'], 'label': rep['label'],
                          'status': rep.get('status'), 'total_chars': rep.get('total_chars')},
                         ensure_ascii=False))
    raise SystemExit(0 if rep.get('status') == 'pass' else 1)


if __name__ == '__main__':
    main()

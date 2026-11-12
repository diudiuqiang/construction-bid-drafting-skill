# -*- coding: utf-8 -*-
"""合规体检（只读，不改文件）：技术标稿件全面检查，输出 A/B/C/D/E 分类报告。

分类口径：
  A 类 必须修正（红线）—— 禁用词命中、疑似废止规范、硬指标缺失、结构断裂
  B 类 应当修正 —— 编号缺号/不连续、缺表名表格、格式残留、初稿注超限
  C 类 建议优化 —— AI 痕迹（warn 级）、风格统一性
  D 类 待澄清 —— 〔初稿注〕内容（需勘察/设计/业主补充资料）
  E 类 提示 —— 一致性建议、待复核项

基线对比：--prev 传入上次体检的 JSON，输出「新增 / 已修复 / 遗留」三类变化。

用法：
  python compliance_check.py --in ./out/施工方案.md \
      --banned ./conf/禁用词.txt --hard ./conf/硬指标.txt \
      --prev ./out/compliance_prev.json \
      --report ./out/compliance.json --md ./out/合规体检报告.md
"""
import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cbd_common as C
import ai_trace


def load_text(path):
    """按扩展名读取文本（md/docx）。"""
    if path.lower().endswith('.docx'):
        return C.docx_full_text(path), 'docx'
    return C.read_text(path), 'md'


def classify(full, banned, hard, notes_limit):
    """执行全部检查并分类。"""
    issues = {'A': [], 'B': [], 'C': [], 'D': [], 'E': []}

    # ---- A 类：红线
    check_text = full
    for prefix in ('不配置', '无', '不设', '不涉及', '不含', '不作', '不采用', '不再'):
        check_text = check_text.replace(prefix + '储能', '  ')
    banned_hits = {}
    for pat in banned:
        try:
            n = len(re.findall(pat, check_text))
        except re.error:
            continue
        if n:
            banned_hits[pat] = n
    for pat, n in banned_hits.items():
        issues['A'].append({'kind': '禁用词', 'text': '%s（%d 处）' % (pat, n)})

    old = C.scan_old_standards(full)
    for s, n, y, ctx in old:
        issues['A'].append({'kind': '疑似废止规范', 'text': '%s %s-%s（…%s…）' % (s, n, y, ctx[:50])})

    hard_rules = C.parse_hard_rules(hard)
    missing_hard = []
    for k, needle, min_n in hard_rules:
        n = full.count(needle)
        if n < min_n:
            if min_n > 1:
                missing_hard.append('%s（需 ≥%d 处，实际 %d 处）' % (k, min_n, n))
            else:
                missing_hard.append(k)
    for k in missing_hard:
        issues['A'].append({'kind': '硬指标缺失', 'text': k})

    heads, struct_a, struct_b = C.check_heading_structure(full)
    for x in struct_a:
        issues['A'].append({'kind': '结构断裂', 'text': x})
    for x in struct_b:
        issues['B'].append({'kind': '编号不衔接', 'text': x})

    # ---- B 类：应当修正
    for x in C.check_numbering(full):
        issues['B'].append({'kind': '编号问题', 'text': x})
    for kind, ch, seq in C.find_dangling_refs(full):
        issues['B'].append({'kind': '悬空引用',
                            'text': '%s %d-%d 被引用但无题注' % (kind, ch, seq)})
    unnamed = C.find_unnamed_tables(full)
    for row in unnamed:
        issues['B'].append({'kind': '缺表名表格', 'text': 'L%s %s' % (row[0], row[1])})
    for x in C.check_format(full):
        issues['B'].append({'kind': '格式残留', 'text': x})
    notes = C.scan_notes(full)
    if len(notes) > notes_limit:
        issues['B'].append({'kind': '初稿注超限',
                            'text': '当前 %d 处，上限 %d 处' % (len(notes), notes_limit)})

    # ---- C 类：AI 痕迹
    ai = ai_trace.full_report(full)
    for lvl, txt in ai['issues']:
        issues['C' if lvl == 'warn' else 'A'].append({'kind': 'AI痕迹', 'text': txt})

    # ---- D 类：待澄清
    for line, t, tail in notes:
        issues['D'].append({'kind': '初稿注', 'text': 'L%d %s' % (line, t)})

    # ---- E 类：提示
    issues['E'].append({'kind': '篇幅',
                        'text': '全文 %d 字，标题 %d 个，表格 %d 个'
                                % (len(full), len(heads), len(C.parse_tables(full)))})
    if not missing_hard and not banned_hits and not old:
        issues['E'].append({'kind': '红线', 'text': '未发现禁用词、废止规范与硬指标缺失'})

    return issues, ai, heads, notes


def fingerprint(issues):
    """把问题列表转成可比较的指纹集合。"""
    fp = set()
    for cat, items in issues.items():
        for it in items:
            fp.add('%s|%s|%s' % (cat, it['kind'], it['text']))
    return fp


def compare_baseline(curr_fp, prev_fp):
    """对比基线：新增 / 已修复 / 遗留。"""
    new = sorted(curr_fp - prev_fp)
    fixed = sorted(prev_fp - curr_fp)
    kept = sorted(curr_fp & prev_fp)
    return {'new': new, 'fixed': fixed, 'kept': kept}


def md_report(issues, ai, heads, notes, meta, baseline=None):
    """生成 Markdown 体检报告。"""
    L = []
    L.append('# 合规体检报告')
    L.append('')
    L.append('| 项目 | 内容 |')
    L.append('|---|---|')
    L.append('| 检查对象 | %s |' % meta['input'])
    L.append('| 文本类型 | %s |' % meta['format'])
    L.append('| 全文 %s | %d 字 |' % ('字符数', meta['chars']))
    L.append('| 标题数 | %d |' % meta['heading_count'])
    L.append('| 表格数 | %d |' % meta['table_count'])
    L.append('| 初稿注 | %d 处（上限 %d） |' % (meta['notes_count'], meta['notes_limit']))
    L.append('| AI 痕迹评分 | %d/100（%s） |' % (ai['score'], ai['level']))
    L.append('')

    total = sum(len(v) for v in issues.values())
    L.append('## 结论')
    L.append('')
    L.append('**共 %d 项问题**：A 类（必须修正）%d 项、B 类（应当修正）%d 项、'
             'C 类（建议优化）%d 项、D 类（待澄清）%d 项、E 类（提示）%d 项。'
             % (total, len(issues['A']), len(issues['B']), len(issues['C']),
                len(issues['D']), len(issues['E'])))
    L.append('')
    if issues['A']:
        L.append('> ⚠️ 存在 A 类红线问题，**不得直接提交**，须全部修正后复检。')
    else:
        L.append('> ✅ 未发现 A 类红线问题。')
    L.append('')

    names = {'A': 'A 类 · 必须修正（红线）', 'B': 'B 类 · 应当修正',
             'C': 'C 类 · 建议优化', 'D': 'D 类 · 待澄清', 'E': 'E 类 · 提示'}
    for cat in ('A', 'B', 'C', 'D', 'E'):
        L.append('## %s' % names[cat])
        L.append('')
        if not issues[cat]:
            L.append('（无）')
            L.append('')
            continue
        L.append('| 类型 | 详情 |')
        L.append('|---|---|')
        for it in issues[cat]:
            L.append('| %s | %s |' % (it['kind'], it['text'].replace('|', '\\|')))
        L.append('')

    L.append('## AI 痕迹明细')
    L.append('')
    d = ai['detail']
    L.append('- 高置信套话：%s' % (json.dumps(d['ai_phrases']['high'], ensure_ascii=False) or '无'))
    L.append('- 中置信连接词：%s' % (json.dumps(d['ai_phrases']['mid'], ensure_ascii=False) or '无'))
    L.append('- 空话套话：%s' % (json.dumps(d['ai_phrases']['cliche'], ensure_ascii=False) or '无'))
    L.append('- 段落长度：%s' % d['paragraph']['detail'])
    L.append('- 句子长度：%s' % d['sentence']['detail'])
    L.append('- 标点习惯：%s' % d['punctuation']['detail'])
    L.append('- 标题对仗：%s' % d['heading']['detail'])
    L.append('- 句首重复：%s' % d['sentence_start']['detail'])
    L.append('')

    if baseline:
        L.append('## 与基线对比')
        L.append('')
        L.append('| 变化 | 数量 |')
        L.append('|---|---|')
        L.append('| 新增问题 | %d |' % len(baseline['new']))
        L.append('| 已修复 | %d |' % len(baseline['fixed']))
        L.append('| 遗留 | %d |' % len(baseline['kept']))
        L.append('')
        if baseline['new']:
            L.append('### 新增问题')
            L.append('')
            for x in baseline['new']:
                L.append('- %s' % x)
            L.append('')
        if baseline['fixed']:
            L.append('### 已修复')
            L.append('')
            for x in baseline['fixed']:
                L.append('- %s' % x)
            L.append('')
        if baseline['kept']:
            L.append('### 遗留问题')
            L.append('')
            for x in baseline['kept']:
                L.append('- %s' % x)
            L.append('')

    L.append('## 处理建议')
    L.append('')
    L.append('1. A 类问题逐项修正，修正后重跑本脚本复检；')
    L.append('2. B 类问题按清单处理，缺表名表格需补表名或确认删除；')
    L.append('3. C 类问题参照 references/AI痕迹识别与消解.md 逐条消解；')
    L.append('4. D 类问题待勘察/设计/业主资料到位后统一替换，替换后重跑；')
    L.append('5. 每次体检后保留本次 JSON 作为下次的 --prev 基线。')
    L.append('')
    return '\n'.join(L)


def main():
    ap = argparse.ArgumentParser(description='合规体检（只读）')
    ap.add_argument('--in', dest='inp', required=True, help='待检文件（md 或 docx）')
    ap.add_argument('--banned', help='禁用词文件（每行一条正则）')
    ap.add_argument('--hard', help='硬指标文件（每行「指标名: 关键字符串」）')
    ap.add_argument('--notes-limit', type=int, default=10, help='初稿注上限（默认 10）')
    ap.add_argument('--prev', help='上次体检 JSON，用于基线对比')
    ap.add_argument('--report', help='体检报告 JSON 输出路径')
    ap.add_argument('--md', dest='md_out', help='Markdown 体检报告输出路径')
    ap.add_argument('--quiet', action='store_true', help='只输出摘要 JSON')
    args = ap.parse_args()

    full, fmt = load_text(args.inp)
    banned = C.read_lines(args.banned)
    hard = C.read_pairs(args.hard)

    issues, ai, heads, notes = classify(full, banned, hard, args.notes_limit)
    curr_fp = fingerprint(issues)

    baseline = None
    if args.prev and os.path.exists(args.prev):
        prev = C.load_json(args.prev, {})
        prev_fp = set(prev.get('fingerprint', []))
        baseline = compare_baseline(curr_fp, prev_fp)

    meta = {
        'input': os.path.abspath(args.inp), 'format': fmt, 'chars': len(full),
        'heading_count': len(heads), 'table_count': len(C.parse_tables(full)),
        'notes_count': len(notes), 'notes_limit': args.notes_limit,
    }
    report = {
        'meta': meta, 'issues': issues, 'ai': ai,
        'counts': {k: len(v) for k, v in issues.items()},
        'fingerprint': sorted(curr_fp),
        'baseline': baseline,
        'status': 'fail' if issues['A'] else ('warn' if issues['B'] or issues['C'] else 'pass'),
    }

    if args.report:
        C.write_text(args.report, json.dumps(report, ensure_ascii=False, indent=2))
    if args.md_out:
        C.write_text(args.md_out, md_report(issues, ai, heads, notes, meta, baseline))

    if not args.quiet:
        print('检查对象: %s（%s，%d 字）' % (meta['input'], fmt, meta['chars']))
        print('状态: %s' % report['status'])
        print('A 类 %d / B 类 %d / C 类 %d / D 类 %d / E 类 %d'
              % (report['counts']['A'], report['counts']['B'], report['counts']['C'],
                 report['counts']['D'], report['counts']['E']))
        print('AI 痕迹: %s（评分 %d/100）' % (ai['level'], ai['score']))
        if issues['A']:
            print('A 类问题:')
            for it in issues['A']:
                print('  - [%s] %s' % (it['kind'], it['text']))
        if issues['B']:
            print('B 类问题（前 10）:')
            for it in issues['B'][:10]:
                print('  - [%s] %s' % (it['kind'], it['text']))
        if baseline:
            print('与基线对比: 新增 %d / 已修复 %d / 遗留 %d'
                  % (len(baseline['new']), len(baseline['fixed']), len(baseline['kept'])))
        if args.report:
            print('JSON 报告: %s' % os.path.abspath(args.report))
        if args.md_out:
            print('Markdown 报告: %s' % os.path.abspath(args.md_out))
    else:
        print(json.dumps({'status': report['status'], 'counts': report['counts'],
                          'ai_level': ai['level'], 'ai_score': ai['score'],
                          'baseline': ({'new': len(baseline['new']),
                                        'fixed': len(baseline['fixed']),
                                        'kept': len(baseline['kept'])} if baseline else None)},
                         ensure_ascii=False))
    raise SystemExit(0 if report['status'] == 'pass' else 1)


if __name__ == '__main__':
    main()

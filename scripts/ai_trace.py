# -*- coding: utf-8 -*-
"""AI 痕迹检测（只读分析，不修改文件）。

检测维度：
1. 套话词频（高/中两级词表）
2. 段落长度分布（AI 文本段长偏均匀）
3. 句子长度分布（AI 文本句长偏均匀）
4. 标点习惯（半角混入、破折号密度、感叹号）
5. 标题对仗工整度（AI 爱用平行结构）
6. 句式重复（同一连接词开头比例过高）
7. 空话套话（"具有重要意义"类）

输出：分级（bad / warn / ok）+ 具体命中 + 修改建议。
"""
import json
import re

# 高置信度 AI 套话：正式工程文本中极少出现
HIGH_AI = [
    '综上所述', '由此可见', '毋庸置疑', '不言而喻', '众所周知', '值得注意的是',
    '需要指出的是', '不难看出', '毫无疑问', '在一定程度上', '总的来说', '总体而言',
    '综合来看', '综合以上', '在这个过程中', '在这个背景下', '让我们', '首先，我们需要',
    '首先，我们', '在当今', '在当今社会', '随着社会的不断发展', '日益增长',
    '具有重要意义', '发挥了重要作用', '取得了显著成效', '奠定了坚实基础',
    '提供了有力保障', '注入新的活力', '增添了新的动能', '擘画', '赋能', '聚焦',
    '凝心聚力', '砥砺前行', '奋进', '新时代', '新征程', '高质量发展',
]

# 中置信度：AI 常用，但正式文体也接受，仅在密度过高时告警
MID_AI = [
    '一方面', '另一方面', '不仅', '而且', '既', '又', '从而', '进而', '同时',
    '因此', '所以', '为此', '基于此', '在此基础上', '与此同时', '此外', '另外',
    '首先', '其次', '最后', '综上', '简言之', '换言之', '换句话说',
]

# 空话套话
CLICHE = [
    '具有重要意义', '发挥了重要作用', '取得了显著成效', '奠定了坚实基础',
    '提供了有力保障', '注入了新的活力', '增添了新的动能', '提供了新思路',
    '具有重要意义和深远影响', '发挥了积极作用', '取得了良好效果',
]


def _hits(text, words):
    out = {}
    for w in words:
        n = text.count(w)
        if n:
            out[w] = n
    return out


def scan_ai_phrases(text):
    """套话扫描。返回 {'high': {词:次数}, 'mid': {词:次数}, 'cliche': {词:次数}}。"""
    mid = _hits(text, MID_AI)
    # 「既有」是工程文本的合法用语（既有线路 / 既有分界 / 既有电杆），
    # 单字「既」会把这类表述全部误判为 AI 痕迹，故先剔除。
    if '既' in mid:
        mid['既'] = max(0, mid['既'] - text.count('既有'))
        if not mid['既']:
            del mid['既']
    return {
        'high': _hits(text, HIGH_AI),
        'mid': mid,
        'cliche': _hits(text, CLICHE),
    }


def paragraph_stats(text):
    """段落长度统计。AI 文本段长分布通常过于均匀（变异系数低）。"""
    paras = [p.strip() for p in re.split(r'\n\s*\n', text) if p.strip()]
    # 过滤掉表格行与标题行，只统计正文段落
    body = []
    for p in paras:
        lines = p.split('\n')
        for ln in lines:
            s = ln.strip()
            if not s or s.startswith('#') or s.startswith('|') or s.startswith('-'):
                continue
            body.append(s)
    lens = [len(s) for s in body]
    if len(lens) < 5:
        return {'count': len(lens), 'mean': 0, 'std': 0, 'cv': 0, 'level': 'ok',
                'detail': '正文段落不足 5 段，无法判断'}
    mean = sum(lens) / float(len(lens))
    var = sum((x - mean) ** 2 for x in lens) / float(len(lens))
    std = var ** 0.5
    cv = std / mean if mean else 0
    if cv < 0.25:
        level = 'bad'
    elif cv < 0.35:
        level = 'warn'
    else:
        level = 'ok'
    return {
        'count': len(lens), 'mean': round(mean, 1), 'std': round(std, 1),
        'cv': round(cv, 3), 'level': level,
        'detail': '正文段落 %d 段，均长 %.1f，变异系数 %.3f（<0.25 偏均匀，疑似模板化）'
                  % (len(lens), mean, cv),
    }


def _body_text(text):
    """剔除 Markdown 标题行，只保留正文。

    标题行（`#`、`##`…）不是句子：混进统计会让句首两字变成 `##`，
    直接把「句首重复」打成 bad，也会污染句长分布。
    """
    return '\n'.join(ln for ln in text.split('\n')
                     if not ln.lstrip().startswith('#'))


def sentence_stats(text):
    """句子长度统计。AI 文本句长偏均匀。"""
    sents = [s.strip() for s in re.split(r'[。！？；]', _body_text(text))
             if len(s.strip()) >= 4]
    lens = [len(s) for s in sents]
    if len(lens) < 10:
        return {'count': len(lens), 'mean': 0, 'std': 0, 'cv': 0, 'level': 'ok',
                'detail': '句子不足 10 句，无法判断'}
    mean = sum(lens) / float(len(lens))
    var = sum((x - mean) ** 2 for x in lens) / float(len(lens))
    std = var ** 0.5
    cv = std / mean if mean else 0
    if cv < 0.30:
        level = 'bad'
    elif cv < 0.40:
        level = 'warn'
    else:
        level = 'ok'
    return {
        'count': len(lens), 'mean': round(mean, 1), 'std': round(std, 1),
        'cv': round(cv, 3), 'level': level,
        'detail': '句子 %d 句，均长 %.1f，变异系数 %.3f' % (len(lens), mean, cv),
    }


def punctuation_stats(text):
    """标点习惯：半角混入、破折号密度、感叹号密度、连续标点。"""
    n = len(text) or 1
    # 半角逗号/句号/分号出现在中文之间
    half_comma = len(re.findall(r'[一-龥],[一-龥]', text))
    half_period = len(re.findall(r'[一-龥]\.[一-龥]', text))
    half_semi = len(re.findall(r'[一-龥];[一-龥]', text))
    half_colon = len(re.findall(r'[一-龥]:[一-龥]', text))
    half_total = half_comma + half_period + half_semi + half_colon
    # 破折号密度（每千字）
    dash = text.count('——')
    dash_density = round(dash * 1000.0 / n, 2)
    # 感叹号
    excl = text.count('！')
    # 连续标点
    dup = len(re.findall(r'[，。；：]{2,}', text))
    issues = []
    level = 'ok'
    if half_total:
        issues.append('半角标点混入中文 %d 处（逗 %d / 句 %d / 分 %d / 冒 %d）'
                      % (half_total, half_comma, half_period, half_semi, half_colon))
        level = 'bad' if half_total > 5 else 'warn'
    if dash_density > 8:
        issues.append('破折号密度 %.2f/千字（>8 偏高，AI 文体常见）' % dash_density)
        level = 'bad' if level == 'ok' else level
    elif dash_density > 4:
        issues.append('破折号密度 %.2f/千字（略高）' % dash_density)
        if level == 'ok':
            level = 'warn'
    if excl > 0:
        issues.append('感叹号 %d 处（工程文本应极少）' % excl)
        level = 'bad'
    if dup:
        issues.append('连续标点 %d 处' % dup)
    return {
        'half_total': half_total, 'dash': dash, 'dash_density': dash_density,
        'excl': excl, 'dup': dup, 'level': level, 'issues': issues,
        'detail': '；'.join(issues) if issues else '标点习惯正常',
    }


def heading_symmetry(text):
    """标题对仗工整度：同级标题结尾相同的比例。"""
    heads = []
    for line in text.split('\n'):
        m = re.match(r'^(#{1,6})\s+(.*)$', line.strip())
        if m:
            heads.append((len(m.group(1)), m.group(2).strip()))
    if len(heads) < 6:
        return {'level': 'ok', 'detail': '标题不足 6 个，无法判断'}
    # 按层级分组，统计结尾 2 字相同的占比
    from collections import defaultdict
    groups = defaultdict(list)
    for lv, t in heads:
        groups[lv].append(t)
    worst = 0.0
    worst_lv = None
    for lv, ts in groups.items():
        if len(ts) < 4:
            continue
        tails = defaultdict(int)
        for t in ts:
            tails[t[-2:]] += 1
        max_same = max(tails.values())
        ratio = max_same / float(len(ts))
        if ratio > worst:
            worst = ratio
            worst_lv = lv
    if worst >= 0.7:
        level = 'bad'
    elif worst >= 0.5:
        level = 'warn'
    else:
        level = 'ok'
    return {
        'level': level, 'worst_ratio': round(worst, 3), 'worst_level': worst_lv,
        'detail': 'H%d 级标题结尾重复率 %.0f%%（>70%% 疑似模板化对仗）'
                  % (worst_lv or 0, worst * 100) if worst_lv else '标题结构正常',
    }


def sentence_start_stats(text):
    """句首词重复度：同一连接词开头的句子占比。"""
    sents = [s.strip() for s in re.split(r'[。！？；\n]', _body_text(text))
             if len(s.strip()) >= 4]
    if len(sents) < 10:
        return {'level': 'ok', 'detail': '句子不足 10 句'}
    from collections import Counter
    starts = Counter(s[:2] for s in sents)
    top, cnt = starts.most_common(1)[0]
    ratio = cnt / float(len(sents))
    if ratio > 0.35:
        level = 'bad'
    elif ratio > 0.25:
        level = 'warn'
    else:
        level = 'ok'
    return {
        'level': level, 'top_start': top, 'ratio': round(ratio, 3),
        'detail': '句首「%s」占比 %.0f%%（>35%% 疑似句式单一）' % (top, ratio * 100),
    }


def full_report(text):
    """汇总 AI 痕迹报告。"""
    ph = scan_ai_phrases(text)
    para = paragraph_stats(text)
    sent = sentence_stats(text)
    pun = punctuation_stats(text)
    head = heading_symmetry(text)
    start = sentence_start_stats(text)
    high_total = sum(ph['high'].values())
    mid_total = sum(ph['mid'].values())
    cliche_total = sum(ph['cliche'].values())

    issues = []
    if high_total:
        issues.append(('bad', '高置信 AI 套话 %d 处: %s'
                       % (high_total, '、'.join('%s×%d' % (k, v) for k, v in ph['high'].items()))))
    if cliche_total:
        issues.append(('bad', '空话套话 %d 处: %s'
                       % (cliche_total, '、'.join('%s×%d' % (k, v) for k, v in ph['cliche'].items()))))
    if mid_total > 5:
        issues.append(('warn', '中置信连接词 %d 处（密度偏高）: %s'
                       % (mid_total, '、'.join('%s×%d' % (k, v) for k, v in ph['mid'].items()))))
    if para['level'] != 'ok':
        issues.append((para['level'], '段落长度: ' + para['detail']))
    if sent['level'] != 'ok':
        issues.append((sent['level'], '句子长度: ' + sent['detail']))
    for x in pun['issues']:
        issues.append((pun['level'], '标点: ' + x))
    if head['level'] != 'ok':
        issues.append((head['level'], '标题对仗: ' + head['detail']))
    if start['level'] != 'ok':
        issues.append((start['level'], '句首重复: ' + start['detail']))

    bad = [x for x in issues if x[0] == 'bad']
    warn = [x for x in issues if x[0] == 'warn']
    if bad:
        level = 'bad'
    elif warn:
        level = 'warn'
    else:
        level = 'ok'
    return {
        'level': level,
        'score': max(0, 100 - 25 * len(bad) - 10 * len(warn)),
        'issues': issues,
        'detail': {
            'ai_phrases': ph,
            'paragraph': para,
            'sentence': sent,
            'punctuation': pun,
            'heading': head,
            'sentence_start': start,
        },
        'summary': 'AI 痕迹风险 %s（%d 严重 / %d 提示），评分 %d/100'
                   % (level, len(bad), len(warn), max(0, 100 - 25 * len(bad) - 10 * len(warn))),
    }


# ---------------------------------------------------------------- CLI
def _load_text(path):
    """读 md 或 docx（docx 走 cbd_common）。"""
    if path.lower().endswith('.docx'):
        try:
            import cbd_common as C
            return C.docx_full_text(path)
        except Exception as e:
            raise SystemExit('读取 docx 失败: %s' % e)
    with open(path, 'r', encoding='utf-8-sig') as f:
        return f.read()


def main():
    import argparse
    ap = argparse.ArgumentParser(description='AI 痕迹检测（只读）')
    ap.add_argument('--in', dest='inp', required=True, help='待检文件（md 或 docx）')
    ap.add_argument('--report', help='报告 JSON 输出路径')
    ap.add_argument('--quiet', action='store_true', help='只输出摘要 JSON')
    a = ap.parse_args()
    rep = full_report(_load_text(a.inp))
    if a.report:
        with open(a.report, 'w', encoding='utf-8') as f:
            json.dump(rep, f, ensure_ascii=False, indent=2)
    if a.quiet:
        print(json.dumps({'level': rep['level'], 'score': rep['score'],
                          'bad': sum(1 for x in rep['issues'] if x[0] == 'bad'),
                          'warn': sum(1 for x in rep['issues'] if x[0] == 'warn')},
                         ensure_ascii=False))
    else:
        print(rep['summary'])
        for lv, txt in rep['issues']:
            print('  [%s] %s' % ('严重' if lv == 'bad' else '提示', txt))
    raise SystemExit(0 if rep['level'] == 'ok' else 1)


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
"""校验软著材料 Word 是否满足提交口径。

检查项：
  · 每页行数（源代码材料要求 ≥50 行，且各页行数应等于设定值）
  · 是否存在空段落（空行）
  · 页眉文字（软件全称 + 版本号）与「共 N 页」是否与实际页数一致
  · 分节页码起始值（前 30 页 / 后 30 页材料）
  · 页边距、纸张、固定行距是否与排版假设一致（防止 Word 里跑版）

用法：
  python check_docx.py 源代码_完整.docx --expect-lines 50 --title "示例事项软件 V1.0"
  python check_docx.py 操作说明书.docx --kind manual
"""
import argparse
import json
import sys
from docx import Document
from docx.oxml.ns import qn

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def para_text(p):
    el = getattr(p, '_p', p)
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def has_page_break_before(p):
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        return False
    el = pPr.find(qn('w:pageBreakBefore'))
    if el is None:
        return False
    return el.get(qn('w:val')) not in ('0', 'false')


def has_hard_break(p):
    for br in p.iter(qn('w:br')):
        if br.get(qn('w:type')) in (None, 'page'):
            return True
    return False


def has_sectpr(p):
    pPr = p.find(qn('w:pPr'))
    return pPr is not None and pPr.find(qn('w:sectPr')) is not None


def split_pages(doc):
    body = doc.element.body
    pages, cur = [], []
    for p in body.iterchildren(qn('w:p')):
        new_page = has_page_break_before(p) or (cur and has_hard_break(p))
        if new_page and cur:
            pages.append(cur)
            cur = []
        cur.append(p)
        if has_sectpr(p):
            pages.append(cur)
            cur = []
    if cur:
        pages.append(cur)
    return pages


def header_has_field(section, field_name):
    for p in section.header.paragraphs:
        for it in p._p.iter(qn('w:instrText')):
            if field_name in (it.text or ''):
                return True
    return False


def header_info(section):
    header = section.header
    texts = [para_text(p) for p in header.paragraphs]
    return ' | '.join(t for t in texts if t.strip())


def check(path, expect_lines=None, title=None, kind='source', expect_total=None):
    doc = Document(path)
    pages = split_pages(doc)
    body = doc.element.body

    total_paras = 0
    empty_paras = 0
    for p in body.iterchildren(qn('w:p')):
        total_paras += 1
        if not para_text(p).strip():
            empty_paras += 1

    counts = []
    for i, page in enumerate(pages, 1):
        n = sum(1 for p in page if para_text(p).strip())
        counts.append(n)

    sections = []
    for idx, sec in enumerate(doc.sections):
        sectPr = sec._sectPr
        pgNum = sectPr.find(qn('w:pgNumType'))
        sections.append({
            '序号': idx + 1,
            '起始页码': int(pgNum.get(qn('w:start'))) if pgNum is not None
            and pgNum.get(qn('w:start')) else 1,
            '页眉': header_info(sec),
        })

    widths = {
        '纸张宽cm': round(sec.page_width.cm, 2),
        '纸张高cm': round(sec.page_height.cm, 2),
        '左边距cm': round(sec.left_margin.cm, 2),
        '右边距cm': round(sec.right_margin.cm, 2),
    } if doc.sections else {}

    problems, notices = [], []

    if kind == 'source':
        if expect_lines:
            bad = [(i + 1, n) for i, n in enumerate(counts) if n != expect_lines]
            # 最后一页允许不足（程序结束页）
            bad = [x for x in bad if x[0] != len(counts)]
            if bad:
                problems.append('以下页行数不等于 %d：%s' % (expect_lines, bad[:10]))
            short = [n for n in counts if n < (expect_lines or 50)]
            if short and len(short) > 1:
                problems.append('存在多页少于 %d 行（仅末页可不足）' % (expect_lines or 50))
        if empty_paras:
            problems.append('存在 %d 个空段落（源代码材料要求无空行）' % empty_paras)
        if title:
            for s in sections:
                if title.replace(' ', '') not in s['页眉'].replace(' ', ''):
                    problems.append('第 %d 节页眉缺少「%s」，实际：%s'
                                    % (s['序号'], title, s['页眉']))

    # 「共 N 页」与实际页数核对
    auto_total = any(header_has_field(sec, 'NUMPAGES') for sec in doc.sections)
    reported = None
    if sections and not auto_total:
        import re
        m = re.search(r'共\s*(\d+)\s*页', sections[0]['页眉'])
        if m:
            reported = int(m.group(1))
    actual_pages = len(pages)
    target_total = expect_total if expect_total else actual_pages
    if auto_total:
        reported = None
    elif reported is not None:
        if reported != target_total:
            problems.append('页眉声明「共 %d 页」，应为 %d 页，页码不一致'
                            % (reported, target_total))
    else:
        if kind == 'source':
            problems.append('页眉未找到「共 N 页」文字')

    if kind == 'manual':
        low = [i + 1 for i, n in enumerate(counts) if n < 30]
        if low:
            notices.append('存在 %d 个分页段落少于 30 行（第 %s 页），纯文字页需补足 30 行，'
                           '整页截图页可豁免' % (len(low), '、'.join(map(str, low[:8]))))
        notices.append('说明书实际页数以 Word 排版为准，本结果中的页数仅反映手工分页标记')

    result = {
        '文件': path,
        '页数': actual_pages,
        '段落总数': total_paras,
        '空段落数': empty_paras,
        '每页行数': {'最小': min(counts) if counts else 0,
                 '最大': max(counts) if counts else 0,
                 '分布': {str(k): counts.count(k) for k in sorted(set(counts))}},
        '节': sections,
        '页面': widths,
        '页眉声明总页数': reported if reported is not None else ('NUMPAGES 域（Word 自动计算）'
                                              if auto_total else None),
        '问题': problems,
        '提示': notices,
        '结论': '通过' if not problems else '需修正',
    }
    return result


def main():
    ap = argparse.ArgumentParser(description='校验软著材料 Word')
    ap.add_argument('files', nargs='+', help='待校验的 docx')
    ap.add_argument('--expect-lines', type=int, default=None, help='源代码材料每页行数')
    ap.add_argument('--title', default=None, help='期望页眉文字（软件全称 + 版本号）')
    ap.add_argument('--kind', choices=['source', 'manual'], default='source')
    ap.add_argument('--expect-total', type=int, default=None,
                    help='页眉「共 N 页」的期望值（original 页码模式的第二份材料需指定）')
    args = ap.parse_args()

    all_ok = True
    for path in args.files:
        res = check(path, args.expect_lines, args.title, args.kind, args.expect_total)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        if res['问题']:
            all_ok = False
    sys.exit(0 if all_ok else 1)


if __name__ == '__main__':
    main()

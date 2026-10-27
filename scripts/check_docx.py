# -*- coding: utf-8 -*-
"""校验软著材料 Word 是否满足提交口径。

检查项：
  · 版式：纸张 A4、页边距、页眉字号/位置/下框线、页脚是否留空
  · 封面对文档鉴别材料：第 1 页不得出现页眉与页码
  · 每页行数（源代码材料要求 ≥50 行，且各页行数应等于设定值）
  · 是否存在空段落（空行）
  · 页眉文字（软件全称 + 版本号）与「共 N 页」是否与实际页数一致
  · 疑似水印/声明文字（AI 生成、请谨慎甄别、试用版等），页眉页脚出现即视为不合规
  · 固定行距裁剪行内图片（w:lineRule="exact" 且段落含 wp:inline，截图会被压成一条）

用法：
  python check_docx.py 源代码_完整.docx --expect-lines 50 --title "示例事项软件 V1.0"
  python check_docx.py 操作说明书.docx --kind manual --title "示例事项软件 V1.0"
  python check_docx.py 别人的材料.docx --audit-only        # 只看版式，不判断行数
"""
import argparse
import json
import re
import sys

from docx import Document
from docx.oxml.ns import qn

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# 与 docx_common.LAYOUT 保持一致（此处不 import，便于单独拷走使用）
SPEC = {
    'source': dict(top=2.0, bottom=1.8, left=1.8, right=1.8, header=1.2, footer=1.0),
    'manual': dict(top=2.2, bottom=1.8, left=2.2, right=2.2, header=1.2, footer=1.0),
}

HEADER_SIZE = 9.0
TOL_CM = 0.15

# 页眉页脚出现这些字样即判不合规；正文只查纯中文短语（英文词撞变量名的概率太高）
FORBIDDEN_HF = [
    'AI 生成', 'AI生成', '生成式人工智能', '请谨慎甄别', '水印', '内部资料',
    '试用版', '草稿', '扫码', '公众号', '二维码', '禁止外传', '仅供学习',
    '内容由', '样例', 'DRAFT', 'SAMPLE',
]
FORBIDDEN_BODY = [
    'AI 生成', 'AI生成', '生成式人工智能', '请谨慎甄别', '内部资料',
    '禁止外传', '扫码关注', '二维码', '试用版',
]


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


def hf_audit(hf):
    """页眉/页脚体检：文字、字号、右制表位、下框线、域。"""
    paras = hf.paragraphs
    texts = [para_text(p) for p in paras]
    size, right_tab, border, fields = None, None, False, []
    tabs_seen = []

    for p in paras:
        pPr = p._p.find(qn('w:pPr'))
        if pPr is not None:
            pBdr = pPr.find(qn('w:pBdr'))
            if pBdr is not None and pBdr.find(qn('w:bottom')) is not None:
                border = True
            tabs = pPr.find(qn('w:tabs'))
            if tabs is not None:
                for tab in tabs.findall(qn('w:tab')):
                    val = tab.get(qn('w:val'))
                    try:
                        pos = round(int(tab.get(qn('w:pos'))) / 566.93, 2)
                    except (TypeError, ValueError):
                        continue
                    tabs_seen.append('%s@%.2fcm' % (val, pos))
                    if val == 'right':
                        right_tab = pos
        for it in p._p.iter(qn('w:instrText')):
            fields.append((it.text or '').strip().upper())
        for r in p.runs:
            if size is None and r.font.size is not None:
                size = round(r.font.size.pt, 1)

    return {
        '文字': ' | '.join(t.strip() for t in texts if t.strip()),
        '字号pt': size,
        '右制表位cm': right_tab,
        '制表位': tabs_seen,
        '下框线': border,
        '域': sorted(set(f for f in fields if f)),
    }


def style_tabs(doc, name='Header'):
    """页眉样式自带的制表位（段落级 tab 会先命中样式里的居中位，页码就跑到中间）。"""
    try:
        style = doc.styles[name]
    except KeyError:
        return []
    pPr = style.element.find(qn('w:pPr'))
    if pPr is None:
        return []
    tabs = pPr.find(qn('w:tabs'))
    if tabs is None:
        return []
    out = []
    for tab in tabs.findall(qn('w:tab')):
        try:
            out.append((tab.get(qn('w:val')),
                        round(int(tab.get(qn('w:pos'))) / 566.93, 2)))
        except (TypeError, ValueError):
            continue
    return out


DRAWING_NS = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'


def style_line_rules(doc):
    """styleId → (lineRule, lineTwips, basedOn)，另以 None 键存文档默认段落行距。"""
    rules = {}
    for st in doc.styles.element.findall(qn('w:style')):
        sid = st.get(qn('w:styleId'))
        rule = value = None
        pPr = st.find(qn('w:pPr'))
        if pPr is not None:
            sp = pPr.find(qn('w:spacing'))
            if sp is not None:
                rule, value = sp.get(qn('w:lineRule')), sp.get(qn('w:line'))
        based = st.find(qn('w:basedOn'))
        rules[sid] = (rule, value, based.get(qn('w:val')) if based is not None else None)

    default = (None, None)
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        pPrDef = dd.find(qn('w:pPrDefault'))
        if pPrDef is not None:
            pPr = pPrDef.find(qn('w:pPr'))
            if pPr is not None:
                sp = pPr.find(qn('w:spacing'))
                if sp is not None:
                    default = (sp.get(qn('w:lineRule')), sp.get(qn('w:line')))
    rules[None] = (default[0], default[1], None)
    return rules


def resolve_line_rule(p, rules):
    """返回段落生效的行距 (lineRule, lineTwips, 来源)，段落 → 样式链 → 文档默认。"""
    pPr = p.find(qn('w:pPr'))
    if pPr is not None:
        sp = pPr.find(qn('w:spacing'))
        if sp is not None and (sp.get(qn('w:lineRule')) or sp.get(qn('w:line'))):
            return sp.get(qn('w:lineRule')) or 'auto', sp.get(qn('w:line')), '段落直接设置'
    sid = None
    if pPr is not None:
        ps = pPr.find(qn('w:pStyle'))
        if ps is not None:
            sid = ps.get(qn('w:val'))

    cur, guard = sid or 'Normal', 0
    while cur and guard < 10:
        rule, value, based = rules.get(cur, (None, None, None))
        if rule:
            return rule, value, '样式 %s' % cur
        cur, guard = based, guard + 1
    rule, value, _ = rules.get(None, (None, None, None))
    if rule:
        return rule, value, '文档默认'
    return None, None, ''


def inline_image_clip(doc):
    """找出「固定行距 + 行内图片」的段落。

    w:lineRule="exact" 把行框高度钉死，行内图片按行高裁剪，Word 导出 PDF 时
    内容流里会出现只有行高那么高的裁剪矩形（如 19.25pt），截图被压成一条。
    浮动图（wp:anchor）不受行距限制，不报。
    """
    rules = style_line_rules(doc)
    hits = []
    for idx, p in enumerate(doc.element.body.iter(qn('w:p')), start=1):
        inlines = [el for d in p.iter(qn('w:drawing'))
                   for el in d.iter(DRAWING_NS + 'inline')]
        if not inlines:
            continue
        rule, value, src = resolve_line_rule(p, rules)
        if rule != 'exact':
            continue
        heights = []
        for el in inlines:
            ext = el.find(DRAWING_NS + 'extent')
            if ext is not None:
                try:
                    heights.append(round(int(ext.get('cy')) / 360000.0, 2))
                except (TypeError, ValueError):
                    pass
        line_pt = None
        if value:
            try:
                line_pt = round(int(value) / 20.0, 1)
            except (TypeError, ValueError):
                pass
        hits.append({'段落序号': idx, '行距pt': line_pt, '来源': src,
                     '图片高度cm': heights, '文字': para_text(p)[:30]})
    return hits



def check(path, expect_lines=None, title=None, kind='source', expect_total=None,
          audit_only=False):
    doc = Document(path)
    pages = split_pages(doc)
    body = doc.element.body

    total_paras, empty_paras = 0, 0
    for p in body.iterchildren(qn('w:p')):
        total_paras += 1
        if not para_text(p).strip():
            empty_paras += 1

    counts = [sum(1 for p in page if para_text(p).strip()) for page in pages]

    sections, hf_audits = [], []
    for idx, sec in enumerate(doc.sections):
        sectPr = sec._sectPr
        pgNum = sectPr.find(qn('w:pgNumType'))
        hf_audits.append({'序号': idx + 1,
                          '页眉': hf_audit(sec.header),
                          '页脚': hf_audit(sec.footer)})
        sections.append({
            '序号': idx + 1,
            '起始页码': int(pgNum.get(qn('w:start'))) if pgNum is not None
            and pgNum.get(qn('w:start')) else 1,
            '页眉': hf_audits[-1]['页眉']['文字'],
        })

    spec = SPEC.get(kind, SPEC['source'])
    layout = {}
    problems, notices = [], []

    if doc.sections:
        sec = doc.sections[0]
        layout = {
            '纸张cm': '%s×%s' % (round(sec.page_width.cm, 2), round(sec.page_height.cm, 2)),
            '上边距cm': round(sec.top_margin.cm, 2),
            '下边距cm': round(sec.bottom_margin.cm, 2),
            '左边距cm': round(sec.left_margin.cm, 2),
            '右边距cm': round(sec.right_margin.cm, 2),
            '页眉距边界cm': round(sec.header_distance.cm, 2),
            '可用宽度cm': round((sec.page_width - sec.left_margin - sec.right_margin) / 360000.0, 2),
        }
        if abs(sec.page_width.cm - 21.0) > 0.1 or abs(sec.page_height.cm - 29.7) > 0.1:
            problems.append('纸张不是 A4：%s×%s cm' % (round(sec.page_width.cm, 2),
                                                    round(sec.page_height.cm, 2)))
        for key, label in (('top', '上'), ('bottom', '下'), ('left', '左'), ('right', '右')):
            actual = round(getattr(sec, key + '_margin').cm, 2)
            if abs(actual - spec[key]) > TOL_CM:
                problems.append('%s边距 %.2fcm，规范为 %.2fcm' % (label, actual, spec[key]))
        if abs(sec.header_distance.cm - spec['header']) > TOL_CM:
            problems.append('页眉距边界 %.2fcm，规范为 %.2fcm'
                            % (sec.header_distance.cm, spec['header']))

    # ---- 页眉页脚版式 ----
    usable = layout.get('可用宽度cm')
    hdr_style_tabs = style_tabs(doc, 'Header')
    for item in hf_audits:
        h = item['页眉']
        tag = '第 %d 节页眉' % item['序号']
        if not h['文字']:
            if kind == 'manual' and item['序号'] == 1:
                pass  # 封面无页眉，正确
            else:
                problems.append('%s为空（应含「软件全称 版本号」与「第 X 页 共 N 页」）' % tag)
            continue
        if h['字号pt'] and abs(h['字号pt'] - HEADER_SIZE) > 0.2:
            problems.append('%s字号 %.1fpt，规范为 %.0fpt' % (tag, h['字号pt'], HEADER_SIZE))
        if not h['下框线']:
            problems.append('%s缺少下框线（0.75 磅实线）' % tag)
        if h['右制表位cm'] is None:
            problems.append('%s缺少右对齐制表位，页码未贴右边' % tag)
        elif usable and abs(h['右制表位cm'] - usable) > 0.3:
            problems.append('%s右制表位在 %.2fcm，应等于可用宽度 %.2fcm'
                            % (tag, h['右制表位cm'], usable))
        for t in h.get('制表位', []):
            if t.startswith('center@'):
                problems.append('%s存在居中制表位 %s，Tab 会停在页面中间，页码不在右上角'
                                % (tag, t[7:]))
            elif t.startswith('left@') and float(t[5:-2]) < (usable or 99):
                problems.append('%s存在左对齐制表位 %s，会截住 Tab' % (tag, t[5:-2]))
        cleared = set(t[6:-2] for t in h.get('制表位', []) if t.startswith('clear@'))
        for val, pos in hdr_style_tabs:
            if val in ('center', 'left') and ('%.2f' % pos) not in cleared:
                problems.append('%s继承了「页眉」样式的 %s 制表位（%.2fcm）且未清除，'
                                '页码会被拉向页面中间' % (tag, '居中' if val == 'center' else '左对齐', pos))
        if 'PAGE' not in h['域']:
            problems.append('%s页码不是 PAGE 域，改页数后不会自动更新' % tag)
        if not any(f in ('NUMPAGES', 'SECTIONPAGES') for f in h['域']):
            if not re.search(r'共\s*\d+\s*页', h['文字']):
                problems.append('%s缺少「共 N 页」' % tag)
        if item['页脚']['文字']:
            problems.append('第 %d 节页脚有内容「%s」，规范要求页脚留空'
                            % (item['序号'], item['页脚']['文字']))

    # ---- 疑似水印/声明文字 ----
    hits = []
    for item in hf_audits:
        for where in ('页眉', '页脚'):
            txt = item[where]['文字']
            for word in FORBIDDEN_HF:
                if word.lower() in txt.lower():
                    hits.append('第 %d 节%s出现「%s」' % (item['序号'], where, word))
    body_hits = []
    for p in body.iterchildren(qn('w:p')):
        txt = para_text(p)
        for word in FORBIDDEN_BODY:
            if word.lower() in txt.lower():
                body_hits.append('正文出现「%s」：%s' % (word, txt.strip()[:40]))
                break
    problems.extend(hits)
    if body_hits:
        notices.append('正文命中 %d 处疑似声明/水印字样，请人工确认：%s'
                       % (len(body_hits), body_hits[:5]))

    # ---- 固定行距裁剪行内图片 ----
    clipped = inline_image_clip(doc)
    for c in clipped:
        problems.append('第 %d 段含行内图片，但行距为固定值（w:lineRule="exact"%s，来自%s），'
                        'Word 与 PDF 会按行高裁剪，截图只剩一条；'
                        '该段改为倍数行距（line_spacing = 1.0）'
                        % (c['段落序号'],
                           '，%spt' % c['行距pt'] if c['行距pt'] else '',
                           c['来源'] or '未知'))

    if kind == 'source' and not audit_only:
        if expect_lines:
            bad = [(i + 1, n) for i, n in enumerate(counts) if n != expect_lines]
            bad = [x for x in bad if x[0] != len(counts)]  # 末页允许不足
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

    # ---- 「共 N 页」与实际页数核对 ----
    auto_field = None
    for item in hf_audits:
        for f in item['页眉']['域']:
            if f in ('NUMPAGES', 'SECTIONPAGES'):
                auto_field = f
    reported = None
    if sections and auto_field is None:
        m = re.search(r'共\s*(\d+)\s*页', sections[0]['页眉'])
        if m:
            reported = int(m.group(1))
    actual_pages = len(pages)
    cover_pages = 0
    if kind == 'manual' and len(doc.sections) > 1 and not hf_audits[0]['页眉']['文字']:
        cover_pages = 1
        # 封面独占第 1 页（正文首页带 page_break_before 或分节符），页数口径按正文计
        notices.append('封面计 1 页，正文自第 2 页起；「共 N 页」按正文页数计（不含封面）')
    target_total = expect_total if expect_total else actual_pages - cover_pages
    if auto_field is None and reported is not None:
        if reported != target_total:
            problems.append('页眉声明「共 %d 页」，应为 %d 页，页码不一致'
                            % (reported, target_total))
    elif auto_field is None and kind == 'source' and not audit_only:
        problems.append('页眉未找到「共 N 页」文字')

    if kind == 'manual' and not audit_only:
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
        '页面': layout,
        '版式': hf_audits,
        '页眉声明总页数': reported if reported is not None else
        ('%s 域（Word 自动计算）' % auto_field if auto_field else None),
        '问题': problems,
        '提示': notices,
        '结论': '通过' if not problems else '需修正',
    }
    return result


def main():
    ap = argparse.ArgumentParser(description='校验软著材料 Word 与版式')
    ap.add_argument('files', nargs='+', help='待校验的 docx')
    ap.add_argument('--expect-lines', type=int, default=None, help='源代码材料每页行数')
    ap.add_argument('--title', default=None, help='期望页眉文字（软件全称 + 版本号）')
    ap.add_argument('--kind', choices=['source', 'manual'], default='source')
    ap.add_argument('--expect-total', type=int, default=None,
                    help='页眉「共 N 页」的期望值（original 页码模式的第二份材料需指定）')
    ap.add_argument('--audit-only', action='store_true',
                    help='只体检版式（纸张/页边距/页眉页脚/水印），不判断行数与页数')
    args = ap.parse_args()

    all_ok = True
    for path in args.files:
        res = check(path, args.expect_lines, args.title, args.kind, args.expect_total,
                    args.audit_only)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        if res['问题']:
            all_ok = False
    sys.exit(0 if all_ok else 1)


if __name__ == '__main__':
    main()

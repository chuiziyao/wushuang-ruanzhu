# -*- coding: utf-8 -*-
"""把 Markdown 版操作说明书转成软著「文档鉴别材料」Word。

要点：
  · 封面独占第 1 页，不显示页眉页码；正文自第 2 页起，页码从 1 重新起算
  · 页眉左＝软件全称 + 版本号，右＝第 X 页 共 N 页（N 用 SECTIONPAGES 域，不含封面）
  · 页脚留空（不写公司名、不写日期、不加水印）
  · 每章（## ）可选另起一页，便于按章补截图
  · 图片不存在时留占位文字并记入《截图清单》，提醒补真实界面截图（不得用 AI 生成图）
  · 文档页数由 Word 排版决定，若超过 60 页，提交时按「前 30 页 + 后 30 页」截取 PDF 页码区间

用法：
  python gen_manual_docx.py --md 操作说明书.md --name "示例事项软件" --version V1.0 \
      --out 操作说明书.docx [--chapter-pagebreak]
"""
import argparse
import os
import re
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from docx.enum.section import WD_SECTION  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

import config as cfgmod  # noqa: E402
from docx_common import (  # noqa: E402
    new_document, setup_page, build_header, build_cover, ensure_empty_hf,
    set_page_number_start, set_run_font, set_update_fields, usable_width_cm,
)

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

BODY_PT = 10.5
MONO_PT = 9.0

INLINE_RE = re.compile(r'(\*\*.+?\*\*|`[^`]+`)')


def add_inline(par, text, size=BODY_PT, cn='宋体', mono='Consolas', base_bold=False):
    for seg in INLINE_RE.split(text):
        if not seg:
            continue
        if seg.startswith('**') and seg.endswith('**') and len(seg) > 4:
            run = par.add_run(seg[2:-2])
            set_run_font(run, cn, cn, size, bold=True)
        elif seg.startswith('`') and seg.endswith('`') and len(seg) > 2:
            run = par.add_run(seg[1:-1])
            set_run_font(run, mono, cn, size - 0.5)
        else:
            run = par.add_run(seg)
            set_run_font(run, cn, cn, size, bold=base_bold)


def body_par(doc, text, first_line_indent=True, size=BODY_PT, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.35
    pf.alignment = align
    if first_line_indent:
        pf.first_line_indent = Pt(size * 2)
    add_inline(par, text, size)
    return par


def heading(doc, text, level):
    sizes = {1: 16, 2: 13.5, 3: 11.5, 4: 10.5}
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.space_before = Pt(12 if level <= 2 else 8)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.3
    pf.keep_with_next = True
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = par.add_run(text)
    set_run_font(run, '黑体', '黑体', sizes.get(level, 10.5), bold=(level <= 3))
    return par


def add_table(doc, rows):
    if not rows:
        return
    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    width = usable_width_cm(doc.sections[0])
    for i, row in enumerate(rows):
        for j in range(cols):
            cell = table.cell(i, j)
            cell.text = ''
            par = cell.paragraphs[0]
            par.paragraph_format.space_before = Pt(1)
            par.paragraph_format.space_after = Pt(1)
            text = row[j] if j < len(row) else ''
            add_inline(par, text, 9.5)
            if i == 0:
                for run in par.runs:
                    run.font.bold = True
            cell.width = Cm(width / cols)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_image(doc, path, caption, base_dir, missing=None):
    full = path if os.path.isabs(path) else os.path.join(base_dir, path)
    width = usable_width_cm(doc.sections[0]) - 1.0
    if os.path.isfile(full):
        par = doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(4)
        par.paragraph_format.space_after = Pt(2)
        # 固定行距（Pt）会把行内图片裁成细条，必须覆盖为倍数行距
        par.paragraph_format.line_spacing = 1.0
        par.add_run().add_picture(full, width=Cm(width))
    else:
        if missing is not None:
            missing.append({'图注': caption or '(无图注)',
                            '约定文件名': os.path.basename(path)})
        par = doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(6)
        par.paragraph_format.space_after = Pt(2)
        run = par.add_run('【此处插入界面截图：%s】' % (caption or os.path.basename(path)))
        set_run_font(run, '宋体', '宋体', 10)
        run.font.italic = True
    if caption:
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_after = Pt(8)
        cap.paragraph_format.line_spacing = 1.0
        run = cap.add_run(caption)
        set_run_font(run, '宋体', '宋体', 9)


def add_code_block(doc, lines):
    for line in lines:
        par = doc.add_paragraph()
        pf = par.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = Pt(13)
        pf.left_indent = Pt(12)
        run = par.add_run(line if line.strip() else ' ')
        set_run_font(run, 'Consolas', '宋体', MONO_PT)


def convert(md_path, out_path, name, version, chapter_pagebreak=False,
            doc_title='用户手册', owner=None, cover=True):
    with open(md_path, encoding='utf-8') as f:
        lines = f.read().split('\n')

    doc = new_document('宋体', '宋体', BODY_PT)
    title = '%s %s' % (name.strip(), version.strip())
    missing = []

    # ---- 第一节：封面（无页眉、无页脚、无页码）----
    cover_sec = doc.sections[0]
    setup_page(cover_sec, kind='manual')
    ensure_empty_hf(cover_sec)
    if cover:
        build_cover(doc, name.strip(), version.strip(), doc_title, owner)
        # ---- 第二节：正文（页码自 1 起，「共 N 页」不含封面）----
        body_sec = doc.add_section(WD_SECTION.NEW_PAGE)
        setup_page(body_sec, kind='manual')
        build_header(body_sec, title, total=1, total_field='SECTIONPAGES',
                     font_ascii='宋体', font_cn='宋体')
        set_page_number_start(body_sec, 1)
    else:
        build_header(cover_sec, title, total=1, total_field='NUMPAGES',
                     font_ascii='宋体', font_cn='宋体')
        body_sec = cover_sec
    set_update_fields(doc)
    base_dir = os.path.dirname(os.path.abspath(md_path))

    i, n = 0, len(lines)
    first_h1_done = False
    while i < n:
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        if stripped.startswith('```'):
            i += 1
            block = []
            while i < n and not lines[i].strip().startswith('```'):
                block.append(lines[i].rstrip())
                i += 1
            i += 1
            add_code_block(doc, block)
            continue

        m = re.match(r'^(#{1,4})\s+(.*)$', stripped)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            par = heading(doc, text, level)
            if level == 1 and not first_h1_done:
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                first_h1_done = True
            if level == 2 and chapter_pagebreak and first_h1_done:
                par.paragraph_format.page_break_before = True
            i += 1
            continue

        # 图片：![说明](路径)
        m = re.match(r'^!\[(.*?)\]\((.*?)\)$', stripped)
        if m:
            add_image(doc, m.group(2), m.group(1), base_dir, missing)
            i += 1
            continue

        # 表格
        if stripped.startswith('|') and stripped.count('|') >= 2:
            rows = []
            while i < n and lines[i].strip().startswith('|'):
                row = lines[i].strip().strip('|')
                cells = [c.strip() for c in row.split('|')]
                if not all(re.match(r'^:?-{2,}:?$', c) for c in cells if c):
                    rows.append(cells)
                i += 1
            add_table(doc, rows)
            continue

        m = re.match(r'^[-*]\s+(.*)$', stripped)
        if m:
            par = body_par(doc, '· ' + m.group(1), first_line_indent=False)
            par.paragraph_format.left_indent = Pt(18)
            i += 1
            continue

        m = re.match(r'^(\d+)[.、)]\s+(.*)$', stripped)
        if m:
            par = body_par(doc, '%s. %s' % (m.group(1), m.group(2)), first_line_indent=False)
            par.paragraph_format.left_indent = Pt(18)
            i += 1
            continue

        if stripped.startswith('>'):
            par = body_par(doc, stripped.lstrip('> '), first_line_indent=False)
            par.paragraph_format.left_indent = Pt(18)
            for run in par.runs:
                run.font.italic = True
            i += 1
            continue

        body_par(doc, stripped)
        i += 1

    doc.save(out_path)
    return out_path, missing


def write_shot_list(md_path, missing, name, version):
    """把缺失的界面截图整理成《截图清单》，交给用户按单补拍。"""
    if not missing:
        return None
    out = os.path.join(os.path.dirname(os.path.abspath(md_path)), '截图清单.md')
    L = ['# 界面截图清单 · %s %s' % (name.strip(), version.strip()), '',
         '说明书里 %d 处图位还没有真实截图，Word 中显示为「【此处插入界面截图：…】」。' % len(missing),
         '按下面清单逐张补拍，放进 Markdown 所在目录的同名路径，重新生成即可。', '',
         '| 序号 | 图注 | 约定文件名 |', '|---|---|---|']
    for i, item in enumerate(missing, 1):
        L.append('| %d | %s | %s |' % (i, item['图注'], item['约定文件名']))
    L += ['', '## 截图要求', '',
          '1. 必须是软件真实运行界面，不得用设计稿、示意图、AI 生成图代替。',
          '2. 画面里不要出现系统通知、微信/QQ 弹窗、个人信息、测试账号密码。',
          '3. 同一套图用同一台设备、同一分辨率截取，避免大小不一。',
          '4. 用 PNG，宽度不低于 1080 像素；横屏界面统一横屏。',
          '5. 浏览器或 IDE 截图时关掉书签栏、控制台、AI 工具水印。', '']
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L))
    return out


def main():
    ap = argparse.ArgumentParser(description='生成软著文档鉴别材料 Word（操作说明书 / 用户手册）')
    ap.add_argument('--md', default=None, help='Markdown 说明书（缺省取当前目录 操作说明书.md）')
    ap.add_argument('--name', default=None, help='软件全称（缺省取 00_基本信息.json）')
    ap.add_argument('--version', default=None, help='版本号（缺省取配置，默认 V1.0）')
    ap.add_argument('--out', default=None, help='输出 docx 路径')
    ap.add_argument('--config', default=None, help='配置文件路径 00_基本信息.json')
    ap.add_argument('--doc-title', default='用户手册',
                    help='封面上的文档名称，默认「用户手册」，也可写「操作说明书」')
    ap.add_argument('--owner', default=None, help='封面上的著作权人全称（缺省取配置）')
    ap.add_argument('--no-cover', action='store_true', help='不生成封面（页眉从第 1 页开始）')
    ap.add_argument('--chapter-pagebreak', action='store_true', help='每章另起一页')
    args = ap.parse_args()

    cfg = cfgmod.load(args.config)
    args.name = args.name or cfgmod.get(cfg, '软件全称')
    args.version = args.version or cfgmod.get(cfg, '版本号', 'V1.0')
    args.owner = args.owner or cfgmod.get(cfg, '申请人名称')
    args.md = args.md or ('操作说明书.md' if os.path.isfile('操作说明书.md') else None)
    if not args.md:
        sys.exit('缺少说明书 Markdown：加 --md 操作说明书.md，或用 examples/操作说明书模板.md 改写')
    if not args.name:
        sys.exit('缺少软件全称：加 --name "全称"，或先 python bootstrap.py --name "全称" ...')

    out = args.out or os.path.splitext(args.md)[0] + '.docx'
    out, missing = convert(args.md, out, args.name, args.version, args.chapter_pagebreak,
                           doc_title=args.doc_title, owner=args.owner,
                           cover=not args.no_cover)
    print('已生成：%s' % out)
    if missing:
        shot = write_shot_list(args.md, missing, args.name, args.version)
        print('待补界面截图 %d 处，清单见：%s' % (len(missing), shot))
    print('提醒：文档页数由 Word 排版决定；若超过 60 页，提交时导出 PDF 后截取'
          '第 1–30 页与最后 30 页（页码保持原文真实页码）。')


if __name__ == '__main__':
    main()

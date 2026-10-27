# -*- coding: utf-8 -*-
"""软著材料 Word 公共组件：页面/页眉/页码域/字体/固定行距。

所有生成器（源代码、操作说明书）共用本模块，保证页眉页脚规则一致：
  · 页眉左：软件全称 + 空格 + 版本号（必须与申请表逐字一致）
  · 页眉右：第 X 页 共 N 页（右上角页码）
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7

# sectPr 子元素顺序（OOXML 规范），插入 pgNumType 时不能乱序
_SECTPR_ORDER = [
    'w:footnotePr', 'w:endnotePr', 'w:type', 'w:pgSz', 'w:pgMar', 'w:paperSrc',
    'w:pgBorders', 'w:lnNumType', 'w:pgNumType', 'w:cols', 'w:formProt',
    'w:vAlign', 'w:noEndnote', 'w:titlePg', 'w:textDirection', 'w:bidi',
    'w:rtlGutter', 'w:docGrid', 'w:printerSettings', 'w:sectPrChange',
]


def set_run_font(run, ascii_font='Consolas', ea_font='宋体', size=9.0, bold=False):
    """同时设置西文与中文字体，避免中文回退成默认字体导致排版漂移。"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), ascii_font)
    rFonts.set(qn('w:hAnsi'), ascii_font)
    rFonts.set(qn('w:eastAsia'), ea_font)
    rFonts.set(qn('w:cs'), ascii_font)


def add_field(paragraph, instr, placeholder='1', size=9.0,
              ascii_font='Consolas', ea_font='宋体'):
    """插入 Word 域（PAGE / NUMPAGES），域结果由 Word 排版时自动计算。"""
    runs = []
    r = paragraph.add_run()
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'begin')
    r._r.append(fc)
    runs.append(r)

    r = paragraph.add_run()
    instr_el = OxmlElement('w:instrText')
    instr_el.set(qn('xml:space'), 'preserve')
    instr_el.text = ' %s ' % instr
    r._r.append(instr_el)
    runs.append(r)

    r = paragraph.add_run()
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'separate')
    r._r.append(fc)
    runs.append(r)

    runs.append(paragraph.add_run(placeholder))

    r = paragraph.add_run()
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'end')
    r._r.append(fc)
    runs.append(r)

    for run in runs:
        set_run_font(run, ascii_font, ea_font, size)
    return runs


def _insert_sectpr_child(sectPr, el):
    tag = el.tag
    order = [qn(t) for t in _SECTPR_ORDER]
    try:
        idx = order.index(tag)
    except ValueError:
        sectPr.append(el)
        return
    for i, child in enumerate(list(sectPr)):
        if child.tag in order and order.index(child.tag) > idx:
            child.addprevious(el)
            return
    sectPr.append(el)


def set_page_number_start(section, start):
    """设置本节起始页码（对应 w:pgNumType @w:start）。"""
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        _insert_sectpr_child(sectPr, pgNumType)
    pgNumType.set(qn('w:start'), str(int(start)))


def set_update_fields(doc):
    """打开文档时让 Word 刷新域，避免页码缓存值显示为占位 1。"""
    settings = doc.settings.element
    if settings.find(qn('w:updateFields')) is None:
        el = OxmlElement('w:updateFields')
        el.set(qn('w:val'), 'true')
        settings.append(el)


def setup_page(section, top_cm=2.2, bottom_cm=1.8, left_cm=2.0, right_cm=2.0,
               header_cm=1.0, footer_cm=1.0):
    section.page_width = Cm(A4_WIDTH_CM)
    section.page_height = Cm(A4_HEIGHT_CM)
    section.top_margin = Cm(top_cm)
    section.bottom_margin = Cm(bottom_cm)
    section.left_margin = Cm(left_cm)
    section.right_margin = Cm(right_cm)
    section.header_distance = Cm(header_cm)
    section.footer_distance = Cm(footer_cm)


def usable_height_pt(section):
    return (section.page_height - section.top_margin - section.bottom_margin) / 12700.0


def usable_width_cm(section):
    return (section.page_width - section.left_margin - section.right_margin) / 360000.0


def build_header(section, title, total_pages=None, size=9.0,
                 font_ascii='Consolas', font_cn='宋体'):
    """页眉：左标题（软件全称+版本号），右「第 X 页 共 N 页」。

    total_pages 给定则写死数字（源代码材料，页数由生成器确定，最稳）；
    为 None 则使用 NUMPAGES 域（说明书，页数随排版变化）。
    """
    header = section.header
    header.is_linked_to_previous = False
    para = header.paragraphs[0]
    para.text = ''
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)

    width = usable_width_cm(section)
    tab_stops = para.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Cm(width), WD_TAB_ALIGNMENT.RIGHT)

    run = para.add_run(title)
    set_run_font(run, font_ascii, font_cn, size)

    run = para.add_run('\t第 ')
    set_run_font(run, font_ascii, font_cn, size)
    add_field(para, 'PAGE', '1', size, font_ascii, font_cn)

    run = para.add_run(' 页 共 ')
    set_run_font(run, font_ascii, font_cn, size)
    if total_pages is None:
        add_field(para, 'NUMPAGES', '1', size, font_ascii, font_cn)
    else:
        run = para.add_run(str(total_pages))
        set_run_font(run, font_ascii, font_cn, size)
    run = para.add_run(' 页')
    set_run_font(run, font_ascii, font_cn, size)

    # 页眉下横线：细实线，视觉上与正文分隔
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '808080')
    pBdr.append(bottom)
    pPr.append(pBdr)
    return para


def new_document(base_font_cn='宋体', base_font_ascii='Consolas', base_size=9.0):
    doc = Document()
    normal = doc.styles['Normal']
    normal.font.size = Pt(base_size)
    normal.font.name = base_font_ascii
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), base_font_ascii)
    rFonts.set(qn('w:hAnsi'), base_font_ascii)
    rFonts.set(qn('w:eastAsia'), base_font_cn)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = Pt(base_size * 1.55)
    return doc


def clone_sectpr_into_last_paragraph(doc):
    """把正文级 sectPr 复制进最后一个段落的 pPr，形成「上一节」的分节符。

    用于「前 30 页 + 后 30 页」材料：两节分别设置起始页码，且不产生空段落
    （python-docx 自带的 add_section 会插入一个空段落，把 50 行挤到 51 行）。
    """
    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    if sectPr is None:
        raise RuntimeError('文档缺少 sectPr')
    paragraphs = body.findall(qn('w:p'))
    if not paragraphs:
        raise RuntimeError('文档没有段落，无法分节')
    last_p = paragraphs[-1]
    pPr = last_p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        last_p.insert(0, pPr)
    import copy
    pPr.append(copy.deepcopy(sectPr))
    return pPr

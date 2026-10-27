# -*- coding: utf-8 -*-
"""软著材料 Word 公共组件：页面、页眉、页脚、封面、字段、字体、固定行距。

版式口径集中在本文件顶部的 LAYOUT / HEADER_* 常量，改一次即全材料生效。
约定（详见 references/05_版式规范_封面页眉页脚.md）：

  · 纸张一律 A4 纵向
  · 页眉左：软件全称 + 空格 + 版本号（与申请表逐字一致），宋体 9pt
  · 页眉右：第 X 页 共 N 页（右上角），宋体 9pt，PAGE 域取页码
  · 页眉下加一条 0.75 磅黑色实线，作为正文分隔
  · 页脚留空，不写公司名、不写日期、不写「仅供…」、不加水印
  · 程序鉴别材料无封面；文档鉴别材料封面独占第 1 页，且封面不显示页眉页码
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7

# 页面版式：source=程序鉴别材料，manual=文档鉴别材料
LAYOUT = {
    'source': dict(top=2.0, bottom=1.8, left=1.8, right=1.8, header=1.2, footer=1.0),
    'manual': dict(top=2.2, bottom=1.8, left=2.2, right=2.2, header=1.2, footer=1.0),
}

HEADER_SIZE = 9.0          # 页眉字号（磅）
HEADER_RULE_SZ = 6         # 页眉下框线粗细，单位 1/8 磅，6 = 0.75 磅
HEADER_RULE_COLOR = '000000'

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


def set_char_spacing(run, pt):
    """字间距（用于封面标题拉字距，视觉上更像正式文档）。"""
    rPr = run._element.get_or_add_rPr()
    el = rPr.find(qn('w:spacing'))
    if el is None:
        el = OxmlElement('w:spacing')
        rPr.append(el)
    el.set(qn('w:val'), str(int(round(pt * 20))))


def add_field(paragraph, instr, placeholder='1', size=9.0,
              ascii_font='Consolas', ea_font='宋体'):
    """插入 Word 域（PAGE / NUMPAGES / SECTIONPAGES），域结果由 Word 排版时计算。"""
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


def setup_page(section, kind='source', **override):
    """按 LAYOUT 设置纸张与页边距；override 可临时覆盖（单位 cm）。"""
    spec = dict(LAYOUT.get(kind, LAYOUT['source']))
    spec.update({k: v for k, v in override.items() if v is not None})
    section.page_width = Cm(A4_WIDTH_CM)
    section.page_height = Cm(A4_HEIGHT_CM)
    section.top_margin = Cm(spec['top'])
    section.bottom_margin = Cm(spec['bottom'])
    section.left_margin = Cm(spec['left'])
    section.right_margin = Cm(spec['right'])
    section.header_distance = Cm(spec['header'])
    section.footer_distance = Cm(spec['footer'])


def usable_height_pt(section):
    return (section.page_height - section.top_margin - section.bottom_margin) / 12700.0


def usable_width_cm(section):
    return (section.page_width - section.left_margin - section.right_margin) / 360000.0


def ensure_empty_hf(section):
    """显式清空本节页眉页脚，防止继承上一节内容。封面必须调用。"""
    for hf in (section.header, section.footer):
        hf.is_linked_to_previous = False
        for p in list(hf.paragraphs):
            p.text = ''


def _clear_style_tabs(doc, names=('Header', 'Footer')):
    """删除页眉/页脚样式自带的中、右制表位。

    python-docx 默认模板的 Header 样式带 8.25cm 居中、16.5cm 右对齐两个制表位，
    页眉里那个 Tab 会先跳到 8.25cm，页码就落在页面中间而不是右上角。必须清掉。
    """
    removed = []
    for name in names:
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        pPr = style.element.find(qn('w:pPr'))
        if pPr is None:
            continue
        tabs = pPr.find(qn('w:tabs'))
        if tabs is not None:
            pPr.remove(tabs)
            removed.append(name)
    return removed


def build_header(section, title, total=None, total_field='literal',
                 size=HEADER_SIZE, font_ascii='宋体', font_cn='宋体', rule=True):
    """页眉：左「软件全称 版本号」，右「第 X 页 共 N 页」。

    total_field:
      · 'literal'      —— 直接写死总页数（源代码材料，页数由生成器确定，最稳）
      · 'NUMPAGES'     —— 全文总页数域（单节文档用）
      · 'SECTIONPAGES' —— 本节总页数域（封面独立成节时用，不含封面页）
    """
    header = section.header
    header.is_linked_to_previous = False
    para = header.paragraphs[0]
    para.text = ''
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)

    pPr = para._p.get_or_add_pPr()

    # 先清掉继承来的制表位（模板自带 4680/9360 twips），再只留一个右对齐停止位
    # 位于右边界；否则 Tab 会停在 8.25cm 的居中位，页码跑到页面中间。
    tabs = pPr.find(qn('w:tabs'))
    if tabs is None:
        tabs = OxmlElement('w:tabs')
        pPr.append(tabs)
    else:
        for child in list(tabs):
            tabs.remove(child)
    for pos in (4680, 9360, 4252):
        clear = OxmlElement('w:tab')
        clear.set(qn('w:val'), 'clear')
        clear.set(qn('w:pos'), str(pos))
        tabs.append(clear)

    width = usable_width_cm(section)
    stop = OxmlElement('w:tab')
    stop.set(qn('w:val'), 'right')
    stop.set(qn('w:pos'), str(int(round(width / 2.54 * 1440))))
    tabs.append(stop)

    run = para.add_run(title)
    set_run_font(run, font_ascii, font_cn, size)

    run = para.add_run('\t第 ')
    set_run_font(run, font_ascii, font_cn, size)
    add_field(para, 'PAGE', '1', size, font_ascii, font_cn)

    run = para.add_run(' 页 共 ')
    set_run_font(run, font_ascii, font_cn, size)
    if total_field == 'literal':
        run = para.add_run(str(total))
        set_run_font(run, font_ascii, font_cn, size)
    else:
        add_field(para, total_field, str(total or 1), size, font_ascii, font_cn)
    run = para.add_run(' 页')
    set_run_font(run, font_ascii, font_cn, size)

    if rule:
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), str(HEADER_RULE_SZ))
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), HEADER_RULE_COLOR)
        pBdr.append(bottom)
        tabs.addprevious(pBdr)   # 按 OOXML 顺序：pBdr 在 tabs 之前
    return para


def build_cover(doc, name, version='V1.0', doc_title='用户手册', owner=None,
                cn_font='宋体', title_font='黑体'):
    """封面：独占一页，不显示页眉、页脚、页码。

    排版层次（自上而下）：著作权人 → 软件全称（主标题）→ 文档名称 → 版本号。
    封面不放日期、不放 Logo、不放网址，避免与申请表日期逻辑打架。
    """
    def blank(size=14, count=1):
        for _ in range(count):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(size)

    def line(text, font, size, bold=False, spacing=0.0, before=0, after=0):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.line_spacing = 1.5
        run = p.add_run(text)
        set_run_font(run, font, font, size, bold=bold)
        if spacing:
            set_char_spacing(run, spacing)
        return p

    blank(count=5)
    if owner:
        line(owner, cn_font, 12, spacing=1.0, after=30)
    else:
        blank(count=2)
    line(name, title_font, 26, bold=True, spacing=2.0, after=6)
    line('版 本 号　%s' % version, cn_font, 12, after=40)
    line(doc_title, title_font, 22, bold=True, spacing=6.0)
    blank(count=4)


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
    _clear_style_tabs(doc)
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

# -*- coding: utf-8 -*-
"""中文公文排版底层工具：页面、字网格、字体、页码、字体嵌入。

所有尺寸以 twips（1/20 磅，1mm = 56.6929tw）为内部单位，
与 Word 的 XML 单位一致，避免换算误差累积。
"""

import os
import uuid

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.opc.rel import Relationships
from docx.opc.part import XmlPart
from docx.oxml import parse_xml
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips

# ------------------------------------------------------------------ 版式常量

MM = 56.6929

A4_W, A4_H = 11906, 16838            # 210mm x 297mm
MARGIN = dict(top=2098, bottom=1984, left=1587, right=1474)   # 37/35/28/26 mm
TEXT_W = A4_W - MARGIN['left'] - MARGIN['right']   # 8845tw = 156mm
TEXT_H = A4_H - MARGIN['top'] - MARGIN['bottom']   # 12756tw = 225mm
LINES_PER_PAGE = 22
CHARS_PER_LINE = 28
LINE_PITCH = int(round(TEXT_H / LINES_PER_PAGE))   # 580tw ≈ 10.23mm

SANHAO = 16          # 三号 = 16pt
CHAR_EM = 320        # 一个三号汉字的理论宽度（twips）：16pt * 20 = 320

# 字号（pt）
SIZE = {'初号': 42, '小初': 36, '一号': 26, '小一': 24, '二号': 22, '小二': 18,
        '三号': 16, '小三': 15, '四号': 14, '小四': 12, '五号': 10.5}

# 公文红：GB/T 9704 要求「大红色」，常用 C00000
RED = 'C00000'
BLACK = '000000'

# 字体族：key -> (Word 中写入的字体名, 本机回退名列表)
FONT_NAMES = {
    'fangsong': ('仿宋_GB2312', ['仿宋', 'FangSong']),
    'kaiti': ('楷体_GB2312', ['楷体', 'KaiTi']),
    'xbs': ('方正小标宋简体', ['宋体', 'SimSun']),
    'heiti': ('黑体', ['SimHei']),
    'songti': ('宋体', ['SimSun']),
    'times': ('Times New Roman', ['宋体']),
    'mono': ('Consolas', ['Courier New']),
}

# 需要嵌入的字体：key -> 子集文件名
EMBED_FONTS = {
    'fangsong': 'FangSong_GB2312.ttf',
    'kaiti': 'KaiTi_GB2312.ttf',
    'xbs': 'FZXiaoBiaoSong-B05S.ttf',
}


# ------------------------------------------------------------------ 字符工具

def char_width(ch):
    """返回字符相对三号汉字宽度（CHAR_EM）的倍数。"""
    cp = ord(ch)
    if cp < 0x80:
        return 0.5                      # 半角
    return 1.0                          # 全角 / 汉字


def measure(text, size_pt=SANHAO):
    """按全角=1字、半角=0.5字估算文本占位，单位「字」。"""
    return sum(char_width(c) for c in text) * size_pt / SANHAO


# ------------------------------------------------------------------ 字体设置

def set_run_font(run, font_key, size_pt, bold=False, color=BLACK,
                 embolden=False, spacing_twips=None):
    """给 run 设置完整字体信息（含 eastAsia，避免被主题字体覆盖）。"""
    name, _fallbacks = FONT_NAMES[font_key]
    ascii_name = 'Times New Roman' if font_key in ('songti', 'times') else name

    run.font.size = Twips(int(size_pt * 20))
    rPr = run._element.get_or_add_rPr()

    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), ascii_name)
    rFonts.set(qn('w:hAnsi'), ascii_name)
    rFonts.set(qn('w:eastAsia'), name)
    rFonts.set(qn('w:cs'), ascii_name)
    # 明确取消主题字体引用
    for attr in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
        if rFonts.get(qn(attr)) is not None:
            del rFonts.attrib[qn(attr)]

    if bold:
        run.font.bold = True
    if embolden:
        _set_embolden(rPr)
    if color and color.upper() != BLACK:
        _set_color(rPr, color)
    if spacing_twips:
        sp = OxmlElement('w:spacing')
        sp.set(qn('w:val'), str(int(spacing_twips)))
        rPr.append(sp)
    return run


def _set_color(rPr, hexcolor):
    color = OxmlElement('w:color')
    color.set(qn('w:val'), hexcolor)
    rPr.append(color)


def _set_embolden(rPr, value=30):
    """伪粗（字符加粗 0.3pt），用于黑体/小标宋这类无粗体字重的字体。"""
    el = OxmlElement('w:em')
    el.set(qn('w:val'), str(value))
    rPr.append(el)


# ------------------------------------------------------------------ 段落设置

def set_exact_line(par, line_tw=LINE_PITCH):
    pPr = par._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    spacing.set(qn('w:line'), str(int(line_tw)))
    spacing.set(qn('w:lineRule'), 'exact')
    # 关闭孤行控制，保证网格严格
    wc = OxmlElement('w:widowControl')
    wc.set(qn('w:val'), '0')
    pPr.append(wc)


def set_indent(par, left=0, right=0, first=0):
    """缩进，单位 twips。"""
    pPr = par._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:left'), str(int(left)))
    ind.set(qn('w:right'), str(int(right)))
    ind.set(qn('w:rightChars'), '0')
    ind.set(qn('w:firstLine'), str(int(first)))
    ind.set(qn('w:firstLineChars'), '0')


def set_align(par, value):
    pPr = par._p.get_or_add_pPr()
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        pPr.append(jc)
    jc.set(qn('w:val'), value)


def no_grid(par):
    """段落不参与文档网格对齐（用于标题、表格等非正文行）。"""
    pPr = par._p.get_or_add_pPr()
    el = pPr.find(qn('w:snapToGrid'))
    if el is None:
        el = OxmlElement('w:snapToGrid')
        pPr.append(el)
    el.set(qn('w:val'), '0')


def set_borders(par, top=None, bottom=None, left=None, right=None):
    """给段落加边框（用于红色分隔线、版记分隔线）。参数为 (sz_eighth_pt, color)。"""
    pPr = par._p.get_or_add_pPr()
    pbdr = pPr.find(qn('w:pBdr'))
    if pbdr is None:
        pbdr = OxmlElement('w:pBdr')
        pPr.append(pbdr)
    for edge, spec in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        if not spec:
            continue
        sz, color = spec
        el = OxmlElement('w:%s' % edge)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(int(sz)))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        pbdr.append(el)


def set_tabs(par, positions):
    """设置制表位。positions: [(twips, align)]，align ∈ left/center/right/decimal。"""
    pPr = par._p.get_or_add_pPr()
    tabs = pPr.find(qn('w:tabs'))
    if tabs is None:
        tabs = OxmlElement('w:tabs')
        pPr.append(tabs)
    for pos, align in positions:
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), align)
        tab.set(qn('w:pos'), str(int(pos)))
        tabs.append(tab)


def keep_with_next(par, on=True):
    pPr = par._p.get_or_add_pPr()
    el = pPr.find(qn('w:keepNext'))
    if el is None:
        el = OxmlElement('w:keepNext')
        pPr.append(el)
    el.set(qn('w:val'), '1' if on else '0')


def mark_paragraph(par, tag):
    """给段落打标记，供后续按标签定位（例如把版记推到页底）。"""
    pPr = par._p.get_or_add_pPr()
    el = pPr.find(qn('w:mark'))
    if el is None:
        el = OxmlElement('w:mark')
        pPr.append(el)
    el.set(qn('w:val'), tag)
    par._gw_tag = tag


def paginate_before_banji(doc):
    """把带 mark="banji" 的段落之前插入一个分页，使其独占最后一面。

    规范要求版记置于公文最后一面底部；实际排版时若正文已接近页尾，
    版记会跨页。用「分页 + 段前空行」保证版记完整落在同一页底部。
    """
    body = doc.element.body
    paras = body.findall(qn('w:p'))
    for idx, p in enumerate(paras):
        pPr = p.find(qn('w:pPr'))
        if pPr is None:
            continue
        mark = pPr.find(qn('w:mark'))
        if mark is None or mark.get(qn('w:val')) != 'banji':
            continue
        # 在该段之前插入分页符
        page_break = OxmlElement('w:p')
        pPr2 = OxmlElement('w:pPr')
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:before'), '0')
        spacing.set(qn('w:after'), '0')
        spacing.set(qn('w:line'), str(LINE_PITCH))
        spacing.set(qn('w:lineRule'), 'exact')
        pPr2.append(spacing)
        page_break.append(pPr2)
        r = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r.append(br)
        page_break.append(r)
        p.addprevious(page_break)
        return True
    return False


def page_break_before(par):
    pPr = par._p.get_or_add_pPr()
    el = pPr.find(qn('w:pageBreakBefore'))
    if el is None:
        el = OxmlElement('w:pageBreakBefore')
        pPr.insert(0, el)
    el.set(qn('w:val'), '1')


# ------------------------------------------------------------------ 页面网格

def setup_section(section, first_page_number=1):
    """A4、页边距、每页 22 行 × 每行 28 字。"""
    section.page_width = Twips(A4_W)
    section.page_height = Twips(A4_H)
    section.top_margin = Twips(MARGIN['top'])
    section.bottom_margin = Twips(MARGIN['bottom'])
    section.left_margin = Twips(MARGIN['left'])
    section.right_margin = Twips(MARGIN['right'])
    section.header_distance = Twips(851)      # 15mm
    section.footer_distance = Twips(397)      # 7mm —— 页码距版心下边缘

    sectPr = section._sectPr
    grid = sectPr.find(qn('w:docGrid'))
    if grid is None:
        grid = OxmlElement('w:docGrid')
        sectPr.append(grid)
    grid.set(qn('w:type'), 'linesAndChars')
    grid.set(qn('w:linePitch'), str(LINE_PITCH))
    # charSpace 单位为 1/4096 磅：每字网格宽 = TEXT_W/28 ≈ 15.80pt，
    # 相对 16pt 字宽需压缩 0.20pt，即 (pitch_pt - 16) * 4096
    pitch_pt = (TEXT_W / CHARS_PER_LINE) / 20.0
    grid.set(qn('w:charSpace'), str(int(round((pitch_pt - 16.0) * 4096))))


def enable_odd_even_pages(doc):
    """页码奇偶分列（单页码居右、双页码居左）需要开启此开关。"""
    settings = doc.settings.element
    for tag in ('w:evenAndOddHeaders', 'w:embedTrueTypeFonts'):
        el = settings.find(qn(tag))
        if el is None:
            settings.append(OxmlElement(tag))


def add_page_field(par):
    """在段落中插入 PAGE 域。"""
    r1 = par.add_run()
    fld = OxmlElement('w:fldChar'); fld.set(qn('w:fldCharType'), 'begin')
    r1._element.append(fld)

    r2 = par.add_run()
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' PAGE '
    r2._element.append(instr)

    r3 = par.add_run()
    fld = OxmlElement('w:fldChar'); fld.set(qn('w:fldCharType'), 'separate')
    r3._element.append(fld)

    r4 = par.add_run('1')

    r5 = par.add_run()
    fld = OxmlElement('w:fldChar'); fld.set(qn('w:fldCharType'), 'end')
    r5._element.append(fld)
    return [r1, r2, r3, r4, r5]


def build_footers(doc, section):
    """单页码居右空一字；双页码居左空一字。4 号半角宋体阿拉伯数字。"""
    indent_one_char = int(SIZE['四号'] * 20)          # 14pt = 280tw

    # 奇数页
    section.footer.is_linked_to_previous = False
    odd = section.footer.paragraphs[0]
    odd.text = ''
    set_align(odd, 'right')
    set_exact_line(odd, 400)
    set_indent(odd, right=indent_one_char)
    for run in add_page_field(odd):
        set_run_font(run, 'times', SIZE['四号'])

    # 偶数页
    section.even_page_footer.is_linked_to_previous = False
    even = section.even_page_footer.paragraphs[0]
    even.text = ''
    set_align(even, 'left')
    set_exact_line(even, 400)
    set_indent(even, left=indent_one_char)
    for run in add_page_field(even):
        set_run_font(run, 'times', SIZE['四号'])


def set_page_number_start(section, start):
    sectPr = section._sectPr
    el = sectPr.find(qn('w:pgNumType'))
    if el is None:
        el = OxmlElement('w:pgNumType')
        sectPr.append(el)
    el.set(qn('w:start'), str(int(start)))


# ------------------------------------------------------------------ 字体嵌入

def new_document():
    """建一个最小 docx：默认模板自带 fontTable.xml，若直接新建会与其重名冲突，
    因此先清空模板中的 fontTable，再由 embed_fonts 写入我们自己的。"""
    doc = Document()
    # 移除模板自带的关系与部件（仅移除文档级 fontTable 关系）
    for rId, rel in list(doc.part.rels.items()):
        if rel.reltype == RT_FONT_TABLE:
            doc.part.drop_rel(rId)
    return doc


def _obfuscate(data, guid_str):
    """按 ECMA-376 odttf 规则混淆：取 GUID 的 16 字节（十六进制串反序），
    与字体数据前 32 字节 XOR 两轮。"""
    hexstr = guid_str.strip('{}').replace('-', '')
    key = bytes.fromhex(hexstr)[::-1]        # 16 字节，反序
    head = bytearray(data[:32])
    for i in range(32):
        head[i] ^= key[i % 16]
    return bytes(head) + data[32:]


def embed_fonts(doc, font_dir):
    """把 assets/fonts 里的字体以 obfuscatedFont 形式嵌入 docx。

    Word 的加载条件：document.xml.rels 指向 fontTable.xml；
    fontTable.xml 中每个 <w:font> 用 <w:embedRegular r:id> 指向字体部件；
    部件位于 /word/fonts/*.odttf，内容类型为 obfuscatedFont；
    settings.xml 含 <w:embedTrueTypeFonts/>。

    注意：odttf 部件需按 fontKey 混淆前 32 字节，否则 Word 解出乱码字形。
    """
    if not os.path.isdir(font_dir):
        return []

    from docx.opc.part import Part

    package = doc.part.package
    entries = []                     # [(字体名, Part, fontKey)]
    for idx, (key, fname) in enumerate(EMBED_FONTS.items(), 1):
        path = os.path.join(font_dir, fname)
        if not os.path.isfile(path):
            continue
        with open(path, 'rb') as fh:
            data = fh.read()
        fkey = '{%s}' % str(uuid.uuid4()).upper()
        part = Part(PackURI('/word/fonts/font%d.odttf' % idx),
                    OBFUSCATED_FONT_CT, _obfuscate(data, fkey), package)
        entries.append((FONT_NAMES[key][0], part, fkey))

    if not entries:
        return []

    # 1) 建立 fontTable 部件，并把字体部件登记为它的关系
    ft_part = XmlPart(PackURI('/word/fontTable.xml'), FONT_TABLE_CT, None, package)
    for i, (_name, part, _fkey) in enumerate(entries, 1):
        ft_part.load_rel(RT_FONT, part, 'rIdFont%d' % i)

    # 2) fontTable.xml：每个 <w:font> 用 embedRegular 指向上面的 rId
    fonts_xml = []
    for i, (name, _part, fkey) in enumerate(entries, 1):
        fonts_xml.append(
            '<w:font w:name="%s"><w:charset w:val="86"/><w:family w:val="auto"/>'
            '<w:pitch w:val="variable"/>'
            '<w:embedRegular r:id="rIdFont%d" w:fontKey="%s"/></w:font>' % (name, i, fkey))
    ft_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<w:fonts xml:space="preserve" %s>%s</w:fonts>'
              % (FONT_TABLE_NS, ''.join(fonts_xml)))
    ft_part._element = parse_xml(ft_xml.encode('utf-8'))

    doc.part.relate_to(ft_part, RT_FONT_TABLE)
    return [e[0] for e in entries]


OBFUSCATED_FONT_CT = 'application/vnd.openxmlformats-officedocument.obfuscatedFont'
FONT_TABLE_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'
FONT_TABLE_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
RT_FONT_TABLE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable'
RT_FONT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'

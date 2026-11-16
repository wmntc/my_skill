#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公文排版主程序：结构化 JSON -> 国标版式 .docx

用法：
    python render_gongwen.py spec.json -o 公文.docx
    python render_gongwen.py spec.json -o out.docx --no-embed
    python render_gongwen.py spec.json --check        # 只做合规自检

spec.json 结构见 references/spec_schema.md。
"""

import argparse
import io
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gw_style as G
from gw_style import LINE_PITCH, RED, SIZE, TEXT_W
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Twips

EM3 = SIZE['三号'] * 20          # 三号一字宽 = 320tw
EM4 = SIZE['四号'] * 20          # 四号一字宽 = 280tw
LINE = LINE_PITCH

# 行数记账：每个段落占用的「网格行数」累加于此，用于把版记推到页底
_used_lines = [0]


def _count(lines):
    _used_lines[0] += int(lines)


def lines_used():
    return _used_lines[0]


def reset_line_counter():
    _used_lines[0] = 0


# ------------------------------------------------------------------ 基础段落

def _par(doc, line_tw=LINE):
    par = doc.add_paragraph()
    G.set_exact_line(par, line_tw)
    G.no_grid(par)
    par.alignment = AL.JUSTIFY
    pf = par.paragraph_format
    pf.first_line_indent = Twips(0)
    pf.left_indent = Twips(0)
    pf.right_indent = Twips(0)
    pf.space_before = Twips(0)
    pf.space_after = Twips(0)
    _count(round(line_tw / float(LINE)))
    return par


def _space_before(par, tw):
    par.paragraph_format.space_before = Twips(int(tw))
    _count(round(int(tw) / float(LINE)))


def _txt(par, s, size_pt, font_key, bold=False, color=G.BLACK, embolden=False):
    if not s:
        return None
    r = par.add_run(s)
    G.set_run_font(r, font_key, size_pt, bold=bold, color=color, embolden=embolden)
    return r


def _blank(doc, n=1, line_tw=LINE):
    for _ in range(int(n)):
        _par(doc, line_tw)


def _chunk(seq, n):
    return [seq[i:i + n] for i in range(0, len(seq), n)]


# ------------------------------------------------------------------ 版头

def render_head(doc, spec):
    """份号 / 密级 / 紧急程度 / 发文机关标志 / 发文字号 / 签发人 / 红色分隔线。"""
    # 左上角三行：份号(3号数字)、密级(3号黑体)、紧急程度(3号黑体)
    top_rows = []
    if spec.get('copy_number'):
        top_rows.append(('copy_number', _pad_copy_number(str(spec['copy_number']))))
    if spec.get('secrecy'):
        top_rows.append(('secrecy', spec['secrecy']))
    if spec.get('urgency'):
        top_rows.append(('urgency', spec['urgency']))

    for kind, value in top_rows:
        par = _par(doc)
        par.alignment = AL.LEFT
        G.set_indent(par, left=0, first=0)
        if kind == 'copy_number':
            _txt(par, value, SIZE['三号'], 'times')
        else:
            _txt(par, value, SIZE['三号'], 'heiti')

    n_top = len(top_rows)
    org_mark = (spec.get('org_mark') or '').strip()

    if org_mark:
        # 发文机关标志自动适配版心宽度：从初号往下缩，保证单行放下
        widest = max(len(l) for l in org_mark.split('\n'))
        text_width_pt = G.measure(org_mark.split('\n')[0], 1000)  # 占位
        avail_pt = TEXT_W / 20.0
        mark_size = SIZE.get(spec.get('org_mark_size', '初号'), SIZE['初号'])
        while mark_size > SIZE['小二'] and widest * mark_size > avail_pt * 0.96:
            mark_size -= 1
        line_tw = int(mark_size * 1.6 * 20)
        lines = org_mark.split('\n')

        # 「距版心上边缘 35mm」：从版心顶量到红头文字上边缘
        need = int(35 * G.MM)
        used = n_top * LINE
        gap = need - used
        first = _par(doc, line_tw)
        first.paragraph_format.space_before = Twips(
            max(gap - int(mark_size * 0.4 * 20), 0))
        first.alignment = AL.CENTER
        _txt(first, lines[0], mark_size, 'xbs', color=RED)

        for extra in lines[1:]:
            p = _par(doc, line_tw)
            p.alignment = AL.CENTER
            _txt(p, extra, mark_size, 'xbs', color=RED)

        # 红头下空两行
        _blank(doc, 2)

    # 发文字号 / 签发人
    doc_number = spec.get('doc_number')
    signer = spec.get('signer')
    is_upward = spec.get('direction') == 'up'

    if doc_number or signer:
        par = _par(doc)
        if is_upward and signer:
            G.set_indent(par, left=EM3, first=0)
            par.alignment = AL.LEFT
            if doc_number:
                _txt(par, doc_number, SIZE['三号'], 'fangsong')
            G.set_tabs(par, [(TEXT_W - EM3, 'right')])
            par.add_run('\t')
            _txt(par, '签发人：', SIZE['三号'], 'fangsong')
            names = signer if isinstance(signer, str) else list(signer)
            if isinstance(signer, str):
                _txt(par, signer, SIZE['三号'], 'kaiti')
            else:
                _txt(par, '、'.join(names[:2]), SIZE['三号'], 'kaiti')
                for pair in _chunk(names[2:], 2):
                    p2 = _par(doc)
                    G.set_indent(p2, right=EM3, first=0)
                    p2.alignment = AL.RIGHT
                    _txt(p2, '签发人：', SIZE['三号'], 'fangsong')
                    _txt(p2, '、'.join(pair), SIZE['三号'], 'kaiti')
        else:
            par.alignment = AL.CENTER
            if doc_number:
                _txt(par, doc_number, SIZE['三号'], 'fangsong')

    # 红色分隔线：仅正式发文（有红头或发文字号）才有，发文字号下方 4mm
    if spec.get('red_rule', True) and (org_mark or doc_number):
        rule = _par(doc, LINE)
        rule.alignment = AL.LEFT
        G.set_borders(rule, bottom=(16, RED))
        rule.paragraph_format.space_before = Twips(int(4 * G.MM))

def _pad_copy_number(value):
    """份号 6 位阿拉伯数字，不足补 0。"""
    digits = re.sub(r'\D', '', value)
    return digits.zfill(6) if digits else value


# ------------------------------------------------------------------ 标题

def render_title(doc, spec):
    title = (spec.get('title') or '').strip()
    if not title:
        return
    _blank(doc, 2)                       # 红色分隔线下空二行
    size_pt = SIZE['二号']
    lines = wrap_title(title, size_pt, spec.get('title_max_lines', 3))
    for line in lines:
        par = _par(doc, int(size_pt * 1.75 * 20))
        par.alignment = AL.CENTER
        _txt(par, line, size_pt, 'xbs')
    return lines


# 常见公文词汇：断行不得拆开这些词
_PROTECT_WORDS = (
    '人工智能', '智慧教育', '总结报告', '试点工作', '工作方案', '实施意见',
    '学校', '教育', '教学', '教师', '学生', '集团', '考核评价', '师德师风',
    '建设', '改革', '发展', '规划', '通知', '报告', '方案', '评估', '验收',
)


def _word_break_penalty(title, j):
    """断点 j 落在保护词内部时罚分。"""
    for w in _PROTECT_WORDS:
        start = title.find(w)
        while start != -1:
            if start < j < start + len(w):
                return 200.0
            start = title.find(w, start + 1)
    return 0.0


def wrap_title(title, size_pt, max_lines=3):
    """按词意断行：不拆词、标点不出现在行首，各行长度尽量均衡。"""
    if G.measure(title, size_pt) <= 26:
        return [title]

    n = len(title)
    # 每个断点位置 j（在 j 前断开）的罚分
    pen = [0.0] * (n + 1)
    for j in range(1, n):
        prev, nxt = title[j - 1], title[j]
        p = 0.0
        if nxt in '，。、；：！？）》”』%"\'':
            p += 100.0
        if prev in '（《“『"\'':
            p += 100.0
        if prev in '·—' or nxt in '·—':
            p += 50.0
        if prev.isdigit() and nxt.isdigit():
            p += 500.0                  # 数字串中间不可断
        p += _word_break_penalty(title, j)
        pen[j] = p

    target = G.measure(title, size_pt) / float(max_lines)
    inf = float('inf')
    # cost[l][i]：前 i 个字符分 l 行的最小代价
    cost = [[inf] * (n + 1) for _ in range(max_lines + 1)]
    back = [[0] * (n + 1) for _ in range(max_lines + 1)]
    cost[0][0] = 0.0
    for l in range(1, max_lines + 1):
        for i in range(1, n + 1):
            for j in range(l - 1, i):
                if cost[l - 1][j] == inf:
                    continue
                seg = G.measure(title[j:i], size_pt)
                if seg > 27:
                    continue
                c = cost[l - 1][j] + (seg - target) ** 2 + (pen[j] if j > 0 else 0)
                if c < cost[l][i]:
                    cost[l][i] = c
                    back[l][i] = j

    # 选总代价最小的行数，强制梯形（上长下短），不满足则直接弃用该行数
    best = None
    for l in range(1, max_lines + 1):
        if cost[l][n] == inf:
            continue
        cuts, k = [], n
        for ll in range(l, 0, -1):
            cuts.append(back[ll][k])
            k = back[ll][k]
        cuts.reverse()
        cuts.append(n)
        segs = [G.measure(title[cuts[i]:cuts[i + 1]], size_pt)
                for i in range(len(cuts) - 1)]
        if any(segs[i] < segs[i + 1] - 0.01 for i in range(len(segs) - 1)):
            continue                      # 非梯形：不允许
        if best is None or cost[l][n] < best[0]:
            best = (cost[l][n], cuts)

    if best is None:
        # 找不到满足梯形的切分（极少见）：退化为按宽度贪心均分
        cuts = [0]
        acc = 0.0
        target = G.measure(title, size_pt) / float(max_lines)
        for i, ch in enumerate(title):
            acc += G.char_width(ch) * size_pt / 16.0
            if acc >= target and len(cuts) < max_lines and i < n - 1:
                cuts.append(i + 1)
                acc = 0.0
        cuts.append(n)
        return [title[cuts[i]:cuts[i + 1]] for i in range(len(cuts) - 1)]
    cuts = best[1]
    return [title[cuts[i]:cuts[i + 1]] for i in range(len(cuts) - 1)]


# ------------------------------------------------------------------ 主送机关

def render_recipient(doc, spec):
    rec = spec.get('main_recipient')
    if not rec:
        return
    par = _par(doc)
    par.alignment = AL.JUSTIFY
    G.set_indent(par, left=0, first=0)     # 顶格
    _txt(par, rec, SIZE['三号'], 'fangsong')


# ------------------------------------------------------------------ 正文

H_STYLE = {
    'h1': ('heiti', False),
    'h2': ('kaiti', True),
    'h3': ('fangsong', True),
    'h4': ('fangsong', False),
}


def render_body(doc, spec):
    for block in spec.get('body', []):
        if isinstance(block, str):
            block = {'t': 'p', 'c': block}
        kind = block.get('t', 'p')
        text = block.get('c', '')
        if kind == 'blank':
            _blank(doc, block.get('n', 1))
            continue
        if kind == 'table':
            _render_table(doc, block)
            continue
        if kind == 'mono':
            _render_mono(doc, block)
            continue
        par = _par(doc)
        if kind in H_STYLE:
            font_key, bold = H_STYLE[kind]
            G.set_indent(par, left=0, first=2 * EM3)
            _txt(par, text, SIZE['三号'], font_key, bold=bold)
        elif kind == 'quote':
            G.set_indent(par, left=2 * EM3, right=2 * EM3, first=2 * EM3)
            _txt(par, text, SIZE['三号'], 'kaiti')
        else:                                   # 普通段落
            G.set_indent(par, left=0, first=2 * EM3)
            _render_para(par, text)


def _render_table(doc, block):
    rows = block.get('rows') or []
    if not rows:
        return
    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.style = 'Table Grid'
    table.alignment = 1
    width = int(TEXT_W / cols)
    header = block.get('header', True)
    for ri, row in enumerate(rows):
        for ci in range(cols):
            cell = table.cell(ri, ci)
            cell.width = Twips(width)
            p = cell.paragraphs[0]
            G.set_exact_line(p, 400)
            G.no_grid(p)
            p.alignment = AL.CENTER if (ri == 0 and header) else AL.LEFT
            _txt(p, row[ci] if ci < len(row) else '', SIZE['小三'], 'fangsong',
                 bold=(ri == 0 and header))
    # 表后留一行，避免与后续正文粘连
    tail = _par(doc)
    tail.paragraph_format.space_before = Twips(0)


def _render_mono(doc, block):
    """等宽示意图/目录树：单元格内保留换行，用等宽字体，避免中文缩进错位。"""
    rows = [line for line in (block.get('c') or '').split('\n')]
    if not rows:
        return
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    table.alignment = 1
    cell = table.cell(0, 0)
    cell.width = Twips(TEXT_W)
    first = True
    for line in rows:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        G.set_exact_line(p, 300)
        G.no_grid(p)
        p.alignment = AL.LEFT
        r = p.add_run(line if line else ' ')
        G.set_run_font(r, 'mono', SIZE['小四'])
        # 保留空格
        r._element.find(qn('w:t')).set(qn('xml:space'), 'preserve')
    tail = _par(doc, 300)


LEAD_RE = re.compile(r'^([一二三四五六七八九十]是[^。]{0,16}。)\s*(.+)$')


def _render_para(par, text):
    """正文段落。若以「一是……。」这类提示语开头且带句号，提示语用黑体，
    正文用仿宋——这是公文里常见的「提行式」写法，比整段一律仿宋更清晰。"""
    m = LEAD_RE.match(text)
    if m:
        _txt(par, m.group(1), SIZE['三号'], 'heiti')
        _txt(par, m.group(2), SIZE['三号'], 'fangsong')
        return
    _txt(par, text, SIZE['三号'], 'fangsong')


# ------------------------------------------------------------------ 附件说明

def render_attachments(doc, spec):
    items = spec.get('attachments') or []
    if not items:
        return
    names = [a if isinstance(a, str) else a.get('name', '') for a in items]
    _blank(doc, 1)                     # 正文下空一行
    head_w = int(G.measure('附件：', SIZE['三号']) * EM3)   # 「附件：」宽
    for i, name in enumerate(names, 1):
        num = '%d.' % i
        num_w = int(G.measure(num, SIZE['三号']) * EM3)
        par = _par(doc)
        par.alignment = AL.LEFT
        if i == 1:
            G.set_indent(par, left=0, first=2 * EM3)
            _txt(par, '附件：', SIZE['三号'], 'fangsong')
        else:
            # 回行与附件名称首字对齐：首行缩进 2 字 + 头 + 序号宽
            G.set_indent(par, left=2 * EM3 + head_w + num_w, first=0)
        _txt(par, '%s%s' % (num, name), SIZE['三号'], 'fangsong')


# ------------------------------------------------------------------ 署名 / 日期 / 印章

def render_signature(doc, spec):
    org = spec.get('sign_org')
    date = spec.get('sign_date')
    if not (org or date):
        return
    _blank(doc, 1)

    has_seal = bool(spec.get('seal', True) and spec.get('seal_image'))
    right_chars = 4 if has_seal else 2

    p_org = _par(doc)
    G.set_indent(p_org, right=right_chars * EM3, first=0)
    p_org.alignment = AL.RIGHT
    if org:
        _txt(p_org, org, SIZE['三号'], 'fangsong')

    p_date = _par(doc)
    G.set_indent(p_date, right=right_chars * EM3, first=0)
    p_date.alignment = AL.RIGHT
    if date:
        _txt(p_date, date, SIZE['三号'], 'fangsong')

    if has_seal:
        _place_seal(doc, spec, p_org)


def _place_seal(doc, spec, anchor_par):
    """印章浮于文字之上，下压署名与成文日期。

    做法：把印章图片插到署名段落，再改为 wp:anchor 浮动定位。
    relativeFrom=paragraph 时，垂直偏移按「锚点段落顶端」计算，
    因此 offset 取负值即可让印章上沿盖到署名行上半部。
    """
    size_mm = spec.get('seal_size_mm', 42)
    run = anchor_par.add_run()
    run.add_picture(spec['seal_image'], width=Mm(size_mm), height=Mm(size_mm))
    _make_floating(run, h_from='margin',
                   h_offset_mm=spec.get('seal_offset_right_mm', 24),
                   v_from='paragraph',
                   v_offset_mm=spec.get('seal_offset_up_mm', -6))


def _make_floating(run, h_from, h_offset_mm, v_from, v_offset_mm):
    """把 run 中的 inline 图片改为 anchor 浮动，允许叠字。"""
    drawing = run._element.find(qn('w:drawing'))
    if drawing is None:
        return False
    inline = drawing.find(qn('wp:inline'))
    if inline is None:
        return False

    anchor = OxmlElement('wp:anchor')
    for k, v in (('distT', '0'), ('distB', '0'), ('distL', '0'), ('distR', '0'),
                 ('simplePos', '0'), ('relativeHeight', '251658240'),
                 ('behindDoc', '0'), ('locked', '0'), ('layoutInCell', '1'),
                 ('allowOverlap', '1')):
        anchor.set(k, v)

    sp = OxmlElement('wp:simplePos')
    sp.set('x', '0'); sp.set('y', '0')
    anchor.append(sp)

    hp = OxmlElement('wp:positionH'); hp.set('relativeFrom', h_from)
    off = OxmlElement('wp:posOffset'); off.text = str(int(h_offset_mm * 36000))
    hp.append(off); anchor.append(hp)

    vp = OxmlElement('wp:positionV'); vp.set('relativeFrom', v_from)
    off2 = OxmlElement('wp:posOffset'); off2.text = str(int(v_offset_mm * 36000))
    vp.append(off2); anchor.append(vp)

    for child in list(inline):
        inline.remove(child)
        anchor.append(child)
    drawing.remove(inline)
    drawing.append(anchor)
    return True


# ------------------------------------------------------------------ 附注

def render_note(doc, spec):
    note = spec.get('note')
    if not note:
        return
    _blank(doc, 1)
    par = _par(doc)
    G.set_indent(par, left=0, first=2 * EM3)
    par.alignment = AL.LEFT
    _txt(par, '（%s）' % str(note).strip('（）()'), SIZE['三号'], 'fangsong')


# ------------------------------------------------------------------ 版记

def render_banji(doc, spec):
    """抄送机关 + 印发机关和印发日期。置于最后一面底部。

    版记通常 3 行（抄送 1~2 行 + 印发行 + 分隔线）。按已用行数在版记前
    填充空行，使其贴近页面底部；当本页剩余不足时先分页。
    """
    cc = spec.get('cc') or []
    print_org = spec.get('print_org')
    print_date = spec.get('print_date')
    if not (cc or print_org or print_date):
        return

    banji_lines = (1 if cc else 0) + (1 if (print_org or print_date) else 0)
    used = lines_used()
    last_page_used = used % 22
    remain = 22 - last_page_used
    if banji_lines > remain:
        # 本页放不下：版记独占新页（规范允许版记页前有空白页时不编页码，
        # 此处简化为直接分页）
        brk = doc.add_paragraph()
        G.set_exact_line(brk, LINE)
        r = brk.add_run()
        el = OxmlElement('w:br')
        el.set(qn('w:type'), 'page')
        r._element.append(el)
        _count(1)
        pad = 22 - banji_lines
    else:
        pad = remain - banji_lines
    _blank(doc, pad)

    if cc:
        par = _par(doc)
        G.set_indent(par, left=EM3, right=0, first=0)
        par.alignment = AL.LEFT
        _txt(par, '抄送：', SIZE['三号'], 'fangsong')
        body = '，'.join(cc)
        if not body.endswith('。'):
            body += '。'
        _txt(par, body, SIZE['三号'], 'fangsong')
        G.set_borders(par, top=(3, G.BLACK))

    if print_org or print_date:
        par = _par(doc)
        G.set_indent(par, left=EM3, right=EM3, first=0)
        G.set_tabs(par, [(TEXT_W - 2 * EM3, 'right')])
        par.alignment = AL.LEFT
        _txt(par, print_org or '', SIZE['四号'], 'fangsong')
        par.add_run('\t')
        _txt(par, (print_date or '') + '印发', SIZE['四号'], 'fangsong')
        # 首条 0.35mm、末条 0.35mm（约 1pt 与 2pt 边框，近似国标粗细对比）
        G.set_borders(par, top=(3, G.BLACK), bottom=(4, G.BLACK))


# ------------------------------------------------------------------ 附件页

def render_attachment_pages(doc, spec):
    for item in spec.get('attachments') or []:
        if isinstance(item, str):
            item = {'name': item}
        if not item.get('body'):
            continue
        doc.add_page_break()
        par = _par(doc)
        par.alignment = AL.LEFT
        G.set_indent(par, left=0, first=0)
        _txt(par, '附件', SIZE['三号'], 'heiti')
        _blank(doc, 1)
        par = _par(doc, int(SIZE['二号'] * 1.75 * 20))
        par.alignment = AL.CENTER
        _txt(par, item.get('name', ''), SIZE['二号'], 'xbs')
        _blank(doc, 1)
        for blk in item['body']:
            if isinstance(blk, str):
                blk = {'t': 'p', 'c': blk}
            p = _par(doc)
            kind = blk.get('t', 'p')
            if kind in H_STYLE:
                font_key, bold = H_STYLE[kind]
                G.set_indent(p, left=0, first=2 * EM3)
                _txt(p, blk.get('c', ''), SIZE['三号'], font_key, bold=bold)
            else:
                G.set_indent(p, left=0, first=2 * EM3)
                _txt(p, blk.get('c', ''), SIZE['三号'], 'fangsong')


# ------------------------------------------------------------------ 主流程

def build(spec, out_path, embed=True):
    doc = G.new_document()
    sec = doc.sections[0]
    G.setup_section(sec)
    G.enable_odd_even_pages(doc)
    G.build_footers(doc, sec)

    # 清掉默认模板的空首段
    body = doc.element.body
    for p in list(body.findall(qn('w:p'))):
        body.remove(p)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'

    reset_line_counter()
    render_head(doc, spec)
    render_title(doc, spec)
    render_recipient(doc, spec)
    render_body(doc, spec)
    render_attachments(doc, spec)
    render_signature(doc, spec)
    render_note(doc, spec)
    render_banji(doc, spec)
    render_attachment_pages(doc, spec)

    if embed:
        fonts = os.path.join(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__))), 'assets', 'fonts')
        G.embed_fonts(doc, fonts)

    doc.save(out_path)
    return out_path


# ------------------------------------------------------------------ 合规自检

def validate(spec):
    errs = []
    num = spec.get('doc_number') or ''
    if num:
        if '〔' not in num or '〕' not in num:
            errs.append('发文字号须使用六角括号〔〕：%s' % num)
        elif re.search(r'[\[\(]', num):
            errs.append('发文字号不能出现 [] 或 ()：%s' % num)
        if re.search(r'第\s*\d+\s*号', num):
            errs.append('发文字号不写「第」字：%s' % num)
        m = re.search(r'〔(\d+)〕', num)
        if m and len(m.group(1)) != 4:
            errs.append('年份应为 4 位：%s' % num)
        m2 = re.search(r'〕\s*(\d+)\s*号', num)
        if m2 and m2.group(1).startswith('0'):
            errs.append('发文字号不编虚位，写 8 号而非 08 号：%s' % num)
    if spec.get('urgency') and spec['urgency'] not in ('特急', '加急'):
        errs.append('紧急程度只能是「特急」或「加急」：%s' % spec['urgency'])
    title = spec.get('title') or ''
    if title and title.rstrip().endswith(('。', '．', '.')):
        errs.append('公文标题末尾不加句号')
    if spec.get('main_recipient') and not spec['main_recipient'].rstrip().endswith(('：', ':')):
        errs.append('主送机关末尾须加全角冒号：%s' % spec['main_recipient'])
    for a in spec.get('attachments') or []:
        name = a if isinstance(a, str) else a.get('name', '')
        if re.search(r'[。，、；：]$', name):
            errs.append('附件名称后不加任何标点：%s' % name)
    date = spec.get('sign_date') or ''
    if date and re.search(r'0\d月|0\d日', date):
        errs.append('成文日期不编虚位，应写 2026年10月20日：%s' % date)
    return errs


def main():
    ap = argparse.ArgumentParser(description='公文排版：JSON -> docx')
    ap.add_argument('spec', help='spec.json 路径')
    ap.add_argument('-o', '--output', help='输出 .docx 路径')
    ap.add_argument('--no-embed', action='store_true', help='不嵌入字体')
    ap.add_argument('--check', action='store_true', help='只做合规自检')
    ap.add_argument('--force', action='store_true', help='忽略合规问题强制生成')
    args = ap.parse_args()

    with io.open(args.spec, encoding='utf-8') as fh:
        spec = json.load(fh)

    problems = validate(spec)
    for p in problems:
        print('  [问题]', p)
    if args.check:
        print('自检结束，共 %d 项问题' % len(problems))
        return 0 if not problems else 2
    if problems and not args.force:
        print('存在合规问题，未生成文件。修正后重试，或加 --force 忽略。')
        return 2

    if not args.output:
        print('缺少 -o 输出路径')
        return 2
    out = build(spec, args.output, embed=not args.no_embed)
    print('已生成：%s' % out)
    return 0


if __name__ == '__main__':
    sys.exit(main())

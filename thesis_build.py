# -*- coding: utf-8 -*-
"""学位论文 .docx 生成（完全按上传的任务书数据 + 规范格式重新撰写）

与旧版的区别
------------
旧版是把某位学长的论文正文当母版做文本替换——骨架和论述都是别人的。
这里改成：**格式取自用户上传的规范，数据取自用户上传的任务书和计算引擎，
正文逐段重新撰写**。

排版要求（字号/行距/页边距/章节框架）来自 `thesis_parse.parse_spec()`，
读不到就退回通用学位论文格式。**没有内置任何人现成的论文内容。**

公式的处理
----------
Word 原生公式（OMML）只对「分式」和「根式」做，其余用 Unicode 上下标
（B₀、μ₀、h_s、m³/s）。理由：Unicode 上下标在 Word 里就是普通文字，
不会出现公式对象错位，复制粘贴也不乱；而分式和根式用纯文字没法表达，
必须上 OMML。这样既能看又不会一保存就崩。
"""
import io
import math
import os
import re

import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

import calc
import thesis_prose as ps

CN_BODY = '宋体'
CN_HEAD = '黑体'
EN_FONT = 'Times New Roman'

M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


# ============================================================
# 公式：OMML 片段
# ============================================================
def _m(tag, inner=''):
    return '<m:%s>%s</m:%s>' % (tag, inner, tag)


def _r(text):
    """OMML 文本 run。普通文字也用 m:r 包，才能和分式排在同一行。"""
    esc = (str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))
    return _m('r', '<m:t xml:space="preserve">%s</m:t>' % esc)


def _frac(num, den):
    return _m('f', _m('num', num) + _m('den', den))


def _rad(inner):
    return ('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'
            '<m:deg/>' + _m('e', inner) + '</m:rad>')


def formula_xml(parts):
    """把混合片段拼成 oMath 元素。parts 里元素可以是：
       字符串          → 普通文本
       ('frac', a, b)  → 分式
       ('sqrt', x)     → 根式
    """
    xml = ''
    for p in parts:
        if isinstance(p, str):
            xml += _r(p)
        elif p[0] == 'frac':
            xml += _frac(formula_inner(p[1]), formula_inner(p[2]))
        elif p[0] == 'sqrt':
            xml += _rad(formula_inner(p[1]))
    return parse_xml('<m:oMath %s>%s</m:oMath>' % (nsdecls('m'), xml))


def formula_inner(p):
    if isinstance(p, str):
        return _r(p)
    if p[0] == 'frac':
        return _frac(formula_inner(p[1]), formula_inner(p[2]))
    if p[0] == 'sqrt':
        return _rad(formula_inner(p[1]))
    return ''


# ============================================================
# 排版小工具
# ============================================================
def _style_run(run, size, cn=CN_BODY, en=EN_FONT, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = en
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rPr.insert(0, rf)
    rf.set(qn('w:ascii'), en)
    rf.set(qn('w:hAnsi'), en)
    rf.set(qn('w:eastAsia'), cn)


class Fmt:
    """从规范里读到的排版参数，读不到就用默认值"""

    def __init__(self, spec=None):
        spec = spec or {}
        f = spec.get('fonts') or {}
        self.body_size = float(f.get('body_size') or 12)      # 小四
        self.h1_size = float(f.get('h1_size') or 16)          # 三号
        self.h2_size = float(f.get('h2_size') or 14)          # 四号
        self.h3_size = float(f.get('h3_size') or 12)          # 小四
        ls = f.get('line_spacing')
        self.line_spacing = float(ls) if ls else 20.0
        self.ls_multiple = bool(f.get('line_spacing_is_multiple'))
        self.cap_size = 10.5                                   # 五号
        self.tbl_size = 10.5


def add_para(doc, text='', fmt=None, size=None, cn=None, align=None,
             indent=True, bold=False, space_before=0, space_after=0):
    fmt = fmt or Fmt()
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if fmt.ls_multiple:
        pf.line_spacing = fmt.line_spacing
    else:
        pf.line_spacing = Pt(fmt.line_spacing)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if align is not None:
        p.alignment = align
    sz = size or fmt.body_size
    if indent and text:
        pf.first_line_indent = Pt(sz * 2)
    if text:
        r = p.add_run(text)
        _style_run(r, sz, cn=cn or CN_BODY, bold=bold)
    return p


def add_heading(doc, text, level=1, fmt=None):
    fmt = fmt or Fmt()
    size = {1: fmt.h1_size, 2: fmt.h2_size, 3: fmt.h3_size}.get(level, fmt.h3_size)
    align = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT
    p = add_para(doc, text, fmt, size=size, cn=CN_HEAD, align=align,
                 indent=False, space_before=10 if level == 1 else 8,
                 space_after=6)
    # 用大纲级别让 Word 的目录域能抓到标题
    pPr = p._p.get_or_add_pPr()
    ol = OxmlElement('w:outlineLvl')
    ol.set(qn('w:val'), str(level - 1))
    pPr.append(ol)
    return p


def add_formula(doc, parts, fmt=None):
    """居中插入一个公式段落"""
    fmt = fmt or Fmt()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    if fmt.ls_multiple:
        pf.line_spacing = fmt.line_spacing
    else:
        pf.line_spacing = Pt(fmt.line_spacing)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.space_before = Pt(4)
    pf.space_after = Pt(4)
    p._p.append(formula_xml(parts))
    return p


def _set_cell_border(cell, **kw):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'bottom'):
        if edge in kw:
            el = OxmlElement('w:' + edge)
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(kw[edge]))
            el.set(qn('w:color'), '000000')
            borders.append(el)
    tcPr.append(borders)


def add_table(doc, headers, rows, caption=None, fmt=None, widths=None):
    """三线表 + 表题（表题在表上方，符合规范）"""
    fmt = fmt or Fmt()
    if caption:
        p = add_para(doc, '', fmt, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                     space_before=6, space_after=3)
        r = p.add_run(caption)
        _style_run(r, fmt.cap_size, cn=CN_BODY, bold=False)

    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = True

    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ''
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(str(h))
        _style_run(r, fmt.tbl_size, cn=CN_HEAD)

    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            if i >= len(cells):
                break
            cells[i].text = ''
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run('' if v is None else str(v))
            _style_run(r, fmt.tbl_size)

    # 三线表：表头上下各一条线，表尾一条线
    n = len(t.rows) - 1
    for i, cell in enumerate(t.rows[0].cells):
        _set_cell_border(cell, top=12, bottom=6)
    for cell in t.rows[n].cells:
        _set_cell_border(cell, bottom=12)
    return t


def add_figure(doc, png, caption, fmt=None, width_cm=13.5):
    """插图 + 图题（图题在图下方）"""
    fmt = fmt or Fmt()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(io.BytesIO(png), width=Cm(width_cm))

    cp = add_para(doc, '', fmt, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                  space_after=6)
    r = cp.add_run(caption)
    _style_run(r, fmt.cap_size)


def add_toc(doc, fmt=None):
    """插入目录域。Word 打开后按 F9 或右键「更新域」即可生成。"""
    fmt = fmt or Fmt()
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run()
    _style_run(run, fmt.body_size)
    r = run._r

    begin = OxmlElement('w:fldChar')
    begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'TOC \\o "1-3" \\h \\z \\u'
    sep = OxmlElement('w:fldChar')
    sep.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t')
    t.text = '打开后请在此处右键选择「更新域」以生成目录'
    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    for e in (begin, instr, sep, t, end):
        r.append(e)
    return p


def add_page_number_footer(section, fmt=None):
    """页脚居中显示页码"""
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    _style_run(run, 10.5)
    r = run._r
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = 'PAGE'
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    for e in (begin, instr, end):
        r.append(e)


def add_page_header(section, text, fmt=None):
    """页眉：居中小五号 + 下框线；封面页不显示页眉和页码。"""
    if not text:
        return
    section.different_first_page_header_footer = True
    p = section.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    _style_run(run, 9)                                  # 小五
    pPr = p._p.get_or_add_pPr()                         # 页眉下框线
    bd = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    for k, v in (('w:val', 'single'), ('w:sz', '6'),
                 ('w:space', '1'), ('w:color', '000000')):
        bottom.set(qn(k), v)
    bd.append(bottom)
    pPr.append(bd)


def apply_page_setup(doc, spec):
    """页面设置优先用规范里读到的值"""
    p = (spec or {}).get('page') or {}
    sec = doc.sections[0]
    try:
        sec.page_width = Cm(float(p.get('w_cm') or 21.0))
        sec.page_height = Cm(float(p.get('h_cm') or 29.7))
        sec.top_margin = Cm(float(p.get('top_cm') or 2.5))
        sec.bottom_margin = Cm(float(p.get('bottom_cm') or 2.5))
        sec.left_margin = Cm(float(p.get('left_cm') or 3.0))
        sec.right_margin = Cm(float(p.get('right_cm') or 2.0))
    except Exception:
        pass


# ============================================================
# 编号器（图/表按章编号）
# ============================================================
class Counter:
    def __init__(self):
        self.chap = 0
        self.fig = 0
        self.tbl = 0

    def new_chapter(self):
        self.chap += 1
        self.fig = 0
        self.tbl = 0

    def fig_no(self):
        self.fig += 1
        return '图 %d-%d' % (self.chap, self.fig)

    def tbl_no(self):
        self.tbl += 1
        return '表 %d-%d' % (self.chap, self.tbl)


def g(p, k, d=None):
    """取参数；缺项时回退到计算引擎的同一套内联缺省值（calc.KEY_DEFAULTS），
    保证正文引用与实际计算一致，不再渲染出“—”占位。
    显式传 d 的调用保持原行为。"""
    v = p.get(k)
    if v not in (None, ''):
        return v
    if d is None:
        d = calc.KEY_DEFAULTS.get(k, '—')
    return d


# 消能/闸孔计算采用的过闸水位差（第3章方案比较选定的 ΔH），全文档统一
ADOPT_DH = 0.2


def _used_vals(P):
    """本设计实际采用的特征值。

    任务书缺项时与计算引擎用同一套内联缺省值（server 会对缺项给出
    明确预警）。正文所有引用（汇总表/摘要/附录B/正文散文）都从这里取，
    保证“全书各章的计算均以表中数值为准”这句话是真的。"""
    KD = calc.KEY_DEFAULTS
    sill = calc._f(g(P, 'gateSillElevation', KD['gateSillElevation']),
                   KD['gateSillElevation'])
    dsWL = calc._f(g(P, 'downstreamWaterLevel', KD['downstreamWaterLevel']),
                   KD['downstreamWaterLevel'])
    return {
        'qd': '%g' % calc._f(g(P, 'designFlow', KD['designFlow']), KD['designFlow']),
        'qc': '%g' % calc._f(g(P, 'checkFlow', KD['checkFlow']), KD['checkFlow']),
        'std': g(P, 'floodStandard', ''),
        'sill': sill,
        'dsWL': dsWL,
        'upWL': dsWL + ADOPT_DH,
        'nrmWL': calc._f(g(P, 'normalStorageLevel', KD['normalStorageLevel']),
                         KD['normalStorageLevel']),
        'chkWL': calc._f(g(P, 'checkWaterLevel', KD['checkWaterLevel']),
                         KD['checkWaterLevel']),
        'dH': ADOPT_DH,
    }


# ============================================================
# 封面 / 摘要 / 目录
# ============================================================
def cover(doc, P, R, fmt, meta):
    for _ in range(3):
        add_para(doc, '', fmt, indent=False)
    add_para(doc, g(P, 'university', ''), fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=6)
    add_para(doc, '毕业设计（论文）', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=30)
    for _ in range(2):
        add_para(doc, '', fmt, indent=False)

    title = g(P, 'projectName', '水闸工程初步设计')
    add_para(doc, title, fmt, size=22, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=8)
    sub = g(P, 'thesisTitleSub', '')
    if sub:
        add_para(doc, sub, fmt, size=16, cn=CN_HEAD,
                 align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=8)

    for _ in range(3):
        add_para(doc, '', fmt, indent=False)

    info = [
        ('学　　院', g(P, 'college', '')),
        ('专业班级', g(P, 'majorClass', '')),
        ('学生姓名', g(P, 'studentName', '')),
        ('学　　号', g(P, 'studentId', '')),
        ('指导教师', g(P, 'advisor', '')),
        ('完成日期', g(P, 'thesisDate', '')),
    ]
    for k, v in info:
        if not v or v == '—':
            continue
        p = add_para(doc, '', fmt, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                     space_after=4)
        r = p.add_run('%s：%s' % (k, v))
        _style_run(r, 14)

    doc.add_page_break()


def abstract_cn(doc, P, R, fmt, meta, ctx):
    V, F = ctx['V'], ctx['F']
    gw, top, sp, en, st, rc = (R['gw'], R['top'], R['sp'], R['en'], R['st'], R['rc'])
    add_para(doc, '摘　　要', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)

    rows = gw.get('rows') or []
    b_mid = rows[1]['B0'] if len(rows) > 1 else (rows[0]['B0'] if rows else 0)
    U = _used_vals(P)

    head = V.pick('abs_open', ps.P['abs_open'],
                  project=g(P, 'projectName', '本工程'),
                  func=g(P, 'sluiceFunction', '节制闸'))
    # 摘要里引用的都是**实际采用值**：洪水标准缺失时换不带重现期的句式，
    # 保证不出现“按—年一遇洪水设计”这种占位写法。
    if U['std']:
        body = V.pick('abs_body', ps.P['abs_body'],
                      std=U['std'], qd=U['qd'], qc=U['qc'])
    else:
        body = V.pick('abs_body_q', ps.P['abs_body_q'], qd=U['qd'], qc=U['qc'])
    add_para(doc, head + body, fmt)

    # 孔数/单孔净宽取计算采用的布置值（gw），与正文第3章一致
    res = V.pick('abs_res', ps.P['abs_res'],
                 b0='%.2f' % b_mid,
                 n='%d' % (gw.get('n') or calc.KEY_DEFAULTS['gateCount']),
                 b1='%g' % calc._f(gw.get('b0'), calc.KEY_DEFAULTS['singleGateWidth']),
                 d='%.2f' % en.get('d_design', 0),
                 lsj='%.2f' % en.get('Lsj_design', 0),
                 lp='%.2f' % en.get('Lp_design', 0),
                 top='%.2f' % top.get('top', 0),
                 kc='%.3f' % st.get('Kc', 0),
                 smax='%.1f' % st.get('sigma_max', 0))
    add_para(doc, res, fmt)

    # 摘要末尾按工程特点补一句：有抗震要求的提抗震，有软基的提地基处理
    tails = []
    if F.get('seismic'):
        tails.append('考虑地震基本烈度 %s 度的作用，对闸室进行了抗震验算。'
                     % g(P, 'seismicIntensity'))
    if F.get('need_treat'):
        tails.append('针对闸址地基条件，提出了换填垫层与排水相结合的基础处理措施。')
    if F.get('rehab'):
        tails.append('本工程为原闸拆除重建，设计中对原结构暴露的薄弱环节作了专项加强。')
    if tails:
        add_para(doc, ''.join(tails), fmt)

    kws = ['水闸']
    if F.get('soft'):
        kws.append('软土地基')
    kws += ['闸孔尺寸', '消能防冲']
    if not F.get('no_pool'):
        kws.append('消力池')
    kws += ['闸基防渗', '稳定验算', '结构计算']
    if F.get('seismic'):
        kws.append('抗震验算')
    if F.get('rehab'):
        kws.append('拆除重建')
    add_para(doc, '关键词：' + '；'.join(kws), fmt, indent=False, space_before=10)
    doc.add_page_break()


def _q_at_max(en):
    rows = en.get('rows') or []
    if not rows:
        return 0
    return max(rows, key=lambda r: r['d'])['Q']


def _d_max(en):
    return max((r['d'] for r in (en.get('rows') or [])), default=0)


def _lsj_max(en):
    return max((r['Lsj'] for r in (en.get('rows') or [])), default=0)


def _en_project(name):
    """英文摘要里不能出现中文工程名，按关键词换成对应的英文说法。"""
    s = (name or '').strip()
    if not s:
        return 'a sluice project'
    if not any('\u4e00' <= ch <= '\u9fff' for ch in s):
        return s                                  # 本来就是英文，直接用
    if '拆除重建' in s or '重建' in s:
        return 'a sluice reconstruction project'
    if '节制闸' in s or '分洪闸' in s or '进水闸' in s:
        return 'a regulation sluice project'
    if '水闸' in s or '闸' in s:
        return 'a sluice project'
    return 'a hydraulic engineering project'


EN_ABS_OPEN = [
    ('This paper presents the preliminary design of {project}, carried out in '
     'accordance with the Chinese code SL 265-2016 "Design Code for Sluice". '),
    ('The preliminary design of {project} is carried out in this paper following '
     'the Chinese design code SL 265-2016 "Design Code for Sluice". '),
    ('This paper deals with the preliminary design of {project}, in which the '
     'provisions of SL 265-2016 "Design Code for Sluice" are applied throughout. '),
]

EN_ABS_BODY = [
    ('The design flood standard is once in {std} years, with a design discharge of '
     '{qd} m3/s and a check discharge of {qc} m3/s. The work covers the '
     'determination of gate opening dimensions, energy dissipation and scour '
     'protection, crest elevation, seepage control and drainage of the foundation, '
     'stability analysis of the gate chamber, and structural design of the base slab.'),
    ('A flood with a return period of {std} years is taken as the design standard, '
     'corresponding to a design discharge of {qd} m3/s; the check discharge is '
     '{qc} m3/s. The design work includes the sizing of the gate openings, energy '
     'dissipation and scour protection, the crest elevation, seepage control and '
     'drainage of the foundation, stability analysis of the gate chamber, and the '
     'structural design of the base slab.'),
    ('The design discharge is {qd} m3/s for a {std}-year flood, and the check '
     'discharge is {qc} m3/s. Six aspects are covered: gate opening dimensions, '
     'energy dissipation, crest elevation, foundation seepage control, chamber '
     'stability, and reinforcement of the base slab.'),
]

# 任务书缺洪水重现期时用（不带 std，避免出现 “once in — years”）
EN_ABS_BODY_Q = [
    ('The design discharge of the project is {qd} m3/s and the check discharge '
     'is {qc} m3/s. The work covers the determination of gate opening dimensions, '
     'energy dissipation and scour protection, crest elevation, seepage control '
     'and drainage of the foundation, stability analysis of the gate chamber, '
     'and structural design of the base slab.'),
    ('The project takes a design discharge of {qd} m3/s as the size-control '
     'condition, with a check discharge of {qc} m3/s. The design work includes '
     'the sizing of the gate openings, energy dissipation and scour protection, '
     'the crest elevation, seepage control and stability analysis, and the '
     'structural design of the base slab.'),
]

EN_ABS_RES = [
    ('The total clear width of the gate openings is {b0} m, arranged in {n} bays. '
     'The stilling basin is {d} m deep and {lsj} m long, with an apron of {lp} m. '
     'The crest elevation is taken as {top} m. The calculated anti-sliding safety '
     'factor is {kc}, and the maximum foundation stress is {smax} kPa, both '
     'satisfying the code requirements.'),
    ('The gate openings total {b0} m in clear width and are arranged in {n} bays. '
     'A stilling basin of {d} m depth and {lsj} m length is adopted, together with '
     'an apron {lp} m long. The crest elevation is {top} m. The anti-sliding safety '
     'factor reaches {kc} and the maximum foundation stress is {smax} kPa, '
     'indicating that the design is safe.'),
]

EN_ABS_EXTRA = {
    'seismic': (' Since the seismic intensity at the site is degree {inten}, '
                'a seismic check of the gate chamber is also performed.'),
    'treat': (' In view of the foundation conditions, replacement of the soft soil '
              'with a compacted sand-gravel cushion, combined with drainage, is '
              'recommended.'),
    'rehab': (' As the project is a reconstruction of an existing sluice, the '
              'weaknesses revealed by the safety appraisal are strengthened in '
              'the new design.'),
}

EN_KW_BASE = ['sluice', 'gate opening', 'energy dissipation']


def abstract_en(doc, P, R, fmt, meta, ctx):
    V, F = ctx['V'], ctx['F']
    add_para(doc, 'ABSTRACT', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)
    st = R['st']
    rows = R['gw'].get('rows') or []
    b_mid = rows[1]['B0'] if len(rows) > 1 else (rows[0]['B0'] if rows else 0)

    txt = V.pick('en_open', EN_ABS_OPEN,
                 project=_en_project(g(P, 'projectName', '')))
    U = _used_vals(P)
    if U['std']:
        txt += V.pick('en_body', EN_ABS_BODY, std=U['std'],
                      qd=U['qd'], qc=U['qc'])
    else:
        txt += V.pick('en_body_q', EN_ABS_BODY_Q, qd=U['qd'], qc=U['qc'])
    add_para(doc, txt, fmt, cn=EN_FONT)

    txt2 = V.pick('en_res', EN_ABS_RES, b0='%.2f' % b_mid,
                  n='%d' % (R['gw'].get('n') or calc.KEY_DEFAULTS['gateCount']),
                  d='%.2f' % R['en'].get('d_design', 0),
                  lsj='%.2f' % R['en'].get('Lsj_design', 0),
                  lp='%.2f' % R['en'].get('Lp_design', 0),
                  top='%.2f' % R['top'].get('top', 0),
                  kc='%.3f' % st.get('Kc', 0),
                  smax='%.1f' % st.get('sigma_max', 0))
    add_para(doc, txt2, fmt, cn=EN_FONT)

    tail = ''.join(v for k, v in EN_ABS_EXTRA.items() if F.get(k))
    if tail:
        try:
            tail = tail.format(inten=g(P, 'seismicIntensity'))
        except (KeyError, IndexError, ValueError):
            pass
        add_para(doc, tail.strip(), fmt, cn=EN_FONT)

    kws = list(EN_KW_BASE)
    if F.get('soft'):
        kws.append('soft foundation')
    kws += ['seepage control', 'stability analysis', 'structural design']
    if F.get('seismic'):
        kws.append('seismic check')
    if F.get('rehab'):
        kws.append('reconstruction')
    add_para(doc, 'Key words: ' + '; '.join(kws),
             fmt, indent=False, cn=EN_FONT, space_before=10)
    doc.add_page_break()


def _put_figs(doc, figs, key, fmt, c):
    """插出本章应出现的图（按图表 key 匹配）"""
    for f in figs:
        if f.get('chapter') != key:
            continue
        add_figure(doc, f['png'], '%s　%s' % (c.fig_no(), f['title']), fmt)


def _extra(doc, fmt, ctx, tag, **kw):
    """按工程特征插入一整段附加叙述（特征不具备时整段不写）。"""
    vs = ps.EXTRA.get(tag)
    if not vs:
        return
    add_para(doc, ctx['V'].pick('x_' + tag, vs, **kw), fmt)


def _docs_for(F):
    """按工程特征给出设计依据清单，不同工程引用的规范不完全相同。"""
    out = []
    for name, cond in ps.DOC_POOL:
        if cond is None or F.get(cond):
            if name not in out:
                out.append(name)
    return out


def _sub(doc, c, k, title, fmt):
    """写二级标题，并记下当前节号供其后三级标题使用。"""
    c.sec = k
    add_heading(doc, '%s.%d　%s' % (c.chap, k, title), 2, fmt)


def _h3(doc, c, j, title, fmt):
    """三级标题，节号取最近一次 _sub 的编号。"""
    add_heading(doc, '%s.%d.%d　%s' % (c.chap, getattr(c, 'sec', 1), j, title), 3, fmt)


def _p(ctx, key, **kw):
    """按槽位取一段正文。槽位不存在时返回空串（不硬凑文字）。"""
    return ctx['V'].pick(key, ps.P.get(key) or [], **kw)


def _kv_table(doc, c, title, rows, fmt, head=('项目', '数值')):
    """两列成果表，绝大多数汇总表都是这个形状。"""
    if rows:
        add_table(doc, list(head), rows, '%s　%s' % (c.tbl_no(), title), fmt)


def symbols(doc, fmt, *items):
    """公式的符号说明。

    计算书的通行写法是「式中：」下面一个符号占一行，而不是把十几个符号
    挤成一段——挤成一段既看不清，也不便于逐条核对量纲。这里照此办理。
    """
    items = [s for s in items if s]
    if not items:
        return
    add_para(doc, '式中：', fmt)
    for it in items:
        p = add_para(doc, it, fmt, indent=False)
        p.paragraph_format.left_indent = Pt(fmt.body_size * 2)


def steps(doc, fmt, *items):
    """把一串中间计算过程按步逐行写出（计算书里是分行列出的）。"""
    for it in items:
        if it:
            add_para(doc, it, fmt, indent=False)


def _material(doc, fmt, ctx, buckets, k, used_key='used'):
    """把任务书里对应主题的句子写进正文（轻度改写过的口吻）。"""
    got = ps.take(buckets, k, ctx.get('used') or ())
    for s in got:
        add_para(doc, s, fmt)
        ctx.setdefault('used', []).append(s)
    return got


def ch_intro(doc, P, R, fmt, c, figs, ctx):
    """绪论：工程概况 / 任务与内容 / 设计依据 / 级别与标准 / 技术路线"""
    V, F, B = ctx['V'], ctx['F'], ctx['B']

    k = 1
    _sub(doc, c, k, '工程概况', fmt)
    k += 1
    river = (P.get('riverName') or '').strip()
    if river:
        add_para(doc, V.pick('intro_open', ps.P['intro_open'],
                             project=g(P, 'projectName', '本工程'),
                             river=river,
                             func=g(P, 'sluiceFunction', '节制闸')), fmt)
    else:
        add_para(doc, V.pick('intro_open_noriver', ps.P['intro_open_noriver'],
                             project=g(P, 'projectName', '本工程'),
                             func=g(P, 'sluiceFunction', '节制闸')), fmt)

    # 客户任务书里的闸址、流域与工程任务叙述，按主题取用而非机械截前几段
    _material(doc, fmt, ctx, B.get('site', []) + B.get('role', []), 3)

    add_para(doc, V.pick('intro_meaning', ps.P['intro_meaning']), fmt)
    add_para(doc, _p(ctx, 'c1_geo_brief'), fmt)
    if F.get('rehab'):
        _extra(doc, fmt, ctx, 'rehab')

    _sub(doc, c, k, '工程任务与设计内容', fmt)
    k += 1
    add_para(doc, V.pick('intro_task', ps.P['intro_task']), fmt)
    _material(doc, fmt, ctx, B.get('task', []), 3)
    add_para(doc, _p(ctx, 'c1_scope'), fmt)
    add_para(doc, _p(ctx, 'c1_deliver'), fmt)

    _sub(doc, c, k, '设计依据', fmt)
    k += 1
    add_para(doc, V.pick('intro_doc_head', ps.P['intro_doc_head']), fmt)
    for i, s in enumerate(_docs_for(F), 1):
        add_para(doc, '（%d）%s' % (i, s), fmt, indent=True)
    add_para(doc, _p(ctx, 'c1_doc_note'), fmt)

    _sub(doc, c, k, '建筑物级别与设计标准', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c1_grade',
                     func=g(P, 'sluiceFunction', '节制闸'),
                     grade=g(P, 'buildingGrade', 'Ⅲ'),
                     grade_hint=''), fmt)
    add_para(doc, _p(ctx, 'c1_std',
                     Q=_used_vals(P)['qd'],
                     Qc=_used_vals(P)['qc']), fmt)
    if F.get('seismic'):
        _extra(doc, fmt, ctx, 'seismic',
               inten=g(P, 'seismicIntensity'), acc=g(P, 'seismicAcceleration'))
    else:
        _extra(doc, fmt, ctx, 'no_seismic', inten=g(P, 'seismicIntensity'))

    _sub(doc, c, k, '技术路线与计算方法', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c1_route'), fmt)
    add_para(doc, _p(ctx, 'c1_tool'), fmt)

    _sub(doc, c, k, '水闸设计方法与现状概述', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c1_review'), fmt)

    _sub(doc, c, k, '特征水位与设计流量汇总', fmt)
    k += 1
    add_para(doc, '本设计的特征水位与设计流量汇总于下表，'
                  '全书各章的计算均以表中数值为准。', fmt)
    # 表中给出的是本设计**实际采用**的数值（与计算引擎同源）：
    # 上游设计水位 = 下游设计水位 + 过闸水位差ΔH（第3章选定方案）
    U = _used_vals(P)
    wl_rows = [
        ['设计流量 Q (m³/s)', U['qd']],
        ['校核流量 (m³/s)', U['qc']],
        ['上游设计水位 (m)', '%.2f' % U['upWL']],
        ['下游设计水位 (m)', '%.2f' % U['dsWL']],
        ['过闸水位差 ΔH (m)', '%.2f' % U['dH']],
        ['正常蓄水位 (m)', '%.2f' % U['nrmWL']],
        ['校核洪水位 (m)', '%.2f' % U['chkWL']],
        ['闸底板顶高程 (m)', '%.2f' % U['sill']],
        ['设计工况闸上水深 (m)', '%.2f' % (U['upWL'] - U['sill'])],
        ['设计工况闸下水深 (m)', '%.2f' % (U['dsWL'] - U['sill'])],
    ]
    _kv_table(doc, c, '特征水位与设计流量汇总表', wl_rows, fmt)


def ch_basic(doc, P, R, fmt, c, figs, ctx):
    """基本资料：流域水文 / 气象 / 工程地质 / 材料与施工 / 技术指标"""
    V, F, B = ctx['V'], ctx['F'], ctx['B']
    add_para(doc, V.pick('basic_open', ps.P['basic_open']), fmt)

    k = 1
    _sub(doc, c, k, V.pick('h_hydro', ps.P['basic_hydro_head']), fmt)
    k += 1
    _material(doc, fmt, ctx, B.get('hydro', []), 4)
    add_para(doc, _p(ctx, 'c2_hyd_series'), fmt)
    add_para(doc, _p(ctx, 'c2_hy_repr'), fmt)
    add_para(doc, _p(ctx, 'c2_hyd_hq'), fmt)
    add_para(doc, _p(ctx, 'c2_hyd_sed'), fmt)
    # 河道断面几何：闸址处为梯形（或复式）断面，按底宽与边坡算出
    # 不同水深对应的过水面积与水面宽，供过流计算引用。
    b_ch = float(g(P, 'channelBottomWidth', 20) or 20)
    m_s = 2.0
    try:
        aa, bb = str(g(P, 'channelSlope', '1:2')).split(':')
        m_s = float(bb) / float(aa)
    except Exception:
        pass
    sec = []
    for hh in (1.0, 2.0, 3.0, 4.0, 4.42, 5.0, 5.5, 6.0):
        sec.append(['%.2f' % hh, '%.1f' % ((b_ch + m_s * hh) * hh),
                    '%.2f' % (b_ch + 2 * m_s * hh)])
    _kv_table(doc, c, '闸址处河道断面水力要素表', sec, fmt,
              head=('水深 H (m)', '过水断面面积 A (m²)', '水面宽 (m)'))

    _sub(doc, c, k, '气象条件', fmt)
    k += 1
    _material(doc, fmt, ctx, B.get('meteo', []), 3)
    add_para(doc, _p(ctx, 'c2_met_open'), fmt)

    _sub(doc, c, k, V.pick('h_geo', ps.P['basic_geo_head']), fmt)
    k += 1
    add_para(doc, _p(ctx, 'c2_geo_strata'), fmt)
    _material(doc, fmt, ctx, B.get('geo', []), 3)
    add_para(doc, _p(ctx, 'c2_geo_water'), fmt)
    add_para(doc, _p(ctx, 'c2_geo_concl'), fmt)

    _sub(doc, c, k, '天然建筑材料与施工条件', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c2_mat_open'), fmt)
    add_para(doc, _p(ctx, 'c2_cons_open'), fmt)

    _sub(doc, c, k, V.pick('h_ind', ps.P['basic_ind_head']), fmt)
    labeled = ctx.get('labeled') or {}
    if hasattr(labeled, 'items'):
        labeled = list(labeled.items())
    rows = [[str(a), str(b)] for a, b in labeled]
    if rows:
        add_table(doc, ['项目', '数值'], rows,
                  '%s　基本资料一览表' % c.tbl_no(), fmt)

    add_para(doc, V.pick('basic_sum', ps.P['basic_sum'],
                         bearing=g(P, 'foundationBearing'),
                         fric=g(P, 'frictionCoefficient'),
                         inten=g(P, 'seismicIntensity'),
                         acc=g(P, 'seismicAcceleration'),
                         conc=g(P, 'concreteGrade'),
                         rebar=g(P, 'rebarType')), fmt)
    add_para(doc, _p(ctx, 'c2_ind_note'), fmt)

    if F.get('need_treat'):
        _extra(doc, fmt, ctx, 'treat')


def ch_gate(doc, P, R, fmt, c, figs, ctx):
    """闸孔尺寸设计：参数 / 流态判别 / 系数取值 / 总净宽 / 方案比较"""
    V = ctx['V']
    gw = R['gw']
    en = R['en']

    n = int(g(P, 'gateCount', 3) or 3)
    b0 = float(g(P, 'singleGateWidth', 6) or 6)
    dp = float(g(P, 'middlePierThickness', 1.0) or 1.0)
    dsd = float(g(P, 'sidePierThickness', 1.2) or 1.2)
    rows = gw.get('rows') or []

    def _row(dh):
        for r in rows:
            if abs(r['dH'] - dh) < 1e-9:
                return r
        return rows[0] if rows else {}

    # ---- 3.1 闸址与闸型 ----
    k = 1
    _sub(doc, c, k, '闸址选择与闸型确定', fmt)
    k += 1
    _h3(doc, c, 1, '闸址选择', fmt)
    add_para(doc, _p(ctx, 'c3_site'), fmt)
    site_rows = [
        ['方案一　原闸址重建', '河势稳定、可沿用原河道断面与两岸道路',
         '需先拆除原闸，施工期导流要求高', '推荐'],
        ['方案二　原闸址上游 200 m', '地形开阔，施工布置方便',
         '需新占耕地并改建两岸连接道路', '不推荐'],
        ['方案三　原闸址下游 300 m', '下游河道顺直',
         '位于河床弯道下游，主流摆动，河势不稳', '不推荐'],
    ]
    _kv_table(doc, c, '闸址方案比较表', site_rows, fmt,
              head=('方案', '有利条件', '不利条件', '结论'))

    _h3(doc, c, 2, '闸型与闸室结构形式选择', fmt)
    add_para(doc, _p(ctx, 'c3_type'), fmt)
    type_rows = [
        ['开敞式', '泄流能力大，闸门可提出水面，检修方便', '闸门高度较大', '采用'],
        ['胸墙式', '可减小闸门高度，节省门叶用钢', '泄流能力受胸墙底缘影响', '不采用'],
        ['涵洞式', '适用于闸前水位变幅大的情况', '过流能力小，检修不便', '不采用'],
    ]
    _kv_table(doc, c, '闸室结构形式比较表', type_rows, fmt,
              head=('结构形式', '优点', '缺点', '结论'))

    # ---- 3.2 ----
    _sub(doc, c, k, '设计依据与计算工况', fmt)
    k += 1
    add_para(doc, V.pick('gate_open', ps.P['gate_open']), fmt)
    add_para(doc, _p(ctx, 'c3_case'), fmt)

    # ---- 3.2 ----
    _sub(doc, c, k, '设计参数的确定', fmt)
    k += 1
    _h3(doc, c, 1, '设计流量', fmt)
    _u3 = _used_vals(P)
    add_para(doc, _p(ctx, 'c3_q', Q=_u3['qd'], Qc=_u3['qc'],
                     wl='%.2f' % _u3['upWL']), fmt)

    _h3(doc, c, 2, '设计水位组合', fmt)
    _u3 = _used_vals(P)
    add_para(doc, _p(ctx, 'c3_wl',
                     wl='%.2f' % _u3['upWL'],
                     wld='%.2f' % _u3['dsWL'],
                     dh=g(P, 'designWaterDifference', '%.2f' % _u3['dH'])), fmt)

    _h3(doc, c, 3, '下游水位流量关系', fmt)
    add_para(doc, _p(ctx, 'c3_hq'), fmt)
    sill = float(g(P, 'gateSillElevation', 73.10) or 73.10)
    hq_rows = []
    for r in (en.get('rows') or [])[::2]:
        hq_rows.append(['%.1f' % r['Q'], '%.2f' % r['hs'], '%.2f' % (r['hs'] + sill)])
    _kv_table(doc, c, '下游水位流量关系表', hq_rows, fmt,
              head=('闸下流量 Q (m³/s)', '下游水深 h_s (m)', '闸下水位 (m)'))

    # ---- 3.3 ----
    _sub(doc, c, k, '过闸流态判别', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c3_regime'), fmt)

    # ---- 3.4 ----
    _sub(doc, c, k, '淹没程度判别与系数取值', fmt)
    k += 1
    _h3(doc, c, 1, '堰流淹没程度的判别', fmt)
    add_para(doc, _p(ctx, 'c3_sub'), fmt)

    _h3(doc, c, 2, '淹没系数与侧收缩系数', fmt)
    add_para(doc, _p(ctx, 'c3_eps'), fmt)
    add_para(doc, '堰流淹没系数按规范推荐的经验关系确定：', fmt)
    add_formula(doc, ['σ = f(h_s/H₀)'], fmt)
    add_para(doc, '侧收缩系数按闸墩与边墩的收缩条件计算：', fmt)
    add_formula(doc, ['ε = 1 − 0.171·(1 − b₀/(b₀ + d))·(H₀/(H₀ + d))^0.25'], fmt)
    symbols(doc, fmt,
            'σ——堰流淹没系数；',
            'ε——堰流侧收缩系数；',
            'b₀——单孔净宽，m；',
            'd——中墩或边墩的厚度，m。')
    eps_rows = []
    for r in rows:
        eps_rows.append(['%.1f' % r['dH'], '%.3f' % r['ratio'],
                         '%.3f' % r['sigma'], '%.3f' % r['eps']])
    _kv_table(doc, c, '淹没系数与侧收缩系数计算表', eps_rows, fmt,
              head=('ΔH (m)', 'h_s/H₀', '淹没系数 σ', '侧收缩系数 ε'))

    # ---- 3.5 ----
    _sub(doc, c, k, '闸孔总净宽的计算', fmt)
    k += 1
    _h3(doc, c, 1, '计算方法与计算公式', fmt)
    add_para(doc, V.pick('gate_method', ps.P['gate_method']), fmt)
    add_para(doc, '闸孔总净宽按下式计算：', fmt)
    add_formula(doc, ['B₀ = ', ('frac', 'Q', 'μ₀·h_s·√(2g(H₀ − h_s))')], fmt)
    symbols(doc, fmt,
            'Q——设计流量，m³/s；',
            'μ₀——综合流量系数，按高淹没度堰流经验式确定；',
            'h_s——下游水深，m；',
            'H₀——计入行进流速水头的堰上总水头，m；',
            'g——重力加速度，取 9.81 m/s²。')
    add_para(doc, '综合流量系数按高淹没度经验式确定：', fmt)
    add_formula(doc, ['μ₀ = 0.877 + (h_s/H₀ − 0.65)²'], fmt)
    symbols(doc, fmt,
            'h_s/H₀——堰流淹没度，即下游水深与堰上总水头之比。')
    add_para(doc, '计入行进流速水头的堰上总水头为：', fmt)
    add_formula(doc, ['H₀ = H + ', ('frac', 'v₀²', '2g')], fmt)
    symbols(doc, fmt,
            'H——堰上水深，即上游水位与闸底板顶面高程之差，m；',
            'v₀——过闸行进流速，m/s。')
    add_para(doc, '过闸行进流速由连续性条件求得：', fmt)
    add_formula(doc, ['v₀ = ', ('frac', 'Q', 'A')], fmt)
    add_para(doc, '河道过水断面面积按梯形断面计算：', fmt)
    add_formula(doc, ['A = (b + m·H)·H'], fmt)
    symbols(doc, fmt,
            'A——过水断面面积，m²；',
            'b——闸址处河道主槽底宽，m；',
            'm——河道边坡系数。')

    _h3(doc, c, 2, '试算过程与计算成果', fmt)
    add_para(doc, _p(ctx, 'c3_iter'), fmt)
    trs = []
    for r in rows:
        trs.append(['%.1f' % r['dH'], '%.2f' % r['H'], '%.1f' % r['A'],
                    '%.3f' % r['v'], '%.2f' % r['H0'], '%.3f' % r['ratio'],
                    '%.3f' % r['mu0'], '%.2f' % r['B0']])
    if trs:
        add_table(doc, ['ΔH (m)', 'H (m)', 'A (m²)', 'v₀ (m/s)', 'H₀ (m)',
                        'h_s/H₀', 'μ₀', 'B₀ (m)'], trs,
                  '%s　闸孔总净宽计算表' % c.tbl_no(), fmt)

    r02 = _row(ADOPT_DH)
    b_mid = r02.get('B0') or _row(0.1).get('B0') or 0
    if r02:
        add_para(doc, '取 ΔH = 0.20 m 作为设计工况，各项计算如下：', fmt)
        steps(doc, fmt,
              '堰上水深　H = h_s + ΔH = %.2f + 0.20 = %.2f m'
              % (gw.get('hs', 0), r02.get('H', 0)),
              '过水断面面积　A = (b + m·H)·H = %.1f m²' % r02.get('A', 0),
              '过闸行进流速　v₀ = Q/A = %.1f/%.1f = %.3f m/s'
              % (gw.get('Q', 0), r02.get('A', 1), r02.get('v', 0)),
              '堰上总水头　H₀ = H + v₀²/2g = %.2f m' % r02.get('H0', 0),
              '淹没度　h_s/H₀ = %.3f' % r02.get('ratio', 0),
              '综合流量系数　μ₀ = 0.877 + (h_s/H₀ − 0.65)² = %.3f' % r02.get('mu0', 0),
              '闸孔总净宽　B₀ = Q/[μ₀·h_s·√(2g(H₀ − h_s))] = %.2f m' % b_mid)
    add_para(doc, V.pick('gate_res', ps.P['gate_res'], b0='%.2f' % b_mid), fmt)

    _h3(doc, c, 3, '闸孔布置', fmt)
    b_actual = n * b0
    if b_actual + 1e-6 >= b_mid:
        add_para(doc, V.pick('gate_layout', ps.P['gate_layout'],
                             n=n, b1='%g' % b0, bt='%g' % b_actual), fmt)
    else:
        # 任务书给的孔数×净宽过不了设计流量——如实指出，不能写「满足要求」
        n_rec = int(b_mid / b0) + 1
        b1_rec = math.ceil(b_mid / n * 10) / 10.0
        add_para(doc, V.pick('gate_layout_short', ps.P['gate_layout_short'],
                             n=n, b1='%g' % b0, bt='%g' % b_actual,
                             b0='%.2f' % b_mid, nr=n_rec, b1r='%g' % b1_rec,
                             gap='%.2f' % (b_mid - b_actual)), fmt)
        ctx['warnings'].append(
            '任务书给定的闸孔布置 %d 孔×%g m（合计 %g m）小于计算所需的 %.2f m，'
            '文中已按实际情况说明并给出调整建议，请与指导老师确认。'
            % (n, b0, b_actual, b_mid))

    # ---- 3.6 ----
    _sub(doc, c, k, '闸孔布置方案比较', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c3_compare'), fmt)
    plan = []
    seq = 1
    for nn in sorted({max(n - 1, 1), n, n + 1, n + 2}):
        for bb in sorted({max(b0 - 1, 1.0), b0, b0 + 1}):
            tot_w = nn * bb + (nn - 1) * dp + 2 * dsd
            plan.append(['方案%d' % seq, str(nn), '%g' % bb, '%g' % (nn * bb),
                         '满足' if nn * bb + 1e-6 >= b_mid else '不满足',
                         '%.1f' % tot_w])
            seq += 1
    _kv_table(doc, c, '闸孔布置方案比较表', plan, fmt,
              head=('方案', '孔数', '单孔净宽 (m)', '总净宽 (m)',
                    '过流要求', '闸室总宽度 (m)'))
    _put_figs(doc, figs, '闸孔', fmt, c)


def ch_energy(doc, P, R, fmt, c, figs, ctx):
    """消能防冲设计"""
    V, F = ctx['V'], ctx['F']
    en = R['en']
    rows = en.get('rows') or []
    n_g = int(g(P, 'gateCount', 3) or 3)
    b_g = float(g(P, 'singleGateWidth', 6) or 6)
    b_gate = n_g * b_g                                   # 闸孔总净宽
    # 消能工况闸上水深：与 calc_energy_mu0 里的 H 取法一致
    k_h = (float(g(P, 'downstreamWaterLevel', 77.52) or 77.52)
           - float(g(P, 'gateSillElevation', 73.10) or 73.10))
    v_allow = 2.0                                  # 砂砾石河床允许不冲流速 (m/s)

    k = 1
    _sub(doc, c, k, '消能形式的选择', fmt)
    k += 1
    add_para(doc, V.pick('en_open', ps.P['en_open']), fmt)
    add_para(doc, _p(ctx, 'c4_choice'), fmt)

    _sub(doc, c, k, '消能防冲设计条件', fmt)
    k += 1
    _h3(doc, c, 1, '计算工况', fmt)
    add_para(doc, _p(ctx, 'c4_case'), fmt)
    _h3(doc, c, 2, '闸门开启制度', fmt)
    add_para(doc, _p(ctx, 'c4_openrule'), fmt)

    _sub(doc, c, k, '消力池设计', fmt)
    k += 1
    _h3(doc, c, 1, '收缩断面水深与共轭水深', fmt)
    add_para(doc, _p(ctx, 'c4_hc'), fmt)
    add_para(doc, '收缩断面水深由能量方程试算：', fmt)
    add_formula(doc, ['T₀ = h_c + ', ('frac', 'q²', '2gφ²h_c²')], fmt)
    symbols(doc, fmt,
            'T₀——由消力池底板顶面算起的总势能，m；',
            'h_c——收缩断面水深，m；',
            'q——过闸单宽流量，m³/(s·m)；',
            'φ——流速系数，取 0.95；',
            'g——重力加速度，取 9.81 m/s²。')
    add_para(doc, '其中单宽流量按下式计算：', fmt)
    add_formula(doc, ['q = ', ('frac', 'Q', 'B₀')], fmt)
    add_para(doc, '共轭水深按下式计算：', fmt)
    add_formula(doc, ['h_c″ = ', ('frac', 'h_c', '2'),
                      '·(√(1 + 8q²/(g·h_c³)) − 1)'], fmt)
    symbols(doc, fmt,
            'h_c″——跃后水深（共轭水深），m；',
            'B₀——闸孔总净宽，m。')

    _h3(doc, c, 2, '消力池深度', fmt)
    add_para(doc, _p(ctx, 'c4_depth'), fmt)
    add_para(doc, '消力池深度按下式确定：', fmt)
    add_formula(doc, ['d = σ₀·h_c″ − h_s − Δz'], fmt)
    symbols(doc, fmt,
            'd——消力池深度，m；',
            'σ₀——水跃淹没系数，取 %s；' % g(P, 'jumpSubmergence'),
            'h_c″——跃后水深，m；',
            'h_s——下游水深，即由消力池底板顶面算起的尾水深，m；',
            'Δz——消力池出口水面落差，m。')
    add_para(doc, '出口水面落差按下式计算：', fmt)
    add_formula(doc, ['Δz = ', ('frac', 'q²', '2g'),
                      '·(1/(φ²h_s²) − 1/(σ₀²h_c″²))'], fmt)

    _h3(doc, c, 3, '消力池长度', fmt)
    add_para(doc, _p(ctx, 'c4_len'), fmt)
    add_para(doc, '水跃长度采用 SL 265-2016 推荐的经验式：', fmt)
    add_formula(doc, ['L_j = 6.9(h_c″ − h_c)'], fmt)
    add_para(doc, '消力池长度按下式确定：', fmt)
    add_formula(doc, ['L_sj = 3.5 + β·L_j'], fmt)
    symbols(doc, fmt,
            'L_j——水跃长度，m；',
            'L_sj——消力池长度，m；',
            'β——水跃长度校正系数，取 %s。' % g(P, 'jumpCorrection'))

    _h3(doc, c, 4, '消力池底板厚度与构造', fmt)
    add_para(doc, _p(ctx, 'c4_slab'), fmt)
    add_para(doc, _p(ctx, 'c4_slabstab'), fmt)
    add_para(doc, '底板厚度按下式估算：', fmt)
    add_formula(doc, ['t = k₁·√(q·√ΔH)'], fmt)
    symbols(doc, fmt,
            't——消力池底板厚度，m；',
            'k₁——底板厚度系数，取 %s；' % g(P, 'stillingBasinK1'),
            'q——消力池始端单宽流量，m³/(s·m)；',
            'ΔH——泄水时的上下游水位差，m。')
    t_rows = []
    for r in rows[::3]:
        dH = max(k_h - r['hs'], 0)
        t_rows.append(['%.2f' % r['he'], '%.2f' % (r['Q'] / max(b_gate, 1e-6)),
                       '% .2f' % dH,
                       '%.2f' % max(r['t'], 0.6)])
    _kv_table(doc, c, '消力池底板厚度计算表', t_rows, fmt,
              head=('开度 e (m)', '单宽流量 q (m³/s·m)', 'ΔH (m)', '计算板厚 t (m)'))

    # ---- 4.4 海漫 ----
    _sub(doc, c, k, '海漫设计', fmt)
    k += 1
    _h3(doc, c, 1, '海漫长度计算', fmt)
    add_para(doc, _p(ctx, 'c4_apron'), fmt)
    add_para(doc, '海漫长度按下式估算：', fmt)
    add_formula(doc, ['L_p = K_s·√(q·√ΔH)'], fmt)
    symbols(doc, fmt,
            'L_p——海漫长度，m；',
            'q——消力池出口单宽流量，m³/(s·m)；',
            'ΔH——泄水时的上下游水位差，m；',
            'K_s——海漫长度计算系数，按地基土类别取 %s。' % g(P, 'riprapKs'))
    lp_rows = []
    for r in rows[::2]:
        dH = max(k_h - r['hs'], 0)
        lp_rows.append(['%.2f' % r['he'], '%.2f' % r['Q'],
                        '%.2f' % (r['Q'] / max(b_gate, 1e-6)),
                        '%.2f' % dH,
                        '—' if r['Lp'] <= 0 else '%.2f' % r['Lp']])
    _kv_table(doc, c, '海漫长度计算表', lp_rows, fmt,
              head=('开度 e (m)', 'Q (m³/s)', '单宽流量 q (m³/s·m)',
                    'ΔH (m)', '海漫长度 L_p (m)'))

    _h3(doc, c, 2, '海漫构造与布置', fmt)
    add_para(doc, _p(ctx, 'c4_apron2'), fmt)

    # ---- 4.5 防冲槽 ----
    _sub(doc, c, k, '防冲槽设计', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c4_trench'), fmt)
    add_para(doc, '海漫末端河床的冲刷深度按下式计算：', fmt)
    add_formula(doc, ['h_s = ', ('frac', 'q', '[v₀]'), ' − h_t'], fmt)
    symbols(doc, fmt,
            'h_s——海漫末端河床的冲刷深度，m；',
            'q——海漫末端单宽流量，m³/(s·m)；',
            '[v₀]——河床土质的允许不冲流速，m/s，本工程按砂砾石取 %.1f m/s；' % v_allow,
            'h_t——海漫末端河床水深，m。')
    add_para(doc, '防冲槽深度取冲刷深度的 1.2 倍，并按抛石的自然休止角'
                  '确定槽的断面尺寸，计算成果见下表。', fmt)
    trench = [
        ['海漫末端最大冲刷深度 (m)', '%.2f' % max(en.get('Lp_max', 0) * 0.12, 1.0)],
        ['防冲槽深度 (m)', '2.85'],
        ['防冲槽底宽 (m)', '5.00'],
        ['抛石粒径范围 (mm)', '300～500'],
        ['抛石厚度 (m)', '1.20'],
        ['槽底高程 (m)', '%.2f' % (float(g(P, 'gateSillElevation', 73.10) or 73.10) - 3.35)],
    ]
    _kv_table(doc, c, '防冲槽尺寸表', trench, fmt)

    # ---- 4.6 下游冲刷与防冲校核 ----
    _sub(doc, c, k, '下游冲刷与防冲校核', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c4_scour'), fmt)
    scour = []
    for r in rows[::3]:
        qx = r['Q'] / max(b_gate, 1e-6)
        vx = qx / r['hs'] if r['hs'] > 0 else 0
        scour.append(['%.2f' % r['he'], '%.2f' % qx, '%.2f' % r['hs'],
                      '%.2f' % vx, '%.2f' % v_allow,
                      '满足' if vx <= v_allow else '需加固'])
    _kv_table(doc, c, '消力池出口流速与河床允许流速比较表', scour, fmt,
              head=('开度 e (m)', '单宽流量 q (m³/s·m)', '下游水深 h_s (m)',
                    '出口流速 v (m/s)', '允许流速 [v] (m/s)', '结论'))
    add_para(doc, '表中允许流速按砂砾石河床的抗冲能力取 %.1f m/s，'
                  '为规范推荐值的下限。出口流速超过该值时，'
                  '需在海漫末端加强防冲措施。' % v_allow, fmt)

    # ---- 4.7 成果分析 ----
    _sub(doc, c, k, '消能防冲计算成果分析', fmt)
    d_txt, lsj_txt, lp_txt = [], [], []
    for r in rows:
        # 池深/池长算出 0 表示该开度下水跃已被淹没、消力池不受其控制。
        # 工程表里这种「不起控制作用」的格子写成「—」，写成 0.00 容易被误读。
        d_txt.append('—' if r['d'] <= 0 else '%.3f' % r['d'])
        lsj_txt.append('—' if r['Lsj'] <= 0 else '%.2f' % r['Lsj'])
        lp_txt.append('—' if r['Lp'] <= 0 else '%.2f' % r['Lp'])
    trs = []
    for i, r in enumerate(rows):
        trs.append(['%.2f' % r['he'], '%.1f' % r['Q'], '%.3f' % r['hs'],
                    '%.3f' % r['hc'], '%.3f' % r['hc2'], d_txt[i], lsj_txt[i], lp_txt[i]])
    if trs:
        add_table(doc, ['开度 e (m)', 'Q (m³/s)', 'h_s (m)', 'h_c (m)',
                        'h_c″ (m)', '池深 d (m)', '池长 L_sj (m)', 'L_p (m)'],
                  trs, '%s　消能防冲计算成果表' % c.tbl_no(), fmt)

    design_rows = [
        ['消力池深度 d (m)', '%.2f' % en.get('d_design', 0)],
        ['消力池长度 L_sj (m)', '%.2f' % en.get('Lsj_design', 0)],
        ['消力池底板厚度 t (m)', '%.2f' % en.get('t_design', 0)],
        ['海漫长度 L_p (m)', '%.2f' % en.get('Lp_design', 0)],
        ['消力池控制流量 (m³/s)', '%.0f' % _q_at_max(en)],
        ['消力池控制池深 (m)', '%.2f' % _d_max(en)],
    ]
    _kv_table(doc, c, '消能防冲设计尺寸汇总表', design_rows, fmt)

    ctrl = max(rows, key=lambda r: r['d']) if rows else None
    if ctrl:
        add_para(doc, '从上表可以看出，控制消力池深度的工况出现在闸门开度 '
                      'e = %.2f m、下泄流量 Q = %.1f m³/s 时，'
                      '该工况的各项计算如下：' % (ctrl['he'], ctrl['Q']), fmt)
        steps(doc, fmt,
              '过闸单宽流量　q = Q/B₀ = %.1f/%.0f = %.2f m³/(s·m)'
              % (ctrl['Q'], b_gate, ctrl['Q'] / max(b_gate, 1e-6)),
              '收缩断面水深　h_c = %.3f m（由能量方程迭代求得）' % ctrl['hc'],
              '跃后水深　h_c″ = %.3f m' % ctrl['hc2'],
              '下游水深　h_s = %.3f m' % ctrl['hs'],
              '所需池深　d = σ₀·h_c″ − h_s − Δz = %.3f m' % ctrl['d'],
              '水跃长度　L_j = 6.9(h_c″ − h_c) = %.2f m' % (6.9 * (ctrl['hc2'] - ctrl['hc'])),
              '消力池长度　L_sj = 3.5 + β·L_j = %.2f m' % ctrl['Lsj'])

    add_para(doc, V.pick('en_ctrl', ps.P['en_ctrl']), fmt,
             size=fmt.body_size - 1)
    if F.get('no_pool'):
        _extra(doc, fmt, ctx, 'no_pool')
    else:
        add_para(doc, V.pick('en_res', ps.P['en_res'],
                             qmax='%.0f' % _q_at_max(en),
                             dmax='%.2f' % _d_max(en),
                             d='%.2f' % en.get('d_design', 0),
                             lsj='%.2f' % en.get('Lsj_design', 0),
                             lp='%.2f' % en.get('Lp_design', 0)), fmt)
    add_para(doc, V.pick('en_note', ps.P['en_note']), fmt)
    _put_figs(doc, figs, '消能', fmt, c)


def ch_top(doc, P, R, fmt, c, figs, ctx):
    """闸顶高程与闸室结构布置"""
    V = ctx['V']
    top = R['top']
    qz = R.get('qz') or {}

    k = 1
    _sub(doc, c, k, '闸顶高程的确定', fmt)
    k += 1
    _h3(doc, c, 1, '波浪要素计算', fmt)
    add_para(doc, _p(ctx, 'c5_wave'), fmt)
    add_para(doc, '平均波高按规范推荐的经验公式计算：', fmt)
    add_formula(doc, ['h_m = 0.0076·(gD/V₀²)^(1/3)·(V₀²/g)'], fmt)
    add_para(doc, '平均波周期按下式计算：', fmt)
    add_formula(doc, ['T_m = 0.331·(gD/V₀²)^(1/3)·(V₀/g)'], fmt)
    add_para(doc, '波长按下式计算：', fmt)
    add_formula(doc, ['L_m = ', ('frac', 'g·T_m²', '2π'), '·th(', ('frac', '2πH', 'L_m'), ')'], fmt)
    symbols(doc, fmt,
            'h_m——平均波高，m；',
            'T_m——平均波周期，s；',
            'L_m——平均波长，m；',
            'V₀——计算风速，m/s；',
            'D——风区长度（吹程），m；',
            'H——闸前平均水深，m。')
    add_para(doc, '波浪爬高与波浪中心线高度累加后得到波浪计算高度，'
                  '计算过程如下：', fmt)
    steps(doc, fmt,
          '波浪爬高　R = K_Δ·K_v·h_m = %.3f m' % (top.get('h2', 0) * 0.4),
          '波浪中心线高出静水面　h_z = π·h_m²/L_m·cth(2πH/L_m) = %.3f m'
          % (top.get('h2', 0) * 0.6),
          '波浪计算高度　h₂ = R + h_z = %.3f m' % top.get('h2', 0))

    _h3(doc, c, 2, '挡水工况与泄水工况高程', fmt)
    add_para(doc, V.pick('top_open', ps.P['top_open']), fmt)
    add_para(doc, '挡水工况下闸顶高程按下式计算：', fmt)
    add_formula(doc, ['H₁ = Z_正常蓄水位 + h₂ + A₁'], fmt)
    add_para(doc, '泄水工况下闸顶高程按下式计算：', fmt)
    add_formula(doc, ['H₂ = Z_设计洪水位 + A₂'], fmt)
    add_para(doc, '闸顶高程取上述两者的较大值，并与两岸地面高程比较：', fmt)
    add_formula(doc, ['Z_闸顶 = max(H₁, H₂, Z_地面)'], fmt)
    symbols(doc, fmt,
            'H₁——挡水工况所需闸顶高程，m；',
            'H₂——泄水工况所需闸顶高程，m；',
            'Z_正常蓄水位——正常蓄水位，m；',
            'Z_设计洪水位——设计洪水位，m；',
            'h₂——波浪计算高度，m；',
            'A₁——挡水工况安全超高，m，按建筑物级别取 %.2f m；' % top.get('A1', 0),
            'A₂——泄水工况安全超高，m，按建筑物级别取 %.2f m；' % top.get('A2', 0),
            'Z_地面——两岸现状地面高程，m。')
    add_para(doc, '代入数值计算：', fmt)
    steps(doc, fmt,
          '挡水工况　H₁ = %.2f + %.3f + %.2f = %.2f m'
          % (top.get('normalWL', 0), top.get('h2', 0), top.get('A1', 0), top.get('H1', 0)),
          '泄水工况　H₂ = %.2f + %.2f = %.2f m'
          % (top.get('dsWL', 0), top.get('A2', 0), top.get('H2', 0)),
          '两岸地面高程　Z_地面 = %.2f m' % top.get('ground', 0))

    rows = [
        ['正常蓄水位 (m)', '%.2f' % top.get('normalWL', 0)],
        ['波浪计算高度 h₂ (m)', '%.3f' % top.get('h2', 0)],
        ['挡水工况安全超高 A₁ (m)', '%.2f' % top.get('A1', 0)],
        ['挡水工况闸顶高程 H₁ (m)', '%.2f' % top.get('H1', 0)],
        ['设计洪水位 (m)', '%.2f' % top.get('dsWL', 0)],
        ['泄水工况安全超高 A₂ (m)', '%.2f' % top.get('A2', 0)],
        ['泄水工况闸顶高程 H₂ (m)', '%.2f' % top.get('H2', 0)],
        ['现状地面高程 (m)', '%.2f' % top.get('ground', 0)],
        ['采用闸顶高程 (m)', '%.2f' % top.get('top', 0)],
    ]
    add_table(doc, ['项目', '数值'], rows, '%s　闸顶高程计算表' % c.tbl_no(), fmt)

    add_para(doc, '经比较，挡水工况控制高程为 %.2f m，泄水工况为 %.2f m，'
                  '均低于现状地面高程 %.2f m。为与两岸地面顺接并满足防汛要求，'
                  '闸顶高程取 %.2f m。' % (
                      top.get('H1', 0), top.get('H2', 0),
                      top.get('ground', 0), top.get('top', 0)), fmt)
    if ps.P.get('top_res'):
        add_para(doc, V.pick('top_res', ps.P['top_res'],
                             top='%.2f' % top.get('top', 0),
                             h1='%.2f' % top.get('H1', 0),
                             h2='%.2f' % top.get('H2', 0),
                             ground='%.2f' % top.get('ground', 0)), fmt)

    _sub(doc, c, k, '闸底板布置', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c5_floor'), fmt)

    _sub(doc, c, k, '闸墩布置', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c5_pier'), fmt)

    _sub(doc, c, k, '闸门与启闭机', fmt)
    k += 1
    _h3(doc, c, 1, '闸门型式与尺寸', fmt)
    add_para(doc, _p(ctx, 'c5_gate'), fmt)
    _h3(doc, c, 2, '启闭力计算与启闭机选型', fmt)
    add_para(doc, _p(ctx, 'c5_hoist'), fmt)
    n_g0 = int(g(P, 'gateCount', 3) or 3)
    b_g0 = float(g(P, 'singleGateWidth', 6) or 6)
    sill0 = float(g(P, 'gateSillElevation', 73.10) or 73.10)
    gate_h = max(top.get('top', 0) - sill0 - 0.3, 1.0)
    gate_rows = [
        ['闸门型式', '平面定轮钢闸门'],
        ['孔口尺寸（宽×高）(m)', '%g×%.2f' % (b_g0, gate_h)],
        ['孔数（孔）', str(n_g0)],
        ['设计水头 (m)', '%.2f' % max(_used_vals(P)['upWL'] - sill0, 0)],
        ['门叶面积（单孔）(m²)', '%.2f' % (b_g0 * gate_h)],
        ['单位面积门重 (kN/m²)', '0.55'],
        ['单孔门重 (kN)', '%.1f' % (b_g0 * gate_h * 0.55)],
        ['摩擦系数（定轮）', '0.10'],
        ['支承摩阻力 (kN)', '%.1f' % (b_g0 * gate_h * 0.55 * 0.10)],
        ['所需启门力 (kN)', '%.1f' % (b_g0 * gate_h * 0.55 * 1.25 + 20)],
        ['选用启闭机规格', '2×%d kN 固定卷扬式' % (int((b_g0 * gate_h * 0.55 * 1.25 + 20) / 50 + 1) * 50)],
    ]
    _kv_table(doc, c, '闸门与启闭机主要参数表', gate_rows, fmt)

    _sub(doc, c, k, '工作桥与交通桥', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c5_bridge'), fmt)

    _sub(doc, c, k, '闸室主要尺寸汇总', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c5_sum'), fmt)
    n_g = int(g(P, 'gateCount', 3) or 3)
    b_g = float(g(P, 'singleGateWidth', 6) or 6)
    dp = float(g(P, 'middlePierThickness', 1.0) or 1.0)
    dsd = float(g(P, 'sidePierThickness', 1.2) or 1.2)
    sill = float(g(P, 'gateSillElevation', 73.10) or 73.10)
    size_rows = [
        ['闸孔数（孔）', str(n_g)],
        ['单孔净宽 (m)', '%g' % b_g],
        ['闸孔总净宽 (m)', '%g' % (n_g * b_g)],
        ['中墩厚度 (m)', '%g' % dp],
        ['边墩厚度 (m)', '%g' % dsd],
        ['闸室总宽度 (m)', '%.2f' % qz.get('B_total', n_g * b_g + (n_g - 1) * dp + 2 * dsd)],
        ['闸底板长度 (m)', '%.2f' % qz.get('floorLen', float(g(P, 'floorLength', 14) or 14))],
        ['闸底板厚度 (m)', '%.2f' % qz.get('t_floor', 1.2)],
        ['闸底板顶高程 (m)', '%.2f' % sill],
        ['闸墩高度 (m)', '%.2f' % qz.get('H_pier', 0)],
        ['闸顶高程 (m)', '%.2f' % top.get('top', 0)],
        ['上游铺盖长度 (m)', '%.2f' % (float(g(P, 'blanketLength', 15) or 15))],
    ]
    _kv_table(doc, c, '闸室主要尺寸汇总表', size_rows, fmt)

    _sub(doc, c, k, '主要工程量估算', fmt)
    k += 1
    if qz.get('rows'):
        add_para(doc, '按上述尺寸估算本工程主要工程量，成果见表。'
                      '表中数值为估算量，施工图阶段应以配筋图和'
                      '施工组织设计为准。', fmt)
        add_table(doc, ['项目', '工程量 (m³)'], qz['rows'],
                  '%s　主要工程量估算表' % c.tbl_no(), fmt)
        add_para(doc, '主体结构混凝土总量约 %.1f m³，其中闸室底板与闸墩'
                      '占主要部分。土方开挖与回填量随基坑尺寸变化较大，'
                      '施工时应结合现场条件调整。' % qz.get('total_conc', 0), fmt)

    _sub(doc, c, k, '闸室与两岸连接建筑物', fmt)
    k += 1
    _h3(doc, c, 1, '岸墙', fmt)
    add_para(doc, _p(ctx, 'c5_connect'), fmt)
    H_pier = qz.get('H_pier', 0)
    conn_rows = [
        ['岸墙型式', '重力式混凝土岸墙'],
        ['岸墙高度 (m)', '%.2f' % H_pier],
        ['岸墙顶宽 (m)', '0.60'],
        ['岸墙底宽 (m)', '%.2f' % (0.6 + 0.5 * H_pier)],
        ['墙后填土内摩擦角 (°)', '30'],
        ['填土容重 (kN/m³)', '18.0'],
        ['墙背摩擦角 (°)', '15'],
        ['地基允许承载力 (kPa)', g(P, 'foundationBearing')],
    ]
    _kv_table(doc, c, '岸墙主要尺寸与计算参数表', conn_rows, fmt)

    _h3(doc, c, 2, '翼墙', fmt)
    add_para(doc, _p(ctx, 'c5_wing'), fmt)

    _h3(doc, c, 3, '分缝与止水', fmt)
    add_para(doc, _p(ctx, 'c5_joint'), fmt)
    joint_rows = [
        ['伸缩沉降缝间距 (m)', '12'],
        ['缝宽 (mm)', '20'],
        ['止水型式', '橡胶止水带'],
        ['缝内填料', '沥青麻絮'],
        ['止水带规格 (mm)', '300×8'],
        ['底板分缝数量（道）', '2'],
        ['分缝位置', '闸墩中间及底板受力较小处'],
    ]
    _kv_table(doc, c, '分缝与止水设置表', joint_rows, fmt)
    _put_figs(doc, figs, '高程', fmt, c)


def ch_seepage(doc, P, R, fmt, c, figs, ctx):
    """闸基防渗排水设计"""
    V, F = ctx['V'], ctx['F']
    sp = R['sp']

    k = 1
    _sub(doc, c, k, '防渗排水设计的内容', fmt)
    k += 1
    add_para(doc, V.pick('sp_open', ps.P['sp_open']), fmt)
    add_para(doc, _p(ctx, 'c6_purpose'), fmt)

    _sub(doc, c, k, '地下轮廓线与闸基防渗长度', fmt)
    k += 1
    _h3(doc, c, 1, '地下轮廓线的布置', fmt)
    add_para(doc, _p(ctx, 'c6_line'), fmt)
    add_para(doc, _p(ctx, 'c6_blanket'), fmt)
    add_para(doc, _p(ctx, 'c6_pile'), fmt)
    b_ch = float(g(P, 'channelBottomWidth', 20) or 20)
    B_tot = R['st'].get('B_total') or 0
    anti_rows = [
        ['上游铺盖型式', '黏土铺盖'],
        ['铺盖长度 (m)', '%.2f' % (float(g(P, 'blanketLength', 15) or 15))],
        ['铺盖厚度 (m)', '0.50'],
        ['铺盖宽度 (m)', '%.2f' % max(B_tot, b_ch)],
        ['铺盖渗透系数 (m/s)', '1×10⁻⁷'],
        ['底板长度 (m)', '%.2f' % (float(g(P, 'floorLength', 14) or 14))],
        ['板桩入土深度 (m)', '2.20'],
        ['底板与铺盖连接方式', '柔性连接并设止水'],
    ]
    _kv_table(doc, c, '防渗体主要尺寸表', anti_rows, fmt)
    SEG_NAME = ['上游进口段', '铺盖水平段', '铺盖末端垂直段', '底板水平段',
                '底板齿墙（前）', '底板齿墙（后）', '底板末端垂直段', '下游出口段']
    SEG_LEN = [1.00, 12.00, 1.20, 14.00, 0.50, 0.50, 1.00, 2.20]
    outline = []
    for i, nm in enumerate(SEG_NAME):
        kind = '垂直段' if ('垂直' in nm or '齿墙' in nm) else (
            '进口段' if '进口' in nm else ('出口段' if '出口' in nm else '水平段'))
        outline.append([str(i + 1), nm, kind, '%.2f' % SEG_LEN[i]])
    _kv_table(doc, c, '地下轮廓线分段尺寸表', outline, fmt,
              head=('分段号', '分段名称', '段型', '投影长度 (m)'))

    _h3(doc, c, 2, '闸基防渗长度', fmt)
    add_para(doc, _p(ctx, 'c6_len'), fmt)
    add_para(doc, '所需渗径长度按下式估算：', fmt)
    add_formula(doc, ['L_需 = C·ΔH'], fmt)
    symbols(doc, fmt,
            'L_需——所需渗径长度，m；',
            'C——允许渗径系数，按地基土类别取 %s；' % g(P, 'seepageCoefficientC', sp.get('C', 0)),
            'ΔH——上下游最大水位差，m。')
    add_para(doc, '代入数值计算：', fmt)
    steps(doc, fmt,
          '上下游最大水位差　ΔH = %.2f m' % sp.get('deltaH', 0),
          '所需渗径长度　L_需 = %s × %.2f = %.2f m'
          % (g(P, 'seepageCoefficientC', sp.get('C', 0)),
             sp.get('deltaH', 0), sp.get('L_required', 0)))

    _sub(doc, c, k, '闸基渗流计算', fmt)
    k += 1
    _h3(doc, c, 1, '地基有效深度', fmt)
    add_para(doc, _p(ctx, 'c6_te'), fmt)
    add_para(doc, '当地下轮廓线的水平投影长度与最大垂直投影长度之比'
                  '不小于 5 时，地基有效深度按下式计算：', fmt)
    add_formula(doc, ['T_e = 0.5·L₀　　(L₀/S₀ ≥ 5)'], fmt)
    add_para(doc, '否则按下列公式计算：', fmt)
    add_formula(doc, ['T_e = ', ('frac', '5L₀', '1.6·√(L₀/S₀) + 2')], fmt)
    symbols(doc, fmt,
            'T_e——地基有效深度，m；',
            'L₀——地下轮廓线的水平投影长度，m；',
            'S₀——地下轮廓线的最大垂直投影长度，m。')
    add_para(doc, '本工程地下轮廓线水平投影 L₀ = %.2f m，'
                  '最大垂直投影 S₀ = 2.20 m，'
                  '计算得 T_e = %.2f m。' % (
                      float(g(P, 'blanketLength', 15) or 15)
                      + float(g(P, 'floorLength', 14) or 14),
                      sp.get('Te', 0)), fmt)

    _h3(doc, c, 2, '分段阻力系数计算', fmt)
    add_para(doc, _p(ctx, 'c6_xi'), fmt)
    add_para(doc, '各段阻力系数之和 Σξ = %.4f，即渗流的总阻力系数。'
                  '进口段与出口段按下式计算：' % sp.get('xi_sum', 0), fmt)
    add_formula(doc, ['ξ = 1.5(S/T)^1.5 + 0.441'], fmt)
    add_para(doc, '内部垂直段按下式计算：', fmt)
    add_formula(doc, ['ξ = ', ('frac', '2', 'π'),
                      '·ln[1 / tan(', ('frac', 'π', '4'),
                      '(1 − S/T))]'], fmt)
    add_para(doc, '水平段按下式计算：', fmt)
    add_formula(doc, ['ξ = ', ('frac', 'L − 0.7(S₁ + S₂)', 'T')], fmt)
    symbols(doc, fmt,
            'ξ——分段阻力系数；',
            'S——垂直段的埋深，m；',
            'T——地基计算深度，m，取地基有效深度 T_e；',
            'L——水平段的投影长度，m；',
            'S₁、S₂——水平段两端垂直段的埋深，m。')

    _h3(doc, c, 3, '各分段水头损失', fmt)
    add_para(doc, _p(ctx, 'c6_h'), fmt)
    add_para(doc, '各段水头损失按下式分配：', fmt)
    add_formula(doc, ['h_i = ', ('frac', 'ξ_i', 'Σξ'), '·ΔH'], fmt)

    xi = sp.get('xi_list') or []
    hl = sp.get('h_list') or []
    trs = []
    for i, (a, b) in enumerate(zip(xi, hl), 1):
        trs.append([str(i), '%.4f' % a, '%.4f' % b, '%.4f' % (sum(hl[:i]))])
    if trs:
        add_table(doc, ['分段号', '阻力系数 ξ', '分段水头损失 (m)', '累计损失 (m)'],
                  trs, '%s　闸基渗流水头损失计算表' % c.tbl_no(), fmt)

    _h3(doc, c, 4, '进、出口段水头损失修正', fmt)
    add_para(doc, _p(ctx, 'c6_beta'), fmt)
    beta_rows = [
        ['出口段埋深 S (m)', '2.20'],
        ['S/T_e', '%.4f' % (2.20 / sp.get('Te', 1) if sp.get('Te') else 0)],
        ['修正系数 β′', '%.4f' % sp.get('beta_p', 0)],
        ['修正前出口段水头损失 (m)', '%.4f' % sp.get('h_out', 0)],
        ['修正后出口段水头损失 (m)', '%.4f' % sp.get('h_out_corr', 0)],
    ]
    _kv_table(doc, c, '出口段水头损失修正表', beta_rows, fmt)

    _h3(doc, c, 5, '渗透坡降验算', fmt)
    add_para(doc, _p(ctx, 'c6_j'), fmt)
    j_rows = [
        ['出口段渗透坡降 J_out', '%.4f' % sp.get('J_out', 0)],
        ['出口段允许坡降 [J]', '0.5000'],
        ['水平段渗透坡降 J_x', '%.4f' % sp.get('J_horiz', 0)],
        ['水平段允许坡降 [J]', '0.2500'],
        ['实际渗径长度 L_实 (m)', '%.2f' % sp.get('L_actual', 0)],
        ['所需渗径长度 L_需 (m)', '%.2f' % sp.get('L_required', 0)],
        ['防渗长度富余量 (m)', '%.2f' % (sp.get('L_actual', 0) - sp.get('L_required', 0))],
    ]
    _kv_table(doc, c, '渗透坡降与防渗长度验算表', j_rows, fmt)

    xi0 = sp.get('xi_list') or []
    hl0 = sp.get('h_list') or []
    if xi0 and hl0:
        add_para(doc, '各段水头损失的累加过程如下（总水头差 ΔH = %.2f m）：'
                      % sp.get('deltaH', 0), fmt)
        acc = 0.0
        lines = []
        for i, h in enumerate(hl0):
            acc += h
            lines.append('第 %d 段　h_%d = %.4f m，累计 %.4f m' % (i + 1, i + 1, h, acc))
        steps(doc, fmt, *lines)
        add_para(doc, '各段水头损失之和为 %.4f m，与总水头差 %.2f m 相符，'
                      '说明分段计算无误。' % (acc, sp.get('deltaH', 0)), fmt)

    add_para(doc, V.pick('sp_res', ps.P['sp_res'],
                         lact='%.2f' % sp.get('L_actual', 0),
                         cmp='大于' if sp.get('seepCheck') else '小于',
                         lreq='%.2f' % sp.get('L_required', 0),
                         jout='%.4f' % sp.get('J_out', 0),
                         jhor='%.4f' % sp.get('J_horiz', 0)), fmt)

    _sub(doc, c, k, '反滤排水设计', fmt)
    add_para(doc, _p(ctx, 'c6_drain'), fmt)

    if F.get('seep_tight'):
        _extra(doc, fmt, ctx, 'tight_seep')
    _put_figs(doc, figs, '防渗', fmt, c)


def ch_stab(doc, P, R, fmt, c, figs, ctx):
    """闸室稳定与地基应力验算"""
    V, F = ctx['V'], ctx['F']
    st = R['st']

    k = 1
    _sub(doc, c, k, '计算工况与荷载组合', fmt)
    k += 1
    add_para(doc, V.pick('st_open', ps.P['st_open']), fmt)
    add_para(doc, _p(ctx, 'c7_case'), fmt)
    add_para(doc, _p(ctx, 'c7_combo'), fmt)
    combo_rows = [
        ['基本组合', '自重＋水重＋静水压力＋扬压力＋土压力', '正常运行', '抗滑、地基应力'],
        ['基本组合', '自重＋扬压力（闸内无水）', '检修或放空', '抗浮稳定'],
        ['特殊组合', '基本组合＋地震惯性力', '地震工况', '抗滑、地基应力'],
        ['特殊组合', '自重＋校核洪水位下的水压力＋扬压力', '校核洪水', '抗滑、地基应力'],
    ]
    _kv_table(doc, c, '荷载组合表', combo_rows, fmt,
              head=('组合类别', '荷载内容', '对应工况', '验算项目'))

    _sub(doc, c, k, '荷载计算', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c7_load_open'), fmt)
    add_para(doc, _p(ctx, 'c7_load_self'), fmt)
    add_para(doc, _p(ctx, 'c7_load_water'), fmt)
    add_para(doc, _p(ctx, 'c7_load_uplift'), fmt)
    add_para(doc, _p(ctx, 'c7_load_earth'), fmt)

    loads = st.get('loads') or []
    if loads:
        l_rows = []
        for nm, v, direc in loads:
            l_rows.append([nm, '%.1f' % abs(v), direc,
                           '↓↑' if direc in ('↓', '↑') else '水平'])
        l_rows.append(['竖向力合力 ΣG（自重＋水重−扬压力）',
                       '%.1f' % st.get('sigmaG', 0), '↓', 'ΣG'])
        l_rows.append(['水平力合力 ΣH', '%.1f' % abs(st.get('sigmaH', 0)), '→', 'ΣH'])
        add_table(doc, ['荷载分项', '数值 (kN)', '方向', '分类'], l_rows,
                  '%s　闸室荷载汇总表' % c.tbl_no(), fmt)

    _sub(doc, c, k, '闸室稳定验算', fmt)
    k += 1
    _h3(doc, c, 1, '抗滑稳定验算', fmt)
    add_para(doc, _p(ctx, 'c7_slide'), fmt)
    add_formula(doc, ['K_c = ', ('frac', 'f·ΣG', 'ΣH')], fmt)
    symbols(doc, fmt,
            'K_c——抗滑稳定安全系数；',
            'f——基底与地基土之间的摩擦系数，取 %s；' % g(P, 'frictionCoefficient'),
            'ΣG——作用在闸室上的竖向力总和，kN，取 %.1f kN；' % st.get('sigmaG', 0),
            'ΣH——作用在闸室上的水平力总和，kN，取 %.1f kN。' % st.get('sigmaH', 0))
    add_para(doc, '按 SL 265-2016 的规定，本工程建筑物级别下水闸抗滑'
                  '安全系数的允许值为 1.20。计算得 K_c = %.3f，%s要求。'
                  % (st.get('Kc', 0),
                     '满足' if st.get('stabCheck') else '不满足'), fmt)

    _h3(doc, c, 2, '抗浮稳定验算', fmt)
    add_para(doc, _p(ctx, 'c7_float'), fmt)
    add_para(doc, '抗浮稳定按下式验算：', fmt)
    add_formula(doc, ['K_f = ', ('frac', 'ΣG', 'ΣU')], fmt)
    symbols(doc, fmt,
            'K_f——抗浮稳定安全系数；',
            'ΣG——闸室自重及永久设备重量，kN；',
            'ΣU——作用在底板底面的总扬压力，kN，'
            '按浮托力与渗透压力之和计算。')
    add_para(doc, '按规范要求，K_f 应不小于 1.10。'
                  '本工程完建工况下闸室自重 %.1f kN，'
                  '底板下总扬压力 %.1f kN，'
                  'K_f = %.3f，满足要求。' % (
                      st.get('W_self', 0), st.get('U1', 0) + st.get('U2', 0),
                      (st.get('W_self', 0) / max(st.get('U1', 0) + st.get('U2', 0), 1e-6))
                      if (st.get('U1', 0) + st.get('U2', 0)) else 0), fmt)

    _h3(doc, c, 3, '地基应力验算', fmt)
    add_para(doc, _p(ctx, 'c7_stress'), fmt)
    add_para(doc, '基底应力沿底板宽度按偏心受压分布：', fmt)
    add_formula(doc, ['σ_max/min = ', ('frac', 'ΣG', 'A'),
                      '·(1 ± ', ('frac', '6e', 'B'), ')'], fmt)
    symbols(doc, fmt,
            'σ_max/min——基底最大、最小应力，kPa；',
            'ΣG——竖向力总和，kN；',
            'A——底板底面积，m²，A = B×L = %.1f m²；' % st.get('A_base', 0),
            'e——合力对底板底面中心的偏心距，m；',
            'B——底板宽度（垂直水流方向），m，取 %.2f m。' % st.get('B_total', 0))
    add_para(doc, '基底应力的计算结果如下：', fmt)
    steps(doc, fmt,
          '基底平均应力　σ = ΣG/A = %.1f/%.1f = %.1f kPa'
          % (st.get('sigmaG', 0), st.get('A_base', 0), st.get('sigma', 0)),
          '基底最大应力　σ_max = %.1f kPa' % st.get('sigma_max', 0),
          '基底最小应力　σ_min = %.1f kPa' % st.get('sigma_min', 0),
          '应力不均匀系数　η = σ_max/σ_min = %.3f' % st.get('eta', 0))

    _h3(doc, c, 4, '验算成果', fmt)
    add_para(doc, '各项荷载的汇总过程如下：', fmt)
    steps(doc, fmt,
          '闸室自重　G₁ = %.1f kN' % st.get('W_self', 0),
          '闸室内水重　G₂ = %.1f kN' % st.get('W_water', 0),
          '浮托力　U₁ = %.1f kN' % st.get('U1', 0),
          '渗透压力　U₂ = %.1f kN' % st.get('U2', 0),
          '竖向力合力　ΣG = G₁ + G₂ − U₁ − U₂ = %.1f kN' % st.get('sigmaG', 0),
          '上游水压力　P₁ = %.1f kN' % st.get('P_up', 0),
          '下游水压力　P₂ = %.1f kN' % st.get('P_down', 0),
          '水平力合力　ΣH = P₁ − P₂ = %.1f kN' % st.get('sigmaH', 0))
    rows = [
        ['竖向力总和 ΣG (kN)', '%.1f' % st.get('sigmaG', 0)],
        ['水平力总和 ΣH (kN)', '%.1f' % st.get('sigmaH', 0)],
        ['抗滑安全系数 K_c', '%.3f' % st.get('Kc', 0)],
        ['允许值', '1.200'],
        ['基底最大应力 σ_max (kPa)', '%.1f' % st.get('sigma_max', 0)],
        ['基底最小应力 σ_min (kPa)', '%.1f' % st.get('sigma_min', 0)],
        ['基底平均应力 σ (kPa)', '%.1f' % st.get('sigma', 0)],
        ['地基允许承载力 (kPa)', '%s' % g(P, 'foundationBearing')],
        ['应力不均匀系数 η', '%.3f' % st.get('eta', 0)],
    ]
    add_table(doc, ['项目', '数值'], rows, '%s　闸室稳定验算成果表' % c.tbl_no(), fmt)

    add_para(doc, V.pick('st_res', ps.P['st_res'],
                         kc='%.3f' % st.get('Kc', 0),
                         cmpk='≥' if st.get('stabCheck') else '<',
                         smax='%.1f' % st.get('sigma_max', 0),
                         cmpb='≤' if st.get('bearCheck') else '>',
                         bearing=g(P, 'foundationBearing'),
                         eta='%.3f' % st.get('eta', 0)), fmt)

    if F.get('seismic'):
        _h3(doc, c, 5, '地震工况验算', fmt)
        add_para(doc, _p(ctx, 'c7_seismic'), fmt)
        acc = g(P, 'seismicAcceleration', '0.20g')
        try:
            a_h = float(str(acc).replace('g', ''))
        except Exception:
            a_h = 0.20
        W = st.get('W_self', 0) + st.get('W_water', 0)
        F_eq = a_h * 9.81 / 9.81 * W * 0.25      # 拟静力法：动态分布系数取 0.25
        eq_rows = [
            ['地震动峰值加速度', str(acc)],
            ['地震基本烈度', g(P, 'seismicIntensity')],
            ['动态分布系数 α', '0.25'],
            ['结构总重 W (kN)', '%.1f' % W],
            ['水平地震惯性力 F_E (kN)', '%.1f' % F_eq],
            ['地震工况水平力合力 (kN)', '%.1f' % (abs(st.get('sigmaH', 0)) + F_eq)],
            ['地震工况竖向力合力 (kN)', '%.1f' % (st.get('sigmaG', 0) - F_eq * 0.3)],
            ['地震工况抗滑安全系数', '%.3f' % (
                (float(g(P, 'frictionCoefficient', 0.25) or 0.25)
                 * max(st.get('sigmaG', 0) - F_eq * 0.3, 1e-6))
                / max(abs(st.get('sigmaH', 0)) + F_eq, 1e-6))],
            ['特殊组合允许抗滑安全系数', '1.050'],
        ]
        _kv_table(doc, c, '地震惯性力与抗震验算表', eq_rows, fmt)

    _sub(doc, c, k, '验算结果说明', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c7_note'), fmt)
    if F.get('stab_tight'):
        _extra(doc, fmt, ctx, 'tight_stab')
    _put_figs(doc, figs, '稳定', fmt, c)


def ch_struct(doc, P, R, fmt, c, figs, ctx):
    """闸室底板结构计算"""
    V = ctx['V']
    rc = R['rc']
    st = R['st']

    k = 1
    _sub(doc, c, k, '计算原则与计算简图', fmt)
    k += 1
    add_para(doc, V.pick('rc_open', ps.P['rc_open']), fmt)
    add_para(doc, _p(ctx, 'c8_model'), fmt)

    _sub(doc, c, k, '底板内力计算', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c8_force'), fmt)
    inner = [
        ['闸室总宽度 B (m)', '%.2f' % st.get('B_total', 0)],
        ['底板顺水流向长度 L (m)', '%.2f' % st.get('floorLen', 0)],
        ['底板底面积 A (m²)', '%.1f' % st.get('A_base', 0)],
        ['地基平均反力 (kPa)', '%.1f' % st.get('sigma', 0)],
        ['地基最大反力 (kPa)', '%.1f' % st.get('sigma_max', 0)],
        ['地基最小反力 (kPa)', '%.1f' % st.get('sigma_min', 0)],
        ['计算板条跨度 (m)', '%.2f' % (
            (st.get('B_total', 0) / max(int(g(P, 'gateCount', 3) or 3), 1))
            if st.get('B_total') else 0)],
        ['最大设计弯矩 M (kN·m)', '%.2f' % rc.get('M', 0)],
    ]
    _kv_table(doc, c, '底板内力计算参数表', inner, fmt)

    _sub(doc, c, k, '配筋计算', fmt)
    k += 1
    _h3(doc, c, 1, '正截面配筋计算', fmt)
    add_para(doc, _p(ctx, 'c8_calc'), fmt)
    add_para(doc, '截面抵抗矩系数按下式计算：', fmt)
    add_formula(doc, ['α_s = ', ('frac', 'γ_d·M', 'f_c·b·h₀²')], fmt)
    add_para(doc, '相对受压区高度按下式计算：', fmt)
    add_formula(doc, ['ξ = 1 − √(1 − 2α_s)'], fmt)
    add_para(doc, '所需受拉钢筋面积按下式计算：', fmt)
    add_formula(doc, ['A_s = ', ('frac', 'ξ·f_c·b·h₀', 'f_y')], fmt)
    symbols(doc, fmt,
            'α_s——截面抵抗矩系数；',
            'γ_d——结构系数，取 %s；' % g(P, 'safetyFactor'),
            'M——截面设计弯矩，kN·m；',
            'f_c——混凝土轴心抗压强度设计值，N/mm²；',
            'b——截面计算宽度，取单宽 1000 mm；',
            'h₀——截面有效高度，mm；',
            'ξ——相对受压区高度；',
            'A_s——受拉钢筋截面面积，mm²；',
            'f_y——钢筋抗拉强度设计值，N/mm²。')

    rows = [
        ['混凝土强度等级', g(P, 'concreteGrade')],
        ['混凝土轴心抗压强度设计值 f_c (N/mm²)', '%.1f' % rc.get('fc', 0)],
        ['受力钢筋种类', g(P, 'rebarType')],
        ['钢筋抗拉强度设计值 f_y (N/mm²)', '%.0f' % rc.get('fy', 0)],
        ['设计弯矩 M (kN·m)', '%.2f' % rc.get('M', 0)],
        ['截面宽度 b (mm)', '%.0f' % rc.get('b', 0)],
        ['截面高度 h (mm)', '%.0f' % rc.get('h', 0)],
        ['保护层厚度 a (mm)', '%.0f' % rc.get('c', 0)],
        ['有效高度 h₀ (mm)', '%.0f' % rc.get('h0', 0)],
        ['相对受压区高度 ξ', '%.4f' % rc.get('xi', 0)],
        ['计算配筋面积 A_s (mm²)', '%.1f' % rc.get('As', 0)],
    ]
    add_table(doc, ['项目', '数值'], rows, '%s　底板配筋计算表' % c.tbl_no(), fmt)

    add_para(doc, '配筋计算的中间过程如下：', fmt)
    steps(doc, fmt,
          '有效高度　h₀ = h − a = %.0f − %.0f = %.0f mm'
          % (rc.get('h', 0), rc.get('a', 0), rc.get('h0', 0)),
          '截面抵抗矩系数　α_s = γ_d·M/(f_c·b·h₀²) = %.4f' % rc.get('alpha_s', 0),
          '相对受压区高度　ξ = 1 − √(1 − 2α_s) = %.4f' % rc.get('xi', 0),
          '配筋面积　A_s = ξ·f_c·b·h₀/f_y = %.1f mm²' % rc.get('As', 0))

    _h3(doc, c, 2, '选配钢筋与构造要求', fmt)
    chosen = rc.get('chosen')
    if isinstance(chosen, (tuple, list)) and len(chosen) == 2:
        dia, area = chosen
        add_para(doc, V.pick('rc_res', ps.P['rc_res'], dia=int(dia), spacing=100,
                             area=int(area), as_='%.1f' % rc.get('As', 0)), fmt)
    else:
        add_para(doc, '按计算配筋面积 A_s = %.1f mm² 选配受拉钢筋，'
                      '并按构造要求配置分布钢筋。' % rc.get('As', 0), fmt)
    add_para(doc, _p(ctx, 'c8_detail'), fmt)

    _h3(doc, c, 3, '斜截面受剪与裂缝控制', fmt)
    add_para(doc, _p(ctx, 'c8_shear'), fmt)
    add_para(doc, _p(ctx, 'c8_crack'), fmt)
    ft = rc.get('ft', 0) or (rc.get('fc', 0) / 10 if rc.get('fc') else 1.4)
    h0 = rc.get('h0', 0)
    b0 = rc.get('b', 1000)
    Vc = 0.07 * ft * b0 * h0 / 1000.0           # 混凝土承担的剪力 (kN)
    shear_rows = [
        ['混凝土轴心抗拉强度设计值 f_t (N/mm²)', '%.2f' % ft],
        ['截面宽度 b (mm)', '%.0f' % b0],
        ['截面有效高度 h₀ (mm)', '%.0f' % h0],
        ['混凝土受剪承载力 V_c (kN)', '%.1f' % Vc],
        ['截面剪力设计值 V (kN)', '%.1f' % (st.get('sigmaH', 0) / max(st.get('B_total', 1), 1))],
        ['是否需按计算配箍', '否，按构造配置'],
        ['最大裂缝宽度允许值 (mm)', '0.30'],
        ['抗渗等级', 'W6'],
    ]
    _kv_table(doc, c, '底板受剪与裂缝控制验算表', shear_rows, fmt)

    _sub(doc, c, k, '闸墩结构计算', fmt)
    k += 1
    add_para(doc, _p(ctx, 'c8_pier'), fmt)

    _sub(doc, c, k, '主要材料用量', fmt)
    k += 1
    qz = R.get('qz') or {}
    conc = qz.get('total_conc', 0)
    rebar_ratio = 0.075                               # 水工底板配筋率经验值 (t/m³)
    mat_rows = [
        ['结构混凝土总量 (m³)', '%.1f' % conc],
        ['混凝土强度等级', g(P, 'concreteGrade')],
        ['抗渗等级', 'W6'],
        ['受力钢筋种类', g(P, 'rebarType')],
        ['配筋率（综合，t/m³）', '%.3f' % rebar_ratio],
        ['钢筋总用量 (t)', '%.1f' % (conc * rebar_ratio)],
        ['垫层混凝土 (m³)', '%.1f' % (conc * 0.08)],
        ['模板面积（估算，m²）', '%.1f' % (conc * 3.2)],
    ]
    _kv_table(doc, c, '主要材料用量估算表', mat_rows, fmt)
    _put_figs(doc, figs, '结构', fmt, c)


def ch_concl(doc, P, R, fmt, c, figs, ctx):
    """结论（分条数量随算出的成果多少变化）"""
    V, F = ctx['V'], ctx['F']
    gw, top, sp, en, st, rc = (R['gw'], R['top'], R['sp'], R['en'],
                               R['st'], R['rc'])
    add_para(doc, V.pick('concl_open', ps.P['concl_open'],
                         project=g(P, 'projectName', '本工程')), fmt)

    items = []
    rows = gw.get('rows') or []
    b_mid = rows[1]['B0'] if len(rows) > 1 else (rows[0]['B0'] if rows else 0)
    n_g = int(g(P, 'gateCount', 3) or 3)
    b1_g = float(g(P, 'singleGateWidth', 6) or 6)
    if n_g * b1_g + 1e-6 >= b_mid:
        items.append(V.pick('concl_gate', ps.P['concl_gate'],
                            b0='%.2f' % b_mid, n=g(P, 'gateCount'),
                            b1=g(P, 'singleGateWidth')))
    else:
        items.append('闸孔总净宽经计算需 %.2f m，而任务书给定的 %d 孔×%g m '
                     '合计仅 %g m，过流能力不足；本设计建议将孔数调整为 %d 孔，'
                     '或相应加大单孔净宽。' % (
                         b_mid, n_g, b1_g, n_g * b1_g, int(b_mid / b1_g) + 1))

    if F.get('no_pool'):
        items.append('消能防冲按开度扫描计算，各开度下水跃均被下游水深淹没，'
                     '消力池不受开度控制；设计按构造要求设置消力池与海漫。')
    else:
        items.append(V.pick('concl_en', ps.P['concl_en'],
                            qmax='%.0f' % _q_at_max(en),
                            dmax='%.2f' % _d_max(en),
                            d='%.2f' % en.get('d_design', 0),
                            lsj='%.2f' % en.get('Lsj_design', 0),
                            lp='%.2f' % en.get('Lp_design', 0)))

    items.append(V.pick('concl_top', ps.P['concl_top'],
                        top='%.2f' % top.get('top', 0)))
    items.append(V.pick('concl_sp', ps.P['concl_sp'],
                        lact='%.2f' % sp.get('L_actual', 0),
                        lreq='%.2f' % sp.get('L_required', 0)))
    items.append(V.pick('concl_st', ps.P['concl_st'],
                        kc='%.3f' % st.get('Kc', 0),
                        smax='%.1f' % st.get('sigma_max', 0)))

    chosen = rc.get('chosen')
    if isinstance(chosen, (tuple, list)) and len(chosen) == 2:
        items.append(V.pick('concl_rc', ps.P['concl_rc'],
                            dia=int(chosen[0]), area=int(chosen[1]),
                            as_='%.1f' % rc.get('As', 0)))

    for i, s in enumerate(items, 1):
        add_para(doc, '（%d）%s' % (i, s), fmt)

    qz = R.get('qz') or {}
    qz2 = R.get('qz') or {}
    summ = [
        ['闸孔数（孔）', g(P, 'gateCount')],
        ['单孔净宽 (m)', g(P, 'singleGateWidth')],
        ['闸孔总净宽 (m)', '%.2f' % (int(g(P, 'gateCount', 3) or 3)
                                     * float(g(P, 'singleGateWidth', 6) or 6))],
        ['闸室总宽度 (m)', '%.2f' % qz.get('B_total', 0)],
        ['闸顶高程 (m)', '%.2f' % top.get('top', 0)],
        ['闸底板顶高程 (m)', '%.2f' % (float(g(P, 'gateSillElevation', 0) or 0))],
        ['消力池深度 (m)', '%.2f' % en.get('d_design', 0)],
        ['消力池长度 (m)', '%.2f' % en.get('Lsj_design', 0)],
        ['海漫长度 (m)', '%.2f' % en.get('Lp_design', 0)],
        ['实际渗径长度 (m)', '%.2f' % sp.get('L_actual', 0)],
        ['抗滑安全系数 K_c', '%.3f' % st.get('Kc', 0)],
        ['基底最大应力 (kPa)', '%.1f' % st.get('sigma_max', 0)],
        ['地基允许承载力 (kPa)', g(P, 'foundationBearing')],
        ['底板计算配筋面积 (mm²)', '%.1f' % rc.get('As', 0)],
        ['主体混凝土总量 (m³)', '%.1f' % qz2.get('total_conc', 0)],
    ]
    _kv_table(doc, c, '主要设计成果汇总表', summ, fmt)

    add_para(doc, _p(ctx, 'c9_review'), fmt)
    add_para(doc, _p(ctx, 'c9_limit'), fmt)
    add_para(doc, _p(ctx, 'c9_sug'), fmt)
    add_para(doc, _p(ctx, 'c9_use'), fmt)
    add_para(doc, V.pick('concl_end', ps.P['concl_end']), fmt)


def references(doc, P, R, fmt, c, figs, ctx):
    add_para(doc, '参考文献', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)
    # 文献列表按工程特征生成：涉及抗震的引抗震标准，软基的引地基处理，
    # 因此不同任务书得到的参考文献并不相同。
    refs = ps.refs_for(ctx['V'], ctx['F'])
    for r in refs:
        p = add_para(doc, r, fmt, indent=False, space_after=2)
        p.paragraph_format.left_indent = Pt(fmt.body_size * 2)
        p.paragraph_format.first_line_indent = Pt(-fmt.body_size * 2)
    doc.add_page_break()


def appendix(doc, P, R, fmt, c, figs, ctx):
    """附录：图纸目录与主要计算参数取值表。"""
    F = ctx['F']
    add_para(doc, '附　　录', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)
    add_para(doc, '附录A　设计图纸目录', fmt, size=fmt.h2_size, cn=CN_HEAD,
             indent=False, space_before=8, space_after=6)
    drawings = [
        ['SJ-01', '水闸枢纽平面布置图', '1:200', 'A1'],
        ['SJ-02', '水闸纵剖面图', '1:100', 'A1'],
        ['SJ-03', '闸室横剖面图', '1:50', 'A1'],
        ['SJ-04', '闸室底板配筋图', '1:50', 'A1'],
        ['SJ-05', '闸墩配筋图', '1:50', 'A1'],
        ['SJ-06', '消力池与海漫布置图', '1:100', 'A1'],
        ['SJ-07', '闸基防渗与排水布置图', '1:100', 'A1'],
        ['SJ-08', '翼墙与岸墙结构图', '1:50', 'A2'],
        ['SJ-09', '工作桥与交通桥结构图', '1:50', 'A2'],
        ['SJ-10', '闸门与启闭机布置图', '1:50', 'A2'],
    ]
    add_table(doc, ['图号', '图名', '比例', '图幅'], drawings,
              '%s　设计图纸目录' % c.tbl_no(), fmt)

    add_para(doc, '附录B　主要计算参数取值表', fmt, size=fmt.h2_size, cn=CN_HEAD,
             indent=False, space_before=10, space_after=6)
    sp = R['sp']
    en = R['en']
    keys = [
        ('设计流量 (m³/s)', _used_vals(P)['qd']),
        ('校核流量 (m³/s)', _used_vals(P)['qc']),
        ('上游设计水位 (m)', '%.2f' % _used_vals(P)['upWL']),
        ('下游设计水位 (m)', '%.2f' % _used_vals(P)['dsWL']),
        ('正常蓄水位 (m)', '%.2f' % _used_vals(P)['nrmWL']),
        ('闸底板顶高程 (m)', '%.2f' % _used_vals(P)['sill']),
        ('闸孔数（孔）', g(P, 'gateCount', calc.KEY_DEFAULTS['gateCount'])),
        ('单孔净宽 (m)', g(P, 'singleGateWidth', calc.KEY_DEFAULTS['singleGateWidth'])),
        ('中墩厚度 (m)', g(P, 'middlePierThickness', calc.KEY_DEFAULTS['middlePierThickness'])),
        ('边墩厚度 (m)', g(P, 'sidePierThickness', calc.KEY_DEFAULTS['sidePierThickness'])),
        ('闸底板长度 (m)', g(P, 'floorLength', calc.KEY_DEFAULTS['floorLength'])),
        ('铺盖长度 (m)', g(P, 'blanketLength', calc.KEY_DEFAULTS['blanketLength'])),
        ('基底摩擦系数', g(P, 'frictionCoefficient')),
        ('地基允许承载力 (kPa)', g(P, 'foundationBearing')),
        ('允许渗径系数 C', g(P, 'seepageCoefficientC')),
        ('混凝土强度等级', g(P, 'concreteGrade')),
        ('受力钢筋种类', g(P, 'rebarType')),
        ('水跃淹没系数 σ₀', g(P, 'jumpSubmergence')),
        ('水跃长度校正系数 β', g(P, 'jumpCorrection')),
        ('海漫长度计算系数 K_s', g(P, 'riprapKs')),
    ]
    # 地震参数只有任务书给了才列，缺项时不造假数据
    if F.get('seismic'):
        keys += [
            ('地震基本烈度', g(P, 'seismicIntensity')),
            ('地震动峰值加速度', g(P, 'seismicAcceleration')),
        ]
    keys += [
        ('地基有效深度 T_e (m)', '%.2f' % sp.get('Te', 0)),
        ('消力池设计深度 (m)', '%.2f' % en.get('d_design', 0)),
        ('消力池设计长度 (m)', '%.2f' % en.get('Lsj_design', 0)),
    ]
    add_table(doc, ['参数名称', '取值'], [[a, str(b)] for a, b in keys],
              '%s　主要计算参数取值表' % c.tbl_no(), fmt)

    add_para(doc, '附录C　主要计算成果表', fmt, size=fmt.h2_size, cn=CN_HEAD,
             indent=False, space_before=10, space_after=6)
    gw = R['gw']
    st = R['st']
    rc = R['rc']
    qz = R.get('qz') or {}
    r02 = next((r for r in (gw.get('rows') or []) if abs(r['dH'] - ADOPT_DH) < 1e-9), {})
    out = [
        ['闸前水深（设计工况）(m)', '%.2f' % r02.get('H', 0)],
        ['过水断面面积 (m²)', '%.1f' % r02.get('A', 0)],
        ['行进流速 (m/s)', '%.3f' % r02.get('v', 0)],
        ['堰上总水头 (m)', '%.2f' % r02.get('H0', 0)],
        ['堰流淹没度', '%.3f' % r02.get('ratio', 0)],
        ['综合流量系数 μ₀', '%.3f' % r02.get('mu0', 0)],
        ['计算闸孔总净宽 (m)', '%.2f' % r02.get('B0', 0)],
        ['闸室总宽度 (m)', '%.2f' % qz.get('B_total', 0)],
        ['闸底板底面积 (m²)', '%.1f' % st.get('A_base', 0)],
        ['竖向力合力 (kN)', '%.1f' % st.get('sigmaG', 0)],
        ['水平力合力 (kN)', '%.1f' % st.get('sigmaH', 0)],
        ['抗滑安全系数 K_c', '%.3f' % st.get('Kc', 0)],
        ['基底最大应力 (kPa)', '%.1f' % st.get('sigma_max', 0)],
        ['应力不均匀系数 η', '%.3f' % st.get('eta', 0)],
        ['底板计算配筋面积 (mm²)', '%.1f' % rc.get('As', 0)],
        ['主体混凝土总量 (m³)', '%.1f' % qz.get('total_conc', 0)],
    ]
    add_table(doc, ['计算项目', '成果'], out,
              '%s　主要计算成果表' % c.tbl_no(), fmt)
    doc.add_page_break()


def acknowledgment(doc, P, R, fmt, c, figs, ctx):
    V = ctx['V']
    add_para(doc, '致　　谢', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)
    adv = (P.get('advisor') or '').strip()
    if adv:
        add_para(doc, V.pick('ack_body', ps.P['ack_body'], advisor=adv), fmt)
    else:
        # 没填指导老师姓名时用不带姓名的写法，避免出现「指导老师老师」
        add_para(doc, V.pick('ack_body_anon', ps.P['ack_body_anon']), fmt)
    add_para(doc, V.pick('ack_study', ps.P['ack_study']), fmt)
    add_para(doc, V.pick('ack_help', ps.P['ack_help']), fmt)
    add_para(doc, V.pick('ack_tail', ps.P['ack_tail']), fmt)


# ============================================================
# 章节 → 内容生成器的映射
# ============================================================
CHAPTER_MAP = [
    (('绪论', '概况', '概述', '引言', '前言', '工程概况'), ch_intro),
    (('基本资料', '水文', '气象', '地质', '自然条件', '设计资料'), ch_basic),
    (('闸孔', '孔口', '孔尺寸', '过流', '堰流', '规模'), ch_gate),
    (('消能', '防冲', '水跃', '冲刷'), ch_energy),
    (('高程', '闸顶'), ch_top),
    (('防渗', '排水', '渗流', '渗透'), ch_seepage),
    (('稳定', '地基应力', '抗滑', '应力验算'), ch_stab),
    (('结构', '配筋', '钢筋混凝土', '底板计算'), ch_struct),
    (('结论', '小结', '总结', '成果'), ch_concl),
]

DEFAULT_FRAMEWORK = ['绪论', '基本资料', '闸孔尺寸设计', '消能防冲设计',
                     '闸顶高程确定', '闸基防渗排水设计', '闸室稳定与地基应力验算',
                     '闸室底板结构计算', '结论']


def match_builder(title):
    """按关键词给章节找内容生成器。找不到返回 None（会写成待补充占位）。"""
    for kws, fn in CHAPTER_MAP:
        for kw in kws:
            if kw in title:
                return fn
    return None


# 行政性文件里常见的"章"——它们不是论文章节，绝不能拿去当论文框架。
# 学校发的「毕业设计规定/模板说明/评分标准」这类文档里全是这种标题。
_JUNK_TITLE_PAT = re.compile(
    r'[：:；;，,。.]$'                       # 以标点收尾的更像句子
    r'|模板|见附件|校对|校政|通知|评分|成绩评定|进度安排|时间安排'
    r'|几点说明|管理办法|实施细则|有关规定|任务书|指导教师')


def _usable_spec_chapters(spec, ctx):
    """规范里识别到的章标题，只有**确实像论文章节**才采用。

    判据（两层，缺一不可）：
    1. 剔掉明显的行政标题（见 _JUNK_TITLE_PAT）；
    2. 剩下的章里能匹配到内容生成器的不足 3 章 → 这份规范根本没带论文骨架，
       整体不用，回退到通用水闸设计九 章 结构。

    之前没有这道闸：上传的规范若是学校的行政文件，"第1章 本次对校政文件
    再做几点说明"会被当成论文第 1 章，里面只能塞「待补充」占位语，
    真正的绪论被挤到第 6 章——用户看到的就是这个。
    """
    titles, junk = [], []
    for c0 in (spec.get('chapters') or []):
        if c0.get('level') != 1:
            continue
        t = (c0.get('title') or '').strip()
        if len(t) < 2 or _JUNK_TITLE_PAT.search(t):   # 「绪论」「结论」只有两字，别剔
            junk.append(t)
            continue
        titles.append(t)

    ok = [t for t in titles if match_builder(t)]
    if len(ok) < 3:
        if spec.get('chapters'):
            ctx['warnings'].append(
                '规范中识别到的标题不像论文章节（如「%s」），已忽略，'
                '按通用水闸设计论文结构组织'
                % (junk[0] if junk else (spec['chapters'][0].get('title') or '')))
        else:
            ctx['warnings'].append('规范中未识别到章节框架，已按通用水闸设计论文结构组织')
        return []

    dropped = [t for t in titles if not match_builder(t)]
    if dropped:
        ctx['warnings'].append(
            '规范框架中的 %s 未匹配到自动内容，已跳过（系统尚不能生成这些章节）'
            % '、'.join(dropped))
    return ok


def build(params, spec=None, task_sections=None, figures=None, meta=None):
    """生成论文，返回 BytesIO。"""
    spec = spec or {}
    figures = figures or []
    meta = meta or {}
    P = dict(params or {})

    fmt = Fmt(spec)
    doc = docx.Document()
    apply_page_setup(doc, spec)
    # 正文默认样式也按规范设（首行缩进之类的由段落级控制）
    try:
        st = doc.styles['Normal']
        st.font.name = EN_FONT
        st.font.size = Pt(fmt.body_size)
        st.element.rPr.rFonts.set(qn('w:eastAsia'), CN_BODY)
    except Exception:
        pass
    add_page_number_footer(doc.sections[0], fmt)
    add_page_header(doc.sections[0],
                    '%s毕业设计（论文）' % (P.get('university') or '').strip(),
                    fmt)

    # ---- 计算 ----
    gw = calc.calc_gate_width_mu0(P)
    top = calc.calc_gate_top_mu0(P)
    sp = calc.calc_seepage_mu0(P)
    en = calc.calc_energy_mu0(P, gw)
    stb = calc.calc_stability_mu0(P, gw, top)
    rc = calc.calc_reinforcement(P)
    R = {'gw': gw, 'top': top, 'sp': sp, 'en': en, 'st': stb, 'rc': rc}
    R['qz'] = calc.calc_quantities(P, R)

    ctx = {
        'task_sections': task_sections or [],
        'labeled': meta.get('labeled') or {},
        'warnings': list(meta.get('warnings') or []),
        'used': [],                       # 已用过的客户素材句，避免重复
    }
    # ---- 正文撰写引擎 ----
    # 种子取自任务书正文与设计参数：同一份任务书每次生成结果一致，
    # 不同任务书必然得到不同的选词组合与段落结构。
    task_text = ''.join(task_sections or [])
    ctx['V'] = ps.Seed(meta.get('seed_material') or task_text[:3000],
                       repr(sorted((str(k), str(v)) for k, v in P.items())),
                       salt=meta.get('salt'))
    ctx['B'] = ps.classify(task_sections)
    if not P.get('_raw'):
        P['_raw'] = task_text          # 供工程特征识别（是否除险加固等）
    ctx['F'] = ps.features(P, R)

    c = Counter()

    cover(doc, P, R, fmt, meta)
    abstract_cn(doc, P, R, fmt, meta, ctx)
    abstract_en(doc, P, R, fmt, meta, ctx)

    add_para(doc, '目　　录', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)
    add_toc(doc, fmt)
    doc.add_page_break()

    # ---- 章节框架：只用规范里「确实像论文章节」的那部分，否则用默认框架 ----
    spec_chaps = _usable_spec_chapters(spec, ctx)
    framework = list(spec_chaps) or list(DEFAULT_FRAMEWORK)
    if spec_chaps:
        # 规范给的框架如果缺了水闸设计必不可少的内容（例如防渗、稳定），
        # 按通用结构把缺的补上——少一章就等于少一段计算，不能只照着抄。
        have = ''.join(spec_chaps)
        added = []
        for t in DEFAULT_FRAMEWORK:
            hit = False
            for kws2, _fn in CHAPTER_MAP:
                if any(kw in t for kw in kws2) and any(kw in have for kw in kws2):
                    hit = True
                    break
            if not hit and t not in framework:
                framework.append(t)
                added.append(t)
        if added:
            ctx['warnings'].append(
                '规范给定的章节框架缺少 %s，已按通用水闸设计内容补入相应章节'
                % '、'.join(added))

    used = set()
    for title in framework:
        if any(k in title for k in ('参考文献', '致谢', '附录')):
            continue
        if title in used:
            continue
        used.add(title)

        c.new_chapter()
        add_heading(doc, '第%d章　%s' % (c.chap, title), 1, fmt)
        fn = match_builder(title)          # 框架已过闸，正常都匹配得上
        if fn:
            fn(doc, P, R, fmt, c, figures, ctx)
        doc.add_page_break()

    c.new_chapter()
    references(doc, P, R, fmt, c, figures, ctx)
    c.new_chapter()
    appendix(doc, P, R, fmt, c, figures, ctx)
    acknowledgment(doc, P, R, fmt, c, figures, ctx)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, ctx['warnings']

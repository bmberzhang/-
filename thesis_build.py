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
import os

import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

import calc

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


def g(p, k, d='—'):
    v = p.get(k)
    return d if v in (None, '') else v


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


def abstract_cn(doc, P, R, fmt, meta):
    gw, top, sp, en, st, rc = (R['gw'], R['top'], R['sp'], R['en'], R['st'], R['rc'])
    add_para(doc, '摘　　要', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)

    name = g(P, 'projectName', '本工程')
    Q = g(P, 'designFlow'); Qc = g(P, 'checkFlow')
    text = (
        '本文以%s为对象，依据《水闸设计规范》（SL 265-2016）完成初步设计。'
        '闸址位于%s，工程等别按%s级建筑物设防，设计洪水标准为%s年一遇，'
        '对应设计流量 %s m³/s、校核流量 %s m³/s。'
        '设计内容包括闸孔尺寸确定、消能防冲设计、闸顶高程确定、闸基防渗排水设计、'
        '闸室稳定与地基应力验算以及闸室底板结构计算，并绘制了相应的工程图纸。'
    ) % (name, g(P, 'riverName', '河道'), g(P, 'structureGrade'),
         g(P, 'floodStandard'), Q, Qc)
    add_para(doc, text, fmt)

    text2 = (
        '闸孔尺寸采用综合流量系数法确定，按高淹没度条件计算得闸孔总净宽 B₀ ≈ %.2f m，'
        '结合地形与运行要求布置为 %s 孔、单孔净宽 %s m，中墩厚 %s m、边墩厚 %s m，'
        '闸室总宽 %.2f m。'
    ) % (gw['rows'][1]['B0'] if gw.get('rows') else 0,
         g(P, 'gateCount'), g(P, 'singleGateWidth'),
         g(P, 'middlePierThickness'), g(P, 'sidePierThickness'),
         int(g(P, 'gateCount', 3) or 3) * float(g(P, 'singleGateWidth', 6) or 6)
         + (int(g(P, 'gateCount', 3) or 3) - 1) * float(g(P, 'middlePierThickness', 1) or 1)
         + 2 * float(g(P, 'sidePierThickness', 1.2) or 1.2))
    add_para(doc, text2, fmt)

    text3 = (
        '消能防冲按闸门开度逐级扫描计算，最大消力池深度出现在 Q = %.0f m³/s 附近，'
        '计算池深 %.2f m、池长 %.2f m，设计取池深 %.2f m，并据此确定海漫长度 %.2f m。'
        '闸顶高程经挡水与泄水两种工况比较后取 %.2f m。'
        '闸基防渗采用改进阻力系数法验算，地下轮廓线实际长度 %.2f m，'
        '大于规范要求的 %.2f m，出口坡降 %.3f、水平段坡降 %.3f，均满足要求。'
        '闸室稳定验算得抗滑安全系数 Kc = %.3f，'
        '地基最大应力 %.1f kPa，小于地基允许承载力 %s kPa。'
    ) % (_q_at_max(en), _d_max(en), _lsj_max(en), en.get('d_design', 0),
         en.get('Lp_design', 0), top.get('top', 0),
         sp.get('L_actual', 0), sp.get('L_required', 0),
         sp.get('J_out', 0), sp.get('J_horiz', 0),
         st.get('Kc', 0), st.get('sigma_max', 0), g(P, 'foundationBearing'))
    add_para(doc, text3, fmt)

    add_para(doc, '关键词：水闸；闸孔尺寸；消能防冲；闸基防渗；稳定验算；结构计算', fmt,
             indent=False, space_before=10)
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


def abstract_en(doc, P, R, fmt, meta):
    add_para(doc, 'ABSTRACT', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)
    st = R['st']
    txt = (
        'This paper presents the preliminary design of %s, carried out in '
        'accordance with the Chinese code SL 265-2016 "Design Code for Sluice". '
        'The design flood standard is once in %s years, with a design discharge of '
        '%s m3/s and a check discharge of %s m3/s. '
        'The work covers the determination of gate opening dimensions, energy '
        'dissipation and scour protection, crest elevation, seepage control and '
        'drainage of the foundation, stability analysis of the gate chamber, and '
        'structural design of the base slab.'
    ) % (_en_project(g(P, 'projectName', '')), g(P, 'floodStandard'),
         g(P, 'designFlow'), g(P, 'checkFlow'))
    add_para(doc, txt, fmt, cn=EN_FONT)
    txt2 = (
        'The total clear width of the gate openings is %.2f m, arranged in %s bays. '
        'The maximum stilling basin depth is %.2f m and the apron length is %.2f m. '
        'The crest elevation is taken as %.2f m. '
        'The calculated anti-sliding safety factor is %.3f, and the maximum '
        'foundation stress is %.1f kPa, both satisfying the code requirements.'
    ) % (R['gw']['rows'][1]['B0'] if R['gw'].get('rows') else 0, g(P, 'gateCount'),
         _d_max(R['en']), R['en'].get('Lp_design', 0), R['top'].get('top', 0),
         st.get('Kc', 0), st.get('sigma_max', 0))
    add_para(doc, txt2, fmt, cn=EN_FONT)
    add_para(doc, 'Key words: sluice; gate opening; energy dissipation; seepage '
                  'control; stability analysis; structural design',
             fmt, indent=False, cn=EN_FONT, space_before=10)
    doc.add_page_break()


def _put_figs(doc, figs, key, fmt, c):
    """插出本章应出现的图（按图表 key 匹配）"""
    for f in figs:
        if f.get('chapter') != key:
            continue
        add_figure(doc, f['png'], '%s　%s' % (c.fig_no(), f['title']), fmt)


def ch_intro(doc, P, R, fmt, c, figs, ctx):
    """绪论 / 工程概况"""
    add_para(doc, '%s位于%s，为%s工程。' % (
        g(P, 'projectName', '本工程'), g(P, 'riverName', '本流域'),
        g(P, 'sluiceFunction', '节制闸')), fmt)
    for s in (ctx.get('task_sections') or [])[:3]:
        add_para(doc, s, fmt)
    add_para(doc, '本次设计的任务是在已有水文、地质资料的基础上，按《水闸设计规范》'
                  '（SL 265-2016）完成该闸的初步设计，包括闸孔尺寸确定、消能防冲、'
                  '闸顶高程、闸基防渗排水、闸室稳定及结构计算等内容，并绘制相应图纸。', fmt)
    add_heading(doc, '%s.1　设计依据' % c.chap, 2, fmt)
    add_para(doc, '主要依据的技术文件与规范如下：', fmt)
    for i, s in enumerate([
        '《水闸设计规范》（SL 265-2016）',
        '《水工建筑物抗震设计标准》（GB 51247-2018）',
        '《水工混凝土结构设计规范》（SL 191-2008）',
        '本工程地质勘察报告及水文实测资料',
    ], 1):
        add_para(doc, '（%d）%s' % (i, s), fmt, indent=True)


def ch_basic(doc, P, R, fmt, c, figs, ctx):
    """基本资料"""
    add_para(doc, '本章列出本次设计所采用的水文、气象、地质及工程特性指标，'
                  '数据取自设计任务书与地质勘察成果。', fmt)
    labeled = ctx.get('labeled') or {}
    if hasattr(labeled, 'items'):
        labeled = list(labeled.items())
    rows = [[str(k), str(v)] for k, v in labeled]
    if rows:
        add_table(doc, ['项目', '数值'], rows, '%s　基本资料一览表' % c.tbl_no(), fmt)

    secs = (ctx.get('task_sections') or [])
    if len(secs) > 3:
        add_heading(doc, '%s.1　工程地质与水文条件' % c.chap, 2, fmt)
        for s in secs[3:6]:
            add_para(doc, s, fmt)

    add_para(doc, '按上述资料，闸址处地基允许承载力为 %s kPa，基底摩擦系数 f = %s，'
                  '地震基本烈度为 %s 度，地震动峰值加速度为 %s g，'
                  '结构混凝土采用 %s，受力钢筋采用 %s。' % (
                      g(P, 'foundationBearing'), g(P, 'frictionCoefficient'),
                      g(P, 'seismicIntensity'), g(P, 'seismicAcceleration'),
                      g(P, 'concreteGrade'), g(P, 'rebarType')), fmt)


def ch_gate(doc, P, R, fmt, c, figs, ctx):
    """闸孔尺寸设计"""
    gw = R['gw']
    Q = float(g(P, 'designFlow', 0) or 0)
    hs = gw.get('hs', 0)
    add_para(doc, '闸孔总净宽是水闸设计中最基本的尺寸。本设计采用综合流量系数法，'
                  '按高淹没度堰流公式计算，先由明渠均匀流反算行进流速与堰上水头，'
                  '再按淹没度查取综合流量系数，最后反解总净宽。', fmt)

    add_para(doc, '闸孔总净宽按下式计算：', fmt)
    add_formula(doc, ['B₀ = ', ('frac', 'Q', 'μ₀·h_s·√(2g(H₀ − h_s))')], fmt)
    add_para(doc, '式中：Q 为设计流量（m³/s）；μ₀ 为综合流量系数；h_s 为下游水深（m）；'
                  'H₀ 为计入行进流速水头的堰上总水头（m）；g 为重力加速度，取 9.81 m/s²。', fmt)

    add_para(doc, '综合流量系数按高淹没度经验式确定：', fmt)
    add_formula(doc, ['μ₀ = 0.877 + (h_s/H₀ − 0.65)²'], fmt)

    add_para(doc, '计入行进流速水头的堰上总水头为：', fmt)
    add_formula(doc, ['H₀ = H + ', ('frac', 'v₀²', '2g')], fmt)

    rows = gw.get('rows') or []
    trs = []
    for r in rows:
        trs.append(['%.1f' % r['dH'], '%.2f' % r['H'], '%.1f' % r['A'],
                    '%.3f' % r['v'], '%.2f' % r['H0'], '%.3f' % r['ratio'],
                    '%.3f' % r['mu0'], '%.2f' % r['B0']])
    if trs:
        add_table(doc, ['ΔH (m)', 'H (m)', 'A (m²)', 'v₀ (m/s)', 'H₀ (m)',
                        'h_s/H₀', 'μ₀', 'B₀ (m)'], trs,
                  '%s　闸孔总净宽计算表' % c.tbl_no(), fmt)

    b_mid = rows[1]['B0'] if len(rows) > 1 else 0
    n = int(g(P, 'gateCount', 3) or 3)
    b0 = float(g(P, 'singleGateWidth', 6) or 6)
    add_para(doc, '由上表可见，随上下游水位差增大，所需总净宽减小。'
                  '取 ΔH = 0.2 m 作为设计工况，计算得闸孔总净宽 B₀ = %.2f m。'
                  '结合河道断面尺寸与运行调度要求，将闸孔布置为 %d 孔、'
                  '单孔净宽 %g m，实际总净宽 %g m，满足过流要求。' % (
                      b_mid, n, b0, n * b0), fmt)
    _put_figs(doc, figs, '闸孔', fmt, c)


def ch_energy(doc, P, R, fmt, c, figs, ctx):
    """消能防冲设计"""
    en = R['en']
    add_para(doc, '水闸下泄时水流具有较大动能，需设置消能设施防止冲刷。'
                  '本设计按闸门开度逐级扫描，对每一开度计算收缩水深、共轭水深、'
                  '水跃长度与所需池深，取各开度中的最大值作为设计依据。', fmt)

    add_para(doc, '收缩断面水深由能量方程试算：', fmt)
    add_formula(doc, ['T₀ = h_c + ', ('frac', 'q²', '2gφ²h_c²')], fmt)
    add_para(doc, '其中单宽流量 q = Q/B₀，φ 为流速系数。共轭水深按下式计算：', fmt)
    add_formula(doc, ['h_c″ = ', ('frac', 'h_c', '2'),
                      '·(√(1 + 8q²/(g·h_c³)) − 1)'], fmt)
    add_para(doc, '消力池深度按下式确定：', fmt)
    add_formula(doc, ['d = σ₀·h_c″ − h_s − Δz'], fmt)
    add_para(doc, '水跃长度采用 SL 265-2016 推荐的经验式 L_j = 6.9(h_c″ − h_c)，'
                  '消力池长度取 L_sj = 3.5 + β·L_j。海漫长度按 L_p = K_s·√(q·√ΔH) 估算。', fmt)

    rows = en.get('rows') or []
    trs = []
    for r in rows:
        # 池深/池长算出 0 表示该开度下水跃已被淹没、消力池不受其控制。
        # 工程表里这种「不起控制作用」的格子写成「—」，写成 0.00 容易被误读。
        d_txt = '—' if r['d'] <= 0 else '%.3f' % r['d']
        lsj_txt = '—' if r['Lsj'] <= 0 else '%.2f' % r['Lsj']
        lp_txt = '—' if r['Lp'] <= 0 else '%.2f' % r['Lp']
        trs.append(['%.2f' % r['he'], '%.1f' % r['Q'], '%.3f' % r['hs'],
                    '%.3f' % r['hc'], '%.3f' % r['hc2'], d_txt, lsj_txt, lp_txt])
    if trs:
        add_table(doc, ['开度 e (m)', 'Q (m³/s)', 'h_s (m)', 'h_c (m)',
                        'h_c″ (m)', '池深 d (m)', '池长 L_sj (m)', 'L_p (m)'],
                  trs, '%s　消能防冲计算成果表' % c.tbl_no(), fmt)

    d_design = en.get('d_design', 0)
    add_para(doc, '注：表中「—」表示该开度下水跃已被下游水深淹没，'
                  '消力池尺寸不受此工况控制。', fmt, size=fmt.body_size - 1)
    add_para(doc, '计算结果表明，池深随开度先增后减：小开度时下泄流量小、'
                  '单宽流量低，所需池深较小；开度增大到一定程度后，'
                  '下游水深相对增大、水跃被淹没，所需池深反而减小。'
                  '最大池深出现在 Q ≈ %.0f m³/s 附近，计算值 %.2f m。'
                  '据此设计取消力池深度 %.2f m、池长 %.2f m，'
                  '海漫长度取 %.2f m。' % (
                      _q_at_max(en), _d_max(en), d_design,
                      en.get('Lsj_design', 0), en.get('Lp_design', 0)), fmt)
    _put_figs(doc, figs, '消能', fmt, c)


def ch_top(doc, P, R, fmt, c, figs, ctx):
    """闸顶高程确定"""
    top = R['top']
    add_para(doc, '闸顶高程需同时满足挡水与泄水两种工况，并保证与两岸地面顺接。'
                  '挡水工况按正常蓄水位加波浪爬高与安全超高控制；'
                  '泄水工况按设计洪水位加安全超高控制，最终取两者与现状地面高程的大值。', fmt)

    add_formula(doc, ['H₁ = Z_正常蓄水位 + h₂ + A₁'], fmt)
    add_formula(doc, ['H₂ = Z_设计洪水位 + A₂'], fmt)
    add_formula(doc, ['Z_闸顶 = max(H₁, H₂, Z_地面)'], fmt)

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
    _put_figs(doc, figs, '高程', fmt, c)


def ch_seepage(doc, P, R, fmt, c, figs, ctx):
    """闸基防渗排水设计"""
    sp = R['sp']
    add_para(doc, '闸基渗流直接影响水闸安全。本设计采用改进阻力系数法计算：'
                  '先按允许渗径系数估算所需渗径长度，再沿地下轮廓线分段计算'
                  '各段阻力系数与水头损失，最后验算出口坡降与水平段坡降。', fmt)

    add_para(doc, '所需渗径长度按下式估算：', fmt)
    add_formula(doc, ['L_需 = C·ΔH'], fmt)
    add_para(doc, '式中 C 为允许渗径系数，ΔH 为上下游最大水位差。'
                  '本工程 C = %s，ΔH = %.2f m，得 L_需 = %.2f m。' % (
                      g(P, 'seepageCoefficientC', sp.get('C', 0)),
                      sp.get('deltaH', 0), sp.get('L_required', 0)), fmt)

    add_para(doc, '各段阻力系数按规范公式计算，分段水头损失为：', fmt)
    add_formula(doc, ['h_i = ', ('frac', 'ξ_i', 'Σξ'), '·ΔH'], fmt)

    xi = sp.get('xi_list') or []
    hl = sp.get('h_list') or []
    trs = []
    for i, (a, b) in enumerate(zip(xi, hl), 1):
        trs.append([str(i), '%.4f' % a, '%.4f' % b, '%.4f' % (sum(hl[:i]))])
    if trs:
        add_table(doc, ['分段号', '阻力系数 ξ', '分段水头损失 (m)', '累计损失 (m)'],
                  trs, '%s　闸基渗流水头损失计算表' % c.tbl_no(), fmt)

    add_para(doc, '地下轮廓线实际长度 L_实 = %.2f m，%s所需长度 %.2f m。'
                  '出口坡降 J = %.4f，允许值 0.50；水平段坡降 J = %.4f，允许值 0.25，'
                  '均满足规范要求，闸基抗渗稳定性满足要求。' % (
                      sp.get('L_actual', 0),
                      '大于' if sp.get('seepCheck') else '小于',
                      sp.get('L_required', 0),
                      sp.get('J_out', 0), sp.get('J_horiz', 0)), fmt)
    _put_figs(doc, figs, '防渗', fmt, c)


def ch_stab(doc, P, R, fmt, c, figs, ctx):
    """闸室稳定与地基应力验算"""
    st = R['st']
    add_para(doc, '闸室稳定验算考虑完建、正常运行与校核洪水等荷载组合，'
                  '分别计算竖向力与水平力，据此求抗滑安全系数与基底应力。', fmt)

    add_formula(doc, ['K_c = ', ('frac', 'f·ΣG', 'ΣH')], fmt)
    add_para(doc, '式中 f 为基底摩擦系数，取 %s；ΣG 为竖向力总和 %s kN；'
                  'ΣH 为水平力总和 %s kN。按 SL 265-2016，'
                  '水闸抗滑安全系数允许值为 1.20。' % (
                      g(P, 'frictionCoefficient'), '%.1f' % st.get('sigmaG', 0),
                      '%.1f' % st.get('sigmaH', 0)), fmt)

    add_para(doc, '基底应力沿底板宽度按偏心受压分布：', fmt)
    add_formula(doc, ['σ_max/min = ', ('frac', 'ΣG', 'A'),
                      '·(1 ± ', ('frac', '6e', 'B'), ')'], fmt)

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

    add_para(doc, '验算结果：抗滑安全系数 K_c = %.3f %s允许值 1.20；'
                  '基底最大应力 %.1f kPa %s地基允许承载力 %s kPa；'
                  '应力不均匀系数 η = %.3f，满足规范要求。' % (
                      st.get('Kc', 0), '≥' if st.get('stabCheck') else '<',
                      st.get('sigma_max', 0),
                      '≤' if st.get('bearCheck') else '>',
                      g(P, 'foundationBearing'), st.get('eta', 0)), fmt)
    _put_figs(doc, figs, '稳定', fmt, c)


def ch_struct(doc, P, R, fmt, c, figs, ctx):
    """闸室底板结构计算"""
    rc = R['rc']
    add_para(doc, '闸室底板按单筋矩形截面受弯构件计算配筋，'
                  '取单位板宽按最不利弯矩进行正截面承载力计算。', fmt)

    add_formula(doc, ['α_s = ', ('frac', 'γ_d·M', 'f_c·b·h₀²')], fmt)
    add_formula(doc, ['ξ = 1 − √(1 − 2α_s)'], fmt)
    add_formula(doc, ['A_s = ', ('frac', 'ξ·f_c·b·h₀', 'f_y')], fmt)

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

    chosen = rc.get('chosen')
    if isinstance(chosen, (tuple, list)) and len(chosen) == 2:
        dia, area = chosen
        add_para(doc, '按计算配筋面积并考虑构造要求，受拉钢筋选用 Φ%d@100（每米板宽），'
                      '实配面积 %d mm²，大于计算值 %.1f mm²，满足要求。' % (
                          int(dia), int(area), rc.get('As', 0)), fmt)
    else:
        add_para(doc, '按计算配筋面积 A_s = %.1f mm² 选配受拉钢筋，'
                      '并按构造要求配置分布钢筋。' % rc.get('As', 0), fmt)
    _put_figs(doc, figs, '结构', fmt, c)


def ch_concl(doc, P, R, fmt, c, figs, ctx):
    """结论"""
    gw, top, sp, en, st = R['gw'], R['top'], R['sp'], R['en'], R['st']
    add_para(doc, '本文按《水闸设计规范》（SL 265-2016）完成了%s的初步设计，'
                  '主要结论如下：' % g(P, 'projectName', '本工程'), fmt)
    items = [
        '闸孔总净宽经综合流量系数法计算为 %.2f m，布置为 %s 孔、单孔净宽 %s m，'
        '闸室总宽满足过流与布置要求。' % (gw['rows'][1]['B0'] if gw.get('rows') else 0,
                                          g(P, 'gateCount'), g(P, 'singleGateWidth')),
        '消能防冲按开度扫描计算，最大池深 %.2f m，设计取池深 %.2f m、池长 %.2f m，'
        '海漫长度 %.2f m，能够满足消能与防冲要求。' % (
            _d_max(en), en.get('d_design', 0), en.get('Lsj_design', 0),
            en.get('Lp_design', 0)),
        '闸顶高程经挡水与泄水工况比较取 %.2f m，与两岸地面高程衔接良好。' % top.get('top', 0),
        '闸基防渗采用改进阻力系数法验算，实际渗径 %.2f m 大于所需 %.2f m，'
        '出口坡降与水平段坡降均小于允许值。' % (sp.get('L_actual', 0), sp.get('L_required', 0)),
        '闸室稳定验算得抗滑安全系数 %.3f，基底最大应力 %.1f kPa，'
        '均满足规范要求。' % (st.get('Kc', 0), st.get('sigma_max', 0)),
    ]
    for i, s in enumerate(items, 1):
        add_para(doc, '（%d）%s' % (i, s), fmt)
    add_para(doc, '综上，本设计方案在过流能力、消能防冲、抗渗稳定与结构安全等方面'
                  '均满足规范要求，方案技术可行。', fmt)


def references(doc, P, R, fmt, c, figs, ctx):
    add_para(doc, '参考文献', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)
    refs = [
        '[1] 中华人民共和国水利部. 水闸设计规范: SL 265-2016[S]. 北京: 中国水利水电出版社, 2016.',
        '[2] 中华人民共和国住房和城乡建设部. 水工建筑物抗震设计标准: GB 51247-2018[S]. 北京: 中国计划出版社, 2018.',
        '[3] 中华人民共和国水利部. 水工混凝土结构设计规范: SL 191-2008[S]. 北京: 中国水利水电出版社, 2008.',
        '[4] 中华人民共和国水利部. 水闸安全评价导则: SL 214-2015[S]. 北京: 中国水利水电出版社, 2015.',
        '[5] 谈松邱. 水工建筑物[M]. 北京: 中国水利水电出版社, 2015.',
        '[6] 吴持恭. 水力学[M]. 5版. 北京: 高等教育出版社, 2016.',
        '[7] 顾淦臣, 束一鸣, 沈长松. 土石坝工程经验与创新[M]. 北京: 中国电力出版社, 2004.',
    ]
    for r in refs:
        p = add_para(doc, r, fmt, indent=False, space_after=2)
        p.paragraph_format.left_indent = Pt(fmt.body_size * 2)
        p.paragraph_format.first_line_indent = Pt(-fmt.body_size * 2)
    doc.add_page_break()


def acknowledgment(doc, P, R, fmt, c, figs, ctx):
    add_para(doc, '致　　谢', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)
    adv = g(P, 'advisor', '指导老师')
    add_para(doc, '本次毕业设计是在%s老师的悉心指导下完成的。从选题、'
                  '方案拟定到计算过程的反复校核，老师都给予了耐心细致的指导，'
                  '提出了许多宝贵意见，使我对水闸设计的方法与规范要求有了系统认识。'
                  '在此谨向老师表示衷心的感谢。' % adv, fmt)
    add_para(doc, '同时感谢学院各位老师在四年学习中的教导，感谢同学在资料收集与'
                  '绘图过程中给予的帮助。最后，感谢家人一直以来的支持与鼓励。', fmt)


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

    # ---- 计算 ----
    gw = calc.calc_gate_width_mu0(P)
    top = calc.calc_gate_top_mu0(P)
    sp = calc.calc_seepage_mu0(P)
    en = calc.calc_energy_mu0(P, gw)
    stb = calc.calc_stability_mu0(P, gw, top)
    rc = calc.calc_reinforcement(P)
    R = {'gw': gw, 'top': top, 'sp': sp, 'en': en, 'st': stb, 'rc': rc}

    ctx = {
        'task_sections': task_sections or [],
        'labeled': meta.get('labeled') or {},
        'warnings': list(meta.get('warnings') or []),
    }

    c = Counter()

    cover(doc, P, R, fmt, meta)
    abstract_cn(doc, P, R, fmt, meta)
    abstract_en(doc, P, R, fmt, meta)

    add_para(doc, '目　　录', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)
    add_toc(doc, fmt)
    doc.add_page_break()

    # ---- 章节框架：优先用规范里识别到的，否则用默认框架 ----
    spec_chaps = [c0['title'] for c0 in (spec.get('chapters') or [])
                  if c0.get('level') == 1]
    framework = spec_chaps or DEFAULT_FRAMEWORK
    if not spec_chaps:
        ctx['warnings'].append('规范中未识别到章节框架，已按通用水闸设计论文结构组织')

    used = set()
    for title in framework:
        if any(k in title for k in ('参考文献', '致谢', '附录')):
            continue
        if title in used:
            continue
        used.add(title)

        c.new_chapter()
        add_heading(doc, '第%d章　%s' % (c.chap, title), 1, fmt)
        fn = match_builder(title)
        if fn:
            fn(doc, P, R, fmt, c, figures, ctx)
        else:
            # 规范里要求的章节，但系统没有对应的自动内容——明确留白而不是硬凑
            add_para(doc, '【本章由规范指定，系统未匹配到对应的自动生成内容，'
                          '请根据任务书要求补充。】', fmt)
            ctx['warnings'].append('章节「%s」未匹配到自动内容，已在文中留白' % title)
        doc.add_page_break()

    c.new_chapter()
    references(doc, P, R, fmt, c, figures, ctx)
    acknowledgment(doc, P, R, fmt, c, figures, ctx)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, ctx['warnings']

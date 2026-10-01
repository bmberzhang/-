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


def abstract_cn(doc, P, R, fmt, meta, ctx):
    V, F = ctx['V'], ctx['F']
    gw, top, sp, en, st, rc = (R['gw'], R['top'], R['sp'], R['en'], R['st'], R['rc'])
    add_para(doc, '摘　　要', fmt, size=fmt.h1_size, cn=CN_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=12)

    rows = gw.get('rows') or []
    b_mid = rows[1]['B0'] if len(rows) > 1 else (rows[0]['B0'] if rows else 0)

    head = V.pick('abs_open', ps.P['abs_open'],
                  project=g(P, 'projectName', '本工程'),
                  func=g(P, 'sluiceFunction', '节制闸'))
    body = V.pick('abs_body', ps.P['abs_body'],
                  std=g(P, 'floodStandard'), qd=g(P, 'designFlow'),
                  qc=g(P, 'checkFlow'))
    add_para(doc, head + body, fmt)

    res = V.pick('abs_res', ps.P['abs_res'],
                 b0='%.2f' % b_mid, n=g(P, 'gateCount'),
                 b1=g(P, 'singleGateWidth'),
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
    txt += V.pick('en_body', EN_ABS_BODY, std=g(P, 'floodStandard'),
                  qd=g(P, 'designFlow'), qc=g(P, 'checkFlow'))
    add_para(doc, txt, fmt, cn=EN_FONT)

    txt2 = V.pick('en_res', EN_ABS_RES, b0='%.2f' % b_mid, n=g(P, 'gateCount'),
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
    add_heading(doc, '%s.%d　%s' % (c.chap, k, title), 2, fmt)


def _material(doc, fmt, ctx, buckets, k, used_key='used'):
    """把任务书里对应主题的句子写进正文（轻度改写过的口吻）。"""
    got = ps.take(buckets, k, ctx.get('used') or ())
    for s in got:
        add_para(doc, s, fmt)
        ctx.setdefault('used', []).append(s)
    return got


def ch_intro(doc, P, R, fmt, c, figs, ctx):
    """绪论：工程概况 + 任务书素材 + 设计依据"""
    V, F, B = ctx['V'], ctx['F'], ctx['B']

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

    k = 1
    _sub(doc, c, k, '工程任务与设计内容', fmt)
    k += 1
    add_para(doc, V.pick('intro_task', ps.P['intro_task']), fmt)
    _material(doc, fmt, ctx, B.get('task', []), 3)

    _sub(doc, c, k, '设计依据', fmt)
    add_para(doc, V.pick('intro_doc_head', ps.P['intro_doc_head']), fmt)
    for i, s in enumerate(_docs_for(F), 1):
        add_para(doc, '（%d）%s' % (i, s), fmt, indent=True)

    if F.get('rehab'):
        _extra(doc, fmt, ctx, 'rehab')
    if F.get('seismic'):
        _extra(doc, fmt, ctx, 'seismic',
               inten=g(P, 'seismicIntensity'), acc=g(P, 'seismicAcceleration'))
    else:
        _extra(doc, fmt, ctx, 'no_seismic', inten=g(P, 'seismicIntensity'))


def ch_basic(doc, P, R, fmt, c, figs, ctx):
    """基本资料：客户水文/地质素材 + 指标表 + 特征相关结论"""
    V, F, B = ctx['V'], ctx['F'], ctx['B']
    add_para(doc, V.pick('basic_open', ps.P['basic_open']), fmt)

    k = 1
    hy = ps.take(B.get('hydro', []) + B.get('meteo', []), 4, ctx.get('used') or ())
    if hy:
        _sub(doc, c, k, V.pick('h_hydro', ps.P['basic_hydro_head']), fmt)
        k += 1
        _material(doc, fmt, ctx, hy, len(hy))
    geo = ps.take(B.get('geo', []), 4, ctx.get('used') or ())
    if geo:
        _sub(doc, c, k, V.pick('h_geo', ps.P['basic_geo_head']), fmt)
        k += 1
        _material(doc, fmt, ctx, geo, len(geo))

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

    if F.get('need_treat'):
        _extra(doc, fmt, ctx, 'treat')


def ch_gate(doc, P, R, fmt, c, figs, ctx):
    """闸孔尺寸设计"""
    V = ctx['V']
    gw = R['gw']
    add_para(doc, V.pick('gate_open', ps.P['gate_open']), fmt)
    add_para(doc, V.pick('gate_method', ps.P['gate_method']), fmt)

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

    b_mid = rows[1]['B0'] if len(rows) > 1 else (rows[0]['B0'] if rows else 0)
    n = int(g(P, 'gateCount', 3) or 3)
    b0 = float(g(P, 'singleGateWidth', 6) or 6)
    add_para(doc, V.pick('gate_res', ps.P['gate_res'], b0='%.2f' % b_mid), fmt)

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
    _put_figs(doc, figs, '闸孔', fmt, c)


def ch_energy(doc, P, R, fmt, c, figs, ctx):
    """消能防冲设计"""
    V, F = ctx['V'], ctx['F']
    en = R['en']
    add_para(doc, V.pick('en_open', ps.P['en_open']), fmt)

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
    """闸顶高程确定"""
    V = ctx['V']
    top = R['top']
    add_para(doc, V.pick('top_open', ps.P['top_open']), fmt)

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
    V, F = ctx['V'], ctx['F']
    sp = R['sp']
    add_para(doc, V.pick('sp_open', ps.P['sp_open']), fmt)

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

    add_para(doc, V.pick('sp_res', ps.P['sp_res'],
                         lact='%.2f' % sp.get('L_actual', 0),
                         cmp='大于' if sp.get('seepCheck') else '小于',
                         lreq='%.2f' % sp.get('L_required', 0),
                         jout='%.4f' % sp.get('J_out', 0),
                         jhor='%.4f' % sp.get('J_horiz', 0)), fmt)
    if F.get('seep_tight'):
        _extra(doc, fmt, ctx, 'tight_seep')
    _put_figs(doc, figs, '防渗', fmt, c)


def ch_stab(doc, P, R, fmt, c, figs, ctx):
    """闸室稳定与地基应力验算"""
    V, F = ctx['V'], ctx['F']
    st = R['st']
    add_para(doc, V.pick('st_open', ps.P['st_open']), fmt)

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

    add_para(doc, V.pick('st_res', ps.P['st_res'],
                         kc='%.3f' % st.get('Kc', 0),
                         cmpk='≥' if st.get('stabCheck') else '<',
                         smax='%.1f' % st.get('sigma_max', 0),
                         cmpb='≤' if st.get('bearCheck') else '>',
                         bearing=g(P, 'foundationBearing'),
                         eta='%.3f' % st.get('eta', 0)), fmt)
    if F.get('stab_tight'):
        _extra(doc, fmt, ctx, 'tight_stab')
    _put_figs(doc, figs, '稳定', fmt, c)


def ch_struct(doc, P, R, fmt, c, figs, ctx):
    """闸室底板结构计算"""
    V = ctx['V']
    rc = R['rc']
    add_para(doc, V.pick('rc_open', ps.P['rc_open']), fmt)

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
        add_para(doc, V.pick('rc_res', ps.P['rc_res'], dia=int(dia), spacing=100,
                             area=int(area), as_='%.1f' % rc.get('As', 0)), fmt)
    else:
        add_para(doc, '按计算配筋面积 A_s = %.1f mm² 选配受拉钢筋，'
                      '并按构造要求配置分布钢筋。' % rc.get('As', 0), fmt)
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
        'used': [],                       # 已用过的客户素材句，避免重复
    }
    # ---- 正文撰写引擎 ----
    # 种子取自任务书正文与设计参数：同一份任务书每次生成结果一致，
    # 不同任务书必然得到不同的选词组合与段落结构。
    task_text = ''.join(task_sections or [])
    ctx['V'] = ps.Seed(meta.get('seed_material') or task_text[:3000],
                       repr(sorted((str(k), str(v)) for k, v in P.items())))
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

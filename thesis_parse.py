# -*- coding: utf-8 -*-
"""任务书 / 规范的 Word 文档解析

两个入口：
    parse_task_book(path) -> {'params': {...}, 'sections': [...], 'fields': [...], 'raw': str}
    parse_spec(path)      -> {'page': {...}, 'fonts': {...}, 'chapters': [...],
                              'has_skeleton': bool, 'notes': [...]}

设计原则
--------
1. **先表格后正文**：任务书的参数绝大多数写在 2 列表格里（"设计流量 | 174 m³/s"），
   表格里 key-value 挨着，匹配比在整段正文里正则更准。所以先扫表格，再扫正文。
2. **抽不到就留空，绝不猜**。宁可让用户在页面上补一个数，也不能编一个数进论文——
   论文里的数据错了是灾难性的。
3. **每条命中都带出处**，前端可以显示"这条是从任务书第 3 张表里读到的"，便于核对。
"""
import os
import re

NUM = r'(\d+(?:[.,]\d+)?)'

# 中文字号 → 磅值
CN_SIZE = {
    '初号': 42, '小初': 36, '一号': 26, '小一': 24, '二号': 22, '小二': 18,
    '三号': 16, '小三': 15, '四号': 14, '小四': 12, '五号': 10.5, '小五': 9,
    '六号': 7.5, '小六': 6.5, '七号': 5.5, '八号': 5,
}


# ============================================================
# 读 docx
# ============================================================
def read_docx(path):
    """读 docx，返回 (段落文本列表, 表格二维列表, Document 对象)"""
    import docx
    d = docx.Document(path)
    paras = [p.text.strip() for p in d.paragraphs if p.text and p.text.strip()]
    tables = []
    for tbl in d.tables:
        rows = []
        for row in tbl.rows:
            cells, seen = [], None
            for c in row.cells:
                t = ' '.join(c.text.split())
                if t != seen:            # 合并单元格会重复出现同一个对象
                    cells.append(t)
                    seen = t
            if any(cells):
                rows.append(cells)
        if rows:
            tables.append(rows)
    return paras, tables, d


def _unquote(s):
    """去掉整体包裹的引号/书名号。

    只剥成对包裹的引号——像「“HG”水闸拆除重建工程」这种书名号只在
    首部出现的情况，整体是工程名的一部分，不能把开引号削掉。
    """
    t = (s or '').strip()
    for a, b in (('“', '”'), ('‘', '’'), ('"', '"'), ('《', '》'), ('「', '」')):
        if len(t) > 2 and t.startswith(a) and t.endswith(b):
            t = t[1:-1]
            break
    return t.strip()


def _norm(s):
    """全角转半角、去千分位、统一负号"""
    if s is None:
        return ''
    out = []
    for ch in str(s):
        code = ord(ch)
        if code == 0x3000:
            out.append(' ')
        elif 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        else:
            out.append(ch)
    t = ''.join(out)
    t = t.replace('－', '-').replace('—', '-')
    t = re.sub(r'(?<=\d)[,，](?=\d{3}\b)', '', t)   # 去掉千分位逗号
    return t.strip()


def _to_float(s):
    s = _norm(s)
    m = re.search(r'-?\d+(?:\.\d+)?', s)
    return float(m.group()) if m else None


# 小标题样式：一、/ 第一章 /（一）/ 1. / 1.1 —— 这些是标题不是正文
_HEAD_LIKE = re.compile(
    r'^\s*(?:[（(]?\s*[一二三四五六七八九十]+\s*[）)]?\s*[、.．]'
    r'|第\s*[一二三四五六七八九十\d]+\s*[章节部分]'
    r'|\d{1,2}(?:\.\d{1,2}){0,2}\s*[、.．)）]?)')


def _is_heading_like(s):
    """判断一段文字是不是小标题（没有句末标点、且以编号开头/很短）。"""
    t = (s or '').strip()
    if not t:
        return False
    if t[-1] in '。；：:':
        return False
    return bool(_HEAD_LIKE.match(t))


def _size_pt(tok):
    """把 '小四号' / '四号' / '12磅' / '12pt' 统一换算成磅值。

    坑：正则的候选里 `小X号` 必须排在 `X号` 前面，否则 '小四号' 会先被
    `[一二三四五六七八九十]号` 咬掉 '四号'，得到 14 磅（应为 12 磅）。
    """
    if not tok:
        return None
    t = tok.strip().replace(' ', '')
    if t in CN_SIZE:
        return CN_SIZE[t]
    if t.endswith('号') and t[:-1] in CN_SIZE:
        return CN_SIZE[t[:-1]]
    m = re.search(r'\d+(?:\.\d+)?', t)
    return float(m.group()) if m else None


# 尺寸候选的正则片段（注意 `小X号` 在前）
_SIZE_PAT = r'(小[一二三四五六七八九十]号|[一二三四五六七八九十]号|\d+(?:\.\d+)?\s*(?:磅|pt|Pt))'


def _kv_from_tables(tables):
    """把表格拍平成 (键, 值, 出处) 列表"""
    out = []
    for ti, rows in enumerate(tables):
        for ri, cells in enumerate(rows):
            for i in range(len(cells) - 1):
                k, v = cells[i], cells[i + 1]
                if k and len(k) <= 26 and v and len(v) <= 60:
                    out.append((k, v, '表%d 第%d行' % (ti + 1, ri + 1)))
    return out


# 只能以文本形式取的参数（不参与数值解析）
TEXT_KEYS = {'concreteGrade', 'rebarType', 'riverName', 'foundationType'}


# ============================================================
# 参数抽取规则
# ============================================================
# key: (显示名, 单位, [正文正则...], [表头关键词...])
RULES = [
    ('floodStandard', '防洪标准', '年一遇',
     [r'(?:设计)?防洪标准[^0-9]{0,10}' + NUM, NUM + r'\s*年\s*一遇'],
     ['防洪标准', '设计标准', '洪水标准', '频率']),
    ('designFlow', '设计流量', 'm³/s',
     [r'设计流量[^0-9]{0,14}' + NUM],
     ['设计流量', '设计洪水流量']),
    ('checkFlow', '校核流量', 'm³/s',
     [r'校核流量[^0-9]{0,14}' + NUM],
     ['校核流量']),
    ('downstreamWaterLevel', '设计洪水位（下游）', 'm',
     [r'设计洪水位[^0-9]{0,14}' + NUM],
     ['设计洪水位', '下游水位']),
    ('normalStorageLevel', '正常蓄水位', 'm',
     [r'正常蓄水位[^0-9]{0,14}' + NUM],
     ['正常蓄水位', '蓄水位']),
    ('gateSillElevation', '闸底板顶高程', 'm',
     [r'(?:闸)?底板(?:顶)?高程[^0-9]{0,14}' + NUM,
      r'闸底高程[^0-9]{0,14}' + NUM],
     ['底板高程', '闸底高程', '底板顶高程']),
    ('groundElevation', '现状地面高程', 'm',
     [r'地面高程[^0-9]{0,14}' + NUM],
     ['地面高程', '滩地高程']),
    ('checkWaterLevel', '校核洪水位', 'm',
     [r'校核洪水位[^0-9]{0,14}' + NUM],
     ['校核洪水位']),
    ('seismicIntensity', '地震基本烈度', '度',
     [r'地震[^。；\n]{0,16}?烈度[^0-9]{0,8}' + NUM,
      NUM + r'\s*度[^。；\n]{0,8}地震'],
     ['地震烈度', '基本烈度', '抗震设防烈度']),
    ('seismicAcceleration', '地震动峰值加速度', 'g',
     [r'加速度[^0-9]{0,12}(0?\.\d+)'],
     ['峰值加速度', '地震加速度']),
    ('foundationBearing', '地基允许承载力', 'kPa',
     [r'承载力[^0-9]{0,14}' + NUM],
     ['承载力', '地基承载力', '允许承载力']),
    ('frictionCoefficient', '基底摩擦系数', '',
     [r'(?:基底|闸室|底板)?摩擦系数[^0-9]{0,10}(0?\.\d+)'],
     ['摩擦系数', '基底摩擦']),
    ('permeabilityCoefficient', '渗透系数', 'cm/s',
     [r'渗透系数[^0-9]{0,14}(\d+(?:\.\d+)?(?:[×xX]\s*10-?\d+)?)'],
     ['渗透系数']),
    ('compressionModulus', '压缩模量', 'MPa',
     [r'压缩模量[^0-9]{0,14}' + NUM],
     ['压缩模量']),
    ('concreteGrade', '混凝土强度等级', '',
     [r'混凝土[^。；\n]{0,10}?(C\d{2})'],
     ['混凝土', '混凝土强度等级']),
    ('rebarType', '钢筋种类', '',
     [r'(HRB\d{3}|HPB\d{3}|RRB\d{3})'],
     ['钢筋', '钢筋种类', '受力钢筋']),
    ('gateCount', '闸孔数', '孔',
     [r'闸孔数[^0-9]{0,10}' + NUM, r'共\s*' + NUM + r'\s*孔', NUM + r'\s*孔'],
     ['闸孔数', '孔数', '孔口数量']),
    ('singleGateWidth', '单孔净宽', 'm',
     [r'单孔净宽[^0-9]{0,10}' + NUM, r'净宽[^0-9]{0,8}' + NUM],
     ['单孔净宽', '孔口净宽', '单孔宽度']),
    ('middlePierThickness', '中墩厚度', 'm',
     [r'中墩[^0-9]{0,10}' + NUM],
     ['中墩厚', '中墩厚度']),
    ('sidePierThickness', '边墩厚度', 'm',
     [r'边墩[^0-9]{0,10}' + NUM],
     ['边墩厚', '边墩厚度', '岸墙厚']),
    ('windSpeed', '设计风速', 'm/s',
     [r'风速[^0-9]{0,12}' + NUM],
     ['风速', '设计风速']),
    ('structureGrade', '建筑物级别', '级',
     [r'[^。；\n]{0,10}建筑物级别[^0-9]{0,8}' + NUM,
      NUM + r'\s*级建筑物'],
     ['建筑物级别', '工程等级', '级别']),
    ('riverName', '所在河流', '',
     [r'位于([^，。；\n]{2,12}河)'],
     ['河流', '所在河流', '河流名称']),
    ('channelBottomWidth', '主河槽底宽', 'm',
     [r'(?:主槽|主河槽|河槽)底宽[^0-9]{0,10}' + NUM],
     ['河槽底宽', '主槽底宽']),
    ('foundationType', '地基类型', '',
     [r'地基[^。；\n]{0,10}?(?:为|是)([^，。；\n]{2,14})'],
     ['地基', '地基类型', '地基土']),
]

_PARAM_KEYS = {r[0] for r in RULES}


def parse_task_book(path):
    """解析任务书，返回抽取结果。

    params 里的值是字符串（与前端表单一致），抽不到的键不出现在结果里。
    """
    paras, tables, doc = read_docx(path)
    kv = _kv_from_tables(tables)
    body = '\n'.join(paras)
    for rows in tables:                       # 表格文字也进正文，兜底用
        for cells in rows:
            body += '\n' + ' '.join(cells)

    params, sources = {}, []
    for key, label, unit, pats, headers in RULES:
        hit = None
        is_text = key in TEXT_KEYS

        # 1) 先在表格里找。文本类参数只认「完全相同」或「以关键词结尾」的匹配——
        #    否则 '地基' 会咬到 '地基允许承载力' 那一行，把地基类型取成 300。
        for h in headers:
            for k, v, src in kv:
                kk = _norm(k)
                strong = (kk == h) or kk.endswith(h)
                if not strong:
                    continue
                if is_text:
                    hit = (v.strip(), src)
                else:
                    fv = _to_float(v)
                    if fv is None:
                        continue
                    hit = (fv, src)
                break
            if hit:
                break

        # 2) 表格没找到，再用正文正则
        if hit is None:
            for pat in pats:
                m = re.search(pat, body)
                if m:
                    g = m.group(1)
                    if is_text:
                        hit = (g.strip(), '正文')
                    else:
                        fv = _to_float(g)
                        if fv is None:
                            continue
                        hit = (fv, '正文')
                    break

        if hit is None or hit[0] in (None, ''):
            continue

        val, src = hit
        if is_text:
            params[key] = str(val)
        elif isinstance(val, float):
            if float(val).is_integer() and key in ('gateCount', 'seismicIntensity',
                                                   'structureGrade', 'floodStandard'):
                params[key] = str(int(val))
            else:
                params[key] = ('%g' % val)
        else:
            params[key] = str(val)
        sources.append({'key': key, 'label': label, 'unit': unit,
                        'value': params[key], 'from': src})

    # 工程概况类叙述段落（长句 + 关键词）
    kw_narr = ('工程概况', '流域', '地形', '地质', '水文', '气象', '闸址', '概况',
               '任务', '设计依据', '规模', '工程任务')
    sections = []
    for p in paras:
        if len(p) < 32:
            continue
        if _is_heading_like(p):        # 一、二、三、/（一）/ 1. 这类小标题不要当正文
            continue
        if any(k in p for k in kw_narr):
            sections.append(p)

    params.setdefault('_doc_ok', True)

    # 河流名：规则表没命中时，从正文里找出现次数最多的「X河」。
    # 任务书常写成「沙河李庄节制闸位于沙河下游…」，规则表要求「位于X河」，
    # 河名后面跟了方位词就抓不到，这里补一刀。
    if not (params.get('riverName') or '').strip():
        bad = ('河床', '河道', '河口', '河水', '河流', '河段', '河槽', '河岸',
               '河势', '河宽', '河堤', '河滩', '河底', '河边', '河渠', '河系')
        cnt = {}
        for p in paras[:25]:
            for m in re.finditer(r'([\u4e00-\u9fa5]{1,4}河)', p):
                w = m.group(1)
                if w in bad or len(w) < 2:
                    continue
                cnt[w] = cnt.get(w, 0) + 1
        if cnt:
            # 出现次数最多者优先；并列时取较长的（更可能是全名）
            best = max(cnt.items(), key=lambda kv: (kv[1], len(kv[0])))[0]
            params['riverName'] = best
            sources.append({'key': 'riverName', 'label': '所在河流', 'unit': '',
                            'value': best, 'from': '正文推断'})

    # 题目：任务书里通常写成「题目：XXX」或在开头有一行带“工程/设计”的标题
    project_name = ''
    for p in paras[:15]:
        m = re.search(r'题\s*目\s*[:：]\s*(.+)', p)
        if m:
            project_name = _unquote(m.group(1))
            break
    if not project_name:
        for p in paras[:8]:
            t = _unquote(p)
            if not (6 <= len(t) <= 40):
                continue
            if t[-1] in '。；，,':           # 标题不会以句读结尾
                continue
            if _is_heading_like(t):         # 「一、」「1.1」这类是章节标题
                continue
            if '工程' in t or '设计' in t:
                project_name = t
                break

    return {
        'params': {k: v for k, v in params.items() if not k.startswith('_')},
        'sources': sources,
        'sections': sections[:40],
        'project_name': project_name,
        'n_tables': len(tables),
        'n_paras': len(paras),
        'raw': body,
    }


# ============================================================
# 规范解析
# ============================================================
_CHAP_PAT = re.compile(
    r'^\s*(?:第\s*([一二三四五六七八九十]+|\d+)\s*[章部分])\s*[、.．]?\s*(.{2,40})$')
_NUM_CHAP_PAT = re.compile(r'^\s*(\d{1,2})(?:\.\d{1,2})*\s+(\S.{1,38})$')


def parse_spec(path):
    """解析毕业设计规范：页面设置、字体要求、章节框架。"""
    paras, tables, doc = read_docx(path)
    out = {'page': {}, 'fonts': {}, 'chapters': [], 'notes': [],
           'has_skeleton': False, 'page_setup_ok': False}

    # ---- 页面设置 ----
    try:
        sec = doc.sections[0]
        from docx.shared import Cm
        out['page']['w_cm'] = round(sec.page_width.cm, 2)
        out['page']['h_cm'] = round(sec.page_height.cm, 2)
        out['page']['top_cm'] = round(sec.top_margin.cm, 2)
        out['page']['bottom_cm'] = round(sec.bottom_margin.cm, 2)
        out['page']['left_cm'] = round(sec.left_margin.cm, 2)
        out['page']['right_cm'] = round(sec.right_margin.cm, 2)
        out['page_setup_ok'] = True
    except Exception:
        pass

    # ---- 样式里的字体（规范若自带样式，直接继承）----
    try:
        normal = doc.styles['Normal'].font
        out['fonts']['body_size'] = float(normal.size.pt) if normal.size else None
        out['fonts']['body_name'] = normal.name
    except Exception:
        pass
    for i, key in ((1, 'h1'), (2, 'h2'), (3, 'h3')):
        try:
            st = doc.styles['Heading %d' % i].font
            if st.size:
                out['fonts'][key + '_size'] = float(st.size.pt)
        except Exception:
            pass

    body = '\n'.join(paras)

    # ---- 从文字要求里读字号 / 行距 ----
    m = re.search(r'正文[^。；\n]{0,24}?' + _SIZE_PAT, body)
    if m:
        pt = _size_pt(m.group(1))
        if pt:
            out['fonts']['body_size'] = pt
        out['notes'].append('规范正文要求：%s' % m.group(0)[:40])
    m = re.search(r'行距[^。；\n]{0,16}?(\d+(?:\.\d+)?)\s*(磅|pt|倍)', body)
    if m:
        out['fonts']['line_spacing'] = float(m.group(1))
        out['fonts']['line_spacing_is_multiple'] = ('倍' in m.group(2))
        out['notes'].append('规范行距要求：%s' % m.group(0)[:40])
    for lv, key in (('一级', 'h1'), ('二级', 'h2'), ('三级', 'h3')):
        m = re.search(lv + r'标题[^。；\n]{0,24}?' + _SIZE_PAT, body)
        if m:
            pt = _size_pt(m.group(1))
            if pt:
                out['fonts'][key + '_size'] = pt

    # ---- 章节框架 ----
    seen = set()
    for p in paras:
        if len(p) > 44:
            continue
        m = _CHAP_PAT.match(p)
        if m:
            title = m.group(2).strip()
            if title and title not in seen:
                seen.add(title)
                out['chapters'].append({'level': 1, 'title': title})
            continue
        m = _NUM_CHAP_PAT.match(p)
        if m:
            level = 1 if '.' not in p.split()[0] else p.split()[0].count('.') + 1
            title = m.group(2).strip()
            if title and title not in seen and level <= 3:
                seen.add(title)
                out['chapters'].append({'level': level, 'title': title})

    # ---- 学校名称（规范标题里通常带，用来做封面抬头）----
    for p in paras[:12]:
        m = re.search(r'([\u4e00-\u9fa5]{2,12}?(?:大学|学院|职业技术学院))', p)
        if m:
            out['university'] = m.group(1)
            break

    # ---- 是否自带论文骨架 ----
    marks = ['摘要', '目录', '参考文献', '致谢', '结论']
    out['has_skeleton'] = sum(1 for k in marks if k in body) >= 3 and len(out['chapters']) >= 5
    out['notes'].append('从规范中识别出 %d 个章节标题' % len(out['chapters']))
    return out


def spec_hint(spec):
    """给前端看的一句话概括"""
    bits = []
    p = spec.get('page') or {}
    if spec.get('page_setup_ok'):
        bits.append('页面 %.1f×%.1f cm' % (p.get('w_cm', 0), p.get('h_cm', 0)))
        bits.append('页边距 %.1f/%.1f/%.1f/%.1f cm' % (
            p.get('top_cm', 0), p.get('bottom_cm', 0),
            p.get('left_cm', 0), p.get('right_cm', 0)))
    f = spec.get('fonts') or {}
    if f.get('body_size'):
        bits.append('正文字号 %.1f 磅' % f['body_size'])
    if f.get('line_spacing'):
        bits.append('行距 %g%s' % (f['line_spacing'], '倍' if f.get('line_spacing_is_multiple') else ' 磅'))
    if spec.get('chapters'):
        bits.append('识别到 %d 个章节标题' % len(spec['chapters']))
    return '；'.join(bits) if bits else '未从规范中读到明确的格式条款，将按通用学位论文格式排版'

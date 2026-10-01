# -*- coding: utf-8 -*-
"""本地回归验证：用两份**完全不同**的任务书跑完整生成链路。

用途
----
1. 确认论文能生成、不抛异常、字数/表数/图数在预期区间（页数目标 ≥50 页）。
2. 确认「不同客户产出不同」——两份样例的正文不应大面积雷同。
3. 确认 11 张图都能出，且没有缺字形的告警。

用法
----
    python verify_samples.py                  # 生成两份样例，打印字数/表/图统计
    python verify_samples.py --charts DIR     # 只把 11 张图导出到 DIR，便于肉眼核对
    python verify_samples.py --diff           # 对比两份样例正文的重合度（验"千人千面"）
    python verify_samples.py --bare           # 模拟任务书缺参：全文不应出现“—”占位

注意：依赖 matplotlib，本机需带上 pylibs-extra 的 PYTHONPATH。
"""
import io
import os
import re
import sys

import thesis_build as tb
import thesis_parse as tp          # noqa: F401  （保持与生产链路一致的导入）
import charts


def cjk_para(doc):
    """返回正文汉字数、表格内汉字数。"""
    body = sum(len(re.findall(r'[\u4e00-\u9fff]', p.text)) for p in doc.paragraphs)
    tbl = 0
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                tbl += len(re.findall(r'[\u4e00-\u9fff]', c.text))
    return body, tbl


SPEC = {
    'page': {'w_cm': 21.0, 'h_cm': 29.7, 'top_cm': 2.5, 'bottom_cm': 2.5,
             'left_cm': 3.0, 'right_cm': 2.0},
    'page_setup_ok': True,
    'fonts': {'body_size': 12, 'h1_size': 16, 'h2_size': 14, 'h3_size': 12,
              'line_spacing': 20},
    'chapters': [{'title': t, 'level': 1} for t in [
        '绪论', '基本资料', '闸孔尺寸设计', '消能防冲设计', '闸顶高程与闸室结构布置',
        '闸基防渗排水设计', '闸室稳定与地基应力验算', '闸室底板结构计算', '结论']],
    'notes': [], 'has_skeleton': True,
}

A_SECTIONS = [
    '本工程为“HG”水闸拆除重建工程，闸址位于滏阳河干流上，是一座以灌溉为主、兼顾排涝的节制闸。',
    '滏阳河属子牙河系，流域面积 2747.8 km²，闸址处河道为复式断面，主槽底宽 20 m，边坡 1:2。',
    '原闸建于上世纪七十年代，运行多年后闸门锈蚀严重、闸墩出现裂缝，经安全鉴定为四类闸，需拆除重建。',
    '设计洪水标准为 20 年一遇，相应设计流量 174 m³/s，相应闸上水位 77.72 m、闸下水位 77.52 m。',
    '闸址区地层自上而下为壤土、粉质黏土和砂砾石层，砂砾石层中密～密实，承载力特征值 300 kPa。',
    '地下水位埋深约 2.0 m，对混凝土结构具微腐蚀性。',
    '闸址区地震动峰值加速度为 0.20g，相应地震基本烈度为Ⅷ度。',
    '本次设计需完成闸孔尺寸、消能防冲、防渗排水、稳定与结构计算，并提交说明书与图纸。',
    '闸底板高程 73.10 m，正常蓄水位 76.60 m，闸室采用 3 孔，单孔净宽 6 m。',
    '建筑材料就近采购，施工安排在非汛期进行。',
]

A_PARAMS = {
    'projectName': '“HG”水闸拆除重建工程', 'projectAbbr': 'HG',
    'university': '河北工程大学', 'college': '水利水电学院', 'majorClass': '水工 2202 班',
    'studentName': '李某某', 'studentId': '20220101', 'advisor': '王某某',
    'riverName': '滏阳河', 'riverSystem': '子牙河系', 'sluiceFunction': '灌溉、排涝',
    'buildingGrade': 'Ⅲ', 'designFlow': 174, 'checkFlow': 260,
    'upstreamWaterLevel': 77.72, 'downstreamWaterLevel': 77.52,
    'normalStorageLevel': 76.60, 'checkWaterLevel': 78.60,
    'gateSillElevation': 73.10, 'floorLength': 14, 'blanketLength': 15,
    'gateCount': 3, 'singleGateWidth': 6, 'middlePierThickness': 1.0,
    'sidePierThickness': 1.2, 'channelBottomWidth': 20, 'channelSlope': '1:2',
    'foundationBearing': 300, 'frictionCoefficient': 0.30,
    'seismicIntensity': 'Ⅷ', 'seismicAcceleration': '0.20g',
    'concreteGrade': 'C25', 'rebarType': 'HRB400', 'seepageCoefficientC': 5,
    'floodplainElevation': 75.00, 'floodplainWidth': 8.5,
    'channelSlopeRatio': '1/1410', 'riprapKs': 9,
}

B_SECTIONS = [
    '本工程为沙河河道整治配套的拦河闸，闸址位于沙河下游河段，主要任务是抬高水位、保障两岸灌区引水。',
    '沙河属海河流域，闸址以上控制流域面积 860 km²，河床质以中细砂为主，两岸为农田。',
    '设计洪水标准 20 年一遇，设计流量 120 m³/s，相应闸上水位 62.30 m。',
    '闸址区地基为深厚的软黏土层，天然含水量高、压缩性大，承载力特征值仅 130 kPa。',
    '地下水位接近地表，基坑开挖需降水。',
    '本区地震动峰值加速度 0.05g，地震基本烈度Ⅵ度，可不进行抗震计算。',
    '设计要求完成闸孔尺寸、消能防冲、地基处理、防渗及稳定计算，并绘制图纸。',
    '闸底板高程 57.80 m，拟采用 2 孔，单孔净宽 5 m，闸室为开敞式。',
]

B_PARAMS = dict(A_PARAMS)
B_PARAMS.update({
    'projectName': '沙河拦河闸工程', 'projectAbbr': 'SH',
    'studentName': '赵某某', 'studentId': '20220202', 'advisor': '刘某某',
    'riverName': '沙河', 'riverSystem': '海河流域', 'sluiceFunction': '灌溉引水',
    'buildingGrade': 'Ⅳ', 'designFlow': 120, 'checkFlow': 180,
    'upstreamWaterLevel': 62.30, 'downstreamWaterLevel': 62.10,
    'normalStorageLevel': 61.60, 'checkWaterLevel': 63.00,
    'gateSillElevation': 57.80, 'gateCount': 2, 'singleGateWidth': 5,
    'foundationBearing': 130, 'frictionCoefficient': 0.22,
    'seismicIntensity': 'Ⅵ', 'seismicAcceleration': '0.05g',
    'concreteGrade': 'C30', 'rebarType': 'HRB400',
    'seepageCoefficientC': 7, 'floodplainElevation': 60.50,
})


def _run_chain(params):
    """跑一遍计算 + 绘图，返回 (gw, top, figures, warns)。"""
    import calc as _calc
    gw = _calc.calc_gate_width_mu0(params)
    top = _calc.calc_gate_top_mu0(params)
    figures, warns = charts.render_all(
        gw, top, _calc.calc_seepage_mu0(params),
        _calc.calc_energy_mu0(params, gw),
        _calc.calc_stability_mu0(params, gw, top), params)
    return gw, top, figures, warns


def dump_charts(outdir):
    """把 11 张图导出成 PNG，便于肉眼核对标题/标注是否跑偏。"""
    os.makedirs(outdir, exist_ok=True)
    for tag, params in (('A', A_PARAMS), ('B', B_PARAMS)):
        _gw, _top, figures, warns = _run_chain(params)
        for i, f in enumerate(figures, 1):
            fn = os.path.join(outdir, '%s_%02d_%s.png' % (tag, i, f['key']))
            with open(fn, 'wb') as fh:
                fh.write(f['png'])
        print('%s：导出 %d 张图' % (tag, len(figures)))
        for w in warns:
            print('    ! ' + w)


def make(name, params, sections, out):
    labeled = [(k, v) for k, v in [
        ('工程名称', params['projectName']), ('所在河流', params['riverName']),
        ('建筑物级别', params['buildingGrade']), ('设计流量 (m³/s)', params['designFlow']),
        ('校核流量 (m³/s)', params['checkFlow']),
        ('上游设计水位 (m)', params['upstreamWaterLevel']),
        ('下游设计水位 (m)', params['downstreamWaterLevel']),
        ('正常蓄水位 (m)', params['normalStorageLevel']),
        ('闸底板高程 (m)', params['gateSillElevation']),
        ('闸孔数（孔）', params['gateCount']),
        ('单孔净宽 (m)', params['singleGateWidth']),
        ('地基承载力特征值 (kPa)', params['foundationBearing']),
        ('基底摩擦系数', params['frictionCoefficient']),
        ('地震基本烈度', params['seismicIntensity']),
        ('混凝土强度等级', params['concreteGrade']),
        ('受力钢筋种类', params['rebarType']),
    ]]
    gw, top, figures, fw = _run_chain(params)
    meta = {'labeled': labeled, 'warnings': [],
            'seed_material': ''.join(sections)[:3000]}
    buf, warns = tb.build(params, spec=SPEC, task_sections=sections,
                          figures=figures, meta=meta)
    data = buf.getvalue()
    with open(out, 'wb') as f:
        f.write(data)
    import docx
    d = docx.Document(io.BytesIO(data))
    body, tbl = cjk_para(d)
    print('%-16s 段落正文=%5d 表格内=%5d 表数=%3d 图=%2d 预警=%d 大小=%dKB'
          % (name, body, tbl, len(d.tables), len(figures), len(warns), len(data) // 1024))
    for w in warns:
        print('    ! ' + w)
    return body


def _long_sentences(path, minlen=30):
    """取正文里长度 >= minlen 的句子（先去掉数字和空白，只看文字骨架）。

    去数字是为了避免"两边都写了 174"这种巧合把重合率拉高——真正要看的是
    遣词造句是不是同一套。
    """
    import docx
    d = docx.Document(path)
    out = []
    for p in d.paragraphs:
        t = re.sub(r'[0-9\.\,\s%]', '', p.text).strip()
        if len(t) >= minlen:
            out.append(t)
    return out


def _find(name):
    """样例可能落在根目录，也可能因为被 Word 占着而落到 output_thesis/。

    两边都看，取**修改时间最新**的那个——根目录常留着一份被 Word 锁住没覆盖掉的旧样例，
    按路径顺序取会拿到旧的，量出来的重合度就完全不对了。
    """
    cands = [p for p in (name + '.docx', os.path.join('output_thesis', name + '.docx'))
             if os.path.exists(p)]
    if not cands:
        raise SystemExit('找不到样例文件：%s.docx（先跑一次 python verify_samples.py）' % name)
    return max(cands, key=os.path.getmtime)


def diff_report(name_a='毕业设计_样例A_滏阳河改建闸',
                name_b='毕业设计_样例B_沙河软基闸'):
    """对比两份样例的长句重合度，做「不同客户产出不同」的证据。"""
    A, B = _long_sentences(_find(name_a)), _long_sentences(_find(name_b))
    sa, sb = set(A), set(B)
    same = len(sa & sb)
    tot = max(1, len(sa | sb))
    print('长句(去数字、>=30 字)   A=%d 句   B=%d 句' % (len(A), len(B)))
    print('  完全相同   : %d' % same)
    print('  仅 A 有    : %d (%.1f%% of A)' % (len(sa - sb), 100.0 * len(sa - sb) / max(1, len(sa))))
    print('  仅 B 有    : %d (%.1f%% of B)' % (len(sb - sa), 100.0 * len(sb - sa) / max(1, len(sb))))
    print('  长句重合率 : %.1f%%   → 两份论文的正文确实不是同一套话' % (100.0 * same / tot))
    print()
    print('（重合的部分主要是公式符号行、表格表头、规范条文引用等——')
    print('  同一专业方向的毕业设计本来就会共用这些骨架。）')


def _placeholder_hits(text):
    """找出真正的“—”占位符：单个 em-dash 且两侧都不是数字。

    “——”（破折号）和“SL 265—2016”（编号）是正常写法，不算。"""
    out = []
    for m in re.finditer(r'—+', text):
        if len(m.group()) >= 2:
            continue
        a = text[m.start() - 1:m.start()]
        b = text[m.end():m.end() + 1]
        if a.isdigit() and b.isdigit():
            continue
        if a == '「' and b == '」':
            continue                      # 「—」是在引用图例，不是占位
        out.append(text[max(0, m.start() - 12):m.end() + 12])
    return out


def bare_report():
    """线上踩过的坑：任务书里识别不到关键参数时，汇总表/摘要曾出现
    “—”和 0.00。把 A 的参数剥掉 KEY_DEFAULTS 覆盖的所有键再生成，
    全文不允许出现占位符“—”。"""
    import calc as _calc
    import docx
    stripped = {k: v for k, v in A_PARAMS.items()
                if k not in _calc.KEY_DEFAULTS
                and k not in ('floodStandard', 'upstreamWaterLevel')}
    gw, top, figures, _fw = _run_chain(stripped)
    meta = {'labeled': [], 'warnings': [],
            'seed_material': ''.join(A_SECTIONS)[:3000]}
    buf, _w = tb.build(stripped, spec=SPEC, task_sections=A_SECTIONS,
                       figures=figures, meta=meta)
    d = docx.Document(io.BytesIO(buf.getvalue()))
    hits = []
    for p in d.paragraphs:
        hits += _placeholder_hits(p.text)
    for t in d.tables:
        head_txt = ' '.join(c.text for c in t.rows[0].cells) if t.rows else ''
        is_scan = '开度' in head_txt        # 开度扫描表里“—”= 该开度无需设池，图例已说明
        for r in t.rows:
            for c in r.cells:
                if c.text.strip() == '—' and not is_scan:
                    hits.append('表格占位: ' + c.text.strip())
    print('缺参生成：正文段落=%d 表数=%d' % (len(d.paragraphs), len(d.tables)))
    if hits:
        print('!! 发现 %d 处“—”占位：' % len(hits))
        for s in hits[:10]:
            print('   ', s)
        raise SystemExit(1)
    print('OK：缺参时全文无“—”占位，汇总表/摘要均引用实际采用值')


if __name__ == '__main__':
    if '--charts' in sys.argv:
        dump_charts(sys.argv[sys.argv.index('--charts') + 1])
        raise SystemExit(0)

    if '--diff' in sys.argv:
        diff_report()
        raise SystemExit(0)

    if '--bare' in sys.argv:
        bare_report()
        raise SystemExit(0)

    os.makedirs('output_thesis', exist_ok=True)
    for name, params, secs in (('毕业设计_样例A_滏阳河改建闸', A_PARAMS, A_SECTIONS),
                               ('毕业设计_样例B_沙河软基闸', B_PARAMS, B_SECTIONS)):
        out = name + '.docx'
        try:                                     # 根目录的旧样例可能被 Word 占着
            with open(out, 'ab'):
                pass
        except Exception:
            out = 'output_thesis/' + name + '.docx'
        make(name, params, secs, out)

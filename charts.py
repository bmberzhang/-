# -*- coding: utf-8 -*-
"""计算曲线图生成（matplotlib → PNG 字节，供插入 docx）

为什么要自带字体
----------------
云服务器（Railway / Docker slim 镜像）里**没有中文字体**，matplotlib 默认字体会把
汉字画成一排方框（tofu）。所以仓库里带了一份 Noto Sans SC 的子集
（`static/fonts/chart_font.ttf`，约 650KB，SIL OFL 协议可自由分发），
绘图时注册进 fontManager 并设为默认字体族，不依赖系统字体。

子集里收录了 ASCII + 希腊字母 + 常用工程符号 + 项目源码出现过的全部汉字。
**图里用到的文字必须来自这个字符集**——所以标签一律写成固定字符串，
不把用户输入的工程名称等任意文本画进图里（图题由 docx 正文的图注承担）。

设计取舍
--------
- 曲线图只画**关系**（尺寸随流量怎么变）；**校核结论**（计算值 vs 允许值）
  放正文表格，不画图——量纲不同的指标塞进同一坐标轴必然把小的压成一条线。
- 柱状图一律从 0 起，不截断纵轴。
- 高程类的对比画成**水平基准线**，而不是柱状（77~79m 的柱子在 0 起点上看着一样高）。
"""
import io
import os

BASE = os.path.dirname(os.path.abspath(__file__))
FONT_PATH = os.path.join(BASE, 'static', 'fonts', 'chart_font.ttf')

_C_MAIN = '#0f5a94'
_C_ALT = '#8a8f98'
_C_THIRD = '#7a5c2e'
_C_LIMIT = '#b03a2e'

DPI = 150
FIG_W = 5.4
FIG_H = 3.1

_FONT_REGISTERED = [False]


def _font(size):
    from matplotlib import font_manager as fm
    if os.path.exists(FONT_PATH):
        return fm.FontProperties(fname=FONT_PATH, size=size)
    return fm.FontProperties(size=size)


def _setup():
    """导入 matplotlib 并设置全局外观。返回 pyplot，失败时抛异常由调用方兜底。

    关键：必须把自带字体**注册进 fontManager 并设为 rcParams 的默认字体族**。
    只给个别 `text(...)` 传 fontproperties 是不够的——刻度标签
    （`set_xticklabels`）不走那块，会静默退回 DejaVu Sans，汉字变成方框，
    只在日志里留一行 "Glyph missing" 警告。
    """
    import matplotlib
    matplotlib.use('Agg')
    from matplotlib import font_manager as fm
    from matplotlib import pyplot as plt

    if os.path.exists(FONT_PATH) and not _FONT_REGISTERED[0]:
        try:
            fm.fontManager.addfont(FONT_PATH)
            family = fm.FontProperties(fname=FONT_PATH).get_name()
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = [family] + list(
                plt.rcParams.get('font.sans-serif', []))
            _FONT_REGISTERED[0] = True
        except Exception:
            pass

    plt.rcParams['axes.unicode_minus'] = False
    plt.rcParams['axes.edgecolor'] = '#4a4a4a'
    plt.rcParams['axes.linewidth'] = 0.7
    return plt


def _frame(ax):
    """统一收边：去掉上/右框，浅网格置底"""
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(True, color='#d8d8d8', linewidth=0.5, alpha=0.9)
    ax.set_axisbelow(True)
    ax.tick_params(labelsize=8, length=3, width=0.6, colors='#333333')


def _label(ax, y, text, color=_C_LIMIT, ha='left', x=None):
    """在水平参考线**上方**留出空隙写标注，避免压在线上"""
    lo, hi = ax.get_ylim()
    ax.text(x if x is not None else ax.get_xlim()[0], y + (hi - lo) * 0.025,
            text, color=color, fontproperties=_font(8), va='bottom', ha=ha)


def _save(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=DPI, bbox_inches='tight',
                facecolor='white', pad_inches=0.08)
    import matplotlib.pyplot as plt
    plt.close(fig)
    return buf.getvalue()


# ============================================================
# 各图
# ============================================================
def fig_gate_width(gw, p):
    """图：闸孔总净宽 B0 随上下游水位差 ΔH 变化"""
    plt = _setup()
    rows = gw.get('rows') or []
    if not rows:
        return None
    xs = ['%.1f' % r['dH'] for r in rows]
    ys = [r['B0'] for r in rows]
    b_total = gw.get('n', 3) * gw.get('b0', 6)

    fig, ax = plt.subplots(figsize=(FIG_W * 0.74, FIG_H * 0.86))
    bars = ax.bar(xs, ys, width=0.4, color=_C_MAIN, alpha=0.85)
    for b, v in zip(bars, ys):
        ax.text(b.get_x() + b.get_width() / 2, v, '%.2f' % v,
                ha='center', va='bottom', fontproperties=_font(8), color='#333333')
    ax.set_ylim(0, max(max(ys), b_total) * 1.3)
    ax.axhline(b_total, color=_C_LIMIT, linestyle='--', linewidth=1)
    _label(ax, b_total, '实际布置总净宽 %.0f m' % b_total, ha='right',
           x=ax.get_xlim()[1])
    ax.set_xlabel('上下游水位差 ΔH (m)', fontproperties=_font(9))
    ax.set_ylabel('闸孔总净宽 B0 (m)', fontproperties=_font(9))
    _frame(ax)
    return _save(fig)


def fig_energy(energy, p):
    """图：消力池尺寸与海漫长度随下泄流量变化（上下两分图）

    不分图的话没法看：池深只有 0.3~0.4m，池长却有十几米，同一个纵轴必然把
    池深压成一条贴着 0 的平线。上下分图各用自己的纵轴，两边都看得清。
    """
    plt = _setup()
    rows = [r for r in (energy.get('rows') or []) if r.get('Q', 0) > 0]
    if not rows:
        return None
    x = [r['Q'] for r in rows]
    d = [r['d'] for r in rows]
    lsj = [r['Lsj'] for r in rows]
    lp = [r['Lp'] for r in rows]

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(FIG_W, FIG_H * 1.7),
                                   sharex=True, gridspec_kw={'hspace': 0.22})

    # 上：消力池深度（设计值取自这里的最大值）
    ax1.plot(x, d, marker='o', ms=3.4, color=_C_MAIN, linewidth=1.6)
    ax1.set_ylabel('消力池深度 d (m)', fontproperties=_font(9))
    _frame(ax1)
    if d:
        i = d.index(max(d))
        ax1.plot([x[i]], [d[i]], marker='o', ms=8, mfc='none',
                 mec=_C_LIMIT, mew=1.3)
        ax1.set_ylim(0, max(d) * 1.55)
        ax1.annotate('最大池深 %.2f m（Q = %.0f m³/s）' % (d[i], x[i]),
                     xy=(x[i], d[i]),
                     xytext=(x[i] + (max(x) - min(x)) * 0.14, max(d) * 1.12),
                     fontproperties=_font(8), color=_C_LIMIT,
                     arrowprops=dict(arrowstyle='-', color=_C_LIMIT, lw=0.7))

    # 下：消力池长度与海漫长度
    ax2.plot(x, lsj, marker='s', ms=3.4, color=_C_ALT, linewidth=1.4,
             linestyle='--', label='消力池长度 Lsj')
    ax2.plot(x, lp, marker='^', ms=3.4, color=_C_THIRD, linewidth=1.4,
             linestyle=':', label='海漫长度 Lp')
    ax2.set_xlabel('下泄流量 Q (m³/s)', fontproperties=_font(9))
    ax2.set_ylabel('长度 (m)', fontproperties=_font(9))
    _frame(ax2)
    ax2.legend(prop=_font(8), frameon=False, loc='upper right')
    return _save(fig)


def fig_seepage(sp, p):
    """图：闸基渗流水头损失沿地下轮廓线分布"""
    plt = _setup()
    h_list = sp.get('h_list') or []
    if len(h_list) < 2:
        return None
    total = sum(h_list) or 1.0
    cum, acc = [0.0], 0.0
    for h in h_list:
        acc += h
        cum.append(acc)
    xs = list(range(len(cum)))

    fig, ax = plt.subplots(figsize=(FIG_W * 0.92, FIG_H * 0.86))
    ax.plot(xs, cum, marker='o', ms=3, color=_C_MAIN, linewidth=1.5)
    ax.fill_between(xs, 0, cum, color=_C_MAIN, alpha=0.08)
    ax.set_ylim(0, total * 1.2)
    ax.axhline(total, color=_C_LIMIT, linestyle='--', linewidth=1)
    _label(ax, total, '总水头差 %.2f m' % total)
    ax.set_xlabel('地下轮廓线分段序号', fontproperties=_font(9))
    ax.set_ylabel('累计水头损失 (m)', fontproperties=_font(9))
    ax.set_xticks(xs)
    ax.set_xticklabels([str(i) for i in xs], fontproperties=_font(8))
    _frame(ax)
    return _save(fig)


def fig_stress(st, p):
    """图：闸基底板地基反力分布（梯形）"""
    plt = _setup()
    smax = st.get('sigma_max')
    smin = st.get('sigma_min')
    bearing = st.get('bearing', 300)
    if smax is None or smin is None:
        return None
    n = int(p.get('gateCount', 3) or 3)
    b0 = float(p.get('singleGateWidth', 6) or 6)
    dp = float(p.get('middlePierThickness', 1) or 1)
    dside = float(p.get('sidePierThickness', 1.2) or 1.2)
    B = n * b0 + (n - 1) * dp + 2 * dside

    fig, ax = plt.subplots(figsize=(FIG_W * 0.92, FIG_H * 0.82))
    ax.plot([0, B], [smin, smax], color=_C_MAIN, linewidth=1.7)
    ax.fill_between([0, B], [smin, smax], color=_C_MAIN, alpha=0.10)
    ax.plot([0], [smin], marker='o', ms=4, color=_C_MAIN)
    ax.plot([B], [smax], marker='o', ms=4, color=_C_MAIN)
    ax.set_xlim(-B * 0.05, B * 1.05)
    ax.set_ylim(0, max(smax, bearing) * 1.22)
    ax.axhline(bearing, color=_C_LIMIT, linestyle='--', linewidth=1)
    _label(ax, bearing, '地基允许承载力 %g kPa' % bearing, ha='right',
           x=ax.get_xlim()[1])
    ax.text(0, smin, ' 最小值 %.1f kPa' % smin, fontproperties=_font(8),
            va='top', color='#333333')
    ax.text(B, smax, '最大值 %.1f kPa ' % smax, fontproperties=_font(8),
            va='bottom', ha='right', color='#333333')
    ax.set_xlabel('闸室底板顺水流方向宽度 (m)', fontproperties=_font(9))
    ax.set_ylabel('地基应力 (kPa)', fontproperties=_font(9))
    _frame(ax)
    return _save(fig)


def fig_levels(top, p):
    """图：闸顶高程控制线（挡水 / 泄水 / 地面 / 采用值）

    两个坑：
    1. 采用闸顶高程常常**就等于现状地面高程**（取 max 的结果），直接画会叠成一条；
    2. 挡水闸顶 77.60 与设计洪水位 77.52 只差 8cm，标注必然压在一起。
    所以先合并同值的项，再把标签按最小间距推开，并用细引线连回真实高程。
    """
    plt = _setup()
    items = [
        ('正常蓄水位', top.get('normalWL')),
        ('挡水工况闸顶 H1', top.get('H1')),
        ('设计洪水位', top.get('dsWL')),
        ('泄水工况闸顶 H2', top.get('H2')),
        ('现状地面高程', top.get('ground')),
        ('采用闸顶高程', top.get('top')),
    ]
    items = [(k, float(v)) for k, v in items if v is not None]
    if len(items) < 2:
        return None
    items.sort(key=lambda t: t[1])

    lo, hi = items[0][1], items[-1][1]
    span = max(hi - lo, 0.5)

    # 1) 合并同值项
    merged = []
    for name, v in items:
        if merged and abs(merged[-1][1] - v) <= span * 0.012:
            merged[-1][0].append(name)
        else:
            merged.append([[name], v])

    # 2) 标签防重叠
    gap = span * 0.115
    placed, prev = [], None
    for names, v in merged:
        y = v if prev is None else max(v, prev + gap)
        placed.append(y)
        prev = y

    fig, ax = plt.subplots(figsize=(FIG_W * 0.98, FIG_H * 0.95))
    for i, ((names, v), ly) in enumerate(zip(merged, placed)):
        is_final = any('采用' in n for n in names)
        color = _C_LIMIT if is_final else _C_MAIN
        ax.hlines(v, 0, 0.70, color=color, linewidth=1.8 if is_final else 1.1,
                  linestyle='-' if is_final else '--')
        # 标签被推开时补一条细引线，免得读者对不上高程
        if abs(ly - v) > span * 0.004:
            ax.plot([0.70, 0.725], [v, ly], color='#b8b8b8', linewidth=0.6)
        ax.text(0.732, ly, '%s  %.2f m' % (' / '.join(names), v),
                fontproperties=_font(8.5), va='center',
                color=color if is_final else '#333333')

    ax.set_xlim(0, 1.45)
    ax.set_ylim(lo - span * 0.16, max(hi, placed[-1]) + span * 0.16)
    ax.set_xticks([])
    ax.set_ylabel('高程 (m)', fontproperties=_font(9))
    ax.grid(True, axis='y', color='#d8d8d8', linewidth=0.5)
    ax.set_axisbelow(True)
    for s in ('top', 'right', 'bottom'):
        ax.spines[s].set_visible(False)
    ax.tick_params(labelsize=8, length=3, width=0.6, colors='#333333')
    return _save(fig)


# ============================================================
# 总入口
# ============================================================
# (key, 生成函数, 图题, 归属章节关键词)
SPECS = [
    ('gate_width', fig_gate_width, '闸孔总净宽随上下游水位差变化关系', '闸孔'),
    ('energy', fig_energy, '消力池尺寸与海漫长度随下泄流量变化关系', '消能'),
    ('levels', fig_levels, '闸顶高程控制高程线', '高程'),
    ('seepage', fig_seepage, '闸基渗流水头损失沿地下轮廓线分布', '防渗'),
    ('stress', fig_stress, '闸基底板地基反力分布', '稳定'),
]


def render_all(gw, top, sp, energy, st, params):
    """生成全部图表，返回 ([{key, title, chapter, png}], [告警])。

    matplotlib 缺失或绘图出错时**不抛异常**——论文该生成还是要生成，
    只是少几张图，并在告警里说明。
    """
    out, warn = [], []
    try:
        _setup()
    except Exception as e:
        return [], ['绘图库不可用，已跳过全部曲线图（%s）' % e]

    args = {'gate_width': (gw, params), 'energy': (energy, params),
            'levels': (top, params), 'seepage': (sp, params),
            'stress': (st, params)}
    import warnings
    for key, fn, title, chapter in SPECS:
        try:
            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter('always')
                png = fn(*args[key])
            missing = sorted({str(w.message) for w in caught
                              if 'missing from font' in str(w.message)})
            if missing:
                warn.append('图「%s」有字符缺少字形：%s' % (title, '；'.join(missing[:3])))
        except Exception as e:
            warn.append('图「%s」绘制失败：%s' % (title, e))
            continue
        if png:
            out.append({'key': key, 'title': title, 'chapter': chapter, 'png': png})
    return out, warn

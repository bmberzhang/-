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


def _rect(ax, x, y, w, h, fc='#eef4fa', ec='#0f5a94', lw=0.9, hatch=None):
    from matplotlib.patches import Rectangle
    ax.add_patch(Rectangle((x, y), w, h, facecolor=fc, edgecolor=ec,
                           linewidth=lw, hatch=hatch, zorder=2))


def _plain(ax):
    ax.set_xticks([])
    ax.set_yticks([])
    for s in ('top', 'right', 'bottom', 'left'):
        ax.spines[s].set_visible(False)


def fig_plan(gw, energy, p):
    """图：闸室平面布置示意图（顺水流方向各段按实际长度成比例）"""
    plt = _setup()
    n = int(float(p.get('gateCount') or 3))
    b0 = float(p.get('singleGateWidth') or 6)
    dp = float(p.get('middlePierThickness') or 1.0)
    dsd = float(p.get('sidePierThickness') or 1.2)
    Lb = float(p.get('blanketLength') or 15)
    Lf = float(p.get('floorLength') or 14)
    Lsj = float(energy.get('Lsj_design') or 15) or 15
    Lp = float(energy.get('Lp_design') or 20) or 20
    Lt = 5.0
    B = n * b0 + (n - 1) * dp + 2 * dsd

    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H * 0.86))
    x = 0.0
    segs = [('上游铺盖', Lb, '#eef4fa'), ('闸室', Lf, '#dcebf7'),
            ('消力池', Lsj, '#eaf3ea'), ('海漫', Lp, '#f3f1e8'),
            ('防冲槽', Lt, '#f0e9e9')]
    for name, ln, fc in segs:
        _rect(ax, x, 0, ln, B, fc=fc)
        ax.text(x + ln / 2, B + B * 0.10, name, ha='center',
                fontproperties=_font(8.5), color='#333333')
        x += ln
    total = x

    # 闸孔与闸墩（平面图里闸孔沿垂直水流方向依次排开）
    y = dsd
    for i in range(n):
        _rect(ax, Lb, y, Lf, b0, fc='#ffffff', ec='#0f5a94', lw=0.8)
        ax.text(Lb + Lf / 2, y + b0 / 2, '闸门', ha='center', va='center',
                fontproperties=_font(8), color='#0f5a94')
        y += b0
        if i < n - 1:
            _rect(ax, Lb, y, Lf, dp, fc='#c9dced', ec='#0f5a94', lw=0.8)
            y += dp

    for x0, lbl in ((0, '0'), (Lb, '%g' % Lb),
                    (Lb + Lf, '%g' % (Lb + Lf)),
                    (Lb + Lf + Lsj, '%g' % (Lb + Lf + Lsj)),
                    (total, '%g' % total)):
        ax.plot([x0, x0], [-B * 0.06, 0], color='#8a8f98', linewidth=0.6)
        ax.text(x0, -B * 0.10, lbl, ha='center', va='top',
                fontproperties=_font(7.5), color='#555555')
    ax.annotate('', xy=(0, -B * 0.20), xytext=(total, -B * 0.20),
                arrowprops=dict(arrowstyle='<->', color='#8a8f98', lw=0.7))
    ax.text(total / 2, -B * 0.27, '顺水流方向长度 (m)', ha='center', va='top',
            fontproperties=_font(8), color='#555555')

    ax.set_xlim(-total * 0.03, total * 1.03)
    ax.set_ylim(-B * 0.42, B * 1.30)
    _plain(ax)
    return _save(fig)


def fig_section(st, top, p):
    """图：闸室横剖面示意图（垂直水流方向）"""
    plt = _setup()
    n = int(float(p.get('gateCount') or 3))
    b0 = float(p.get('singleGateWidth') or 6)
    dp = float(p.get('middlePierThickness') or 1.0)
    dsd = float(p.get('sidePierThickness') or 1.2)
    B = n * b0 + (n - 1) * dp + 2 * dsd
    tf = float(p.get('floorThickness') or 1.2)
    z_top = float(top.get('top') or 0)
    sill = float(p.get('gateSillElevation') or 0)
    H = max(z_top - sill, 1.0)

    fig, ax = plt.subplots(figsize=(FIG_W * 0.92, FIG_H * 0.90))
    _rect(ax, 0, -tf, B, tf, fc='#c9dced')            # 底板
    x = dsd
    for i in range(n):                                 # 边墩 + 各孔
        _rect(ax, x, 0, b0, H, fc='#ffffff', ec='#0f5a94', lw=0.7)
        ax.plot([x + b0 * 0.12, x + b0 * 0.12], [0, H * 0.86],
                color='#0f5a94', linewidth=1.0)
        ax.text(x + b0 / 2, H * 0.45, '闸门', ha='center', va='center',
                fontproperties=_font(8), color='#0f5a94')
        x += b0
        if i < n - 1:
            _rect(ax, x, 0, dp, H, fc='#c9dced')
            ax.text(x + dp / 2, H * 0.70, '中墩', ha='center', va='center',
                    rotation=90, fontproperties=_font(7.5), color='#0f5a94')
            x += dp
    _rect(ax, 0, 0, dsd, H, fc='#c9dced')
    ax.text(dsd / 2, H * 0.70, '边墩', ha='center', va='center',
            rotation=90, fontproperties=_font(7.5), color='#0f5a94')
    _rect(ax, B - dsd, 0, dsd, H, fc='#c9dced')

    ax.text(B / 2, -tf * 2.1, '闸底板厚 %.2f m' % tf, ha='center', va='top',
            fontproperties=_font(8), color='#333333')
    ax.annotate('', xy=(B * 0.02, H * 1.06), xytext=(0.02 * B, H * 1.06))
    ax.plot([0, B], [H * 1.06, H * 1.06], color='#8a8f98', linewidth=0.7)
    ax.annotate('', xy=(0, H * 1.10), xytext=(B, H * 1.10),
                arrowprops=dict(arrowstyle='<->', color='#8a8f98', lw=0.7))
    ax.text(B / 2, H * 1.13, '闸室总宽度 %.2f m' % B, ha='center', va='bottom',
            fontproperties=_font(8), color='#555555')
    ax.set_xlim(-B * 0.06, B * 1.06)
    ax.set_ylim(-tf * 3.0, H * 1.35)
    _plain(ax)
    return _save(fig)


def fig_profile(gw, top, energy, p):
    """图：沿水流方向的纵剖面示意图（铺盖—闸室—消力池—海漫—防冲槽）"""
    plt = _setup()
    Lb = float(p.get('blanketLength') or 15)
    Lf = float(p.get('floorLength') or 14)
    Lsj = float(energy.get('Lsj_design') or 15) or 15
    Lp = float(energy.get('Lp_design') or 20) or 20
    Lt = 5.0
    sill = float(p.get('gateSillElevation') or 0)
    z_up = float(p.get('upstreamWaterLevel') or sill + 4)
    z_dn = float(p.get('downstreamWaterLevel') or sill + 3)
    tf = float(p.get('floorThickness') or 1.2)
    d_pool = float(energy.get('d_design') or 0.5)

    fig, ax = plt.subplots(figsize=(FIG_W * 1.02, FIG_H * 0.94))
    _rect(ax, 0, sill - 0.5, Lb, 0.5, fc='#f3f1e8')                 # 铺盖
    _rect(ax, Lb, sill - tf, Lf, tf, fc='#c9dced')                  # 闸室底板
    _rect(ax, Lb + Lf, sill - tf - d_pool, Lsj, tf + d_pool, fc='#eaf3ea')  # 消力池
    _rect(ax, Lb + Lf + Lsj, sill - tf, Lp, 0.35, fc='#f3f1e8')     # 海漫
    _rect(ax, Lb + Lf + Lsj + Lp, sill - tf - 2.85, Lt, 2.85 + 0.35, fc='#f0e9e9')

    total = Lb + Lf + Lsj + Lp + Lt
    y_lo, y_hi = sill - tf - 3.4, z_up + 1.5
    span = y_hi - y_lo
    ax.plot([0, total], [z_up, z_up], color=_C_MAIN, linestyle='--', linewidth=1.0)
    ax.plot([0, total], [z_dn, z_dn], color=_C_ALT, linestyle='--', linewidth=1.0)
    # 上下游水位只差零点几米，在 20 m 的量程里几乎是同一条线：
    # 一个标签写在上方、另一个压到下方，各带一条细引线，免得读成一行字
    ax.text(0.5, z_up + span * 0.02, '上游水位 %.2f m' % z_up, ha='left',
            va='bottom', fontproperties=_font(8), color=_C_MAIN)
    ax.plot([0.5, 0.5], [z_up, z_up + span * 0.018], color=_C_MAIN, linewidth=0.6)
    ax.text(total - 0.5, z_dn - span * 0.055, '下游水位 %.2f m' % z_dn, ha='right',
            va='top', fontproperties=_font(8), color=_C_ALT)
    ax.plot([total - 0.5, total - 0.5], [z_dn, z_dn - span * 0.05],
            color=_C_ALT, linewidth=0.6)
    for xx, lbl in ((Lb * 0.5, '铺盖'), (Lb + Lf * 0.5, '闸室底板'),
                    (Lb + Lf + Lsj * 0.5, '消力池'), (Lb + Lf + Lsj + Lp * 0.5, '海漫'),
                    (Lb + Lf + Lsj + Lp + Lt * 0.5, '防冲槽')):
        ax.text(xx, sill - tf - 1.5, lbl, ha='center', va='top',
                fontproperties=_font(8), color='#333333')
    ax.set_xlim(-1, total + 1)
    ax.set_ylim(y_lo, y_hi)
    ax.set_ylabel('高程 (m)', fontproperties=_font(9))
    ax.set_xlabel('顺水流方向距离 (m)', fontproperties=_font(9))
    ax.grid(True, color='#d8d8d8', linewidth=0.5)
    ax.set_axisbelow(True)
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
    ax.tick_params(labelsize=8, length=3, width=0.6, colors='#333333')
    return _save(fig)


def fig_seep_profile(sp, p):
    """图：闸基防渗布置与分段水头损失示意"""
    plt = _setup()
    Lb = float(p.get('blanketLength') or 15)
    Lf = float(p.get('floorLength') or 14)
    sill = float(p.get('gateSillElevation') or 0)
    tf = float(p.get('floorThickness') or 1.2)
    hs = sp.get('h_list') or []
    if not hs:
        return None
    Te = float(sp.get('Te') or 10)

    # 上下两个分图共用横坐标：上面画地下轮廓线，下面画各段水头损失。
    # 画在同一坐标系里会看成两座建筑物。
    fig, (ax, ax2) = plt.subplots(
        2, 1, sharex=True, figsize=(FIG_W, FIG_H * 1.05),
        gridspec_kw={'height_ratios': [1.5, 1.0], 'hspace': 0.18})
    _rect(ax, 0, sill - 0.6, Lb, 0.6, fc='#f3f1e8')     # 铺盖
    _rect(ax, Lb, sill - tf, Lf, tf, fc='#c9dced')       # 底板
    ax.plot([Lb + Lf, Lb + Lf], [sill - tf, sill - tf - 2.2], color='#7a5c2e',
            linewidth=2.4)                                # 板桩
    ax.plot([Lb, Lb], [sill - tf, sill - tf - 0.8], color='#7a5c2e', linewidth=2.4)
    ax.text(Lb + Lf + 0.4, sill - tf - 1.8, '板桩', fontproperties=_font(8),
            color='#7a5c2e')
    ax.text(Lb * 0.5, sill - 0.15, '铺盖', ha='center', va='bottom',
            fontproperties=_font(8), color='#333333')
    ax.text(Lb + Lf * 0.5, sill - tf * 0.5, '闸底板', ha='center', va='center',
            fontproperties=_font(8), color='#0f5a94')
    ax.set_ylim(sill - 3.6, sill + 0.8)
    ax.set_ylabel('高程 (m)', fontproperties=_font(9))
    ax.grid(True, color='#d8d8d8', linewidth=0.5)
    ax.set_axisbelow(True)
    ax.tick_params(labelsize=8, length=3, width=0.6, colors='#333333')

    x, seg_len = 0.0, [1.0, 12.0, 1.2, 14.0, 0.5, 0.5, 1.0, 2.2]
    for i, h in enumerate(hs):
        w = seg_len[i] if i < len(seg_len) else 1.0
        _rect(ax2, x, 0, w, h, fc='#dcebf7', ec='#0f5a94', lw=0.7)
        x += w
    ax2.plot([0, x], [0, 0], color='#4a4a4a', linewidth=0.7)
    ax2.set_ylim(0, max(hs) * 1.35)
    ax2.set_ylabel('水头损失 (m)', fontproperties=_font(9))
    ax2.set_xlabel('顺水流方向距离 (m)', fontproperties=_font(9))
    ax2.grid(True, color='#d8d8d8', linewidth=0.5)
    ax2.set_axisbelow(True)
    ax2.tick_params(labelsize=8, length=3, width=0.6, colors='#333333')
    for a in (ax, ax2):
        for s in ('top', 'right'):
            a.spines[s].set_visible(False)
    return _save(fig)


def fig_loads(st, p):
    """图：闸室荷载分项（竖向力向下为正，水平力按方向分列）"""
    plt = _setup()
    loads = st.get('loads') or []
    if not loads:
        return None
    names, vals = [], []
    for nm, v, _d in loads:
        names.append(nm.split('（')[0][:8])
        vals.append(abs(v))
    order = sorted(range(len(vals)), key=lambda i: vals[i])
    names = [names[i] for i in order]
    vals = [vals[i] for i in order]

    fig, ax = plt.subplots(figsize=(FIG_W * 0.98, FIG_H * 0.92))
    bars = ax.barh(names, vals, color=_C_MAIN, alpha=0.85, height=0.6)
    for b, v in zip(bars, vals):
        ax.text(v, b.get_y() + b.get_height() / 2, ' %.0f' % v,
                va='center', fontproperties=_font(8), color='#333333')
    ax.set_xlim(0, max(vals) * 1.22)
    ax.set_xlabel('荷载 (kN)', fontproperties=_font(9))
    ax.grid(True, axis='x', color='#d8d8d8', linewidth=0.5)
    ax.set_axisbelow(True)
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
    ax.tick_params(labelsize=8, length=3, width=0.6, colors='#333333')
    return _save(fig)


def fig_hq(energy, p):
    """图：闸下水位流量关系曲线"""
    plt = _setup()
    rows = energy.get('rows') or []
    if len(rows) < 3:
        return None
    sill = float(p.get('gateSillElevation') or 0)
    xs = [r['Q'] for r in rows]
    ys = [r['hs'] + sill for r in rows]

    fig, ax = plt.subplots(figsize=(FIG_W * 0.78, FIG_H * 0.88))
    ax.plot(xs, ys, color=_C_MAIN, linewidth=1.6, marker='o', markersize=3)
    ax.fill_between(xs, min(ys) - 0.2, ys, color=_C_MAIN, alpha=0.10)
    ax.set_xlabel('闸下流量 Q (m³/s)', fontproperties=_font(9))
    ax.set_ylabel('闸下水位 (m)', fontproperties=_font(9))
    _frame(ax)
    return _save(fig)


# ============================================================
# 总入口
# ============================================================
# (key, 生成函数, 图题, 归属章节关键词)
SPECS = [
    ('gate_width', fig_gate_width, '闸孔总净宽随上下游水位差变化关系', '闸孔'),
    ('plan', fig_plan, '闸室平面布置示意图', '闸孔'),
    ('hq', fig_hq, '闸下水位流量关系曲线', '闸孔'),
    ('energy', fig_energy, '消力池尺寸与海漫长度随下泄流量变化关系', '消能'),
    ('levels', fig_levels, '闸顶高程控制高程线', '高程'),
    ('profile', fig_profile, '水闸纵剖面布置示意图', '高程'),
    ('section', fig_section, '闸室横剖面示意图', '高程'),
    ('seepage', fig_seepage, '闸基渗流水头损失沿地下轮廓线分布', '防渗'),
    ('seep_profile', fig_seep_profile, '闸基防渗布置与分段水头损失示意图', '防渗'),
    ('stress', fig_stress, '闸基底板地基反力分布', '稳定'),
    ('loads', fig_loads, '闸室荷载分项图', '稳定'),
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
            'stress': (st, params),
            'plan': (gw, energy, params), 'hq': (energy, params),
            'profile': (gw, top, energy, params),
            'section': (st, top, params),
            'seep_profile': (sp, params), 'loads': (st, params)}
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

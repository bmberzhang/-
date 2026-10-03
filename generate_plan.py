"""
水闸枢纽平面布置图生成器
========================
输入：参数（P 字典）——长度/高程单位为【米】，绘图时 ×1000 转为 mm
输出：DXF 文件（AutoCAD 用）+ SVG 预览（网页用）

设计约定：
- 平面图，X = 水流方向（向右为下游），Y = 横向（向上游左侧为正）
- 闸孔沿 Y 方向依次排列：边墩 / 孔 / 中墩 / 孔 / … / 边墩
- 翼墙为 1/4 圆弧 + 竖直挡墙，上下游各 2 道
- 单位：图纸最终为 mm（CAD 中 1 单位 = 1 mm）
"""

import math

import ezdxf
from ezdxf import units
from ezdxf.enums import TextEntityAlignment

K = 1000.0          # 米 → 毫米
PAD = 2000.0        # SVG 视口外扩（mm）

# 标准图框模块（同目录）
try:
    import frame as _FRAME
except ImportError:                                     # 以文件路径加载时的兜底
    import importlib.util as _ilu
    import os as _os
    _spec = _ilu.spec_from_file_location(
        "frame", _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "frame.py"))
    _FRAME = _ilu.module_from_spec(_spec)
    _spec.loader.exec_module(_FRAME)

NOTES = _FRAME.NOTES_MAIN


# ============================================================
# 1. 参数表（长度单位为米）
# ============================================================
P = {
    # ① 工程与出图
    "proj": "滏阳河XG水闸重建工程",
    "title": "水闸枢纽平面布置图",
    "fn": "水闸平面布置图",
    "cn": 0,        # 中文构件注记
    "cen": 1,       # 中心线
    "frm": 1,       # 左右贯通外框竖线
    # ② 河道与底板
    "riverW": 20.0, "bed": 73.10,
    # ③ 闸室
    "nBay": 3, "bayW": 6.0, "midP": 1.0, "sideP": 1.2,
    "gateL": 14.0, "wb": 6.775, "mh": 1,
    "rd": 0.65, "rt": 0.30, "rs": 0.20,
    "wd": 2.85, "wt": 0.80, "ws": 0.30,
    # ④ 上游铺盖
    "apL": 15.0, "apE": 73.10,
    # ⑤ 消力池
    "stL": 15.5, "stE": 73.10, "tooth": 0.50, "rise": 0.50,
    # ⑥ 海漫
    "hmH": 10.0, "hmD": 1.0, "hmS": "", "hmEw": 20.0,
    # ⑦ 防冲槽
    "fcD": 2.85, "fcX": 5.0, "fcY": 14.0, "fcN": 2.0, "fcE": "",
    # ⑧ 圆弧翼墙
    "wgon": 1, "wgR": 10.0, "wgT": 0.30,
    # ⑨ 滩地与地面带
    "tdon": 1, "tdE": 75.0, "tdW": 8.5, "gdE": 78.8, "gdW": 1.0, "smn": 5,
    # ⑩ 高程与细部标注
    "slope": 2.0, "hmN": 20.0,
    "el": 1, "sm": 1, "hs": 1, "gs": 1,
    "dim": 1,       # 长度尺寸标注（底部顺流链/左侧宽链/右侧带宽）
    # ⑪ 标准图框
    "frmOn": 1,     # 是否绘制标准图框
    "sheet": "A2",  # 图纸幅面（A4/A3/A2/A1，加长用 A3x3 形式）
    "sc": 250,      # 出图比例分母（250 = 1:250）
    "dwgno": "XG-SG-01",       # 图号
    "drafter": "张旭",          # 制图
    "checker": "樊晶晶",        # 审核
}


# ============================================================
# 2. 页面字段表： (key, 名称, 分组, 默认值, 单位)
#    单位含义：m=米 / 比率 / 孔 / 个 / 文本 / 勾选
# ============================================================
FIELDS = [
    # ① 工程与出图
    ("proj", "工程名称", "工程与出图", "滏阳河XG水闸重建工程", "文本"),
    ("title", "图名", "工程与出图", "水闸枢纽平面布置图", "文本"),
    ("fn", "DXF 文件名", "工程与出图", "水闸平面布置图", "文本"),
    ("cn", "中文构件注记", "工程与出图", 0, "勾选"),
    ("cen", "中心线", "工程与出图", 1, "勾选"),
    ("frm", "左右贯通外框竖线", "工程与出图", 1, "勾选"),
    # ② 河道与底板
    ("riverW", "河道（河底）宽", "河道与底板", 20.0, "m"),
    ("bed", "闸底板（河底）高程", "河道与底板", 73.10, "m"),
    # ③ 闸室
    ("nBay", "闸孔数", "闸室", 3, "孔"),
    ("bayW", "单孔净宽", "闸室", 6.0, "m"),
    ("midP", "中墩厚", "闸室", 1.0, "m"),
    ("sideP", "边墩厚", "闸室", 1.2, "m"),
    ("gateL", "闸室顺流长（底板）", "闸室", 14.0, "m"),
    ("wb", "工作桥顺流长（自上游端）", "闸室", 6.775, "m"),
    ("mh", "中墩迎水面半圆头", "闸室", 1, "勾选"),
    ("rd", "检修门槽距上游端", "闸室", 0.65, "m"),
    ("rt", "检修门槽厚（顺流）", "闸室", 0.30, "m"),
    ("rs", "检修支承入槽深", "闸室", 0.20, "m"),
    ("wd", "工作门槽距上游端", "闸室", 2.85, "m"),
    ("wt", "工作门槽厚（顺流）", "闸室", 0.80, "m"),
    ("ws", "工作支承入槽深", "闸室", 0.30, "m"),
    # ④ 上游铺盖
    ("apL", "铺盖长", "上游铺盖", 15.0, "m"),
    ("apE", "铺盖顶高程", "上游铺盖", 73.10, "m"),
    # ⑤ 消力池
    ("stL", "消力池长", "消力池", 15.5, "m"),
    ("stE", "池底板高程", "消力池", 73.10, "m"),
    ("tooth", "末端齿墙（坎）顺流宽", "消力池", 0.50, "m"),
    ("rise", "齿墙顶高出底板", "消力池", 0.50, "m"),
    # ⑥ 海漫
    ("hmH", "水平段长", "海漫", 10.0, "m"),
    ("hmD", "斜坡段落差（首末高差）", "海漫", 1.0, "m"),
    ("hmS", "斜坡段长（顺流投影）", "海漫", "", "m"),
    ("hmEw", "斜坡末端宽（自闸室宽收窄至）", "海漫", 20.0, "m"),
    # ⑦ 防冲槽
    ("fcD", "槽深", "防冲槽", 2.85, "m"),
    ("fcX", "槽底顺流宽", "防冲槽", 5.0, "m"),
    ("fcY", "槽底横向长", "防冲槽", 14.0, "m"),
    ("fcN", "放坡 1:n", "防冲槽", 2.0, "比率"),
    ("fcE", "槽底高程（留空自动）", "防冲槽", "", "m"),
    # ⑧ 圆弧翼墙
    ("wgon", "画翼墙（上/下游各 2 道）", "圆弧翼墙", 1, "勾选"),
    ("wgR", "圆弧半径", "圆弧翼墙", 10.0, "m"),
    ("wgT", "墙厚", "圆弧翼墙", 0.30, "m"),
    # ⑨ 滩地与地面带
    ("tdon", "画滩地（不勾则无滩地，地面带照画）", "滩地与地面带", 1, "勾选"),
    ("tdE", "滩地高程", "滩地与地面带", 75.0, "m"),
    ("tdW", "滩地平台宽", "滩地与地面带", 8.5, "m"),
    ("gdE", "地面高程", "滩地与地面带", 78.8, "m"),
    ("gdW", "地面平台宽", "滩地与地面带", 1.0, "m"),
    ("smn", "示坡线数量（自动均布，0=不画）", "滩地与地面带", 5, "个"),
    # ⑩ 高程与细部标注
    ("slope", "滩地/地面边坡（示坡）1:n", "高程与细部标注", 2.0, "比率"),
    ("hmN", "海漫斜坡底坡（示坡）1:n", "高程与细部标注", 20.0, "比率"),
    ("el", "高程绿框标", "高程与细部标注", 1, "勾选"),
    ("sm", "滩地/地面 1:n 示坡线", "高程与细部标注", 1, "勾选"),
    ("hs", "海漫斜段 1:n 示坡", "高程与细部标注", 1, "勾选"),
    ("gs", "闸墩检修/工作门槽", "高程与细部标注", 1, "勾选"),
    ("dim", "长度尺寸标注（底部顺流链/左侧宽链/右侧带宽）", "高程与细部标注", 1, "勾选"),
]

GROUPS = [
    ("工程与出图", "其他"), ("河道与底板", "结构"), ("闸室", "结构"),
    ("上游铺盖", "结构"), ("消力池", "结构"), ("海漫", "结构"),
    ("防冲槽", "结构"), ("圆弧翼墙", "防护"), ("滩地与地面带", "场地"),
    ("高程与细部标注", "标注"),
]

TEXT_KEYS = {"proj", "title", "fn"}
CHECK_KEYS = {"cn", "cen", "frm", "mh", "wgon", "tdon", "el", "sm", "hs", "gs", "dim"}
OPT_KEYS = {"hmS", "fcE"}          # 可留空 = 按公式自动
LAYERS = {
    "轮廓":   ("轮廓", 7),
    "中心线": ("中心线", 1),
    "高程标注": ("高程标注", 3),
    "示坡线": ("示坡线", 6),
    "尺寸标注": ("尺寸标注", 7),
    # 标准图框（配色沿用用户样图：图幅线绿、图框线蓝、标题栏/文字白）
    "图幅":   ("图幅", 3),
    "图框":   ("图框", 5),
    "标题栏": ("标题栏", 7),
}
LAYER_OF = {"out": "轮廓", "cen": "中心线", "elev": "高程标注", "show": "示坡线", "dim": "尺寸标注",
            "border": "图幅", "frame": "图框", "tblock": "标题栏", "note": "标题栏"}
THICK_LAYERS = {"轮廓", "图框"}


def _num(v, default=""):
    """数值容错：空串/None/非数 → 返回 default（保持原 JS 的 "留空=自动" 语义）"""
    if v is None or v == "":
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


# ============================================================
# 3. 几何主流程
#    返回 (G, N, U, TH)
#      G  图元列表：t=1 直线 / 2 圆弧 / 3 文字 / 4 填充矩形 / 5 填充多边形
#      N  关键节点坐标字典
#      U  示坡线基准长（mm）  TH 基准字高（mm）
# ============================================================
def compute_geo(p):
    G = []
    N = {}
    state = {"U": 1695.0, "TH": 930.0}

    def m(v):
        return float(v) * K

    def L(x1, y1, x2, y2, lay, d=0):
        G.append({"t": 1, "x1": x1, "y1": y1, "x2": x2, "y2": y2, "l": lay, "d": 1 if d else 0})

    def R(x0, y0, x1, y1, lay, fc=None):
        L(x0, y0, x1, y0, lay)
        L(x1, y0, x1, y1, lay)
        L(x1, y1, x0, y1, lay)
        L(x0, y1, x0, y0, lay)
        if fc:
            G.append({"t": 4, "x0": x0, "y0": y0, "x1": x1, "y1": y1, "fc": fc})

    def A(cx, cy, r, a0, a1, lay, d=0):
        G.append({"t": 2, "cx": cx, "cy": cy, "r": r,
                  "a0": a0, "a1": a1, "l": lay, "d": 1 if d else 0})

    def T(s, x, y, h, lay, rot=0):
        G.append({"t": 3, "s": str(s), "x": x, "y": y, "h": h, "l": lay, "ro": rot})

    def EB(s, x, y, h=None):
        """高程绿框标：矩形框 + 居中文字"""
        if h is None:
            h = round(state["TH"] * 0.617)
        w = round((len(str(s)) * 360 + 500) * h / 550)
        b = round(710 * h / 550)
        R(x - w / 2, y - b / 2, x + w / 2, y + b / 2, "elev")
        T(s, x, y, h, "elev")

    # ---- 纵向（顺流）控制点 ----
    riverW = m(p["riverW"]); YR = riverW / 2
    gateW = m(p["nBay"] * p["bayW"] + (p["nBay"] - 1) * p["midP"] + 2 * p["sideP"])
    YS = gateW / 2
    Xap = m(p["apL"]); Xch = Xap + m(p["gateL"]); Xba = Xch + m(p["stL"])
    hmS = _num(p.get("hmS"), None)
    if hmS is None:
        hmS = p["hmD"] * p["hmN"]              # 留空则按 落差 × 坡比 自动
    Xh1 = Xba + m(p["hmH"]); Xhd = Xh1 + m(hmS)
    YR2 = m(p["hmEw"]) / 2
    slopePr = m(p["fcD"]) * p["fcN"]
    Xs1 = Xhd + slopePr; Xs2 = Xs1 + m(p["fcX"]); Xs3 = Xs2 + slopePr
    # ---- 横向（翼墙/滩地）控制点 ----
    wgR = m(p["wgR"]) if p["wgon"] else 0
    wgT = m(p["wgT"]) if p["wgon"] else 0
    upEnd = Xap - wgR
    dnStart = Xch + wgR
    if p["tdon"]:
        Yb1 = YR + (p["tdE"] - p["bed"]) * p["slope"] * K
        Yb2 = Yb1 + m(p["tdW"])
        Yd1 = Yb2 + (p["gdE"] - p["tdE"]) * p["slope"] * K
    else:
        # 无滩地：地面带从河槽边直接放坡上地面高程（地面始终必须画出）
        Yb1 = Yb2 = YR
        Yd1 = YR + (p["gdE"] - p["bed"]) * p["slope"] * K
    Yd2 = Yd1 + m(p["gdW"])

    N.update(YR=YR, YS=YS, Xap=Xap, Xch=Xch, Xba=Xba, Xh1=Xh1, Xhd=Xhd,
             Xs1=Xs1, Xs2=Xs2, Xs3=Xs3, Yb1=Yb1, Yb2=Yb2, Yd1=Yd1, Yd2=Yd2,
             upEnd=upEnd, dnStart=dnStart)

    U = Xs3 / 56.0          # 图幅大 → 示坡线长/字高大，全图比例恒定
    TH = Xs3 / 102.0
    state["U"] = U
    state["TH"] = TH

    if p["cen"]:
        L(-2000, 0, Xs3 + 2000, 0, "cen", 1)

    # ---- 上游铺盖 / 闸室 / 消力池 ----
    R(0, -YR, Xap, YR, "out", "#f3e2bd")            # 铺盖
    R(Xap, -YS, Xch, YS, "out", "#cfd8e6")          # 闸室底板
    L(Xap, -YR, Xap, -YS, "out"); L(Xap, YR, Xap, YS, "out")
    L(Xch, -YR, Xch, -YS, "out"); L(Xch, YR, Xch, YS, "out")

    # ---- 闸墩分界 ----
    pierWs = []
    for b in range(0, p["nBay"] + 1):
        pierWs.append(m(p["sideP"]) if (b == 0 or b == p["nBay"]) else m(p["midP"]))
    edges = []; yy = -YS
    for i, w in enumerate(pierWs):
        yy += w; edges.append(yy)
        if i < len(pierWs) - 1:
            yy += m(p["bayW"]); edges.append(yy)
    N["edges"] = edges; N["pierWs"] = pierWs

    piers = []; acc = -YS
    for ip, w in enumerate(pierWs):
        piers.append([acc, acc + w]); acc += w
        if ip < len(pierWs) - 1:
            acc += m(p["bayW"])
    N["piers"] = piers

    # ---- 门槽位置 ----
    g1x = Xap + m(p["rd"]); g1y = g1x + m(p["rt"])
    g2x = Xap + m(p["wd"]); g2y = g2x + m(p["wt"])
    N.update(g1x=g1x, g1y=g1y, g2x=g2x, g2y=g2y)

    def segs(x_start, x_end):
        """把 [x_start, x_end] 按门槽区间打断，门槽处不画墩分界线"""
        br = [[g1x, g1y], [g2x, g2y]]; out = []; x0 = x_start
        for b0, b1 in br:
            if x0 < b0:
                out.append([x0, b0])
            x0 = max(x0, b1)
        if x0 < x_end:
            out.append([x0, x_end])
        return out

    for d, (lo, hi) in enumerate(piers):
        if 0 < d < len(piers) - 1:
            # 中墩：上下两条边都是闸孔侧分界
            for yc2 in (lo, hi):
                s0 = Xap + pierWs[d] / 2
                for s in segs(s0, Xch):
                    L(s[0], yc2, s[1], yc2, "out")
        else:
            # 边墩：只画靠闸孔那一侧
            surf2 = lo if hi > 0 else hi
            for s in segs(Xap, Xch):
                L(s[0], surf2, s[1], surf2, "out")

    # ---- 工作桥分界 ----
    wbX = Xap + m(p["wb"]); N["wbX"] = wbX
    L(wbX, -YS, wbX, YS, "out")

    # ---- 中墩迎水面半圆头 ----
    if p["mh"]:
        for d2 in range(1, len(piers) - 1):
            rc = pierWs[d2] / 2
            yc = (piers[d2][0] + piers[d2][1]) / 2
            A(Xap + rc, yc, rc, 90, 270, "out")

    # ---- 圆弧翼墙 + 竖直挡墙 ----
    if p["wgon"]:
        cy0 = YR + wgR
        for w in [(Xap, cy0, 180, 270), (Xap, -cy0, 90, 180),
                  (Xch, cy0, 270, 360), (Xch, -cy0, 0, 90)]:
            A(w[0], w[1], wgR, w[2], w[3], "out")
            A(w[0], w[1], wgR - wgT, w[2], w[3], "out")
        up1 = Xap - wgR; up2 = up1 + wgT
        dn1 = Xch + wgR - wgT; dn2 = Xch + wgR
        for sg in (-1, 1):
            yt = (YR + wgR) * sg; yb = Yd1 * sg
            L(up1, yt, up1, yb, "out"); L(up2, yt, up2, yb, "out")
            L(dn1, yt, dn1, yb, "out"); L(dn2, yt, dn2, yb, "out")

    # ---- 门槽 U 形（朝闸孔侧开口） ----
    if p["gs"]:
        faces = []
        for d3 in range(len(piers)):
            if d3 == 0:
                faces.append([piers[d3][1], -1])
            elif d3 == len(piers) - 1:
                faces.append([piers[d3][0], 1])
            else:
                faces.append([piers[d3][1], -1])
                faces.append([piers[d3][0], 1])
        for ys, dirn in faces:
            for x0, Ld, Dp in ((g1x, m(p["rt"]), m(p["rs"])), (g2x, m(p["wt"]), m(p["ws"]))):
                L(x0, ys, x0, ys + dirn * Dp, "out")
                L(x0, ys + dirn * Dp, x0 + Ld, ys + dirn * Dp, "out")
                L(x0 + Ld, ys, x0 + Ld, ys + dirn * Dp, "out")

    # ---- 消力池 / 海漫 / 防冲槽 ----
    R(Xch, -YR, Xba, YR, "out", "#fbe5e5")
    L(Xba, -YR, Xba, YR, "out")
    R(Xba, -YS, Xh1, YS, "out", "#e6efd6")
    L(Xh1, YS, Xhd, YR2, "out"); L(Xh1, -YS, Xhd, -YR2, "out")
    G.append({"t": 5, "pts": [[Xh1, YS], [Xhd, YR2], [Xhd, -YR2], [Xh1, -YS]], "fc": "#e6efd6"})
    R(Xhd, -YR, Xs3, YR, "out", "#ece6f6")
    if p["tooth"] > 0:
        L(Xba - m(p["tooth"]), -YR, Xba - m(p["tooth"]), YR, "out", 1)
    by = m(p["fcY"]) / 2
    G.append({"t": 4, "x0": Xs1, "y0": -by, "x1": Xs2, "y1": by, "fc": "#cfc4e4"})
    L(Xs1, -by, Xs1, by, "out", 1); L(Xs2, -by, Xs2, by, "out", 1)
    L(Xs1, by, Xs2, by, "out", 1); L(Xs1, -by, Xs2, -by, "out", 1)
    L(Xs1, by, Xhd, YR, "out", 1); L(Xs2, by, Xs3, YR, "out", 1)
    L(Xs1, -by, Xhd, -YR, "out", 1); L(Xs2, -by, Xs3, -YR, "out", 1)

    # ---- 滩地 / 地面带（滩地可无，地面带始终绘制）----
    for sg2 in (-1, 1):
        upE = upEnd; dnS = dnStart
        if p["tdon"]:
            L(0, sg2 * Yb1, upE, sg2 * Yb1, "out")
            L(0, sg2 * Yb2, upE, sg2 * Yb2, "out")
            L(Xap, sg2 * Yb1, Xch, sg2 * Yb1, "out")
            L(Xap, sg2 * Yb2, Xch, sg2 * Yb2, "out")
            L(dnS, sg2 * Yb1, Xs3, sg2 * Yb1, "out")
            L(dnS, sg2 * Yb2, Xs3, sg2 * Yb2, "out")
        L(0, sg2 * Yd1, Xs3, sg2 * Yd1, "out")
        L(0, sg2 * Yd2, Xs3, sg2 * Yd2, "out")
        L(0, sg2 * Yd1, 0, sg2 * Yd2, "out")
        L(Xs3, sg2 * Yd1, Xs3, sg2 * Yd2, "out")

    if p["frm"]:
        ym = Yd2
        L(0, ym, 0, -ym, "out"); L(Xs3, ym, Xs3, -ym, "out")

    # ---- 高程标 ----
    if p["el"]:
        hmE = p["stE"] + p["rise"]
        fcE = _num(p.get("fcE"), None)
        if fcE is None:
            fcE = hmE - (hmS / p["hmN"]) - p["fcD"]
        yA = 0.65 * YR; yG = 0.625 * YS
        for yy2 in (yA, 0, -yA):
            EB(f2(p["apE"]), Xap / 2, yy2)
        for yy2 in (yG, 0, -yG):
            EB(f2(p["bed"]), (Xap + Xch) / 2, yy2)
        for yy2 in (yA, 0, -yA):
            EB(f2(p["stE"]), (Xch + Xba) / 2, yy2)
        for yy2 in (yG, 0, -yG):
            EB(f2(hmE), (Xba + Xh1) / 2, yy2)
        EB(f2(fcE), (Xs1 + Xs2) / 2, 0)
        if p["tdon"]:
            EB(f2(p["tdE"]), Xba + m(0.5), (Yb1 + Yb2) / 2)
            EB(f2(p["tdE"]), Xba + m(0.5), -(Yb1 + Yb2) / 2)
        EB(f2(p["gdE"]), Xba + m(0.5), (Yd1 + Yd2) / 2)
        EB(f2(p["gdE"]), Xba + m(0.5), -(Yd1 + Yd2) / 2)

    # ---- 海漫斜坡 1:n 示坡 ----
    if p["hs"]:
        yo0 = min(2500, YS * 0.4)
        st = U * 0.156
        ln = [U, 0.56 * U, U, 0.56 * U, U]
        for yo in (yo0, -yo0):
            for q2 in range(5):
                L(Xh1, yo - q2 * st, Xh1 + ln[q2], yo - q2 * st, "show")
            T("1:" + num2str(p["hmN"]), Xh1 + U * 0.56, yo + TH * 0.4, TH * 0.45, "show")

    # ---- 滩地/地面 1:n 示坡线 ----
    if p["sm"] and p["smn"] > 0:
        slT = max(0.0, (Yb1 - YR) * 0.85)   # 滩地坡长（无滩地时为 0，不画滩地示坡）
        slG = max(0.0, (Yd1 - (Yb2 if p["tdon"] else YR)) * 0.85)
        lnT = [min(v, slT) for v in (U, U / 2, U, U / 2, U)] if p["tdon"] else []
        lnG = [min(v, slG) for v in (2 * U, U, 2 * U, U, 2 * U)]

        def brush(cx, edge, dir_in, lens, wid):
            for i2 in range(5):
                x = cx - wid * 2 + i2 * wid
                L(x, edge, x, edge + dir_in * lens[i2], "show")

        sk = [[upEnd, Xap], [Xch, dnStart]] if p["wgon"] else []
        gap = Xs3 / max(1, p["smn"])
        cx = gap * 0.5
        while cx < Xs3:
            ok = True
            for seg in sk:
                if seg[0] < cx < seg[1]:
                    ok = False
            if ok:
                if lnT:
                    brush(cx, Yb1, -1, lnT, U * 0.22); brush(cx, -Yb1, 1, lnT, U * 0.22)
                    T("1:" + num2str(p["slope"]), cx + U * 0.75, Yb1 - U * 0.5, TH * 0.45, "show")
                    T("1:" + num2str(p["slope"]), cx + U * 0.75, -Yb1 + U * 0.5, TH * 0.45, "show")
                brush(cx, Yd1, -1, lnG, U * 0.22); brush(cx, -Yd1, 1, lnG, U * 0.22)
                T("1:" + num2str(p["slope"]), cx + U * 0.75, Yd1 - U, TH * 0.45, "show")
                T("1:" + num2str(p["slope"]), cx + U * 0.75, -Yd1 + U, TH * 0.45, "show")
            cx += gap

    # ---- 长度尺寸标注：底部顺流尺寸链 + 总长 / 左侧闸室宽度链 + 总宽 / 右侧滩地地面带宽 ----
    if p["dim"]:
        def tk(x, y):
            L(x - U * 0.18, y - U * 0.18, x + U * 0.18, y + U * 0.18, "dim")

        def dV(y1, y2, x):
            L(x, y1, x, y2, "dim")
            tk(x, y1); tk(x, y2)
            T(str(int(round(y2 - y1))), x - U * 0.45, (y1 + y2) / 2, TH * 0.5, "dim", 90)

        yB = -(max(YS, YR2, Yd2) + U * 1.6)
        bx = [0, Xap, Xch, Xba, Xh1, Xhd, Xs3]
        L(0, yB, Xs3, yB, "dim")
        for v in bx:
            tk(v, yB)
        for bi in range(len(bx) - 1):
            T(str(int(round(bx[bi + 1] - bx[bi]))), (bx[bi] + bx[bi + 1]) / 2, yB + U * 0.55, TH * 0.5, "dim")
        L(0, yB - U * 1.6, Xs3, yB - U * 1.6, "dim"); tk(0, yB - U * 1.6); tk(Xs3, yB - U * 1.6)
        T(str(int(round(Xs3))), Xs3 / 2, yB - U * 1.05, TH * 0.5, "dim")
        xL = min(0, upEnd) - U * 1.6
        L(xL, -YS, xL, YS, "dim")
        vb = [-YS]
        for vp in range(len(piers)):
            vb.append(piers[vp][1])
            if vp < len(piers) - 1:
                vb.append(piers[vp][1] + m(p["bayW"]))
        for v in vb:
            tk(xL, v)
        for vi in range(len(vb) - 1):
            T(str(int(round(vb[vi + 1] - vb[vi]))), xL - U * 0.45, (vb[vi] + vb[vi + 1]) / 2, TH * 0.5, "dim", 90)
        L(xL - U * 1.6, -YS, xL - U * 1.6, YS, "dim"); tk(xL - U * 1.6, -YS); tk(xL - U * 1.6, YS)
        T(str(int(round(2 * YS))), xL - U * 2.05, 0, TH * 0.5, "dim", 90)
        xR = Xs3 + U * 1.6
        if p["tdon"]:
            dV(Yb1, Yb2, xR); dV(-Yb2, -Yb1, xR)
            dV(Yd1, Yd2, xR + U * 1.4); dV(-Yd2, -Yd1, xR + U * 1.4)
        else:
            dV(Yd1, Yd2, xR); dV(-Yd2, -Yd1, xR)

    # ---- 中文构件注记 ----
    if p["cn"]:
        yt = Yd2 + TH * 2.7
        T("上游铺盖", m(0.02), yt, TH, "out")
        T("闸室", (Xap + Xch) / 2, yt, TH, "out")
        T("消力池", (Xch + Xba) / 2, yt, TH, "out")
        T("海漫", (Xba + Xh1) / 2, yt, TH, "out")
        T("防冲槽", (Xhd + Xs3) / 2, yt, TH, "out")

    N["totalX"] = max(Xs3, Xhd)

    # ---- 标准图框（图幅边界 + 图框线 + 标题栏 + 说明文字）----
    if p.get("frmOn", 1):
        _add_frame(G, p)

    return G, N, U, TH


def f2(v):
    """高程两位小数"""
    return ("-" if v < 0 else "") + "{:.2f}".format(abs(v))


def num2str(v):
    """坡比数字：整数不带小数点"""
    try:
        fv = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(fv)) if fv == int(fv) else str(fv)


def bbox(G):
    x0 = y0 = 1e18
    x1 = y1 = -1e18
    for e in G:
        if e["t"] == 1:
            x0 = min(x0, e["x1"], e["x2"]); x1 = max(x1, e["x1"], e["x2"])
            y0 = min(y0, e["y1"], e["y2"]); y1 = max(y1, e["y1"], e["y2"])
        elif e["t"] == 2:
            x0 = min(x0, e["cx"] - e["r"]); x1 = max(x1, e["cx"] + e["r"])
            y0 = min(y0, e["cy"] - e["r"]); y1 = max(y1, e["cy"] + e["r"])
        elif e["t"] == 4:
            x0 = min(x0, e["x0"]); x1 = max(x1, e["x1"])
            y0 = min(y0, e["y0"]); y1 = max(y1, e["y1"])
        elif e["t"] == 5:
            for px, py in e["pts"]:
                x0 = min(x0, px); x1 = max(x1, px)
                y0 = min(y0, py); y1 = max(y1, py)
        else:
            x0 = min(x0, e["x"] - e["h"] * 3); x1 = max(x1, e["x"] + e["h"] * 3)
            y0 = min(y0, e["y"] - e["h"]); y1 = max(y1, e["y"] + e["h"])
    return x0, y0, x1, y1


def _add_frame(G, p):
    """把标准图框追加进图元列表 G（与图形同处一个模型坐标系，单位 mm）"""
    sc = _num(p.get("sc"), 250)
    items, fb = _FRAME.build_frame(
        bbox(G),
        sheet=p.get("sheet", "A2"),
        scale=sc,
        info={          # 标题栏内容一律留空，只保留表格线与栏目名（人工填写）
            "proj": "", "title": "", "ratio": "", "no": "", "drafter": "", "checker": "",
        },
        notes=NOTES,
    )
    for it in items:
        if it[0] == "L":
            G.append({"t": 1, "x1": it[1], "y1": it[2], "x2": it[3], "y2": it[4],
                      "l": it[5], "d": 0})
        else:
            G.append({"t": 3, "s": it[1], "x": it[2], "y": it[3], "h": it[4],
                      "l": it[5], "ro": 0, "al": it[6]})
    return fb


# ============================================================
# 4. DXF 输出
# ============================================================
def setup_doc():
    doc = ezdxf.new("R2013", units=units.MM)
    msp = doc.modelspace()
    doc.styles.add("CN", font="simhei.ttf")
    doc.styles.add("数字字体", font="gbenor.shx")
    doc.header["$DWGCODEPAGE"] = "ANSI_936"
    from ezdxf.tools import standards
    standards.setup_linetypes(doc)
    for name, (desc, color) in LAYERS.items():
        doc.layers.add(name=name, color=color)
    for layer in doc.layers:
        lw = 50 if layer.dxf.name in THICK_LAYERS else 18
        try:
            layer.dxf.lineweight = lw
        except Exception:
            pass
    return doc, msp


def generate_dxf(p, out_path):
    G, N, U, TH = compute_geo(p)
    doc, msp = setup_doc()
    for e in G:
        if e["t"] not in (1, 2, 3):
            continue          # t=4 / t=5 为预览着色，不进 DXF
        lay = LAYER_OF.get(e["l"], "轮廓")
        if e["t"] == 1:
            attribs = {"layer": lay}
            if e["d"]:
                attribs["linetype"] = "DASHED"
            elif e["l"] == "cen":
                attribs["linetype"] = "CENTER"
            msp.add_line((e["x1"], e["y1"]), (e["x2"], e["y2"]), dxfattribs=attribs)
        elif e["t"] == 2:
            attribs = {"layer": lay}
            if e["d"]:
                attribs["linetype"] = "DASHED"
            msp.add_arc((e["cx"], e["cy"]), e["r"], e["a0"], e["a1"], dxfattribs=attribs)
        elif e["t"] == 3:
            style = "数字字体" if e["l"] == "dim" else "CN"
            t = msp.add_text(e["s"], dxfattribs={
                "layer": lay, "height": e["h"], "style": style,
            })
            t.set_placement((e["x"], e["y"]),
                            align=TextEntityAlignment.MIDDLE_LEFT if e.get("al") == "l"
                            else TextEntityAlignment.MIDDLE_CENTER)
            if e.get("ro"):
                t.dxf.rotation = e["ro"]
    doc.saveas(out_path)


# ============================================================
# 5. SVG 预览输出
# ============================================================
SVG_PX_PER_MM = 0.022
PREVIEW_COLOR = {"cen": "#cc0000", "elev": "#1e8a3a", "show": "#c71585", "dim": "#555555",
                 "border": "#1e7a3c", "frame": "#1f4fa8", "tblock": "#333333", "note": "#333333"}
PREVIEW_SW = {"cen": 1.1, "show": 1.6, "border": 1.0, "frame": 2.0}


def generate_svg(p, out_path):
    G, N, U, TH = compute_geo(p)
    x0, y0, x1, y1 = bbox(G)
    x0 -= PAD; y0 -= PAD; x1 += PAD; y1 += PAD
    scale = SVG_PX_PER_MM
    w_px = int((x1 - x0) * scale) + 40
    h_px = int((y1 - y0) * scale) + 40

    def MX(x):
        return 20 + (x - x0) * scale

    def MY(y):
        return 20 + (y1 - y) * scale

    out = ['<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d" '
           'style="background:#fafafa;font-family:SimHei,Microsoft YaHei,sans-serif;">'
           % (w_px, h_px, w_px, h_px)]

    # 先铺底色块，再压线，避免填充盖住线条
    for e in G:
        if e["t"] == 4:
            out.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s" opacity="0.5"/>'
                       % (MX(e["x0"]), MY(e["y1"]), (e["x1"] - e["x0"]) * scale,
                          (e["y1"] - e["y0"]) * scale, e["fc"]))
        elif e["t"] == 5:
            pts = " ".join("%.1f,%.1f" % (MX(px), MY(py)) for px, py in e["pts"])
            out.append('<polygon points="%s" fill="%s" opacity="0.5"/>' % (pts, e["fc"]))

    for e in G:
        if e["t"] == 1:
            col = PREVIEW_COLOR.get(e["l"], "#2b2b2b")
            sw = PREVIEW_SW.get(e["l"], 1.4)
            dash = ' stroke-dasharray="10,7"' if e["d"] or e["l"] == "cen" else ""
            out.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="%.1f"%s/>'
                       % (MX(e["x1"]), MY(e["y1"]), MX(e["x2"]), MY(e["y2"]), col, sw, dash))
        elif e["t"] == 2:
            col = PREVIEW_COLOR.get(e["l"], "#2b2b2b")
            a0, a1 = e["a0"], e["a1"]
            if a1 < a0:
                a1 += 360
            n = max(2, int(math.ceil((a1 - a0) / 12.0)))
            pts = []
            for i in range(n + 1):
                a = math.radians(a0 + (a1 - a0) * i / n)
                pts.append("%.1f,%.1f" % (MX(e["cx"] + e["r"] * math.cos(a)),
                                          MY(e["cy"] + e["r"] * math.sin(a))))
            out.append('<polyline points="%s" fill="none" stroke="%s" stroke-width="1.4"/>'
                       % (" ".join(pts), col))
        elif e["t"] == 3:
            col = PREVIEW_COLOR.get(e["l"], "#2b2b2b")
            font = "Arial, sans-serif" if e["l"] == "dim" else "SimHei, Microsoft YaHei, sans-serif"
            rot = ' transform="rotate(-%.1f %.1f %.1f)"' % (e["ro"], MX(e["x"]), MY(e["y"])) if e.get("ro") else ""
            anchor = "start" if e.get("al") == "l" else "middle"
            out.append('<text x="%.1f" y="%.1f" font-size="%.1f" font-family="%s" fill="%s" text-anchor="%s" '
                       'dominant-baseline="middle"%s>%s</text>'
                       % (MX(e["x"]), MY(e["y"]), max(7.0, e["h"] * scale), font, col, anchor, rot,
                          e["s"].replace("&", "&amp;").replace("<", "&lt;")))

    out.append('</svg>')
    with open(out_path, "w", encoding="utf-8") as fp:
        fp.write("".join(out))
    return N

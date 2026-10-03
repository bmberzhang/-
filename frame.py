# -*- coding: utf-8 -*-
"""
标准图框模块（图幅边界 + 图框线 + 标题栏 + 说明文字）
=====================================================
供 平面布置图(generate_plan.py) / 纵剖面图(generate_drawing.py) 共用。

设计要点
--------
1. 图纸幅面（A3/A2/A1… 及加长幅面）按【出图比例】放大到模型空间，与图纸内容同处一个坐标系，
   这样 DXF 里打印/量取都是真实图纸尺寸。
2. 图纸内容在内框中的定位：**先为左下角说明区和右下角标题栏预留高度，再在剩余空间里居中**，
   避免说明文字压到图形。
3. 图元以轻量元组返回，调用方自行转成 DXF 实体或 SVG：

       ("L", x1, y1, x2, y2, kind)          直线
       ("T", text, x, y, h, kind, align)    文字，align: "c" 居中 / "l" 左对齐（垂直居中）

   kind ∈ {"border", "frame", "tblock", "note"}
       border : 图幅边界线（细）
       frame  : 图框线（粗）
       tblock : 标题栏格线
       note   : 说明文字 / 标题栏文字
"""

# ============================================================
# 幅面
# ============================================================
# 基本幅面（短边, 长边），mm —— GB/T 14689
BASE = {
    "A4": (210, 297),
    "A3": (297, 420),
    "A2": (420, 594),
    "A1": (594, 841),
    "A0": (841, 1189),
}

MARGIN_BIND = 25.0      # 装订边（左侧），mm
MARGIN = 5.0            # 其余三边，mm

# 图纸说明（图框左下角，两图共用）
NOTES_MAIN = [
    "1、图中高程为1985国家高程；高程以m计，其余尺寸均以mm计；",
    "2、排水管采用A110PVC排水管，间距为2m，呈梅花状布置；三层反滤自下而上依次为粗砂垫层厚0.2m；"
    "碎石垫层厚0.2m；碎石头垫层厚0.3m；排水管外包两层反滤土工布，土工布规格为400g/㎡；",
    "3、上游铺盖、混凝土护坡、闸室、消力池底板、下游护底、左右岸两侧挡墙混凝土强度等级为C30F150W6；"
    "工作桥、检修板强度等级为C30F150，交通桥及铺装层混凝土强度等级为C40，垫层混凝土强度等级为C15；",
    "4、回填土压实度不小于0.99；",
    "5、结构缝宽20mm，缝内填充闭孔泡沫塑料板；橡胶止水采用651（300×8）型橡胶止水带，"
    "布置位置距离迎水面200mm；",
    "6、栏杆采用不锈钢栏杆。",
    "7、水闸末端与下游渠道自然顺连。",
]

# 标题栏外形尺寸（图纸 mm）
TB_W, TB_H = 180.0, 56.0
# 标题栏格线位置（相对标题栏左下角），取自样图比例
TB_C = (22.86, 70.02, 103.32, 138.42, 163.80)   # 竖向分割线
TB_R = (14.56, 29.12, 42.70)                    # 横向分割线（自底向上）


def sheet_size(name):
    """解析幅面 → (宽, 高)，横向放置（长边水平）。
    支持 "A3"、"A2"、"A1"，以及加长幅面 "A3x3"（A3 加长，长边 = 短边×3 = 891）。"""
    s = str(name or "A3").upper().replace("×", "X").replace("*", "X").replace("-", "X")
    n = 1
    if "X" in s:
        base, _, ns = s.partition("X")
        try:
            n = int(ns)
        except ValueError:
            n = 1
    else:
        base = s
    if base not in BASE:
        base = "A3"
    short, long_ = BASE[base]
    if n <= 1:
        return float(long_), float(short)
    # 加长幅面：长边 = 短边 × n（GB/T 14689），另一方向仍为基本幅面的长边
    return float(short * n), float(long_)


# ============================================================
# 说明文字折行
# ============================================================
def _vis_width(s):
    """按字高为 1 的视觉宽度估算（中文 1.0，西文 0.55）"""
    return sum(1.0 if ord(c) > 127 else 0.55 for c in s)


def wrap_text(s, max_w):
    """把一行文字按最大视觉宽度 max_w 折行，优先在标点处断开"""
    if _vis_width(s) <= max_w:
        return [s]
    out, buf, w = [], "", 0.0
    for ch in s:
        cw = 1.0 if ord(ch) > 127 else 0.55
        if w + cw > max_w and buf:
            cut = len(buf)
            for k in range(len(buf), max(1, len(buf) - 12), -1):
                if buf[k - 1] in "；，、。；":
                    cut = k
                    break
            out.append(buf[:cut])
            buf = buf[cut:]
            w = _vis_width(buf)
        buf += ch
        w += cw
    if buf:
        out.append(buf)
    return out or [""]


# ============================================================
# 主函数
# ============================================================
def build_frame(bbox, sheet="A3", scale=100, info=None, notes=None,
                note_h=3.0, note_gap=2.0):
    """生成标准图框的全部图元。

    参数
    ----
    bbox   : (x0, y0, x1, y1) 图纸内容的包围盒（模型坐标，mm）
    sheet  : 幅面名，如 "A3" / "A2" / "A3x3"
    scale  : 出图比例分母（100 表示 1:100）
    info   : dict —— proj 工程名 / title 图名 / ratio 比例 / no 图号 /
                     drafter 制图 / checker 审核
    notes  : 说明文字列表（每条一行，过长自动折行）

    返回
    ----
    (items, frame_bbox)  frame_bbox 为该图框在模型坐标下的外边界
    """
    info = dict(info or {})
    notes = [str(x) for x in (notes or []) if str(x).strip()]

    W, H = sheet_size(sheet)
    S = float(scale)
    x0, y0, x1, y1 = [float(v) for v in bbox]

    ix0, iy0 = MARGIN_BIND, MARGIN                  # 内框左下（图纸坐标）
    ix1, iy1 = W - MARGIN, H - MARGIN               # 内框右上

    # ---- 说明文字折行 ----
    usable_w = (ix1 - ix0) - 8.0 - (TB_W + 10.0)    # 横向避开标题栏
    max_chars = max(16.0, usable_w / note_h)
    lines = []
    if notes:
        lines.append("说明：")
        for s in notes:
            lines.extend(wrap_text(s, max_chars))
    line_h = note_h + note_gap
    note_block = (len(lines) * line_h + 8.0) if lines else 0.0

    # ---- 底部预留高度（说明区 / 标题栏 取大者）----
    reserve = max(note_block + 8.0, TB_H + 12.0)

    # ---- 内容定位：在扣除预留后的内框里居中 ----
    cw = (x1 - x0) / S          # 内容在图面上的宽
    ch = (y1 - y0) / S
    avail_w = (ix1 - ix0)
    avail_h = (iy1 - iy0) - reserve
    ux = ix0 + max(0.0, (avail_w - cw) / 2.0)
    uy = iy0 + reserve + max(0.0, (avail_h - ch) / 2.0)

    X0 = x0 - ux * S            # 图纸坐标 → 模型坐标
    Y0 = y0 - uy * S

    def MX(u):
        return X0 + u * S

    def MY(v):
        return Y0 + v * S

    out = []

    def L(u1, v1, u2, v2, kind):
        out.append(("L", MX(u1), MY(v1), MX(u2), MY(v2), kind))

    def T(s, u, v, h, kind, align="c"):
        s = "" if s is None else str(s)
        if not s.strip():          # 空文字不产生实体（标题栏留空时用）
            return
        out.append(("T", s, MX(u), MY(v), h * S, kind, align))

    def RECT(u1, v1, u2, v2, kind):
        L(u1, v1, u2, v1, kind)
        L(u2, v1, u2, v2, kind)
        L(u2, v2, u1, v2, kind)
        L(u1, v2, u1, v1, kind)

    # ---- 图幅边界（外框）+ 图框线（内框）----
    RECT(0.0, 0.0, W, H, "border")
    RECT(ix0, iy0, ix1, iy1, "frame")

    # ---- 标题栏（贴内框右下角）----
    tx, ty = ix1 - TB_W, iy0
    c, r = TB_C, TB_R
    RECT(tx, ty, ix1, ty + TB_H, "tblock")
    for cx_ in c[:3]:                                    # 下半部分竖线
        L(tx + cx_, ty, tx + cx_, ty + r[1], "tblock")
    for cx_ in c[3:]:                                    # 上半部分竖线
        L(tx + cx_, ty + r[1], tx + cx_, ty + TB_H, "tblock")
    L(tx, ty + r[0], tx + c[2], ty + r[0], "tblock")     # 制图 / 审核 分界
    L(tx, ty + r[1], tx + TB_W, ty + r[1], "tblock")     # 上 / 下半分界
    L(tx + c[3], ty + r[2], tx + TB_W, ty + r[2], "tblock")   # 比例 / 图号 分界

    # 标签
    th_lab, th_val = 4.0, 4.0
    T("制 图", tx + c[0] / 2, ty + (r[0] + r[1]) / 2, th_lab, "note")
    T("审 核", tx + c[0] / 2, ty + r[0] / 2, th_lab, "note")
    T("比 例", tx + (c[3] + c[4]) / 2, ty + (r[2] + TB_H) / 2, th_lab, "note")
    T("图 号", tx + (c[3] + c[4]) / 2, ty + (r[1] + r[2]) / 2, th_lab, "note")

    # 值
    vx = tx + (c[4] + TB_W) / 2
    T(info.get("ratio", ""), vx, ty + (r[2] + TB_H) / 2, th_val, "note")
    T(info.get("no", ""), vx, ty + (r[1] + r[2]) / 2, th_val, "note")
    T(info.get("drafter", ""), tx + (c[0] + c[1]) / 2, ty + (r[0] + r[1]) / 2, th_val, "note")
    T(info.get("checker", ""), tx + (c[0] + c[1]) / 2, ty + r[0] / 2, th_val, "note")

    # 图名格（左上大格）：工程名 + 图名
    box_cx = tx + c[3] / 2
    box_v0, box_v1 = ty + r[1], ty + TB_H
    bh = box_v1 - box_v0
    proj = info.get("proj", "")
    title = info.get("title", "")
    th_name = 6.0
    if _vis_width(title) * th_name > c[3] - 8.0:         # 图名过长自动缩小
        th_name = max(3.5, (c[3] - 8.0) / max(1.0, _vis_width(title)))
    if proj:
        T(proj, box_cx, box_v0 + bh * 0.26, 3.6, "note")
        T(title, box_cx, box_v0 + bh * 0.68, th_name, "note")
    else:
        T(title, box_cx, (box_v0 + box_v1) / 2, th_name, "note")

    # ---- 说明文字（左下角，自下而上排）----
    if lines:
        n = len(lines)
        nx = ix0 + 8.0
        for i, s in enumerate(lines):
            v = iy0 + 8.0 + (n - 1 - i + 0.5) * line_h
            T(s, nx, v, note_h, "note", "l")

    return out, (MX(0.0), MY(0.0), MX(W), MY(H))


def merge_bbox(items, bbox):
    """把图元并入包围盒（用于 SVG 视口扩展）"""
    x0, y0, x1, y1 = bbox
    for it in items:
        if it[0] == "L":
            xs, ys = (it[1], it[3]), (it[2], it[4])
        else:
            xs, ys = (it[2],), (it[3],)
        x0 = min(x0, min(xs)); x1 = max(x1, max(xs))
        y0 = min(y0, min(ys)); y1 = max(y1, max(ys))
    return x0, y0, x1, y1

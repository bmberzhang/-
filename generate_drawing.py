"""
水闸纵剖面图生成器
==================
输入：参数（P 字典）
输出：DXF 文件（AutoCAD 用）+ SVG 预览（网页用）

设计约定：
- 纵剖面图，X = 沿程位置（向右为下游），Y = 高程（向上为正，AutoCAD 标准）
- 铺盖、底板、消力池底板【顶面相平】（同一 Y），仅厚度不同
- 齿墙【垂直朝下】矩形墙，深度方向是 -Y
- 单位：米
"""

import os
import math
import ezdxf
from ezdxf import units
from ezdxf.enums import TextEntityAlignment


ALIGN_MAP = {
    "LEFT": TextEntityAlignment.LEFT,
    "CENTER": TextEntityAlignment.CENTER,
    "RIGHT": TextEntityAlignment.RIGHT,
    "MIDDLE": TextEntityAlignment.MIDDLE,
    "TOPLEFT": TextEntityAlignment.TOP_LEFT,
    "TOPCENTER": TextEntityAlignment.TOP_CENTER,
    "TOPRIGHT": TextEntityAlignment.TOP_RIGHT,
    "MIDDLELEFT": TextEntityAlignment.MIDDLE_LEFT,
    "MIDDLECENTER": TextEntityAlignment.MIDDLE_CENTER,
    "MIDDLERIGHT": TextEntityAlignment.MIDDLE_RIGHT,
    "BOTTOMLEFT": TextEntityAlignment.BOTTOM_LEFT,
    "BOTTOMCENTER": TextEntityAlignment.BOTTOM_CENTER,
    "BOTTOMRIGHT": TextEntityAlignment.BOTTOM_RIGHT,
}


# ============================================================
# 1. 参数表（毫米 mm）
#    注：长度/厚度/高程全部以 mm 为单位；仅有"坡率/比例"保留为比率
# ============================================================
P = {
    # 铺盖
    "pg_len": 15000.0, "pg_h": 500,
    "pg_cd": 500, "pg_cw": 500,
    # 底板
    "db_len": 14000.0, "db_h": 1200,
    "db_cd": 1000, "db_cw": 1000,
    # 消力池
    "xl_len": 15000.0, "xl_h": 600,
    "xl_cd": 500, "xl_cw": 500,
    # 反滤层
    "fl_gravel": 200, "fl_stone": 300, "fl_sand": 200,
    # 海漫
    "hm_total": 20000.0, "hm_horiz": 10000.0,
    "hm_stone": 500, "hm_cushion": 100,
    "hm_slope": 0.1,  # 1:10（比率，保留）
    "hm_cd": 500, "hm_cw": 500,
    # 防冲槽
    "fcc_d": 2850, "fcc_bw": 5000,
    "fcc_rip": 400, "fcc_m": 2.0,  # 坡率，保留
    # 闸门（检修闸门 + 工作闸门）
    "gate1_x": 650, "gate1_w": 300,   # 检修闸门：左端距底板左端 650mm，宽 300mm
    "gate2_gap": 1900, "gate2_w": 800,  # 工作闸门：左端距检修闸门右端 1900mm，宽 800mm
    # 排架（中心线 = 闸门中心线）
    "pa_w": 400,            # 排架立柱宽 400mm
    "pa_beam_w": 4600,      # 排架横梁总长度 4600mm（对齐工作桥总宽）
    "pa_beam_h": 300,       # 排架横梁厚度 300mm
    # 工作桥（中心线 = 闸门中心线）
    "br_w": 4600,           # 工作桥总宽 4600mm
    "br_col_w": 400,        # 工作桥底端（支撑腿）宽 400mm
    "br_deck_h": 200,       # 工作桥桥面板厚度 200mm
    # 交通桥（排架右端与底板右端之间的中心处）
    "tb_w": 4600,           # 交通桥宽度（沿水流方向）
    "tb_deck_top": 78900.0, # 交通桥桥面顶高程 78.900m = 78900mm
    "tb_cushion": 140,      # 垫层+橡胶支座总高 140mm
    "tb_slab": 500,         # 空心板板厚 500mm
    "tb_overlay": 100,      # 板上混凝土整体厚 100mm
    # 启闭机房（工作桥上方）
    "house_w": 4600,        # 房屋宽度
    "house_wall_h": 2000,   # 墙体高度 2000mm
    "house_roof_h": 1500,   # 人字顶高度 1500mm
    "house_door_w": 900,    # 门宽 900mm
    "house_door_h": 1600,   # 门高 1600mm
    # 高程（绝对高程 mm，原 m × 1000）
    "el_pg": 73100,        # 铺盖顶 = 底板顶（73.100m）
    "el_bank": 75000,      # 滩地 75.000m
    "el_wl": 76600,        # 正常蓄水位 76.600m
    "el_gate_top": 78800,  # 闸顶 78.800m
    "el_trestle": 84800,   # 排架 84.800m
    "el_bridge": 85500,    # 工作桥 85.500m
    # 绘图选项
    "title": "水闸纵剖面图",
    "scale_text": "1:100",
    "unit_text": "mm",
}


# ============================================================
# 2. 几何布局（Y 向上为正，AutoCAD 标准）
# ============================================================
def compute_layout(p):
    L = {}

    # ---- X 边界 ----
    L["x_pg1"], L["x_pg2"] = 0, p["pg_len"]
    L["x_db1"], L["x_db2"] = L["x_pg2"], L["x_pg2"] + p["db_len"]
    L["x_xl1"], L["x_xl2"] = L["x_db2"], L["x_db2"] + p["xl_len"]
    L["x_xl_c2"] = L["x_xl2"] + p["xl_cw"]
    L["x_hm1"] = L["x_xl_c2"]
    L["x_hm_h2"] = L["x_hm1"] + p["hm_horiz"]
    L["x_hm2"] = L["x_hm1"] + p["hm_total"]
    L["x_hm_c2"] = L["x_hm2"] + p["hm_cw"]
    fcc_side = p["fcc_d"] * p["fcc_m"]
    L["x_fcc1"] = L["x_hm2"]   # 防冲槽与海漫末端衔接，不留缝
    L["x_fcc2"] = L["x_fcc1"] + p["fcc_bw"] + 2 * fcc_side
    L["fcc_side"] = fcc_side
    L["x_fcc_b1"] = L["x_fcc1"] + fcc_side
    L["x_fcc_b2"] = L["x_fcc_b1"] + p["fcc_bw"]

    # ---- Y 边界（向上为正，AutoCAD 标准）----
    # 顶面参考：铺盖/底板/消力池顶 = 3850mm
    L["y_pg_t"] = 3850
    L["y_pg_b"] = L["y_pg_t"] - p["pg_h"]
    L["y_db_t"] = L["y_pg_t"]
    L["y_db_b"] = L["y_db_t"] - p["db_h"]
    L["y_xl_t"] = L["y_pg_t"]
    L["y_xl_b"] = L["y_xl_t"] - p["xl_h"]

    # 海漫：水平段顶与凸起齿墙顶持平（比铺盖顶高 500mm）
    HM_RISE = 500
    drop = p["hm_slope"] * (p["hm_total"] - p["hm_horiz"])
    L["y_hm_t1"] = L["y_pg_t"] + HM_RISE            # 4350
    L["y_hm_cushion_t1"] = L["y_hm_t1"] - p["hm_stone"]   # 3850
    L["y_hm_b1"] = L["y_hm_t1"] - (p["hm_stone"] + p["hm_cushion"])  # 3750
    L["y_hm_t2"] = L["y_hm_t1"] - drop              # 3350
    L["y_hm_cushion_t2"] = L["y_hm_t2"] - p["hm_stone"]   # 2850
    L["y_hm_b2"] = L["y_hm_t2"] - (p["hm_stone"] + p["hm_cushion"])  # 2750

    # 防冲槽：顶与海漫末端顶齐平（连接不留缝）；底 = 顶 - 深
    L["y_fcc_t"] = L["y_hm_t2"]                     # 3350
    L["y_fcc_b"] = L["y_fcc_t"] - p["fcc_d"]        # 500
    L["y_rip_t"] = L["y_fcc_b"] + p["fcc_rip"]      # 900

    # 反滤层（消力池底板下方，从上到下：砾石→碎石→粗砂）
    L["y_fl_gravel_t"] = L["y_xl_b"]
    L["y_fl_gravel_b"] = L["y_fl_gravel_t"] - p["fl_gravel"]
    L["y_fl_stone_t"] = L["y_fl_gravel_b"]
    L["y_fl_stone_b"] = L["y_fl_stone_t"] - p["fl_stone"]
    L["y_fl_sand_t"] = L["y_fl_stone_b"]
    L["y_fl_sand_b"] = L["y_fl_sand_t"] - p["fl_sand"]

    # ---- 全部 Y 转为绝对高程（偏移 = 铺盖顶高程 - 原相对值 3850）----
    off = p["el_pg"] - 3850
    for k in list(L.keys()):
        if k.startswith("y"):
            L[k] += off

    # 高程特征线（绝对高程）
    L["y_el_pg"] = p["el_pg"]          # 铺盖顶/底板顶 73.1
    L["y_el_bank"] = p["el_bank"]      # 滩地 75.0
    L["y_el_wl"] = p["el_wl"]          # 正常蓄水位 76.6
    L["y_el_gate_top"] = p["el_gate_top"]  # 闸顶 78.8
    L["y_el_trestle"] = p["el_trestle"]    # 排架 84.8
    L["y_el_bridge"] = p["el_bridge"]      # 工作桥 85.5

    # 闸门 X（在底板上）
    L["x_g1a"] = L["x_db1"] + p["gate1_x"]
    L["x_g1b"] = L["x_g1a"] + p["gate1_w"]
    L["x_g2a"] = L["x_g1b"] + p["gate2_gap"]
    L["x_g2b"] = L["x_g2a"] + p["gate2_w"]
    # 交通桥 Y（垫层+支座 0.14 → 空心板 0.5 → 板顶混凝土 0.1 → 桥面顶 78.90）
    L["y_tb_top"] = p["tb_deck_top"]                                  # 78.90
    L["y_tb_overlay_btm"] = L["y_tb_top"] - p["tb_overlay"]           # 78.80
    L["y_tb_slab_btm"] = L["y_tb_overlay_btm"] - p["tb_slab"]         # 78.30
    L["y_tb_btm"] = L["y_tb_slab_btm"] - p["tb_cushion"]              # 78.16
    # 闸门中心线（排架/工作桥共用轴线）
    x_gc = (L["x_g2a"] + L["x_g2b"]) / 2
    L["x_gc"] = x_gc
    # 排架：两根立柱（宽 pa_w）对称于中心线，间距 = 横梁宽 pa_beam_w
    half = p["pa_beam_w"] / 2
    L["x_pa1a"] = x_gc - half
    L["x_pa1b"] = L["x_pa1a"] + p["pa_w"]
    L["x_pa2a"] = x_gc + half - p["pa_w"]
    L["x_pa2b"] = x_gc + half
    L["x_pab1"] = x_gc - half          # 排架横梁
    L["x_pab2"] = x_gc + half
    # 交通桥 X（排架右端与底板右端之间的中心处）
    L["x_tb_c"] = (L["x_pa2b"] + L["x_db2"]) / 2
    L["x_tb1"] = L["x_tb_c"] - p["tb_w"] / 2
    L["x_tb2"] = L["x_tb_c"] + p["tb_w"] / 2
    # 工作桥：底端腿（宽 br_col_w）、桥面（总宽 br_w），均以中心线对称
    L["x_brc1"] = x_gc - p["br_col_w"] / 2
    L["x_brc2"] = x_gc + p["br_col_w"] / 2
    L["x_br1"] = x_gc - p["br_w"] / 2
    L["x_br2"] = x_gc + p["br_w"] / 2
    # 启闭机房（工作桥正上方，中心对齐闸门中心线）
    L["x_house1"] = x_gc - p["house_w"] / 2
    L["x_house2"] = x_gc + p["house_w"] / 2
    L["y_house_btm"] = L["y_el_bridge"]                        # 85.5
    L["y_house_wall_top"] = L["y_house_btm"] + p["house_wall_h"]    # 87.5
    L["y_house_roof_top"] = L["y_house_wall_top"] + p["house_roof_h"]  # 89.0

    return L


# ============================================================
# 3. DXF 生成
# ============================================================
# 颜色约定：结构线=红色(1)+粗线(0.5mm)；垫层/反滤层=绿色(3)+细线(0.18mm)
STRUC_COLOR = 1      # 红
PAD_COLOR = 3        # 绿
STRUC_LW = 50        # 0.5 mm（1/100 mm 单位）
PAD_LW = 18          # 0.18 mm
TEXT_COLOR = 7       # 黑（文字）
RC_REBAR_COLOR = 130 # 蓝（钢筋混凝土-钢筋）
RC_CONC_COLOR = 131  # 天蓝（钢筋混凝土-混凝土）

LAYERS = {
    "铺盖":   ("铺盖", STRUC_COLOR),
    "底板":   ("底板", STRUC_COLOR),
    "消力池底板":   ("消力池底板", STRUC_COLOR),
    "反滤层-砾石": ("反滤层-砾石", PAD_COLOR),
    "反滤层-碎石":  ("反滤层-碎石", PAD_COLOR),
    "反滤层-粗砂":   ("反滤层-粗砂", PAD_COLOR),
    "海漫-浆砌石":  ("海漫-浆砌石", STRUC_COLOR),
    "海漫-干砌石":    ("海漫-干砌石", STRUC_COLOR),
    "海漫-粗砂垫层":   ("海漫-粗砂垫层", PAD_COLOR),
    "防冲槽":  ("防冲槽", STRUC_COLOR),
    "防冲槽-堆石":   ("防冲槽-堆石", STRUC_COLOR),
    "齿墙":    ("齿墙", STRUC_COLOR),
    "闸门":      ("闸门", STRUC_COLOR),
    "轮廓":   ("轮廓", STRUC_COLOR),
    "尺寸标注":       ("尺寸标注", STRUC_COLOR),
    "文字":      ("文字", TEXT_COLOR),
    "标题":     ("标题", TEXT_COLOR),
    # 钢筋混凝土填充双图层（ANS31 钢筋 + AR-CONC 混凝土）
    "钢筋混凝土-钢筋":  ("钢筋混凝土-钢筋", RC_REBAR_COLOR),
    "钢筋混凝土-混凝土":   ("钢筋混凝土-混凝土", RC_CONC_COLOR),
}

# 细线层（垫层/反滤层）
PAD_LAYERS = {"反滤层-砾石", "反滤层-碎石", "反滤层-粗砂", "海漫-粗砂垫层"}


def setup_doc():
    doc = ezdxf.new("R2013", units=units.M)
    msp = doc.modelspace()
    # 中文文字样式（黑体，Windows 自带）
    doc.styles.add("CN", font="simhei.ttf")
    doc.header["$DWGCODEPAGE"] = "ANSI_936"
    # 加载标准线型（DASHED 水位虚线），否则 AutoCAD 因引用未定义线型打不开文件
    from ezdxf.tools import standards
    standards.setup_linetypes(doc)
    for name, (desc, color) in LAYERS.items():
        doc.layers.add(name=name, color=color)
    # 线宽：结构层粗线，垫层/反滤层细线
    for layer in doc.layers:
        if layer.dxf.name in PAD_LAYERS:
            layer.dxf.lineweight = PAD_LW
        else:
            layer.dxf.lineweight = STRUC_LW
    return doc, msp


def add_polyline(msp, points, layer, closed=False, color=None):
    attribs = {"layer": layer}
    if color is not None:
        attribs["color"] = color
    msp.add_lwpolyline(points, close=closed, dxfattribs=attribs)


def add_text(msp, x, y, text, layer="文字", height=0.4, align=None, color=None):
    attribs = {"layer": layer, "height": height, "style": "CN"}
    if color is not None:
        attribs["color"] = color
    t = msp.add_text(text, dxfattribs=attribs)
    a = ALIGN_MAP.get((align or "BOTTOMLEFT").upper(), TextEntityAlignment.BOTTOM_LEFT)
    t.set_placement((x, y), align=a)
    return t


def add_rect(msp, x1, y1, x2, y2, layer):
    """矩形（y1 < y2）"""
    add_polyline(msp, [(x1, y1), (x2, y1), (x2, y2), (x1, y2)], layer, closed=True)


def add_rect_no_bottom(msp, x1, y1, x2, y2, layer):
    """矩形不画底边（顶、左、右 3 条；底边在齿墙处单独画分段）"""
    add_polyline(msp, [(x2, y1), (x2, y2), (x1, y2), (x1, y1)], layer, closed=False)


def add_hatch(msp, points, layer, scale=400, rebar_layer="钢筋混凝土-钢筋", conc_layer="钢筋混凝土-混凝土"):
    """钢筋混凝土填充：双层叠加两个图层
    - RC_REBAR 层：ANSI31（45°钢筋斜线）
    - RC_CONC  层：AR-CONC（混凝土骨料散点）
    两个图层可分别开关/改色，符合钢筋混凝土断面标准画法。
    AR-CONC 图案定义直接嵌入 DXF（13 条图案线），
    避免依赖用户 CAD 的 PAT 库而显示不出来。
    scale 默认 400（mm）：ANSI31 线间距 400mm，AR-CONC 骨料 200mm，1:100 图上不密不疏。
    """
    # ANSI31 钢筋斜线（AutoCAD 标准图案，一般都有）
    hatch = msp.add_hatch(dxfattribs={"layer": rebar_layer})
    hatch.set_pattern_fill("ANSI31", scale=scale)
    hatch.paths.add_polyline_path(points, is_closed=True)
    # AR-CONC 混凝土骨料（嵌入定义）
    # 注意：set_pattern_definition 会把 solid_fill 置 1（纯色模式），
    # 必须显式设回 0（图案模式），否则 AutoCAD 解析为纯色填充导致黑屏/不显示
    hatch = msp.add_hatch(dxfattribs={"layer": conc_layer})
    conc_def = _load_arconc()
    hatch.set_pattern_definition(conc_def)
    hatch.dxf.solid_fill = 0
    hatch.dxf.pattern_name = "AR-CONC"
    hatch.dxf.pattern_scale = max(scale // 2, 60)
    hatch.paths.add_polyline_path(points, is_closed=True)
    return hatch


# AR-CONC 标准图案定义（缓存在模块级，避免每次加载）
_ARCONC_DEF = None


def _load_arconc():
    global _ARCONC_DEF
    if _ARCONC_DEF is None:
        from ezdxf.tools.pattern import load
        _ARCONC_DEF = load()["AR-CONC"]
    return _ARCONC_DEF


def add_cutoff_left(msp, x1, y_top, w, d, layer="齿墙"):
    """左端齿墙：向下(d) → 向右延伸(w) → 45°斜线向右上回到底面（四边形，不画顶边）
    同时填充钢筋混凝土图案。"""
    pts = [(x1, y_top), (x1, y_top - d), (x1 + w, y_top - d), (x1 + w + d, y_top)]
    add_hatch(msp, pts, layer, scale=400)
    add_polyline(msp, pts, layer, closed=False)


def add_cutoff_right(msp, x2, y_top, w, d, layer="齿墙"):
    """右端齿墙：向下(d) → 向左延伸(w) → 45°斜线向左上回到底面（四边形，不画顶边）
    同时填充钢筋混凝土图案。"""
    pts = [(x2, y_top), (x2, y_top - d), (x2 - w, y_top - d), (x2 - w - d, y_top)]
    add_hatch(msp, pts, layer, scale=400)
    add_polyline(msp, pts, layer, closed=False)


def add_dim_h(msp, x1, x2, y, text, layer="尺寸标注", color=STRUC_COLOR):
    add_polyline(msp, [(x1, y), (x2, y)], layer, color=color)
    add_polyline(msp, [(x1, y - 100), (x1, y + 300)], layer, color=color)
    add_polyline(msp, [(x2, y - 100), (x2, y + 300)], layer, color=color)
    add_text(msp, (x1 + x2) / 2, y + 50, text, layer=layer, height=350, color=color)


def generate_dxf(p, out_path):
    L = compute_layout(p)
    doc, msp = setup_doc()

    # 1. 铺盖（顶面 y_pg_t，底面 y_pg_b）
    add_rect_no_bottom(msp, L["x_pg1"], L["y_pg_b"], L["x_pg2"], L["y_pg_t"], "铺盖")
    # 混凝土填充图案（铺盖）
    add_hatch(msp, [
        (L["x_pg1"], L["y_pg_b"]), (L["x_pg2"], L["y_pg_b"]),
        (L["x_pg2"], L["y_pg_t"]), (L["x_pg1"], L["y_pg_t"]),
    ], "铺盖", scale=400)
    # 铺盖底边（两端齿墙位置断开，避免和齿墙顶边重合产生横线）
    add_polyline(msp, [(L["x_pg1"] + p["pg_cw"] + p["pg_cd"], L["y_pg_b"]),
                        (L["x_pg2"] - p["pg_cw"] - p["pg_cd"], L["y_pg_b"])], "铺盖")
    add_cutoff_left(msp, L["x_pg1"], L["y_pg_b"], p["pg_cw"], p["pg_cd"])
    add_cutoff_right(msp, L["x_pg2"], L["y_pg_b"], p["pg_cw"], p["pg_cd"])

    # 2. 底板
    add_rect_no_bottom(msp, L["x_db1"], L["y_db_b"], L["x_db2"], L["y_db_t"], "底板")
    # 混凝土填充图案（底板）
    add_hatch(msp, [
        (L["x_db1"], L["y_db_b"]), (L["x_db2"], L["y_db_b"]),
        (L["x_db2"], L["y_db_t"]), (L["x_db1"], L["y_db_t"]),
    ], "底板", scale=400)
    # 底板底边（两端齿墙位置断开）
    add_polyline(msp, [(L["x_db1"] + p["db_cw"] + p["db_cd"], L["y_db_b"]),
                        (L["x_db2"] - p["db_cw"] - p["db_cd"], L["y_db_b"])], "底板")
    add_cutoff_left(msp, L["x_db1"], L["y_db_b"], p["db_cw"], p["db_cd"])
    add_cutoff_right(msp, L["x_db2"], L["y_db_b"], p["db_cw"], p["db_cd"])

    # 3. 消力池底板（延伸到 x_xl_c2=44.5 含凸起齿墙下方，右端面与凸起齿墙右边缘连上；底板底边在齿墙位置断开）
    add_rect_no_bottom(msp, L["x_xl1"], L["y_xl_b"], L["x_xl_c2"], L["y_xl_t"], "消力池底板")
    # 钢筋混凝土填充（消力池）
    add_hatch(msp, [
        (L["x_xl1"], L["y_xl_b"]), (L["x_xl_c2"], L["y_xl_b"]),
        (L["x_xl_c2"], L["y_xl_t"]), (L["x_xl1"], L["y_xl_t"]),
    ], "消力池底板", scale=400)
    add_polyline(msp, [(L["x_xl1"] + p["xl_cw"] + p["xl_cd"], L["y_xl_b"]),
                        (L["x_xl_c2"] - p["xl_cw"] - p["xl_cd"], L["y_xl_b"])], "消力池底板")
    # 消力池上游端齿墙：四边形（向下+延伸+45°斜面）
    add_cutoff_left(msp, L["x_xl1"], L["y_xl_b"], p["xl_cw"], p["xl_cd"])
    # 消力池右下角齿墙：四边形，放在消力池最右端 x_xl_c2（与海漫齐平，不留缝）
    add_cutoff_right(msp, L["x_xl_c2"], L["y_xl_b"], p["xl_cw"], p["xl_cd"])
    # 消力池顶面凸起齿墙：长方形向上凸起（不画底边，底边与消力池顶面重合）+ 填充
    add_hatch(msp, [
        (L["x_xl2"], L["y_xl_t"]), (L["x_xl2"], L["y_xl_t"] + p["xl_cd"]),
        (L["x_xl_c2"], L["y_xl_t"] + p["xl_cd"]), (L["x_xl_c2"], L["y_xl_t"]),
    ], "齿墙", scale=400)
    add_polyline(
        msp,
        [(L["x_xl2"], L["y_xl_t"]), (L["x_xl2"], L["y_xl_t"] + p["xl_cd"]),
         (L["x_xl_c2"], L["y_xl_t"] + p["xl_cd"]), (L["x_xl_c2"], L["y_xl_t"])],
        "齿墙", closed=False,
    )

    # 4. 反滤层（消力池下方，三层；砾石/碎石层避开两端齿墙，粗砂层全宽包住齿墙底）
    fl_x1, fl_x2 = L["x_xl1"], L["x_xl_c2"]
    fl_xa = L["x_xl1"] + p["xl_cw"] + p["xl_cd"]
    fl_xb = L["x_xl_c2"] - p["xl_cw"] - p["xl_cd"]
    add_rect(msp, fl_xa, L["y_fl_gravel_b"], fl_xb, L["y_fl_gravel_t"], "反滤层-砾石")
    add_rect(msp, fl_xa, L["y_fl_stone_b"], fl_xb, L["y_fl_stone_t"], "反滤层-碎石")
    add_rect(msp, fl_x1, L["y_fl_sand_b"], fl_x2, L["y_fl_sand_t"], "反滤层-粗砂")

    # 5. 海漫水平段（浆砌石，不画底边；底边在齿墙位置断开）
    add_rect_no_bottom(msp, L["x_hm1"], L["y_hm_cushion_t1"], L["x_hm_h2"], L["y_hm_t1"], "海漫-浆砌石")
    # 混凝土填充图案（海漫水平段）
    add_hatch(msp, [
        (L["x_hm1"], L["y_hm_cushion_t1"]), (L["x_hm_h2"], L["y_hm_cushion_t1"]),
        (L["x_hm_h2"], L["y_hm_t1"]), (L["x_hm1"], L["y_hm_t1"]),
    ], "海漫-浆砌石", scale=400)
    add_polyline(msp, [(L["x_hm1"] + p["hm_cw"] + p["hm_cd"], L["y_hm_cushion_t1"]),
                        (L["x_hm_h2"] - p["hm_cw"] - p["hm_cd"], L["y_hm_cushion_t1"])], "海漫-浆砌石")
    # 海漫斜坡段（干砌石，不画底边；底边为斜线，在齿墙位置断开）
    add_polyline(
        msp,
        [(L["x_hm_h2"], L["y_hm_t1"]), (L["x_hm2"], L["y_hm_t2"]),
         (L["x_hm2"], L["y_hm_cushion_t2"]), (L["x_hm_h2"], L["y_hm_cushion_t1"])],
        "海漫-干砌石", closed=True,
    )
    # 混凝土填充图案（海漫斜坡段）
    add_hatch(msp, [
        (L["x_hm_h2"], L["y_hm_t1"]), (L["x_hm2"], L["y_hm_t2"]),
        (L["x_hm2"], L["y_hm_cushion_t2"]), (L["x_hm_h2"], L["y_hm_cushion_t1"]),
    ], "海漫-干砌石", scale=400)
    # 海漫齿墙（垫层已去掉；齿墙移到石层底，不留缝）
    add_cutoff_left(msp, L["x_hm1"], L["y_hm_cushion_t1"], p["hm_cw"], p["hm_cd"])
    add_cutoff_right(msp, L["x_hm_h2"], L["y_hm_cushion_t1"], p["hm_cw"], p["hm_cd"])
    # 海漫下游齿墙（倾斜段：斜线终点贴合石层底斜线，不留缝）+ 填充
    add_hatch(msp, [
        (L["x_hm2"], L["y_hm_cushion_t2"]),
        (L["x_hm2"], L["y_hm_cushion_t2"] - p["hm_cd"]),
        (L["x_hm2"] - p["hm_cw"], L["y_hm_cushion_t2"] - p["hm_cd"]),
        (L["x_hm2"] - p["hm_cw"] - p["hm_cd"],
         L["y_hm_cushion_t2"] + p["hm_slope"] * (p["hm_cw"] + p["hm_cd"]))],
        "齿墙", scale=400,
    )
    add_polyline(
        msp,
        [(L["x_hm2"], L["y_hm_cushion_t2"]),
         (L["x_hm2"], L["y_hm_cushion_t2"] - p["hm_cd"]),
         (L["x_hm2"] - p["hm_cw"], L["y_hm_cushion_t2"] - p["hm_cd"]),
         (L["x_hm2"] - p["hm_cw"] - p["hm_cd"],
          L["y_hm_cushion_t2"] + p["hm_slope"] * (p["hm_cw"] + p["hm_cd"]))],
        "齿墙", closed=False,
    )

    # 6. 防冲槽
    add_polyline(
        msp,
        [(L["x_fcc1"], L["y_fcc_t"]),
         (L["x_fcc_b1"], L["y_fcc_b"]),
         (L["x_fcc_b2"], L["y_fcc_b"]),
         (L["x_fcc2"], L["y_fcc_t"])],
        "防冲槽", closed=True,
    )
    # 堆石（边坡内侧 + 槽底）
    add_polyline(
        msp,
        [(L["x_fcc1"], L["y_fcc_t"]),
         (L["x_fcc_b1"], L["y_fcc_b"]),
         (L["x_fcc_b1"], L["y_rip_t"]),
         (L["x_fcc1"], L["y_fcc_t"] - p["fcc_rip"])],
        "防冲槽-堆石", closed=True,
    )
    add_polyline(
        msp,
        [(L["x_fcc2"], L["y_fcc_t"]),
         (L["x_fcc_b2"], L["y_fcc_b"]),
         (L["x_fcc_b2"], L["y_rip_t"]),
         (L["x_fcc2"], L["y_fcc_t"] - p["fcc_rip"])],
        "防冲槽-堆石", closed=True,
    )
    add_polyline(
        msp,
        [(L["x_fcc_b1"], L["y_fcc_b"]),
         (L["x_fcc_b2"], L["y_fcc_b"]),
         (L["x_fcc_b2"], L["y_rip_t"]),
         (L["x_fcc_b1"], L["y_rip_t"])],
        "防冲槽-堆石", closed=True,
    )

    # 7. 尺寸标注（挪到图最下方，防冲槽底之下）
    dim_y_top = L["y_fcc_b"] - 2500
    dim_y_total = L["y_fcc_b"] - 4000
    add_dim_h(msp, L["x_pg1"], L["x_pg2"], dim_y_top, f"铺盖 {p['pg_len']:.0f} mm")
    add_dim_h(msp, L["x_db1"], L["x_db2"], dim_y_top, f"底板 {p['db_len']:.0f} mm")
    add_dim_h(msp, L["x_xl1"], L["x_xl_c2"], dim_y_top, f"消力池 {p['xl_len'] + p['xl_cw']:.0f} mm")
    add_dim_h(msp, L["x_hm1"], L["x_hm2"], dim_y_top, f"海漫 {p['hm_total']:.0f} mm")
    add_dim_h(msp, L["x_fcc1"], L["x_fcc2"], dim_y_top, f"防冲槽 {L['x_fcc2'] - L['x_fcc1']:.0f} mm")
    add_dim_h(msp, L["x_pg1"], L["x_fcc2"], dim_y_total, f"总长 {L['x_fcc2']:.0f} mm")

    # 7.5 闸门（检修闸门 + 工作闸门，在底板上，顶到闸顶高程）
    add_rect(msp, L["x_g1a"], L["y_db_t"], L["x_g1b"], L["y_el_gate_top"], "闸门")
    add_rect(msp, L["x_g2a"], L["y_db_t"], L["x_g2b"], L["y_el_gate_top"], "闸门")

    # 7.6 排架 + 工作桥（保留各构件形状，仅去掉相邻构件重叠的小线段），轴线 = 闸门中心线
    pa_beam_yc = (L["y_el_gate_top"] + L["y_el_trestle"]) / 2  # 横梁在立柱高度中心 ≈81.8
    # 排架立柱（两段完整矩形，78.8~84.8）+ 填充
    add_rect(msp, L["x_pa1a"], L["y_el_gate_top"], L["x_pa1b"], L["y_el_trestle"], "闸门")
    add_hatch(msp, [
        (L["x_pa1a"], L["y_el_gate_top"]), (L["x_pa1b"], L["y_el_gate_top"]),
        (L["x_pa1b"], L["y_el_trestle"]), (L["x_pa1a"], L["y_el_trestle"]),
    ], "闸门", scale=400)
    add_rect(msp, L["x_pa2a"], L["y_el_gate_top"], L["x_pa2b"], L["y_el_trestle"], "闸门")
    add_hatch(msp, [
        (L["x_pa2a"], L["y_el_gate_top"]), (L["x_pa2b"], L["y_el_gate_top"]),
        (L["x_pa2b"], L["y_el_trestle"]), (L["x_pa2a"], L["y_el_trestle"]),
    ], "闸门", scale=400)
    # 排架横梁（只画顶边和底边两条横线，只在两立柱之间 16.35~20.15，立柱内部分不画）+ 填充
    add_hatch(msp, [
        (L["x_pa1b"], pa_beam_yc - p["pa_beam_h"] / 2), (L["x_pa2a"], pa_beam_yc - p["pa_beam_h"] / 2),
        (L["x_pa2a"], pa_beam_yc + p["pa_beam_h"] / 2), (L["x_pa1b"], pa_beam_yc + p["pa_beam_h"] / 2),
    ], "闸门", scale=400)
    add_polyline(msp, [(L["x_pa1b"], pa_beam_yc - p["pa_beam_h"] / 2),
                        (L["x_pa2a"], pa_beam_yc - p["pa_beam_h"] / 2)], "闸门")
    add_polyline(msp, [(L["x_pa1b"], pa_beam_yc + p["pa_beam_h"] / 2),
                        (L["x_pa2a"], pa_beam_yc + p["pa_beam_h"] / 2)], "闸门")
    # 工作桥：总高 0.7m（84.8~85.5），桥面板 0.2m 厚，两端底端 0.4m 宽
    br_btm = L["y_el_trestle"]                    # 84.8 = 排架顶（工作桥底=排架顶，高度由高程推导）
    br_deck_btm = L["y_el_bridge"] - p["br_deck_h"]  # 85.3
    # 桥面板填充（+ 顶边 + 左右缘 + 底边中段；底边两端与左/右腿顶重合的短横线不画）
    add_hatch(msp, [
        (L["x_br1"], br_deck_btm), (L["x_br2"], br_deck_btm),
        (L["x_br2"], L["y_el_bridge"]), (L["x_br1"], L["y_el_bridge"]),
    ], "闸门", scale=400)
    add_polyline(msp, [(L["x_br1"], L["y_el_bridge"]), (L["x_br2"], L["y_el_bridge"])], "闸门")  # 顶边
    add_polyline(msp, [(L["x_br1"], br_deck_btm), (L["x_br1"], L["y_el_bridge"])], "闸门")  # 左缘
    add_polyline(msp, [(L["x_br2"], br_deck_btm), (L["x_br2"], L["y_el_bridge"])], "闸门")  # 右缘
    add_polyline(msp, [(L["x_br1"] + p["br_col_w"], br_deck_btm),
                        (L["x_br2"] - p["br_col_w"], br_deck_btm)], "闸门")  # 底边中段（两腿之间）
    # 左底端（画左/底/右 3 边，顶边与桥面板底边重合不重复画）+ 填充
    add_hatch(msp, [
        (L["x_br1"], br_btm), (L["x_br1"] + p["br_col_w"], br_btm),
        (L["x_br1"] + p["br_col_w"], br_deck_btm), (L["x_br1"], br_deck_btm),
    ], "闸门", scale=400)
    add_polyline(msp, [
        (L["x_br1"], br_deck_btm),
        (L["x_br1"], br_btm),
        (L["x_br1"] + p["br_col_w"], br_btm),
        (L["x_br1"] + p["br_col_w"], br_deck_btm),
    ], "闸门", closed=False)
    # 右底端（画右/底/左 3 边，顶边不重复画）+ 填充
    add_hatch(msp, [
        (L["x_br2"] - p["br_col_w"], br_btm), (L["x_br2"], br_btm),
        (L["x_br2"], br_deck_btm), (L["x_br2"] - p["br_col_w"], br_deck_btm),
    ], "闸门", scale=400)
    add_polyline(msp, [
        (L["x_br2"], br_deck_btm),
        (L["x_br2"], br_btm),
        (L["x_br2"] - p["br_col_w"], br_btm),
        (L["x_br2"] - p["br_col_w"], br_deck_btm),
    ], "闸门", closed=False)

    # 7.7 交通桥（排架右端与底板右端之间的中心处；桥面顶 78.90）
    add_rect(msp, L["x_tb1"], L["y_tb_btm"], L["x_tb2"], L["y_tb_slab_btm"], "闸门")          # 垫层+橡胶支座
    add_rect(msp, L["x_tb1"], L["y_tb_slab_btm"], L["x_tb2"], L["y_tb_overlay_btm"], "闸门")  # 空心板
    # 空心板空心孔（2 个，细线）
    tb_hw = (L["x_tb2"] - L["x_tb1"]) / 3
    for i in (1, 2):
        hx1 = L["x_tb1"] + tb_hw * i - 0.12
        hx2 = L["x_tb1"] + tb_hw * i + 0.12
        add_rect(msp, hx1, L["y_tb_slab_btm"] + 0.08, hx2, L["y_tb_overlay_btm"] - 0.08, "文字")
    add_rect(msp, L["x_tb1"], L["y_tb_overlay_btm"], L["x_tb2"], L["y_tb_top"], "闸门")      # 板顶混凝土

    # 7.8 启闭机房（工作桥正上方；矩形墙体 + 人字顶 + 门）
    add_rect(msp, L["x_house1"], L["y_house_btm"], L["x_house2"], L["y_house_wall_top"], "闸门")  # 墙体
    add_polyline(msp, [
        (L["x_house1"], L["y_house_wall_top"]),
        (L["x_gc"], L["y_house_roof_top"]),
        (L["x_house2"], L["y_house_wall_top"]),
    ], "闸门", closed=True)  # 人字屋顶

    # 闸顶高程线（红色实线，从铺盖最左端到防冲槽最右端，与各结构顶面一样长）
    msp.add_lwpolyline(
        [(L["x_pg1"], L["y_el_gate_top"]), (L["x_fcc2"], L["y_el_gate_top"])],
        dxfattribs={"layer": "尺寸标注", "color": STRUC_COLOR},
    )
    # 结构分界垂直线（从各结构顶面到闸顶高程 78.8 相交，仅此一段）
    add_polyline(msp, [(L["x_pg2"], L["y_pg_t"]), (L["x_pg2"], L["y_el_gate_top"])], "尺寸标注")    # 铺盖/底板
    add_polyline(msp, [(L["x_db2"], L["y_db_t"]), (L["x_db2"], L["y_el_gate_top"])], "尺寸标注")    # 底板/消力池
    add_polyline(msp, [(L["x_xl_c2"], L["y_hm_t1"]), (L["x_xl_c2"], L["y_el_gate_top"])], "尺寸标注")  # 消力池/海漫
    add_polyline(msp, [(L["x_hm2"], L["y_fcc_t"]), (L["x_hm2"], L["y_el_gate_top"])], "尺寸标注")    # 海漫/防冲槽

    # 上游翼墙光滑曲线：底板左上角 (x_db1, y_db_t) 向铺盖方向延伸，
    # 弯曲上升交到闸顶高程；终点在铺盖中点 (x_pg1 + x_pg2)/2
    # 注意：用 Catmull-Rom 多点折线（与 SVG 及求交函数同源同形），
    # 避免 SPLINE 与 Catmull-Rom 形状差异导致水平线端点"超出曲线"
    up_curve_ctrl = [
        (L["x_db1"],                  L["y_db_t"]),                        # 起点：底板左上角
        (L["x_db1"] - 2000,          L["y_db_t"] + 150),                  # 微抬
        (L["x_db1"] - 4000,          L["y_db_t"] + 500),                  # 缓升
        (L["x_db1"] - 5500,          L["y_db_t"] + 1500),                 # 渐陡
        ((L["x_pg1"] + L["x_pg2"]) / 2, L["y_el_gate_top"]),              # 终点：铺盖中点 + 闸顶
    ]
    msp.add_lwpolyline(
        _catmull_rom(up_curve_ctrl, 200),
        dxfattribs={"layer": "尺寸标注", "color": STRUC_COLOR},
    )

    # 下游翼墙光滑曲线：海漫左上角 (x_hm1, y_hm_t1) 向海漫方向延伸，
    # 弯曲上升交到闸顶高程；终点在海漫水平段末端 (x_hm_h2)
    down_curve_ctrl = [
        (L["x_hm1"],                 L["y_hm_t1"]),                        # 起点：海漫左上角
        (L["x_hm1"] + 2000,          L["y_hm_t1"] + 150),                  # 微抬
        (L["x_hm1"] + 4000,          L["y_hm_t1"] + 500),                  # 缓升
        (L["x_hm1"] + 6000,          L["y_hm_t1"] + 1500),                 # 渐陡
        (L["x_hm_h2"],               L["y_el_gate_top"]),                  # 终点：海漫水平段末端 + 闸顶
    ]
    msp.add_lwpolyline(
        _catmull_rom(down_curve_ctrl, 200),
        dxfattribs={"layer": "尺寸标注", "color": STRUC_COLOR},
    )

    # 滩地高程水平线：从最左到最右，但被两条翼墙曲线截断——两交点之间不画，只保留两端
    bank_y = L["y_el_bank"]                      # 75000
    up_bank_x = _spline_intersect_y(up_curve_ctrl, bank_y)
    down_bank_x = _spline_intersect_y(down_curve_ctrl, bank_y)
    if up_bank_x is not None and down_bank_x is not None:
        add_polyline(msp, [(L["x_pg1"], bank_y), (up_bank_x, bank_y)], "尺寸标注")
        add_polyline(msp, [(down_bank_x, bank_y), (L["x_fcc2"], bank_y)], "尺寸标注")

    # 正常蓄水位水平线（实线，被两翼墙曲线截断——两交点之间不画，只保留两端）
    wl_y = L["y_el_wl"]                          # 76600
    up_wl_x = _spline_intersect_y(up_curve_ctrl, wl_y)
    down_wl_x = _spline_intersect_y(down_curve_ctrl, wl_y)
    if up_wl_x is not None and down_wl_x is not None:
        add_polyline(msp, [(L["x_pg1"], wl_y), (up_wl_x, wl_y)], "尺寸标注")
        add_polyline(msp, [(down_wl_x, wl_y), (L["x_fcc2"], wl_y)], "尺寸标注")

    # 正常蓄水位线虚线段已删除（原 0~x_g2a 段与新加的截断式实线在 0~x_up_wl 区间重叠，导致视觉重影）
    # 现在 y=76.6 处只有新加的两段实线截断式水平线（被两翼墙曲线截断）


    # 高程标注（右侧集中标注）
    el_x1 = L["x_fcc2"] + 1200
    el_x2 = L["x_fcc2"] + 4200

    def add_el(y, label):
        add_polyline(msp, [(el_x1, y), (el_x2, y)], "尺寸标注")
        add_polyline(msp, [(el_x1, y - 150), (el_x1, y + 150)], "尺寸标注")
        add_text(msp, el_x1 + 100, y + 100, label, layer="文字", height=350)

    add_el(L["y_el_pg"], f"▽{p['el_pg']:.0f} 铺盖顶（底板顶）")
    add_el(L["y_el_bank"], f"▽{p['el_bank']:.0f} 滩地")
    add_el(L["y_el_wl"], f"▽{p['el_wl']:.0f} 正常蓄水位")
    add_el(L["y_el_gate_top"], f"▽{p['el_gate_top']:.0f} 闸顶")
    add_el(L["y_el_trestle"], f"▽{p['el_trestle']:.0f} 排架")
    add_el(L["y_el_bridge"], f"▽{p['el_bridge']:.0f} 工作桥")

    # 8. 尺寸标注文字已在上方（构件名/材料/齿墙等文字按用户要求全部删除）

    doc.saveas(out_path, encoding="utf-8")
    return L


# ============================================================
# 4. SVG 生成
# ============================================================
# SVG 1mm → 0.022px（与原 m→22px 等价，即原比例 1:100 ≈ 屏幕 1mm=0.22px）
SVG_PX_PER_MM = 0.022
SVG_PX_PER_M = SVG_PX_PER_MM  # 保留旧名以免破坏其他位置用法


def _catmull_rom(ctrl, n=32):
    """Catmull-Rom 样条插值：给定控制点列表，返回 n 个插值点（含首尾），生成光滑曲线"""
    if len(ctrl) < 3:
        return ctrl
    pts = []
    segs = len(ctrl) - 1
    per = max(2, n // segs)
    for i in range(segs):
        p0 = ctrl[i - 1] if i > 0 else ctrl[i]
        p1 = ctrl[i]
        p2 = ctrl[i + 1]
        p3 = ctrl[i + 2] if i + 2 < len(ctrl) else p2
        for j in range(per):
            t = j / per
            t2, t3 = t * t, t * t * t
            x = 0.5 * ((2 * p1[0]) + (-p0[0] + p2[0]) * t + (2 * p0[0] - 5 * p1[0] + 4 * p2[0] - p3[0]) * t2 + (-p0[0] + 3 * p1[0] - 3 * p2[0] + p3[0]) * t3)
            y = 0.5 * ((2 * p1[1]) + (-p0[1] + p2[1]) * t + (2 * p0[1] - 5 * p1[1] + 4 * p2[1] - p3[1]) * t2 + (-p0[1] + 3 * p1[1] - 3 * p2[1] + p3[1]) * t3)
            pts.append((x, y))
    pts.append(ctrl[-1])
    return pts


def _spline_intersect_y(ctrl, y_level, n=2048):
    """求样条曲线（由控制点定义）与水平线 y=y_level 的交点 x。
    用 Catmull-Rom 密集采样 + 线性插值求交，返回第一个交点的 x；若无交点返回 None。
    n 默认 2048：80m 范围段长约 40mm，线性插值后误差 < 1mm。"""
    pts = _catmull_rom(ctrl, n=n)
    for i in range(len(pts) - 1):
        (x1, y1), (x2, y2) = pts[i], pts[i + 1]
        if (y1 - y_level) * (y2 - y_level) <= 0:  # 跨越水平线
            if y2 == y1:
                return x1
            t = (y_level - y1) / (y2 - y1)
            return x1 + t * (x2 - x1)
    return None


def generate_svg(p, out_path):
    L = compute_layout(p)

    margin_l, margin_r = 90, 80
    margin_t, margin_b = 60, 60
    width_m = L["x_fcc2"]  # 这里 x_fcc2 已为 mm 单位
    # SVG Y 轴向下，但我们要在 SVG 顶显示结构顶 → 把结构翻转
    y_top = max(L["y_pg_t"], L["y_hm_t1"], L["y_el_bridge"]) + 2000
    height_m = y_top - (L["y_fcc_b"] - 4800) + 1000  # 底部留出尺寸标注区（含下移后的总长标注）
    width_px = int(width_m * SVG_PX_PER_MM) + margin_l + margin_r
    height_px = int(height_m * SVG_PX_PER_MM) + margin_t + margin_b

    def M(x, y):
        # CAD y 向上 → SVG y 向下：翻转
        svg_y = (y_top - y) * SVG_PX_PER_MM
        return (margin_l + x * SVG_PX_PER_MM, margin_t + svg_y)

    parts = []
    parts.append(
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width_px}" height="{height_px}" '
        f'viewBox="0 0 {width_px} {height_px}" style="background:#fafafa;font-family:SimHei,Microsoft YaHei,sans-serif;">'
    )
    # ===== 混凝土填充图案（模拟 AR-CONC：45°斜线 + 散点）=====
    # 钢筋混凝土填充：浅底 + 45°钢筋斜线（粗）+ 混凝土骨料散点
    # 单元 1600×1600mm，斜线每 320mm 一条（比 v53 的 800/160 更疏）
    parts.append(
        '<defs>'
        '<pattern id="conc-fill" patternUnits="userSpaceOnUse" width="1600" height="1600" patternTransform="rotate(45)">'
        # 浅色底（钢筋混凝土断面底色）
        '<rect width="1600" height="1600" fill="#e8e8e8"/>'
        # 钢筋斜线（每 320mm 一条，粗、深）
        '<line x1="0" y1="0" x2="1600" y2="0" stroke="#24458a" stroke-width="3"/>'
        '<line x1="0" y1="320" x2="1600" y2="320" stroke="#24458a" stroke-width="3"/>'
        '<line x1="0" y1="640" x2="1600" y2="640" stroke="#24458a" stroke-width="3"/>'
        '<line x1="0" y1="960" x2="1600" y2="960" stroke="#24458a" stroke-width="3"/>'
        '<line x1="0" y1="1280" x2="1600" y2="1280" stroke="#24458a" stroke-width="3"/>'
        # 混凝土骨料散点
        '<circle cx="120" cy="120" r="6" fill="#24458a"/>'
        '<circle cx="520" cy="80" r="5" fill="#24458a"/>'
        '<circle cx="1000" cy="160" r="6" fill="#24458a"/>'
        '<circle cx="1400" cy="60" r="5" fill="#24458a"/>'
        '<circle cx="240" cy="440" r="6" fill="#24458a"/>'
        '<circle cx="760" cy="400" r="5" fill="#24458a"/>'
        '<circle cx="1240" cy="520" r="6" fill="#24458a"/>'
        '<circle cx="1520" cy="360" r="5" fill="#24458a"/>'
        '<circle cx="160" cy="800" r="6" fill="#24458a"/>'
        '<circle cx="680" cy="720" r="5" fill="#24458a"/>'
        '<circle cx="1160" cy="840" r="6" fill="#24458a"/>'
        '<circle cx="1480" cy="760" r="5" fill="#24458a"/>'
        '<circle cx="400" cy="1120" r="6" fill="#24458a"/>'
        '<circle cx="920" cy="1080" r="5" fill="#24458a"/>'
        '<circle cx="1320" cy="1240" r="6" fill="#24458a"/>'
        '<circle cx="120" cy="1400" r="6" fill="#24458a"/>'
        '<circle cx="600" cy="1360" r="5" fill="#24458a"/>'
        '<circle cx="1080" cy="1480" r="6" fill="#24458a"/>'
        # 反向短斜线（骨料纹理）
        '<line x1="400" y1="240" x2="500" y2="340" stroke="#24458a" stroke-width="2.5"/>'
        '<line x1="900" y1="180" x2="1000" y2="280" stroke="#24458a" stroke-width="2.5"/>'
        '<line x1="320" y1="600" x2="420" y2="700" stroke="#24458a" stroke-width="2.5"/>'
        '<line x1="840" y1="560" x2="940" y2="660" stroke="#24458a" stroke-width="2.5"/>'
        '<line x1="1300" y1="680" x2="1400" y2="780" stroke="#24458a" stroke-width="2.5"/>'
        '<line x1="520" y1="960" x2="620" y2="1060" stroke="#24458a" stroke-width="2.5"/>'
        '<line x1="1040" y1="920" x2="1140" y2="1020" stroke="#24458a" stroke-width="2.5"/>'
        '<line x1="240" y1="1240" x2="340" y2="1340" stroke="#24458a" stroke-width="2.5"/>'
        '<line x1="760" y1="1200" x2="860" y2="1300" stroke="#24458a" stroke-width="2.5"/>'
        '<line x1="1240" y1="1400" x2="1340" y2="1500" stroke="#24458a" stroke-width="2.5"/>'
        '</pattern>'
        '</defs>'
    )

    COL = {
        "PG_FILL": "url(#conc-fill)", "DB_FILL": "url(#conc-fill)", "XL_FILL": "#7faedd",
        "FL_GRAVEL": "#7ddc7d", "FL_STONE": "#e6d96a", "FL_SAND": "#e89090",
        "HM_STONE": "url(#conc-fill)", "HM_DRY": "url(#conc-fill)", "HM_CUSH": "#f0a060",
        "FCC_FILL": "#b8d8e8", "FCC_RIP": "#a87a52", "CUTOFF": "#404040",
        "GATE": "#c8b8e0", "OUTLINE": "#000", "DIM": "#e60000", "TEXT": "#000",
    }

    STRUC_STROKE = "#e60000"  # 结构线：红色粗线
    STRUC_SW = 2.5
    PAD_STROKE = "#00a000"    # 垫层/反滤层线：绿色细线
    PAD_SW = 1.0

    def poly(points, fill, stroke=STRUC_STROKE, sw=STRUC_SW):
        pts = " ".join(f"{x:.1f},{y:.1f}" for x, y in points)
        return f'<polygon points="{pts}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'

    def poly_no_top(points, fill, stroke=STRUC_STROKE, sw=STRUC_SW):
        """多边形填充 + 只描前 3 条边（最后一条闭合边=顶边不描，与底板底边重合）"""
        pts = " ".join(f"{x:.1f},{y:.1f}" for x, y in points)
        s = f'<polygon points="{pts}" fill="{fill}" stroke="none"/>'
        s += f'<polyline points="{pts}" fill="none" stroke="{stroke}" stroke-width="{sw}"/>'
        return s

    def line(p1, p2, stroke=STRUC_STROKE, sw=STRUC_SW):
        return f'<line x1="{p1[0]:.1f}" y1="{p1[1]:.1f}" x2="{p2[0]:.1f}" y2="{p2[1]:.1f}" stroke="{stroke}" stroke-width="{sw}"/>'

    def text(x, y, s, size=14, color="#000", anchor="start", rotate=None):
        t = f'<text x="{x:.1f}" y="{y:.1f}" font-size="{size}" fill="{color}" text-anchor="{anchor}"'
        if rotate is not None:
            t += f' transform="rotate({rotate} {x:.1f} {y:.1f})"'
        t += f'>{s}</text>'
        return t

    def rect(x1, y1, x2, y2, fill, stroke=STRUC_STROKE, sw=STRUC_SW):
        p1, p2 = M(x1, y1), M(x2, y2)
        # 在 SVG 坐标里 x1 对应左边，x2 对应右边
        # y1 (CAD 较低) → SVG 较大值（在 SVG 中更靠下）
        # y2 (CAD 较高) → SVG 较小值（在 SVG 中更靠上）
        x_left, x_right = min(p1[0], p2[0]), max(p1[0], p2[0])
        y_top_svg, y_bot_svg = min(p1[1], p2[1]), max(p1[1], p2[1])
        return f'<rect x="{x_left:.1f}" y="{y_top_svg:.1f}" width="{x_right-x_left:.1f}" height="{y_bot_svg-y_top_svg:.1f}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'

    def rect_no_bottom(x1, y1, x2, y2, fill):
        """矩形不画底边（顶、左、右 3 边 polygon 填充，底边单独 line 画）"""
        pts = [M(x2, y1), M(x2, y2), M(x1, y2), M(x1, y1)]
        pts_str = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
        return f'<polygon points="{pts_str}" fill="{fill}" stroke="none"/>'

    def dim_h(x1, x2, y_m, label):
        p1, p2 = M(x1, y_m), M(x2, y_m)
        parts.append(line(p1, p2, COL["DIM"], STRUC_SW))
        for px in (p1, p2):
            parts.append(line((px[0], px[1] - 10), (px[0], px[1] + 10), COL["DIM"], STRUC_SW))
        mid_x = (p1[0] + p2[0]) / 2
        parts.append(text(mid_x, p1[1] - 14, label, size=26, color=COL["DIM"], anchor="middle"))

    def dim_v(x_m, y1, y2, label):
        p1, p2 = M(x_m, y1), M(x_m, y2)
        parts.append(line(p1, p2, COL["DIM"], STRUC_SW))
        for py in (p1, p2):
            parts.append(line((py[0] - 10, py[1]), (py[0] + 10, py[1]), COL["DIM"], STRUC_SW))
        mid_y = (p1[1] + p2[1]) / 2
        parts.append(text(p1[0] - 14, mid_y, label, size=26, color=COL["DIM"], anchor="end", rotate=-90))

    # ----- 构件 -----
    # 铺盖（不画底边，底边在齿墙位置断开单独画）
    parts.append(rect_no_bottom(L["x_pg1"], L["y_pg_b"], L["x_pg2"], L["y_pg_t"], COL["PG_FILL"]))
    parts.append(line(M(L["x_pg1"] + p["pg_cw"] + p["pg_cd"], L["y_pg_b"]),
                     M(L["x_pg2"] - p["pg_cw"] - p["pg_cd"], L["y_pg_b"]), "#000"))  # 底边（断开）
    parts.append(line(M(L["x_pg1"], L["y_pg_b"]), M(L["x_pg1"], L["y_pg_t"]), "#000"))  # 左边
    parts.append(line(M(L["x_pg1"], L["y_pg_t"]), M(L["x_pg2"], L["y_pg_t"]), "#000"))  # 顶边
    parts.append(line(M(L["x_pg2"], L["y_pg_b"]), M(L["x_pg2"], L["y_pg_t"]), "#000"))  # 右边
    parts.append(poly_no_top(  # 铺盖左端齿墙（不画顶边）
        [M(L["x_pg1"], L["y_pg_b"]),
         M(L["x_pg1"], L["y_pg_b"] - p["pg_cd"]),
         M(L["x_pg1"] + p["pg_cw"], L["y_pg_b"] - p["pg_cd"]),
         M(L["x_pg1"] + p["pg_cw"] + p["pg_cd"], L["y_pg_b"])],
        COL["CUTOFF"],
    ))
    parts.append(poly_no_top(  # 铺盖右端齿墙（不画顶边）
        [M(L["x_pg2"], L["y_pg_b"]),
         M(L["x_pg2"], L["y_pg_b"] - p["pg_cd"]),
         M(L["x_pg2"] - p["pg_cw"], L["y_pg_b"] - p["pg_cd"]),
         M(L["x_pg2"] - p["pg_cw"] - p["pg_cd"], L["y_pg_b"])],
        COL["CUTOFF"],
    ))

    # 底板（不画底边，底边在齿墙位置断开）
    parts.append(rect_no_bottom(L["x_db1"], L["y_db_b"], L["x_db2"], L["y_db_t"], COL["DB_FILL"]))
    parts.append(line(M(L["x_db1"] + p["db_cw"] + p["db_cd"], L["y_db_b"]),
                     M(L["x_db2"] - p["db_cw"] - p["db_cd"], L["y_db_b"]), "#000"))  # 底边（断开）
    parts.append(line(M(L["x_db1"], L["y_db_b"]), M(L["x_db1"], L["y_db_t"]), "#000"))  # 左边
    parts.append(line(M(L["x_db1"], L["y_db_t"]), M(L["x_db2"], L["y_db_t"]), "#000"))  # 顶边
    parts.append(line(M(L["x_db2"], L["y_db_b"]), M(L["x_db2"], L["y_db_t"]), "#000"))  # 右边
    parts.append(poly_no_top(  # 底板左端齿墙（不画顶边，深 1m）
        [M(L["x_db1"], L["y_db_b"]),
         M(L["x_db1"], L["y_db_b"] - p["db_cd"]),
         M(L["x_db1"] + p["db_cw"], L["y_db_b"] - p["db_cd"]),
         M(L["x_db1"] + p["db_cw"] + p["db_cd"], L["y_db_b"])],
        COL["CUTOFF"],
    ))
    parts.append(poly_no_top(  # 底板右端齿墙（不画顶边，深 1m）
        [M(L["x_db2"], L["y_db_b"]),
         M(L["x_db2"], L["y_db_b"] - p["db_cd"]),
         M(L["x_db2"] - p["db_cw"], L["y_db_b"] - p["db_cd"]),
         M(L["x_db2"] - p["db_cw"] - p["db_cd"], L["y_db_b"])],
        COL["CUTOFF"],
    ))

    # 消力池底板（延伸到 x_xl_c2=44.5 含凸起齿墙下方，右端面与凸起齿墙右边缘连上）
    parts.append(rect_no_bottom(L["x_xl1"], L["y_xl_b"], L["x_xl_c2"], L["y_xl_t"], COL["XL_FILL"]))
    # 消力池底板底边（两端齿墙位置断开）
    parts.append(line(M(L["x_xl1"] + p["xl_cw"] + p["xl_cd"], L["y_xl_b"]),
                     M(L["x_xl_c2"] - p["xl_cw"] - p["xl_cd"], L["y_xl_b"]), "#000"))
    parts.append(line(M(L["x_xl1"], L["y_xl_b"]), M(L["x_xl1"], L["y_xl_t"]), "#000"))  # 左边
    parts.append(line(M(L["x_xl1"], L["y_xl_t"]), M(L["x_xl_c2"], L["y_xl_t"]), "#000"))  # 顶边（延伸到 44.5）
    parts.append(line(M(L["x_xl_c2"], L["y_xl_b"]), M(L["x_xl_c2"], L["y_xl_t"]), "#000"))  # 右端面（连接凸起齿墙）

    # 反滤层（先画，齿墙后画盖在上面 → 反滤层包住齿墙；砾石/碎石层避开齿墙投影）
    fl_x1, fl_x2 = L["x_xl1"], L["x_xl_c2"]
    fl_xa = L["x_xl1"] + p["xl_cw"] + p["xl_cd"]
    fl_xb = L["x_xl_c2"] - p["xl_cw"] - p["xl_cd"]
    parts.append(rect(fl_xa, L["y_fl_gravel_b"], fl_xb, L["y_fl_gravel_t"], COL["FL_GRAVEL"], stroke=PAD_STROKE, sw=PAD_SW))
    parts.append(rect(fl_xa, L["y_fl_stone_b"], fl_xb, L["y_fl_stone_t"], COL["FL_STONE"], stroke=PAD_STROKE, sw=PAD_SW))
    parts.append(rect(fl_x1, L["y_fl_sand_b"], fl_x2, L["y_fl_sand_t"], COL["FL_SAND"], stroke=PAD_STROKE, sw=PAD_SW))

    # 消力池上游端齿墙：四边形（不画顶边）
    parts.append(poly_no_top(
        [M(L["x_xl1"], L["y_xl_b"]),
         M(L["x_xl1"], L["y_xl_b"] - p["xl_cd"]),
         M(L["x_xl1"] + p["xl_cw"], L["y_xl_b"] - p["xl_cd"]),
         M(L["x_xl1"] + p["xl_cw"] + p["xl_cd"], L["y_xl_b"])],
        COL["CUTOFF"],
    ))
    # 消力池右下角齿墙：四边形，放在消力池最右端 x_xl_c2（与海漫齐平不留缝）
    parts.append(poly_no_top(
        [M(L["x_xl_c2"], L["y_xl_b"]),
         M(L["x_xl_c2"], L["y_xl_b"] - p["xl_cd"]),
         M(L["x_xl_c2"] - p["xl_cw"], L["y_xl_b"] - p["xl_cd"]),
         M(L["x_xl_c2"] - p["xl_cw"] - p["xl_cd"], L["y_xl_b"])],
        COL["CUTOFF"],
    ))
    # 消力池顶面凸起齿墙：长方形向上凸起（不画底边）
    parts.append(poly_no_top(
        [M(L["x_xl2"], L["y_xl_t"]),
         M(L["x_xl2"], L["y_xl_t"] + p["xl_cd"]),
         M(L["x_xl_c2"], L["y_xl_t"] + p["xl_cd"]),
         M(L["x_xl_c2"], L["y_xl_t"])],
        COL["CUTOFF"],
    ))

    # 海漫水平段石层（不画底边；底边在齿墙位置断开）
    parts.append(poly(
        [M(L["x_hm_h2"], L["y_hm_t1"]), M(L["x_hm1"], L["y_hm_t1"]),
         M(L["x_hm1"], L["y_hm_cushion_t1"]), M(L["x_hm1"] + p["hm_cw"] + p["hm_cd"], L["y_hm_cushion_t1"])],
        COL["HM_STONE"], stroke="none",
    ))
    parts.append(poly(
        [M(L["x_hm_h2"] - p["hm_cw"] - p["hm_cd"], L["y_hm_cushion_t1"]), M(L["x_hm_h2"], L["y_hm_cushion_t1"])],
        COL["HM_STONE"], stroke="none",
    ))
    parts.append(line(M(L["x_hm1"] + p["hm_cw"] + p["hm_cd"], L["y_hm_cushion_t1"]),
                     M(L["x_hm_h2"] - p["hm_cw"] - p["hm_cd"], L["y_hm_cushion_t1"]), "#000"))
    parts.append(poly(
        [M(L["x_hm_h2"], L["y_hm_t1"]), M(L["x_hm2"], L["y_hm_t2"]),
         M(L["x_hm2"], L["y_hm_cushion_t2"]), M(L["x_hm_h2"], L["y_hm_cushion_t1"])],
        COL["HM_DRY"],
    ))
    # 海漫齿墙（垫层已去掉；齿墙移到石层底，不留缝）
    # 海漫上游端齿墙（从石层底向下）
    parts.append(poly_no_top(
        [M(L["x_hm1"], L["y_hm_cushion_t1"]),
         M(L["x_hm1"], L["y_hm_cushion_t1"] - p["hm_cd"]),
         M(L["x_hm1"] + p["hm_cw"], L["y_hm_cushion_t1"] - p["hm_cd"]),
         M(L["x_hm1"] + p["hm_cw"] + p["hm_cd"], L["y_hm_cushion_t1"])],
        COL["CUTOFF"],
    ))
    # 海漫水平段右端 = 倾斜段左端 齿墙（从石层底向下）
    parts.append(poly_no_top(
        [M(L["x_hm_h2"], L["y_hm_cushion_t1"]),
         M(L["x_hm_h2"], L["y_hm_cushion_t1"] - p["hm_cd"]),
         M(L["x_hm_h2"] - p["hm_cw"], L["y_hm_cushion_t1"] - p["hm_cd"]),
         M(L["x_hm_h2"] - p["hm_cw"] - p["hm_cd"], L["y_hm_cushion_t1"])],
        COL["CUTOFF"],
    ))
    # 海漫下游端齿墙（倾斜段：斜线终点贴合石层底斜线，不留缝）
    parts.append(poly_no_top(
        [M(L["x_hm2"], L["y_hm_cushion_t2"]),
         M(L["x_hm2"], L["y_hm_cushion_t2"] - p["hm_cd"]),
         M(L["x_hm2"] - p["hm_cw"], L["y_hm_cushion_t2"] - p["hm_cd"]),
         M(L["x_hm2"] - p["hm_cw"] - p["hm_cd"],
           L["y_hm_cushion_t2"] + p["hm_slope"] * (p["hm_cw"] + p["hm_cd"]))],
        COL["CUTOFF"],
    ))

    parts.append(poly(
        [M(L["x_fcc1"], L["y_fcc_t"]), M(L["x_fcc_b1"], L["y_fcc_b"]),
         M(L["x_fcc_b2"], L["y_fcc_b"]), M(L["x_fcc2"], L["y_fcc_t"])],
        COL["FCC_FILL"],
    ))
    parts.append(poly(
        [M(L["x_fcc1"], L["y_fcc_t"]), M(L["x_fcc_b1"], L["y_fcc_b"]),
         M(L["x_fcc_b1"], L["y_rip_t"]), M(L["x_fcc1"], L["y_fcc_t"] - p["fcc_rip"])],
        COL["FCC_RIP"],
    ))
    parts.append(poly(
        [M(L["x_fcc2"], L["y_fcc_t"]), M(L["x_fcc_b2"], L["y_fcc_b"]),
         M(L["x_fcc_b2"], L["y_rip_t"]), M(L["x_fcc2"], L["y_fcc_t"] - p["fcc_rip"])],
        COL["FCC_RIP"],
    ))
    parts.append(poly(
        [M(L["x_fcc_b1"], L["y_fcc_b"]), M(L["x_fcc_b2"], L["y_fcc_b"]),
         M(L["x_fcc_b2"], L["y_rip_t"]), M(L["x_fcc_b1"], L["y_rip_t"])],
        COL["FCC_RIP"],
    ))

    # ---- 尺寸（挪到图最下方，防冲槽底之下）----
    dim_h(L["x_pg1"], L["x_pg2"], L["y_fcc_b"] - 2500, f"铺盖 {p['pg_len']:.0f}mm")
    dim_h(L["x_db1"], L["x_db2"], L["y_fcc_b"] - 2500, f"底板 {p['db_len']:.0f}mm")
    dim_h(L["x_xl1"], L["x_xl_c2"], L["y_fcc_b"] - 2500, f"消力池 {p['xl_len'] + p['xl_cw']:.0f}mm")
    dim_h(L["x_hm1"], L["x_hm2"], L["y_fcc_b"] - 2500, f"海漫 {p['hm_total']:.0f}mm")
    dim_h(L["x_fcc1"], L["x_fcc2"], L["y_fcc_b"] - 2500, f"防冲槽 {L['x_fcc2'] - L['x_fcc1']:.0f}mm")
    dim_h(L["x_pg1"], L["x_fcc2"], L["y_fcc_b"] - 4000, f"总长 {L['x_fcc2']:.0f}mm")

    dim_v(L["x_pg1"] - 1500, L["y_pg_b"], L["y_pg_t"], f"{p['pg_h']:.0f}mm")
    dim_v(L["x_db1"] - 1500, L["y_db_b"], L["y_db_t"], f"{p['db_h']:.0f}mm")
    dim_v(L["x_xl1"] - 1500, L["y_xl_b"], L["y_xl_t"], f"{p['xl_h']:.0f}mm")
    dim_v(L["x_hm1"] - 1500, L["y_hm_b1"], L["y_hm_t1"], f"{p['hm_stone']+p['hm_cushion']:.0f}mm")
    dim_v(L["x_fcc1"] - 1500, L["y_fcc_b"], L["y_fcc_t"], f"{p['fcc_d']:.0f}mm")

    # 闸门（检修闸门 + 工作闸门，在底板上，顶到闸顶高程）
    parts.append(rect(L["x_g1a"], L["y_db_t"], L["x_g1b"], L["y_el_gate_top"], COL["GATE"]))
    parts.append(rect(L["x_g2a"], L["y_db_t"], L["x_g2b"], L["y_el_gate_top"], COL["GATE"]))

    # 排架 + 工作桥（画成整体外轮廓 polygon，内部线条去掉），轴线 = 闸门中心线
    pa_beam_yc = (L["y_el_gate_top"] + L["y_el_trestle"]) / 2  # 横梁在立柱高度中心 ≈81.8
    # 排架立柱（两段完整矩形，78.8~84.8）+ 钢筋混凝土填充
    parts.append(rect(L["x_pa1a"], L["y_el_gate_top"], L["x_pa1b"], L["y_el_trestle"], "url(#conc-fill)"))
    parts.append(rect(L["x_pa2a"], L["y_el_gate_top"], L["x_pa2b"], L["y_el_trestle"], "url(#conc-fill)"))
    # 排架横梁（只画顶边和底边两条横线，两端竖线与立柱外缘重合不再重复画）+ 填充
    parts.append(rect(L["x_pa1b"], pa_beam_yc - p["pa_beam_h"] / 2, L["x_pa2a"], pa_beam_yc + p["pa_beam_h"] / 2,
                      "url(#conc-fill)", stroke="none"))
    parts.append(line(M(L["x_pa1b"], pa_beam_yc - p["pa_beam_h"] / 2),
                     M(L["x_pa2a"], pa_beam_yc - p["pa_beam_h"] / 2)))
    parts.append(line(M(L["x_pa1b"], pa_beam_yc + p["pa_beam_h"] / 2),
                     M(L["x_pa2a"], pa_beam_yc + p["pa_beam_h"] / 2)))
    # 工作桥：总高 0.7m（84.8~85.5），桥面板 0.2m 厚，两端底端 0.4m 宽
    br_btm = L["y_el_trestle"]                    # 84.8 = 排架顶（工作桥底=排架顶，高度由高程推导）
    br_deck_btm = L["y_el_bridge"] - p["br_deck_h"]  # 85.3
    # 桥面板填充 +（顶边 + 左右缘 + 底边中段；底边两端与左/右腿顶重合的短横线不画）
    parts.append(rect(L["x_br1"], br_deck_btm, L["x_br2"], L["y_el_bridge"], "url(#conc-fill)", stroke="none"))
    parts.append(line(M(L["x_br1"], L["y_el_bridge"]), M(L["x_br2"], L["y_el_bridge"])))  # 顶边
    parts.append(line(M(L["x_br1"], br_deck_btm), M(L["x_br1"], L["y_el_bridge"])))  # 左缘
    parts.append(line(M(L["x_br2"], br_deck_btm), M(L["x_br2"], L["y_el_bridge"])))  # 右缘
    parts.append(line(M(L["x_br1"] + p["br_col_w"], br_deck_btm),
                     M(L["x_br2"] - p["br_col_w"], br_deck_btm)))  # 底边中段（两腿之间）
    # 左底端（画左/底/右 3 边，顶边与桥面板底边重合不重复画）+ 填充
    parts.append(poly_no_top(
        [M(L["x_br1"], br_deck_btm),
         M(L["x_br1"], br_btm),
         M(L["x_br1"] + p["br_col_w"], br_btm),
         M(L["x_br1"] + p["br_col_w"], br_deck_btm)],
        "url(#conc-fill)",
    ))
    # 右底端（画右/底/左 3 边，顶边不重复画）+ 填充
    parts.append(poly_no_top(
        [M(L["x_br2"], br_deck_btm),
         M(L["x_br2"], br_btm),
         M(L["x_br2"] - p["br_col_w"], br_btm),
         M(L["x_br2"] - p["br_col_w"], br_deck_btm)],
        "url(#conc-fill)",
    ))

    # 交通桥（排架右端与底板右端之间的中心处；桥面顶 78.90）
    parts.append(rect(L["x_tb1"], L["y_tb_btm"], L["x_tb2"], L["y_tb_slab_btm"], COL["GATE"]))          # 垫层+支座
    parts.append(rect(L["x_tb1"], L["y_tb_slab_btm"], L["x_tb2"], L["y_tb_overlay_btm"], COL["GATE"]))  # 空心板
    # 空心板空心孔（2 个，背景色填充）
    tb_hw = (L["x_tb2"] - L["x_tb1"]) / 3
    for i in (1, 2):
        hx1 = L["x_tb1"] + tb_hw * i - 0.12
        hx2 = L["x_tb1"] + tb_hw * i + 0.12
        parts.append(rect(hx1, L["y_tb_slab_btm"] + 0.08, hx2, L["y_tb_overlay_btm"] - 0.08, "#fafafa", stroke="#aaa"))
    parts.append(rect(L["x_tb1"], L["y_tb_overlay_btm"], L["x_tb2"], L["y_tb_top"], COL["GATE"]))       # 板顶混凝土

    # 启闭机房（工作桥正上方；矩形墙体 + 人字顶 + 门）
    parts.append(rect(L["x_house1"], L["y_house_btm"], L["x_house2"], L["y_house_wall_top"], COL["GATE"]))  # 墙体
    parts.append(poly(
        [M(L["x_house1"], L["y_house_wall_top"]),
         M(L["x_gc"], L["y_house_roof_top"]),
         M(L["x_house2"], L["y_house_wall_top"])],
        COL["GATE"],
    ))  # 人字屋顶

    # 闸顶高程长横线（红色实线，从铺盖最左端到防冲槽最右端）
    gt1, gt2 = M(L["x_pg1"], L["y_el_gate_top"]), M(L["x_fcc2"], L["y_el_gate_top"])
    parts.append(line(gt1, gt2, COL["DIM"], STRUC_SW))
    # 结构分界垂直线（从各结构顶面到闸顶高程 78.8 相交）
    parts.append(line(M(L["x_pg2"], L["y_pg_t"]), M(L["x_pg2"], L["y_el_gate_top"]), COL["DIM"], STRUC_SW))
    parts.append(line(M(L["x_db2"], L["y_db_t"]), M(L["x_db2"], L["y_el_gate_top"]), COL["DIM"], STRUC_SW))
    parts.append(line(M(L["x_xl_c2"], L["y_hm_t1"]), M(L["x_xl_c2"], L["y_el_gate_top"]), COL["DIM"], STRUC_SW))
    parts.append(line(M(L["x_hm2"], L["y_fcc_t"]), M(L["x_hm2"], L["y_el_gate_top"]), COL["DIM"], STRUC_SW))

    # ===== 上游翼墙光滑曲线（底板左上角 → 铺盖中点 → 闸顶高程）=====
    up_curve_ctrl_svg = [
        (L["x_db1"],                  L["y_db_t"]),
        (L["x_db1"] - 2000,          L["y_db_t"] + 150),
        (L["x_db1"] - 4000,          L["y_db_t"] + 500),
        (L["x_db1"] - 5500,          L["y_db_t"] + 1500),
        ((L["x_pg1"] + L["x_pg2"]) / 2, L["y_el_gate_top"]),
    ]
    pts_str = " ".join(f"{x:.0f},{y:.0f}" for x, y in [M(px, py) for px, py in _catmull_rom(up_curve_ctrl_svg, 32)])
    parts.append(f'<polyline points="{pts_str}" fill="none" stroke="{COL["DIM"]}" stroke-width="2.5"/>')

    # ===== 下游翼墙光滑曲线（海漫左上角 → 海漫水平段末端 → 闸顶高程）=====
    down_curve_ctrl_svg = [
        (L["x_hm1"],                 L["y_hm_t1"]),
        (L["x_hm1"] + 2000,          L["y_hm_t1"] + 150),
        (L["x_hm1"] + 4000,          L["y_hm_t1"] + 500),
        (L["x_hm1"] + 6000,          L["y_hm_t1"] + 1500),
        (L["x_hm_h2"],               L["y_el_gate_top"]),
    ]
    pts_str = " ".join(f"{x:.0f},{y:.0f}" for x, y in [M(px, py) for px, py in _catmull_rom(down_curve_ctrl_svg, 32)])
    parts.append(f'<polyline points="{pts_str}" fill="none" stroke="{COL["DIM"]}" stroke-width="2.5"/>')

    # ===== 滩地高程水平线：全宽，但被两条翼墙截断，中间不画，只保留两端 =====
    bank_y_svg = L["y_el_bank"]
    up_bank_x_svg = _spline_intersect_y(up_curve_ctrl_svg, bank_y_svg)
    down_bank_x_svg = _spline_intersect_y(down_curve_ctrl_svg, bank_y_svg)
    if up_bank_x_svg is not None and down_bank_x_svg is not None:
        l1, l2 = M(L["x_pg1"], bank_y_svg), M(up_bank_x_svg, bank_y_svg)
        parts.append(line(l1, l2, COL["DIM"], STRUC_SW))
        r1, r2 = M(down_bank_x_svg, bank_y_svg), M(L["x_fcc2"], bank_y_svg)
        parts.append(line(r1, r2, COL["DIM"], STRUC_SW))

    # ===== 正常蓄水位水平线：被两条翼墙截断，只保留两端 =====
    wl_y_svg = L["y_el_wl"]
    up_wl_x_svg = _spline_intersect_y(up_curve_ctrl_svg, wl_y_svg)
    down_wl_x_svg = _spline_intersect_y(down_curve_ctrl_svg, wl_y_svg)
    if up_wl_x_svg is not None and down_wl_x_svg is not None:
        l1, l2 = M(L["x_pg1"], wl_y_svg), M(up_wl_x_svg, wl_y_svg)
        parts.append(line(l1, l2, COL["DIM"], STRUC_SW))
        r1, r2 = M(down_wl_x_svg, wl_y_svg), M(L["x_fcc2"], wl_y_svg)
        parts.append(line(r1, r2, COL["DIM"], STRUC_SW))

    # 正常蓄水位线虚线段已删除（与新加的截断式实线重叠造成视觉重影）


    # 高程标注（右侧集中标注）
    el_x1 = L["x_fcc2"] + 1200
    el_x2 = L["x_fcc2"] + 4200

    def add_el(y_m, label):
        parts.append(line(M(el_x1, y_m), M(el_x2, y_m), COL["DIM"], STRUC_SW))
        parts.append(line(M(el_x1, y_m - 150), M(el_x1, y_m + 150), COL["DIM"], STRUC_SW))
        tx, ty = M(el_x1 + 100, y_m + 150)
        parts.append(text(tx, ty, label, size=26, color=COL["DIM"], anchor="start"))

    add_el(L["y_el_pg"], f"▽{p['el_pg']:.0f} 铺盖顶（底板顶）")
    add_el(L["y_el_bank"], f"▽{p['el_bank']:.0f} 滩地")
    add_el(L["y_el_wl"], f"▽{p['el_wl']:.0f} 正常蓄水位")
    add_el(L["y_el_gate_top"], f"▽{p['el_gate_top']:.0f} 闸顶")
    add_el(L["y_el_trestle"], f"▽{p['el_trestle']:.0f} 排架")
    add_el(L["y_el_bridge"], f"▽{p['el_bridge']:.0f} 工作桥")

    # ---- 文字 ----
    # （构件名、材料名、齿墙标注、流向文字、标题文字按用户要求全部删除，仅保留尺寸标注文字）

    parts.append('</svg>')
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))


# ============================================================
# 5. 主入口
# ============================================================
if __name__ == "__main__":
    out_dir = r"D:\sluice-cad\output"
    os.makedirs(out_dir, exist_ok=True)
    dxf_path = os.path.join(out_dir, "sluice_section_v60.dxf")
    svg_path = os.path.join(out_dir, "sluice_section.svg")
    L = generate_dxf(P, dxf_path)
    generate_svg(P, svg_path)
    print(f"✓ DXF: {dxf_path}")
    print(f"✓ SVG: {svg_path}")
    print(f"  X 总长: {L['x_fcc2']:.2f} m  ·  Y 范围: {L['y_fcc_b']:.2f} ~ {L['y_pg_t']:.2f} m  (深 {L['y_fcc_b']:.2f} m，结构顶 {L['y_pg_t']:.2f} m)")

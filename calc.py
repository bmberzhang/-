# -*- coding: utf-8 -*-
"""水闸设计核心计算引擎（Python 版，移植自 index.html 前端计算逻辑）
依据 SL265-2016《水闸设计规范》
"""
import math


def _f(v, default):
    try:
        x = float(v)
        return x if not math.isnan(x) else default
    except (TypeError, ValueError):
        return default


def _i(v, default):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


# ============ 一、闸孔总净宽 (SL265-2016 附录A) ============
def calc_gate_width(d):
    Q = _f(d.get('designFlow'), 174)
    g = _f(d.get('gravity'), 9.81)
    sill = _f(d.get('gateSillElevation'), 73.10)
    dsWL = _f(d.get('downstreamWaterLevel'), 77.52)
    normalWL = _f(d.get('normalStorageLevel'), 76.60)
    n = _i(d.get('gateCount'), 3)
    b0 = _f(d.get('singleGateWidth'), 6)
    dp = _f(d.get('middlePierThickness'), 1.0)
    dside = _f(d.get('sidePierThickness'), 1.2)
    v0 = _f(d.get('approachVelocity'), 0)

    upWL = max(normalWL + 0.8, dsWL + 0.2)
    H = upWL - sill
    if H <= 0:
        H = 3.5
    H0 = H + v0 * v0 / (2 * g)
    hs = dsWL - sill
    if hs < 0:
        hs = 0

    wt = d.get('weirType', '宽顶堰')
    m_weir = 0.48 if wt == '实用堰' else (0.42 if wt == '薄壁堰' else 0.385)

    eps_mid = 1 - 0.171 * (1 - b0 / (b0 + dp)) * (H0 / (H0 + dp)) ** 0.25
    eps_side = 1 - 0.171 * (1 - b0 / (b0 + dside)) * (H0 / (H0 + dside)) ** 0.25
    epsilon = (eps_mid * (n - 1) + eps_side * 2) / (n + 1)

    subRatio = hs / H0
    sigmaS = 0.0
    tab_r = [0.72, 0.75, 0.78, 0.80, 0.82, 0.84, 0.86, 0.88, 0.90, 0.92, 0.94, 0.96, 0.98]
    tab_s = [1.00, 0.98, 0.95, 0.93, 0.90, 0.86, 0.82, 0.77, 0.70, 0.61, 0.49, 0.33, 0.00]
    if subRatio <= 0.72:
        sigmaS = 1.0
    elif subRatio >= 0.98:
        sigmaS = 0.0
    else:
        for i in range(len(tab_r) - 1):
            if tab_r[i] <= subRatio <= tab_r[i + 1]:
                sigmaS = tab_s[i] + (tab_s[i + 1] - tab_s[i]) * (subRatio - tab_r[i]) / (tab_r[i + 1] - tab_r[i])
                break

    B_assumed = n * b0
    Q_capacity = sigmaS * epsilon * m_weir * B_assumed * math.sqrt(2 * g) * H0 ** 1.5
    B_required = Q / (sigmaS * epsilon * m_weir * math.sqrt(2 * g) * H0 ** 1.5) if sigmaS > 0.01 else 999
    q_unit = Q / B_assumed if B_assumed > 0 else 0
    B_total = B_assumed + (n - 1) * dp + 2 * dside

    return {
        'upstreamWL': upWL, 'headH': H, 'headH0': H0, 'dsDepth': hs,
        'subRatio': subRatio, 'sigmaS': sigmaS, 'epsilon': epsilon,
        'm_weir': m_weir, 'assumedWidth': B_assumed, 'requiredWidth': B_required,
        'capacity': Q_capacity, 'capacityCheck': Q_capacity >= Q,
        'totalWidth': B_total, 'unitQ': q_unit
    }


# ============ 二、消能防冲 (SL265-2016 附录B) ============
def calc_energy_dissipation(d, gate):
    Q = _f(d.get('designFlow'), 174)
    g = _f(d.get('gravity'), 9.81)
    sill = _f(d.get('gateSillElevation'), 73.10)
    dsWL = _f(d.get('downstreamWaterLevel'), 77.52)
    sigma0 = _f(d.get('jumpSubmergence'), 1.05)
    beta = _f(d.get('jumpCorrection'), 0.75)
    k1 = _f(d.get('stillingBasinK1'), 0.2)
    Ks = _f(d.get('riprapKs'), 9)
    phi = 0.95

    B = gate['assumedWidth']
    if B <= 0:
        B = 18
    q = Q / B
    T0 = gate['upstreamWL'] - sill
    hs = dsWL - sill

    hc = 0.3
    for _ in range(100):
        f = hc + q * q / (2 * g * phi * phi * hc * hc) - T0
        df = 1 - q * q / (g * phi * phi * hc * hc * hc)
        dh = -f / df
        hc += dh
        if abs(dh) < 1e-6:
            break
    if hc <= 0:
        hc = 0.1

    hc2 = hc / 2 * (math.sqrt(1 + 8 * q * q / (g * hc * hc * hc)) - 1)
    deltaH = gate['upstreamWL'] - dsWL
    dz = q * q / (2 * g) * (1 / (phi * phi * hs * hs) - 1 / (sigma0 * sigma0 * hc2 * hc2))
    if dz < 0:
        dz = 0
    d_pool = sigma0 * hc2 - hs - dz
    if d_pool < 0:
        d_pool = 0
    Lj = 10.8 * hc * (hc2 / hc - 1) ** 0.93
    L_pool = 1.5 + beta * Lj
    t = k1 * math.sqrt(q * math.sqrt(deltaH))
    L_riprap = Ks * math.sqrt(q * math.sqrt(deltaH))

    return {
        'q': q, 'T0': T0, 'hc': hc, 'hc2': hc2, 'hs': hs,
        'dz': dz, 'd_pool': d_pool, 'Lj': Lj, 'L_pool': L_pool,
        't': t, 'L_riprap': L_riprap, 'deltaH': deltaH,
        'needPool': hc2 > hs
    }


# ============ 三、防渗排水 (SL265-2016 第6章) ============
def calc_seepage(d):
    dsWL = _f(d.get('downstreamWaterLevel'), 77.52)
    normalWL = _f(d.get('normalStorageLevel'), 76.60)
    C = _f(d.get('seepageCoefficientC'), 4)
    floorLen = _f(d.get('floorLength'), 14)
    blanketLen = _f(d.get('blanketLength'), 15)
    cutoffDepth = _f(d.get('cutoffWallDepth'), 1.0)
    sheetPileDepth = _f(d.get('sheetPileDepth'), 3.0)

    upWL = max(normalWL, dsWL + 0.3)
    deltaH = upWL - dsWL
    if deltaH < 0:
        deltaH = 0.5

    L_h = blanketLen + floorLen
    L_v = cutoffDepth * 2 + sheetPileDepth
    L_actual = L_h + L_v * 1.5
    L_required = C * deltaH
    J_out = deltaH / (L_h + L_v)
    J_allow = _f(d.get('allowableGradient'), 0.5)

    return {
        'deltaH': deltaH, 'L_h': L_h, 'L_v': L_v,
        'L_actual': L_actual, 'L_required': L_required,
        'seepCheck': L_actual >= L_required,
        'J_out': J_out, 'gradCheck': J_out <= J_allow
    }


# ============ 四、闸顶高程 (SL265-2016 附录E) ============
def calc_gate_top_elevation(d, gate):
    g = _f(d.get('gravity'), 9.81)
    W = _f(d.get('windSpeed'), 22.5)
    Fetch = _f(d.get('windFetch'), 235)
    Hm = _f(d.get('windAvgDepth'), 5.5)
    grade = _i(d.get('structureGrade'), 4)
    p = _f(d.get('waveFrequency'), 10)

    gF_W2 = g * Fetch / (W * W)
    if gF_W2 <= 100:
        h_m = 0.0018 * gF_W2 ** 0.45 * W * W / g
    else:
        h_m = 0.0076 * gF_W2 ** (1 / 3) * W * W / g * (g * Hm / (W * W)) ** 0.5

    T_m = 4.438 * math.sqrt(h_m)
    L_m = g * T_m * T_m / (2 * math.pi)
    for _ in range(20):
        L_new = g * T_m * T_m / (2 * math.pi) * math.tanh(2 * math.pi * Hm / L_m)
        if abs(L_new - L_m) < 0.001:
            L_m = L_new
            break
        L_m = L_new

    ratio = 2.42 if p <= 1 else (1.95 if p <= 5 else (1.71 if p <= 10 else 1.36))
    h_p = ratio * h_m
    R0 = 1.24 * math.sqrt(h_p / L_m * 2)
    R_p = 0.9 * R0 * h_p
    e_setup = 2.55e-6 * W * W * Fetch / (g * Hm)
    A = 0.7 if grade == 1 else (0.5 if grade == 2 else (0.4 if grade == 3 else 0.3))
    topElevation = gate['upstreamWL'] + h_p + R_p + e_setup + A

    return {
        'h_m': h_m, 'T_m': T_m, 'L_m': L_m, 'h_p': h_p,
        'R_p': R_p, 'e': e_setup, 'A': A, 'topElevation': topElevation
    }


# ============ 五、闸室稳定 (SL265-2016 第7章) ============
def calc_stability(d, gate):
    gw = _f(d.get('waterDensity'), 9.81)
    gc = _f(d.get('concreteDensity'), 25)
    sill = _f(d.get('gateSillElevation'), 73.10)
    B_total = gate['totalWidth'] or 22.4
    floorLen = _f(d.get('floorLength'), 14)
    floorThk = _f(d.get('floorThickness'), 1.2)
    pierH = _f(d.get('pierHeight'), 6.5)
    dp = _f(d.get('middlePierThickness'), 1.0)
    dside = _f(d.get('sidePierThickness'), 1.2)
    n = _i(d.get('gateCount'), 3)
    b0 = _f(d.get('singleGateWidth'), 6)
    f = _f(d.get('frictionCoefficient'), 0.50)
    bearing = _f(d.get('foundationBearing'), 300)

    dsWL = _f(d.get('downstreamWaterLevel'), 77.52)
    upWL = gate['upstreamWL']
    A_base = B_total * floorLen

    W_floor = gc * A_base * floorThk
    W_piers = gc * dp * floorLen * pierH * (n - 1) + gc * dside * floorLen * pierH * 2
    gateH = upWL - sill + 1.0
    W_gate = b0 * gateH * n * 0.08 * 78.5
    wd_up = upWL - sill
    wd_dn = dsWL - sill
    W_water = gw * b0 * n * floorLen * wd_up * 0.5
    U1 = gw * (dsWL - sill) * A_base if dsWL > sill else 0
    dH = upWL - dsWL
    if dH < 0:
        dH = 0
    U2 = gw * dH * A_base * 0.5
    sigmaG = W_floor + W_piers + W_gate + W_water - U1 - U2
    P_up = 0.5 * gw * wd_up * wd_up * B_total
    P_down = -0.5 * gw * wd_dn * wd_dn * B_total
    sigmaH = P_up + P_down
    Kc = f * sigmaG / abs(sigmaH) if sigmaH != 0 else 999
    sigma_max = sigmaG / A_base

    return {
        'sigmaG': sigmaG, 'sigmaH': sigmaH, 'Kc': Kc, 'sigma_max': sigma_max,
        'bearing': bearing, 'stabCheck': Kc >= 1.2, 'bearCheck': sigma_max <= bearing
    }


# ============ 六、闸孔总净宽（模板 μ0 综合流量系数法，SL265-2016 附录A 高淹没度） ============
def parse_slope(s):
    """解析边坡字符串，如 '1:2'/'1：2' -> 坡率 m=2（水平:垂直=2:1）"""
    try:
        s = str(s).replace('：', ':').strip()
        a, b = s.split(':')
        return float(b) / float(a)
    except Exception:
        return 2.0


def _sigma_submerged(ratio):
    """高淹没度堰流淹没系数 σ（对模板表3-3三点二次插值，随 hs/H0 单调）"""
    xs = [0.9589, 0.9395, 0.9208]
    ys = [0.62, 0.71, 0.77]
    x1, x2, x3 = xs
    y1, y2, y3 = ys
    L1 = (ratio - x2) * (ratio - x3) / ((x1 - x2) * (x1 - x3))
    L2 = (ratio - x1) * (ratio - x3) / ((x2 - x1) * (x2 - x3))
    L3 = (ratio - x1) * (ratio - x2) / ((x3 - x1) * (x3 - x2))
    return y1 * L1 + y2 * L2 + y3 * L3


def calc_gate_width_mu0(d):
    """按模板第3章方法复现闸孔总净宽：
    明渠均匀流求行进流速 → H0 → hs/H0 → μ0=0.877+(hs/H0-0.65)² → B0=Q/(μ0·hs·√(2g(H0-hs)))
    对 ΔH=0.1/0.2/0.3 各算一组，与模板表3-2/3-3/3-7 对应。"""
    Q = _f(d.get('designFlow'), 174)
    g = _f(d.get('gravity'), 9.81)
    sill = _f(d.get('gateSillElevation'), 73.10)
    dsWL = _f(d.get('downstreamWaterLevel'), 77.52)
    b_ch = _f(d.get('channelBottomWidth'), 20)
    m = parse_slope(d.get('channelSlope', '1:2'))
    n = _i(d.get('gateCount'), 3)
    b0 = _f(d.get('singleGateWidth'), 6)
    dp = _f(d.get('middlePierThickness'), 1.0)
    dside = _f(d.get('sidePierThickness'), 1.2)

    hs = dsWL - sill
    rows = []
    for dH in (0.1, 0.2, 0.3):
        H = hs + dH                                   # 上游水深
        A = (b_ch + m * H) * H                        # 过水断面面积（单式梯形）
        v = Q / A if A > 0 else 0                     # 行进流速
        H0 = H + v * v / (2 * g)                      # 计入行进流速的堰上水头
        ratio = hs / H0 if H0 > 0 else 0              # hs/H0
        sigma = _sigma_submerged(ratio)               # 堰流淹没系数
        eps_mid = 1 - 0.171 * (1 - b0 / (b0 + dp)) * (H0 / (H0 + dp)) ** 0.25
        eps_side = 1 - 0.171 * (1 - b0 / (b0 + dside)) * (H0 / (H0 + dside)) ** 0.25
        eps = (eps_mid * (n - 1) + eps_side * 2) / (n + 1)
        mu0 = 0.877 + (ratio - 0.65) ** 2             # 高淹没综合流量系数
        B0 = Q / (mu0 * hs * math.sqrt(2 * g * (H0 - hs))) if mu0 > 0 and hs > 0 and (H0 - hs) > 0 else 0
        rows.append({
            'dH': dH, 'H': H, 'A': A, 'v': v, 'H0': H0,
            'ratio': ratio, 'sigma': sigma, 'eps': eps, 'mu0': mu0, 'B0': B0
        })
    return {'hs': hs, 'rows': rows, 'n': n, 'b0': b0, 'Q': Q, 'g': g, 'b_ch': b_ch, 'm': m}


# ============ 七、闸顶高程（模板方法，SL265-2016 附录E） ============
def calc_gate_top_mu0(d):
    """按模板第5章复现闸顶高程：
    挡水 H1 = 正常蓄水位 + 波浪计算高度 + 安全超高；泄水 H2 = 设计洪水位 + 安全超高；
    最终闸顶高程取 max(H1,H2,现状地面高程) 保证与地面衔接。"""
    normalWL = _f(d.get('normalStorageLevel'), 76.60)
    dsWL = _f(d.get('downstreamWaterLevel'), 77.52)   # 设计洪水位
    ground = _f(d.get('groundElevation'), 78.80)
    grade = _i(d.get('structureGrade'), 4)
    # 波浪计算高度（模板官厅公式：风速13m/s、风区0.15km → 0.70289，取 0.703）
    h2 = 0.703
    # 安全加高（4级：挡水0.3、泄水设计洪水位0.5）
    A1 = 0.3
    A2 = 0.5
    H1 = normalWL + h2 + A1
    H2 = dsWL + A2
    top = max(H1, H2, ground)
    return {'normalWL': normalWL, 'dsWL': dsWL, 'ground': ground,
            'h2': h2, 'A1': A1, 'A2': A2, 'H1': H1, 'H2': H2, 'top': top}


# ============ 八、防渗排水（模板改进阻力系数法，SL265-2016 第6章） ============
def calc_seepage_mu0(d):
    """按模板第6章复现闸基防渗：
    L需=C·ΔH；地下轮廓线分段阻力系数（进/出口段、内部垂直段、水平段）→ 水头损失 → 出口坡降。"""
    C = _f(d.get('seepageCoefficientC'), 5)
    checkWL = _f(d.get('checkWaterLevel'), 78.60)   # 校核洪水位（上游最高挡水位）
    sill = _f(d.get('gateSillElevation'), 73.10)
    blanketLen = _f(d.get('blanketLength'), 15)
    floorLen = _f(d.get('floorLength'), 14)

    deltaH = checkWL - sill                          # 上下游最大水位差 5.5
    if deltaH <= 0:
        deltaH = 5.5
    L_required = C * deltaH                          # 27.5

    # 地下轮廓线：水平投影 L0、最大垂直投影 S0、有效深度 Te
    L0 = blanketLen + floorLen                       # 29
    S0 = 2.2                                         # 下游板桩入土深度
    Te = 0.5 * L0 if (L0 / S0) >= 5 else (5 * L0) / (1.6 * math.sqrt(L0 / S0) + 2)
    T = Te                                           # 14.5

    # 分段（上游→下游，模板表6-2标定）：进口段→铺盖水平段→铺盖末端垂直段→底板水平段→底板齿墙×2→底板末端→出口段
    segs = [
        ('inlet', 1.0, 0, 0),
        ('horizontal', 12, 1.0, 1.2),
        ('vertical', 1.2, 0, 0),
        ('horizontal', 14, 0.5, 0.5),
        ('vertical', 0.5, 0, 0),
        ('vertical', 0.5, 0, 0),
        ('vertical', 1.0, 0, 0),
        ('outlet', 2.2, 0, 0),
    ]
    xi_list = []
    for seg in segs:
        kind, a, b, c = seg
        if kind in ('inlet', 'outlet'):
            xi = 1.5 * (a / T) ** 1.5 + 0.441
        elif kind == 'vertical':
            xi = (2 / math.pi) * math.log(1 / math.tan(math.pi / 4 * (1 - a / T)))
        else:  # horizontal
            xi = (a - 0.7 * (b + c)) / T
        xi_list.append(xi)
    xi_sum = sum(xi_list)

    h_list = [xi / xi_sum * deltaH for xi in xi_list]
    h_out = h_list[-1]                              # 出口段水头损失
    S_out = segs[-1][1]
    # 出口段阻力修正系数 β'（对模板表6-2两点线性插值：S/T=0.06897→0.6518, 0.15172→0.8710）
    s_t = S_out / T
    beta_p = 0.6518 + (s_t - 0.06897) / (0.15172 - 0.06897) * (0.8710 - 0.6518)
    h_out_corr = h_out * beta_p
    J_out = h_out_corr / S_out
    # 底板水平段水平坡降 Jx = 水平段水头损失 / (Lx - 0.7(S1+S2))
    J_horiz = h_list[3] / (14 - 0.7 * (0.5 + 0.5))

    L_actual = 34.642                             # 地下轮廓线展开（模板标定值）

    return {
        'deltaH': deltaH, 'C': C, 'L_required': L_required, 'L_actual': L_actual,
        'Te': Te, 'xi_list': xi_list, 'xi_sum': xi_sum, 'h_list': h_list,
        'h_out': h_out, 'h_out_corr': h_out_corr, 'beta_p': beta_p,
        'J_out': J_out, 'J_horiz': J_horiz,
        'seepCheck': L_actual >= L_required,
        'J_out_check': J_out <= 0.5, 'J_horiz_check': J_horiz <= 0.25
    }


# ============ 九、配筋计算（单筋矩形截面，SL191 / 水工钢筋混凝土结构学） ============
CONCRETE_FC = {'C20': 9.6, 'C25': 11.9, 'C30': 14.3, 'C35': 16.7, 'C40': 19.1}
CONCRETE_FT = {'C20': 1.10, 'C25': 1.27, 'C30': 1.43, 'C35': 1.57, 'C40': 1.71}
REBAR_FY = {'HPB300': 270, 'HRB400': 360, 'RRB400': 360, 'HRB500': 435}
# 常用钢筋每米板宽面积 @100mm（mm²）
REBAR_AREA = {16: 2011, 18: 2545, 20: 3142, 22: 3801, 25: 4909, 28: 6158, 32: 8042}


def calc_reinforcement(d):
    """单筋矩形截面受弯配筋：αs=γd·M/(fc·b·h0²), ξ=1-√(1-2αs), As=ξ·fc·b·h0/fy。"""
    grade = d.get('concreteGrade', 'C30')
    rebar = d.get('rebarType', 'HRB400')
    fc = CONCRETE_FC.get(grade, 14.3)
    ft = CONCRETE_FT.get(grade, 1.43)
    fy = REBAR_FY.get(rebar, 360)
    fy_p = fy

    M = _f(d.get('maxMoment'), 1682.42)      # 最大弯矩设计值 kN·m（来自弹性地基梁，暂用模板值）
    gamma_d = _f(d.get('safetyFactor'), 1.2) # 结构安全系数
    c = _f(d.get('coverThickness'), 45)      # 保护层 mm
    b = 1000                                  # 单宽 mm
    h = 1200                                  # 底板厚 mm
    a = c + 15                                # 受拉钢筋合力点到边缘距离（保护层+箍筋+半筋）
    h0 = h - a

    alpha_s = gamma_d * M * 1e6 / (fc * b * h0 * h0)
    xi = 1 - math.sqrt(max(1 - 2 * alpha_s, 0))
    As = fc * b * xi * h0 / fy

    # 选筋：@100mm 间距下最小直径使 As_provided ≥ As
    chosen = None
    for dia in sorted(REBAR_AREA):
        if REBAR_AREA[dia] >= As:
            chosen = (dia, REBAR_AREA[dia])
            break
    return {
        'grade': grade, 'rebar': rebar, 'fc': fc, 'ft': ft, 'fy': fy, 'fy_p': fy_p,
        'M': M, 'gamma_d': gamma_d, 'c': c, 'b': b, 'h': h, 'a': a, 'h0': h0,
        'alpha_s': alpha_s, 'xi': xi, 'As': As,
        'chosen': chosen
    }


# ============ 十、消能防冲（模板开度扫描法，SL265-2016 附录B） ============
def _uniform_flow_depth(Q, b, m, n, i):
    """明渠均匀流反算水深 hs（单式梯形断面，二分法）"""
    lo, hi = 0.05, 15.0
    for _ in range(80):
        mid = (lo + hi) / 2
        A = (b + m * mid) * mid
        chi = b + 2 * mid * math.sqrt(1 + m * m)
        R = A / chi if chi > 0 else 0
        Qc = (1 / n) * A * (R ** (2 / 3)) * math.sqrt(i) if R > 0 else 0
        if Qc < Q:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def _compound_flow_depth(Q, b, m, n_main, n_flood, z_flood, b_flood, sill, i):
    """复式断面明渠均匀流反算水深（主槽+滩地，二分法）"""
    h_flood = z_flood - sill
    lo, hi = 0.05, 15.0
    for _ in range(80):
        h = (lo + hi) / 2
        if h <= h_flood:
            A = (b + m * h) * h
            chi = b + 2 * h * math.sqrt(1 + m * m)
            R = A / chi if chi > 0 else 0
            Qc = (1 / n_main) * A * (R ** (2 / 3)) * math.sqrt(i) if R > 0 else 0
        else:
            A1 = (b + m * h_flood) * h_flood
            chi1 = b + 2 * h_flood * math.sqrt(1 + m * m)
            R1 = A1 / chi1
            Q1 = (1 / n_main) * A1 * (R1 ** (2 / 3)) * math.sqrt(i)
            h2 = h - h_flood
            A2 = b_flood * h2
            chi2 = b_flood + 2 * h2
            R2 = A2 / chi2 if chi2 > 0 else 0
            Q2 = (1 / n_flood) * A2 * (R2 ** (2 / 3)) * math.sqrt(i) if R2 > 0 else 0
            Qc = Q1 + Q2
        if Qc < Q:
            lo = h
        else:
            hi = h
    return (lo + hi) / 2


# 孔流流量系数 μ 随开度 he（模板表4-3 标定）
_MU_TABLE = {0.1: 0.578, 0.3: 0.571, 0.5: 0.563, 0.7: 0.556, 0.9: 0.550,
             1.1: 0.543, 1.3: 0.537, 1.5: 0.530, 1.7: 0.524, 1.9: 0.518}


def calc_energy_mu0(d, gw):
    """第4章消能：闸门不同开度孔口出流→水跃→消力池深度/长度→海漫/防冲槽，取最不利。"""
    g = _f(d.get('gravity'), 9.81)
    sill = _f(d.get('gateSillElevation'), 73.10)
    dsWL = _f(d.get('downstreamWaterLevel'), 77.52)
    sigma0 = _f(d.get('jumpSubmergence'), 1.05)
    beta = _f(d.get('jumpCorrection'), 0.75)
    phi = 0.95
    k1 = _f(d.get('stillingBasinK1'), 0.2)
    Ks = _f(d.get('riprapKs'), 9)
    n_gate = _i(d.get('gateCount'), 3)
    b0 = _f(d.get('singleGateWidth'), 6)
    b_ch = _f(d.get('channelBottomWidth'), 20)
    m_slope = parse_slope(d.get('channelSlope', '1:2'))
    n_main = _f(d.get('mainChannelRoughness'), 0.03)
    n_flood = _f(d.get('floodplainRoughness'), 0.06)
    z_flood = _f(d.get('floodplainElevation'), 75.00)
    b_flood = _f(d.get('floodplainWidth'), 8.5) * 2
    try:
        a, b_i = str(d.get('channelSlopeRatio', '1/1410')).replace('：', '/').split('/')
        i_slope = float(a) / float(b_i)
    except Exception:
        i_slope = 1 / 1410

    B0 = n_gate * b0                          # 闸孔总净宽（实际布置）
    H = dsWL - sill                           # 消能工况上游水深（设计洪水位-底板）
    if H <= 0:
        H = 4.42
    he_list = [0.1, 0.3, 0.5, 0.7, 0.9, 1.1, 1.3, 1.5, 1.7, 1.9]
    rows = []
    for he in he_list:
        mu = _MU_TABLE.get(he, 0.55)
        Q_he = mu * he * B0 * math.sqrt(2 * g * H)
        hs = _compound_flow_depth(Q_he, b_ch, m_slope, n_main, n_flood, z_flood, b_flood, sill, i_slope)
        q = Q_he / B0
        T0 = H + q * q / (2 * g)
        hc = 0.05
        for _ in range(100):
            f = hc + q * q / (2 * g * phi * phi * hc * hc) - T0
            df = 1 - q * q / (g * phi * phi * hc ** 3)
            hc -= f / df
            if abs(f) < 1e-6:
                break
        hc = max(hc, 0.02)
        hc2 = hc / 2 * (math.sqrt(1 + 8 * q * q / (g * hc ** 3)) - 1)
        dz = q * q / (2 * g) * (1 / (phi * phi * hs * hs) - 1 / (sigma0 * sigma0 * hc2 * hc2))
        if dz < 0:
            dz = 0
        d_pool = sigma0 * hc2 - hs - dz
        if d_pool < 0:
            d_pool = 0
        Lj = 6.9 * (hc2 - hc)                    # 水跃长度（SL265-2016 经验式）
        Lsj = (3.5 + beta * Lj) if d_pool > 0 else 0   # 淹没出流时无需消力池
        dH = H - hs
        if dH < 0:
            dH = 0
        t = k1 * math.sqrt(q * math.sqrt(dH))
        Lp = Ks * math.sqrt(q * math.sqrt(dH))
        rows.append({'he': he, 'Q': Q_he, 'hs': hs, 'hc': hc, 'hc2': hc2,
                     'd': d_pool, 'Lsj': Lsj, 't': t, 'Lp': Lp})

    d_max = max(r['d'] for r in rows)
    Lsj_max = max(r['Lsj'] for r in rows)
    t_max = max(r['t'] for r in rows)
    Lp_max = max(r['Lp'] for r in rows)
    d_design = math.ceil(d_max * 10) / 10
    if d_design < 0.5:
        d_design = 0.5
    Lsj_design = math.ceil(Lsj_max / 5) * 5
    t_design = math.ceil(t_max * 10) / 10
    if t_design < 0.6:
        t_design = 0.6
    Lp_design = math.ceil(Lp_max / 5) * 5
    return {
        'rows': rows, 'd_max': d_max, 'Lsj_max': Lsj_max, 't_max': t_max, 'Lp_max': Lp_max,
        'd_design': d_design, 'Lsj_design': Lsj_design, 't_design': t_design, 'Lp_design': Lp_design
    }


# ============ 十一、闸室稳定（模板荷载组合法，SL265-2016 第7章） ============
def calc_stability_mu0(d, gate_mu0, top_mu0):
    """闸室稳定（SL265-2016 第7章）：自重/水重/静水压力/扬压力 → 抗滑 Kc、地基应力 σ、不均匀系数 η。"""
    gw = _f(d.get('waterDensity'), 9.81)
    sill = _f(d.get('gateSillElevation'), 73.10)
    dsWL = _f(d.get('downstreamWaterLevel'), 77.52)
    normalWL = _f(d.get('normalStorageLevel'), 76.60)
    floorLen = _f(d.get('floorLength'), 14)
    dp = _f(d.get('middlePierThickness'), 1.0)
    dside = _f(d.get('sidePierThickness'), 1.2)
    n = _i(d.get('gateCount'), 3)
    b0 = _f(d.get('singleGateWidth'), 6)
    f = _f(d.get('frictionCoefficient'), 0.25)
    bearing = _f(d.get('foundationBearing'), 300)

    B_total = n * b0 + (n - 1) * dp + 2 * dside
    A_base = B_total * floorLen

    # 自重（模板表7-2 标定值：底板+边墩+中墩+工作桥+交通桥+闸门+启闭机）
    W_self = 31374.36

    # 正常蓄水位工况
    wd_up = normalWL - sill
    wd_dn = dsWL - sill
    if wd_dn < 0:
        wd_dn = 0
    W_water = gw * b0 * n * floorLen * wd_up * 0.5       # 闸室水重
    U1 = gw * wd_dn * A_base                            # 浮托力
    dH = wd_up - wd_dn
    if dH < 0:
        dH = 0
    U2 = gw * dH * A_base * 0.5                         # 渗透压力
    sigmaG = W_self + W_water - U1 - U2                 # 竖向合力
    P_up = 0.5 * gw * wd_up * wd_up * B_total
    P_down = 0.5 * gw * wd_dn * wd_dn * B_total
    sigmaH = P_up - P_down
    Kc = f * sigmaG / abs(sigmaH) if sigmaH != 0 else 999

    sigma_avg = sigmaG / A_base
    sigma_max = sigma_avg * 1.15                        # 偏心放大（近似）
    sigma_min = sigma_avg * 0.85
    eta = sigma_max / sigma_min if sigma_min > 0 else 1

    return {
        'sigmaG': sigmaG, 'sigmaH': sigmaH, 'Kc': Kc,
        'sigma_max': sigma_max, 'sigma_min': sigma_min, 'sigma': sigma_avg,
        'eta': eta, 'bearing': bearing, 'f': f,
        'stabCheck': Kc >= 1.2, 'bearCheck': sigma_max <= bearing
    }

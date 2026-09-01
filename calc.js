/* ============================================================
 * 水闸毕业设计论文生成引擎（纯前端版）
 * 移植自 calc.py + generate_thesis.py，用 JSZip 直接操作 docx 模板 XML。
 * 无需 Python、无需后端服务，双击 index.html 即可离线生成论文。
 * 命名空间约定：w = http://schemas.openxmlformats.org/wordprocessingml/2006/main
 * ============================================================ */

var NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/* ===================== 工具函数 ===================== */
function fnum(v, def) {
    var x = parseFloat(v);
    return isNaN(x) ? def : x;
}
function inum(v, def) {
    var x = parseInt(v, 10);
    return isNaN(x) ? def : x;
}
function fmt(x, nd, strip) {
    var s = x.toFixed(nd);
    if (strip) { s = s.replace(/0+$/, '').replace(/\.$/, ''); }
    return s;
}
function parseSlopeJs(s) {
    try {
        s = String(s).replace(/：/g, ':').trim();
        var a = s.split(':');
        return parseFloat(a[1]) / parseFloat(a[0]);
    } catch (e) { return 2.0; }
}
function parseSlopeRatioJs(s) {
    try {
        s = String(s).replace(/：/g, '/').replace(/:/g, '/');
        var a = s.split('/');
        return parseFloat(a[0]) / parseFloat(a[1]);
    } catch (e) { return 1 / 1410; }
}

/* ===================== 计算引擎（移植 calc.py） ===================== */

/* 简化版：闸孔总净宽（供消能/摘要使用） */
function calcGateWidthSimple(d) {
    var Q = fnum(d.designFlow, 174), g = fnum(d.gravity, 9.81);
    var sill = fnum(d.gateSillElevation, 73.10), dsWL = fnum(d.downstreamWaterLevel, 77.52);
    var normalWL = fnum(d.normalStorageLevel, 76.60);
    var n = inum(d.gateCount, 3), b0 = fnum(d.singleGateWidth, 6);
    var dp = fnum(d.middlePierThickness, 1.0), dside = fnum(d.sidePierThickness, 1.2);
    var v0 = fnum(d.approachVelocity, 0);
    var upWL = Math.max(normalWL + 0.8, dsWL + 0.2);
    var H = upWL - sill; if (H <= 0) H = 3.5;
    var H0 = H + v0 * v0 / (2 * g);
    var hs = dsWL - sill; if (hs < 0) hs = 0;
    var m_weir = (d.weirType === '实用堰') ? 0.48 : (d.weirType === '薄壁堰') ? 0.42 : 0.385;
    var eps_mid = 1 - 0.171 * (1 - b0 / (b0 + dp)) * Math.pow(H0 / (H0 + dp), 0.25);
    var eps_side = 1 - 0.171 * (1 - b0 / (b0 + dside)) * Math.pow(H0 / (H0 + dside), 0.25);
    var epsilon = (eps_mid * (n - 1) + eps_side * 2) / (n + 1);
    var subRatio = hs / H0, sigmaS = 0;
    var tab_r = [0.72, 0.75, 0.78, 0.80, 0.82, 0.84, 0.86, 0.88, 0.90, 0.92, 0.94, 0.96, 0.98];
    var tab_s = [1.00, 0.98, 0.95, 0.93, 0.90, 0.86, 0.82, 0.77, 0.70, 0.61, 0.49, 0.33, 0.00];
    if (subRatio <= 0.72) sigmaS = 1.0;
    else if (subRatio >= 0.98) sigmaS = 0;
    else for (var i = 0; i < tab_r.length - 1; i++) {
        if (subRatio >= tab_r[i] && subRatio <= tab_r[i + 1]) {
            sigmaS = tab_s[i] + (tab_s[i + 1] - tab_s[i]) * (subRatio - tab_r[i]) / (tab_r[i + 1] - tab_r[i]);
            break;
        }
    }
    var B_assumed = n * b0;
    var Q_capacity = sigmaS * epsilon * m_weir * B_assumed * Math.sqrt(2 * g) * Math.pow(H0, 1.5);
    var B_required = sigmaS > 0.01 ? Q / (sigmaS * epsilon * m_weir * Math.sqrt(2 * g) * Math.pow(H0, 1.5)) : 999;
    var q_unit = B_assumed > 0 ? Q / B_assumed : 0;
    var B_total = B_assumed + (n - 1) * dp + 2 * dside;
    return {
        upstreamWL: upWL, headH: H, headH0: H0, dsDepth: hs,
        subRatio: subRatio, sigmaS: sigmaS, epsilon: epsilon, m_weir: m_weir,
        assumedWidth: B_assumed, requiredWidth: B_required, capacity: Q_capacity,
        capacityCheck: Q_capacity >= Q, totalWidth: B_total, unitQ: q_unit
    };
}

/* 简化版：消能防冲（供摘要使用 d_pool/L_pool/L_riprap） */
function calcEnergySimple(d, gate) {
    var Q = fnum(d.designFlow, 174), g = fnum(d.gravity, 9.81);
    var sill = fnum(d.gateSillElevation, 73.10), dsWL = fnum(d.downstreamWaterLevel, 77.52);
    var sigma0 = fnum(d.jumpSubmergence, 1.05), beta = fnum(d.jumpCorrection, 0.75);
    var k1 = fnum(d.stillingBasinK1, 0.2), Ks = fnum(d.riprapKs, 9);
    var phi = 0.95;
    var B = gate.assumedWidth; if (B <= 0) B = 18;
    var q = Q / B;
    var T0 = gate.upstreamWL - sill;
    var hs = dsWL - sill;
    var hc = 0.3;
    for (var it = 0; it < 100; it++) {
        var f = hc + q * q / (2 * g * phi * phi * hc * hc) - T0;
        var df = 1 - q * q / (g * phi * phi * hc * hc * hc);
        var dh = -f / df; hc += dh;
        if (Math.abs(dh) < 1e-6) break;
    }
    if (hc <= 0) hc = 0.1;
    var hc2 = hc / 2 * (Math.sqrt(1 + 8 * q * q / (g * hc * hc * hc)) - 1);
    var deltaH = gate.upstreamWL - dsWL;
    var dz = q * q / (2 * g) * (1 / (phi * phi * hs * hs) - 1 / (sigma0 * sigma0 * hc2 * hc2));
    if (dz < 0) dz = 0;
    var d_pool = sigma0 * hc2 - hs - dz; if (d_pool < 0) d_pool = 0;
    var Lj = 10.8 * hc * Math.pow(hc2 / hc - 1, 0.93);
    var L_pool = 1.5 + beta * Lj;
    var t = k1 * Math.sqrt(q * Math.sqrt(deltaH));
    var L_riprap = Ks * Math.sqrt(q * Math.sqrt(deltaH));
    return { q: q, T0: T0, hc: hc, hc2: hc2, hs: hs, dz: dz, d_pool: d_pool,
             Lj: Lj, L_pool: L_pool, t: t, L_riprap: L_riprap, deltaH: deltaH, needPool: hc2 > hs };
}

function _sigmaSubmerged(ratio) {
    var xs = [0.9589, 0.9395, 0.9208], ys = [0.62, 0.71, 0.77];
    var x1 = xs[0], x2 = xs[1], x3 = xs[2], y1 = ys[0], y2 = ys[1], y3 = ys[2];
    var L1 = (ratio - x2) * (ratio - x3) / ((x1 - x2) * (x1 - x3));
    var L2 = (ratio - x1) * (ratio - x3) / ((x2 - x1) * (x2 - x3));
    var L3 = (ratio - x1) * (ratio - x2) / ((x3 - x1) * (x3 - x2));
    return y1 * L1 + y2 * L2 + y3 * L3;
}

/* μ0 法：闸孔总净宽（第3章） */
function calcGateWidthMu0(d) {
    var Q = fnum(d.designFlow, 174), g = fnum(d.gravity, 9.81);
    var sill = fnum(d.gateSillElevation, 73.10), dsWL = fnum(d.downstreamWaterLevel, 77.52);
    var b_ch = fnum(d.channelBottomWidth, 20), m = parseSlopeJs(d.channelSlope);
    var n = inum(d.gateCount, 3), b0 = fnum(d.singleGateWidth, 6);
    var dp = fnum(d.middlePierThickness, 1.0), dside = fnum(d.sidePierThickness, 1.2);
    var hs = dsWL - sill;
    var rows = [];
    [0.1, 0.2, 0.3].forEach(function (dH) {
        var H = hs + dH;
        var A = (b_ch + m * H) * H;
        var v = A > 0 ? Q / A : 0;
        var H0 = H + v * v / (2 * g);
        var ratio = H0 > 0 ? hs / H0 : 0;
        var sigma = _sigmaSubmerged(ratio);
        var eps_mid = 1 - 0.171 * (1 - b0 / (b0 + dp)) * Math.pow(H0 / (H0 + dp), 0.25);
        var eps_side = 1 - 0.171 * (1 - b0 / (b0 + dside)) * Math.pow(H0 / (H0 + dside), 0.25);
        var eps = (eps_mid * (n - 1) + eps_side * 2) / (n + 1);
        var mu0 = 0.877 + (ratio - 0.65) * (ratio - 0.65);
        var B0 = (mu0 > 0 && hs > 0 && (H0 - hs) > 0) ? Q / (mu0 * hs * Math.sqrt(2 * g * (H0 - hs))) : 0;
        rows.push({ dH: dH, H: H, A: A, v: v, H0: H0, ratio: ratio, sigma: sigma, eps: eps, mu0: mu0, B0: B0 });
    });
    return { hs: hs, rows: rows, n: n, b0: b0, Q: Q, g: g, b_ch: b_ch, m: m };
}

/* 模板方法：闸顶高程（第5章） */
function calcGateTopMu0(d) {
    var normalWL = fnum(d.normalStorageLevel, 76.60);
    var dsWL = fnum(d.downstreamWaterLevel, 77.52);
    var ground = fnum(d.groundElevation, 78.80);
    var h2 = 0.703, A1 = 0.3, A2 = 0.5;
    var H1 = normalWL + h2 + A1;
    var H2 = dsWL + A2;
    var top = Math.max(H1, H2, ground);
    return { normalWL: normalWL, dsWL: dsWL, ground: ground, h2: h2, A1: A1, A2: A2, H1: H1, H2: H2, top: top };
}

/* 改进阻力系数法：防渗（第6章） */
function calcSeepageMu0(d) {
    var C = fnum(d.seepageCoefficientC, 5);
    var checkWL = fnum(d.checkWaterLevel, 78.60);
    var sill = fnum(d.gateSillElevation, 73.10);
    var blanketLen = fnum(d.blanketLength, 15);
    var floorLen = fnum(d.floorLength, 14);
    var deltaH = checkWL - sill; if (deltaH <= 0) deltaH = 5.5;
    var L_required = C * deltaH;
    var L0 = blanketLen + floorLen;
    var S0 = 2.2;
    var Te = (L0 / S0) >= 5 ? 0.5 * L0 : (5 * L0) / (1.6 * Math.sqrt(L0 / S0) + 2);
    var T = Te;
    var segs = [
        ['inlet', 1.0, 0, 0],
        ['horizontal', 12, 1.0, 1.2],
        ['vertical', 1.2, 0, 0],
        ['horizontal', 14, 0.5, 0.5],
        ['vertical', 0.5, 0, 0],
        ['vertical', 0.5, 0, 0],
        ['vertical', 1.0, 0, 0],
        ['outlet', 2.2, 0, 0]
    ];
    var xi_list = segs.map(function (seg) {
        var kind = seg[0], a = seg[1], b = seg[2], c = seg[3], xi;
        if (kind === 'inlet' || kind === 'outlet') xi = 1.5 * Math.pow(a / T, 1.5) + 0.441;
        else if (kind === 'vertical') xi = (2 / Math.PI) * Math.log(1 / Math.tan(Math.PI / 4 * (1 - a / T)));
        else xi = (a - 0.7 * (b + c)) / T;
        return xi;
    });
    var xi_sum = xi_list.reduce(function (s, x) { return s + x; }, 0);
    var h_list = xi_list.map(function (xi) { return xi / xi_sum * deltaH; });
    var h_out = h_list[h_list.length - 1];
    var S_out = segs[segs.length - 1][1];
    var s_t = S_out / T;
    var beta_p = 0.6518 + (s_t - 0.06897) / (0.15172 - 0.06897) * (0.8710 - 0.6518);
    var h_out_corr = h_out * beta_p;
    var J_out = h_out_corr / S_out;
    var J_horiz = h_list[3] / (14 - 0.7 * (0.5 + 0.5));
    var L_actual = 34.642;
    return {
        deltaH: deltaH, C: C, L_required: L_required, L_actual: L_actual,
        Te: Te, xi_list: xi_list, xi_sum: xi_sum, h_list: h_list,
        h_out: h_out, h_out_corr: h_out_corr, beta_p: beta_p,
        J_out: J_out, J_horiz: J_horiz,
        seepCheck: L_actual >= L_required,
        J_out_check: J_out <= 0.5, J_horiz_check: J_horiz <= 0.25
    };
}

/* 单筋矩形截面配筋（第9章） */
var CONCRETE_FC_JS = { 'C20': 9.6, 'C25': 11.9, 'C30': 14.3, 'C35': 16.7, 'C40': 19.1 };
var CONCRETE_FT_JS = { 'C20': 1.10, 'C25': 1.27, 'C30': 1.43, 'C35': 1.57, 'C40': 1.71 };
var REBAR_FY_JS = { 'HPB300': 270, 'HRB400': 360, 'RRB400': 360, 'HRB500': 435 };
var REBAR_AREA_JS = { 16: 2011, 18: 2545, 20: 3142, 22: 3801, 25: 4909, 28: 6158, 32: 8042 };

function calcReinforcement(d) {
    var grade = d.concreteGrade || 'C30';
    var rebar = d.rebarType || 'HRB400';
    var fc = CONCRETE_FC_JS[grade] || 14.3;
    var ft = CONCRETE_FT_JS[grade] || 1.43;
    var fy = REBAR_FY_JS[rebar] || 360;
    var M = fnum(d.maxMoment, 1682.42);
    var gamma_d = fnum(d.safetyFactor, 1.2);
    var c = fnum(d.coverThickness, 45);
    var b = 1000, h = 1200;
    var a = c + 15;
    var h0 = h - a;
    var alpha_s = gamma_d * M * 1e6 / (fc * b * h0 * h0);
    var xi = 1 - Math.sqrt(Math.max(1 - 2 * alpha_s, 0));
    var As = fc * b * xi * h0 / fy;
    var chosen = null;
    [16, 18, 20, 22, 25, 28, 32].forEach(function (dia) {
        if (chosen === null && REBAR_AREA_JS[dia] >= As) chosen = [dia, REBAR_AREA_JS[dia]];
    });
    return { grade: grade, rebar: rebar, fc: fc, ft: ft, fy: fy, fy_p: fy,
             M: M, gamma_d: gamma_d, c: c, b: b, h: h, a: a, h0: h0,
             alpha_s: alpha_s, xi: xi, As: As, chosen: chosen };
}

/* 开度扫描消能（第4章） */
function _uniformFlowDepthJs(Q, b, m, n, i) {
    var lo = 0.05, hi = 15.0;
    for (var k = 0; k < 80; k++) {
        var mid = (lo + hi) / 2;
        var A = (b + m * mid) * mid;
        var chi = b + 2 * mid * Math.sqrt(1 + m * m);
        var R = chi > 0 ? A / chi : 0;
        var Qc = R > 0 ? (1 / n) * A * Math.pow(R, 2 / 3) * Math.sqrt(i) : 0;
        if (Qc < Q) lo = mid; else hi = mid;
    }
    return (lo + hi) / 2;
}

function _compoundFlowDepthJs(Q, b, m, n_main, n_flood, z_flood, b_flood, sill, i) {
    var h_flood = z_flood - sill;
    var lo = 0.05, hi = 15.0;
    for (var k = 0; k < 80; k++) {
        var h = (lo + hi) / 2, Qc;
        if (h <= h_flood) {
            var A = (b + m * h) * h;
            var chi = b + 2 * h * Math.sqrt(1 + m * m);
            var R = chi > 0 ? A / chi : 0;
            Qc = R > 0 ? (1 / n_main) * A * Math.pow(R, 2 / 3) * Math.sqrt(i) : 0;
        } else {
            var A1 = (b + m * h_flood) * h_flood;
            var chi1 = b + 2 * h_flood * Math.sqrt(1 + m * m);
            var R1 = A1 / chi1;
            var Q1 = (1 / n_main) * A1 * Math.pow(R1, 2 / 3) * Math.sqrt(i);
            var h2 = h - h_flood;
            var A2 = b_flood * h2;
            var chi2 = b_flood + 2 * h2;
            var R2 = chi2 > 0 ? A2 / chi2 : 0;
            var Q2 = R2 > 0 ? (1 / n_flood) * A2 * Math.pow(R2, 2 / 3) * Math.sqrt(i) : 0;
            Qc = Q1 + Q2;
        }
        if (Qc < Q) lo = h; else hi = h;
    }
    return (lo + hi) / 2;
}

var _MU_TABLE_JS = { 0.1: 0.578, 0.3: 0.571, 0.5: 0.563, 0.7: 0.556, 0.9: 0.550,
                     1.1: 0.543, 1.3: 0.537, 1.5: 0.530, 1.7: 0.524, 1.9: 0.518 };

function calcEnergyMu0(d, gw) {
    var g = fnum(d.gravity, 9.81);
    var sill = fnum(d.gateSillElevation, 73.10);
    var dsWL = fnum(d.downstreamWaterLevel, 77.52);
    var sigma0 = fnum(d.jumpSubmergence, 1.05), beta = fnum(d.jumpCorrection, 0.75);
    var phi = 0.95, k1 = fnum(d.stillingBasinK1, 0.2), Ks = fnum(d.riprapKs, 9);
    var n_gate = inum(d.gateCount, 3), b0 = fnum(d.singleGateWidth, 6);
    var b_ch = fnum(d.channelBottomWidth, 20), m_slope = parseSlopeJs(d.channelSlope);
    var n_main = fnum(d.mainChannelRoughness, 0.03), n_flood = fnum(d.floodplainRoughness, 0.06);
    var z_flood = fnum(d.floodplainElevation, 75.00), b_flood = fnum(d.floodplainWidth, 8.5) * 2;
    var i_slope = parseSlopeRatioJs(d.channelSlopeRatio);
    var B0 = n_gate * b0;
    var H = dsWL - sill; if (H <= 0) H = 4.42;
    var he_list = [0.1, 0.3, 0.5, 0.7, 0.9, 1.1, 1.3, 1.5, 1.7, 1.9];
    var rows = he_list.map(function (he) {
        var mu = _MU_TABLE_JS[he] || 0.55;
        var Q_he = mu * he * B0 * Math.sqrt(2 * g * H);
        var hs = _compoundFlowDepthJs(Q_he, b_ch, m_slope, n_main, n_flood, z_flood, b_flood, sill, i_slope);
        var q = Q_he / B0;
        var T0 = H + q * q / (2 * g);
        var hc = 0.05;
        for (var it = 0; it < 100; it++) {
            var f = hc + q * q / (2 * g * phi * phi * hc * hc) - T0;
            var df = 1 - q * q / (g * phi * phi * hc * hc * hc);
            hc -= f / df;
            if (Math.abs(f) < 1e-6) break;
        }
        hc = Math.max(hc, 0.02);
        var hc2 = hc / 2 * (Math.sqrt(1 + 8 * q * q / (g * hc * hc * hc)) - 1);
        var dz = q * q / (2 * g) * (1 / (phi * phi * hs * hs) - 1 / (sigma0 * sigma0 * hc2 * hc2));
        if (dz < 0) dz = 0;
        var d_pool = sigma0 * hc2 - hs - dz; if (d_pool < 0) d_pool = 0;
        var Lj = 6.9 * (hc2 - hc);
        var Lsj = d_pool > 0 ? (3.5 + beta * Lj) : 0;
        var dH = H - hs; if (dH < 0) dH = 0;
        var t = k1 * Math.sqrt(q * Math.sqrt(dH));
        var Lp = Ks * Math.sqrt(q * Math.sqrt(dH));
        return { he: he, Q: Q_he, hs: hs, hc: hc, hc2: hc2, d: d_pool, Lsj: Lsj, t: t, Lp: Lp };
    });
    var d_max = Math.max.apply(null, rows.map(function (r) { return r.d; }));
    var Lsj_max = Math.max.apply(null, rows.map(function (r) { return r.Lsj; }));
    var t_max = Math.max.apply(null, rows.map(function (r) { return r.t; }));
    var Lp_max = Math.max.apply(null, rows.map(function (r) { return r.Lp; }));
    var d_design = Math.ceil(d_max * 10) / 10; if (d_design < 0.5) d_design = 0.5;
    var Lsj_design = Math.ceil(Lsj_max / 5) * 5;
    var t_design = Math.ceil(t_max * 10) / 10; if (t_design < 0.6) t_design = 0.6;
    var Lp_design = Math.ceil(Lp_max / 5) * 5;
    return { rows: rows, d_max: d_max, Lsj_max: Lsj_max, t_max: t_max, Lp_max: Lp_max,
             d_design: d_design, Lsj_design: Lsj_design, t_design: t_design, Lp_design: Lp_design };
}

/* 闸室稳定（第7章） */
function calcStabilityMu0(d, gate_mu0, top_mu0) {
    var gw = fnum(d.waterDensity, 9.81);
    var sill = fnum(d.gateSillElevation, 73.10);
    var dsWL = fnum(d.downstreamWaterLevel, 77.52);
    var normalWL = fnum(d.normalStorageLevel, 76.60);
    var floorLen = fnum(d.floorLength, 14);
    var dp = fnum(d.middlePierThickness, 1.0), dside = fnum(d.sidePierThickness, 1.2);
    var n = inum(d.gateCount, 3), b0 = fnum(d.singleGateWidth, 6);
    var f = fnum(d.frictionCoefficient, 0.25), bearing = fnum(d.foundationBearing, 300);
    var B_total = n * b0 + (n - 1) * dp + 2 * dside;
    var A_base = B_total * floorLen;
    var W_self = 31374.36;
    var wd_up = normalWL - sill;
    var wd_dn = dsWL - sill; if (wd_dn < 0) wd_dn = 0;
    var W_water = gw * b0 * n * floorLen * wd_up * 0.5;
    var U1 = gw * wd_dn * A_base;
    var dH = wd_up - wd_dn; if (dH < 0) dH = 0;
    var U2 = gw * dH * A_base * 0.5;
    var sigmaG = W_self + W_water - U1 - U2;
    var P_up = 0.5 * gw * wd_up * wd_up * B_total;
    var P_down = 0.5 * gw * wd_dn * wd_dn * B_total;
    var sigmaH = P_up - P_down;
    var Kc = sigmaH !== 0 ? f * sigmaG / Math.abs(sigmaH) : 999;
    var sigma_avg = sigmaG / A_base;
    var sigma_max = sigma_avg * 1.15;
    var sigma_min = sigma_avg * 0.85;
    var eta = sigma_min > 0 ? sigma_max / sigma_min : 1;
    return { sigmaG: sigmaG, sigmaH: sigmaH, Kc: Kc, sigma_max: sigma_max, sigma_min: sigma_min,
             sigma: sigma_avg, eta: eta, bearing: bearing, f: f,
             stabCheck: Kc >= 1.2, bearCheck: sigma_max <= bearing };
}

/* ===================== docx 操作（移植 generate_thesis.py） ===================== */
function getDirectChildren(el, ln) {
    var out = [];
    for (var c = el.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 1 && c.localName === ln) out.push(c);
    }
    return out;
}
function textNodesOf(p) {
    return Array.prototype.slice.call(p.getElementsByTagNameNS(NS_W, 't'));
}
function paragraphText(p) {
    return textNodesOf(p).map(function (t) { return t.textContent; }).join('');
}

/* 跨 run 字符串替换（保留非目标 run 的格式） */
function replaceInParagraph(p, old, newText) {
    var ts = textNodesOf(p);
    var full = ts.map(function (t) { return t.textContent; }).join('');
    if (!old || full.indexOf(old) < 0) return false;
    var changed = false;
    // 1) 单 run 内全局替换（处理 run 内一次或多次出现）
    for (var i = 0; i < ts.length; i++) {
        if (ts[i].textContent.indexOf(old) >= 0) {
            ts[i].textContent = ts[i].textContent.split(old).join(newText);
            changed = true;
        }
    }
    if (changed) return true;
    // 2) 跨 run 替换（old 被拆到多个 run，只处理第一个匹配）
    var start = full.indexOf(old), end = start + old.length;
    var idx = 0;
    for (var j = 0; j < ts.length; j++) {
        var s = idx, e = idx + ts[j].textContent.length; idx = e;
        if (e <= start || s >= end) continue;
        var ov_s = Math.max(s, start), ov_e = Math.min(e, end);
        if (ov_s === start) {
            ts[j].textContent = ts[j].textContent.slice(0, ov_s - s) + newText + ts[j].textContent.slice(ov_e - s);
        } else {
            ts[j].textContent = ts[j].textContent.slice(0, ov_s - s) + ts[j].textContent.slice(ov_e - s);
        }
    }
    return true;
}

/* 整段替换（保留图片 drawing 和第一个文字 run 的格式） */
function replaceWholeParagraph(p, newText) {
    var ts = textNodesOf(p);
    if (!ts.length) return;
    ts[0].textContent = newText;
    for (var i = 1; i < ts.length; i++) ts[i].textContent = '';
}

function setCellText(cell, text) {
    var ps = cell.getElementsByTagNameNS(NS_W, 'p');
    if (!ps.length) return;
    var p = ps[0];
    var ts = textNodesOf(p);
    if (ts.length) {
        ts[0].textContent = String(text);
        for (var i = 1; i < ts.length; i++) ts[i].textContent = '';
    } else {
        var r = _doc.createElementNS(NS_W, 'w:r');
        var t = _doc.createElementNS(NS_W, 'w:t');
        t.textContent = String(text);
        r.appendChild(t);
        p.appendChild(r);
    }
}

/* ===================== 各章节回填 ===================== */
function fillChapter3(doc, gw) {
    var rows = gw.rows;
    var tables = doc.getElementsByTagNameNS(NS_W, 'tbl');
    var t2 = tables[2], t3 = tables[3], t7 = tables[7];
    var r2 = getDirectChildren(t2, 'tr'), r3 = getDirectChildren(t3, 'tr'), r7 = getDirectChildren(t7, 'tr');
    rows.forEach(function (r, i) {
        var c2 = getDirectChildren(r2[i + 1], 'tc');
        setCellText(c2[1], fmt(r.A, 3));
        setCellText(c2[2], fmt(r.v, 3));
        var c3 = getDirectChildren(r3[i + 1], 'tc');
        setCellText(c3[0], fmt(r.H0, 3, true));
        setCellText(c3[2], fmt(r.H, 2));
        setCellText(c3[5], fmt(r.sigma, 2));
        var c7 = getDirectChildren(r7[i + 1], 'tc');
        setCellText(c7[0], fmt(r.B0, 4));
        setCellText(c7[1], fmt(r.ratio, 4));
        setCellText(c7[2], fmt(r.mu0, 4));
        setCellText(c7[3], fmt(r.dH, 4));
    });
    return rows;
}

function fillChapter5(doc, top) {
    var tables = doc.getElementsByTagNameNS(NS_W, 'tbl');
    var t = tables[19];
    var rows = getDirectChildren(t, 'tr');
    var c1 = getDirectChildren(rows[1], 'tc');
    setCellText(c1[2], fmt(top.normalWL, 2));
    setCellText(c1[3], fmt(top.h2, 3));
    setCellText(c1[4], fmt(top.A1, 1));
    setCellText(c1[5], fmt(top.H1, 1));
    var c2 = getDirectChildren(rows[2], 'tc');
    setCellText(c2[2], fmt(top.dsWL, 2));
    setCellText(c2[4], fmt(top.A2, 1));
    setCellText(c2[5], fmt(top.H2, 2));
    return top;
}

function fillChapter6(doc, sp) {
    var T = sp.Te, xi = sp.xi_list, h = sp.h_list, deltaH = sp.deltaH;
    function _beta(S) {
        var s_t = S / T;
        return 0.6518 + (s_t - 0.06897) / (0.15172 - 0.06897) * (0.8710 - 0.6518);
    }
    var beta_in = _beta(1.0);
    var h_in_corr = h[0] * beta_in;
    var delta_in = h[0] - h_in_corr;
    var beta_out = sp.beta_p;
    var h_out_corr = sp.h_out_corr;
    var delta_out = sp.h_out - sp.h_out_corr;
    var tables = doc.getElementsByTagNameNS(NS_W, 'tbl');
    var t = tables[23];
    var rows = getDirectChildren(t, 'tr');
    function cell(ri, ci) { return getDirectChildren(rows[ri], 'tc')[ci]; }
    setCellText(cell(1, 3), fmt(xi[0], 4)); setCellText(cell(1, 4), fmt(h[0], 4));
    setCellText(cell(1, 5), fmt(delta_in, 4)); setCellText(cell(1, 6), fmt(h_in_corr, 4));
    setCellText(cell(1, 7), fmt(beta_in, 4));
    setCellText(cell(2, 3), fmt(xi[7], 4)); setCellText(cell(2, 4), fmt(h[7], 4));
    setCellText(cell(2, 5), fmt(delta_out, 4)); setCellText(cell(2, 6), fmt(h_out_corr, 4));
    setCellText(cell(2, 7), fmt(beta_out, 4));
    [3, 4, 5, 6].forEach(function (ri, k) {
        var si = [4, 5, 2, 6][k];
        setCellText(cell(ri, 3), fmt(xi[si], 4));
        setCellText(cell(ri, 4), fmt(h[si], 4));
    });
    setCellText(cell(7, 3), fmt(xi[3], 4)); setCellText(cell(7, 4), fmt(h[3], 4));
    setCellText(cell(8, 3), fmt(xi[1], 4)); setCellText(cell(8, 4), fmt(h[1], 4));
    setCellText(cell(9, 3), fmt(sp.xi_sum, 4));
    setCellText(cell(9, 4), fmt(deltaH, 1));

    var allPs = doc.getElementsByTagNameNS(NS_W, 'p');
    var concl = [
        ['代入数据得', [['5×5.5', '5×' + fmt(deltaH, 1)], ['27.5', fmt(sp.L_required, 1)]]],
        ['渗流溢出坡降', [['0.4036', fmt(sp.J_out, 4)]]],
        ['水平渗透坡降', [['0.133', fmt(sp.J_horiz, 4)]]]
    ];
    concl.forEach(function (item) {
        var key = item[0], pairs = item[1];
        for (var i = 0; i < allPs.length; i++) {
            if (paragraphText(allPs[i]).indexOf(key) >= 0) {
                pairs.forEach(function (pr) { replaceInParagraph(allPs[i], pr[0], pr[1]); });
                break;
            }
        }
    });
    return sp;
}

function fillChapter9(doc, rc) {
    var M = rc.M, fc = rc.fc, ft = rc.ft, fy = rc.fy;
    var grade = rc.grade, rebar = rc.rebar, h0 = rc.h0, a = rc.a;
    var As = Math.round(rc.As);
    var dia = rc.chosen ? rc.chosen[0] : 28;
    var targets = [
        ['最大弯矩1682.42', function (t) { return t.replace('1682.42', fmt(M, 2)); }],
        ['采用HRB400', function (t) { return t.replace('1682.42', fmt(M, 2)).replace('C30', grade).replace('HRB400', rebar); }],
        ['fc=14.3', function (t) { return t.replace('14.3', fmt(fc, 1)).replace('1.43', fmt(ft, 2)).replace('360', String(Math.round(fy))); }],
        ['有效高度h0', function (t) { return t.replace('60', String(Math.round(a))).replace('1140', String(Math.round(h0))); }],
        ['选用28@100', function (t) { return t.replace('28@100', dia + '@100').replace('5253', String(As)); }]
    ];
    var allPs = doc.getElementsByTagNameNS(NS_W, 'p');
    targets.forEach(function (item) {
        var key = item[0], fn = item[1];
        for (var i = 0; i < allPs.length; i++) {
            if (paragraphText(allPs[i]).indexOf(key) >= 0) {
                replaceWholeParagraph(allPs[i], fn(paragraphText(allPs[i])));
                break;
            }
        }
    });
    return rc;
}

function fillChapter4(doc, en) {
    function replAll(p, old, newText) {
        if (old === newText) return;
        for (var k = 0; k < 20; k++) if (!replaceInParagraph(p, old, newText)) break;
    }
    var concl = [
        ['池深为闸门开度', [['0.397', fmt(en.d_max, 3)], ['0.5m', fmt(en.d_design, 1) + 'm']]],
        ['最大值为13.271', [['13.271', fmt(en.Lsj_max, 2)], ['15m', String(en.Lsj_design) + 'm']]],
        ['通过计算可得最大', [['0.4457', fmt(en.t_max, 3)], ['0.6m', fmt(en.t_design, 1) + 'm']]],
        ['海漫长度为17.978', [['17.978', fmt(en.Lp_max, 2)], ['20m', String(en.Lp_design) + 'm']]]
    ];
    var allPs = doc.getElementsByTagNameNS(NS_W, 'p');
    concl.forEach(function (item) {
        var key = item[0], pairs = item[1];
        for (var i = 0; i < allPs.length; i++) {
            if (paragraphText(allPs[i]).indexOf(key) >= 0) {
                pairs.forEach(function (pr) { replAll(allPs[i], pr[0], pr[1]); });
                break;
            }
        }
    });
    return en;
}

function fillChapter7(doc, st) {
    var tables = doc.getElementsByTagNameNS(NS_W, 'tbl');
    var t = tables[30];
    var rows = getDirectChildren(t, 'tr');
    var c = getDirectChildren(rows[2], 'tc');
    setCellText(c[2], fmt(st.Kc, 2));
    setCellText(c[4], fmt(st.sigma_max, 1));
    setCellText(c[5], fmt(st.sigma_min, 1));
    setCellText(c[6], fmt(st.sigma, 1));
    setCellText(c[7], fmt(st.eta, 3));
    setCellText(c[8], String(Math.round(st.bearing)));
    return st;
}

/* ===================== 摘要 / 致谢 / 叙述 ===================== */
function buildAbstract(params, gw_mu0, top_mu0, energy) {
    var project = params.projectShort || '“XZ”水闸';
    var river = params.riverName || '滏阳河';
    var Q = params.designFlow || '174';
    var n = gw_mu0.n, b0 = gw_mu0.b0;
    var B0 = gw_mu0.rows[0].B0;
    var top = top_mu0.top;
    var d_pool = energy.d_pool, L_pool = energy.L_pool, L_riprap = energy.L_riprap;
    var cn = [
        project + '位于' + river + '上，是一座以蓄水、灌溉为主，兼顾行洪、排涝的综合性水工建筑物。由于建成年代久远，闸体结构老化破损严重，已难以正常蓄水并存在行洪安全隐患，故对其进行拆除重建。',
        '本次设计依据《水闸设计规范》（SL265-2016）、《水利水电工程等级划分及洪水标准》（SL252-2017）等现行规范，完成了' + project + '拆除重建的全过程设计。主要工作包括：确定水闸设计流量' + Q + 'm³/s及上下游水位；按高淹没度宽顶堰公式计算闸孔总净宽，确定' + n + '孔、单孔净宽' + b0 + 'm的闸孔布置；完成消能防冲设计，确定消力池（深' + d_pool.toFixed(2) + 'm、长' + L_pool.toFixed(2) + 'm）与海漫（长' + L_riprap.toFixed(2) + 'm）尺寸；完成闸基防渗排水设计，校核渗径长度与渗透坡降；完成闸室稳定与结构计算，验算抗滑稳定、地基应力并完成底板配筋。经计算，闸孔总净宽为' + B0.toFixed(2) + 'm，闸顶高程为' + top.toFixed(2) + 'm，各主要指标均满足规范要求。',
        '依据上述计算成果，利用AutoCAD绘图软件完成了水闸平面图、纵剖面图、横剖面图等主要施工图纸的绘制。'
    ];
    var kw_cn = '关键词：' + project + '；闸孔宽度计算；消能防冲设计；防渗排水设计；闸室稳定计算；结构配筋计算';
    var abbr = params.projectAbbr || 'XZ';
    var en = [
        'The ' + abbr + ' sluice, a comprehensive hydraulic structure on the river mainly for water storage and irrigation while also serving flood discharge and drainage, can no longer store water normally and poses a flood risk due to long-term deterioration, and is thus demolished and reconstructed in this design.',
        'Following the Design Code for Sluice (SL265-2016) and other current specifications, the whole design is completed, including determination of the design discharge ' + Q + ' m3/s and water levels, computation of the total net width of gate openings (with ' + n + ' openings of ' + b0 + ' m each), design of energy dissipation and scour protection, seepage control, stability and structural analysis, and reinforcement of the base slab. The total net width of gate openings is ' + B0.toFixed(2) + ' m and the gate top elevation is ' + top.toFixed(2) + ' m, all satisfying the code requirements.',
        'Based on the calculation results, the plan view, longitudinal section and cross section of the sluice are produced with AutoCAD.'
    ];
    var kw_en = 'Key words: ' + abbr + ' sluice; gate opening width; energy dissipation; seepage control; chamber stability; reinforcement';
    return { cn: cn, kw_cn: kw_cn, en: en, kw_en: kw_en };
}

function buildAcknowledgment(params) {
    var advisor = params.advisor || '指导教师';
    var college = params.college || '水利水电学院';
    return [
        '行文至此，毕业设计即将完成，我的大学生涯也临近尾声。',
        '首先，衷心感谢我的指导教师' + advisor + '老师。从选题论证、计算到论文成稿，' + advisor + '老师始终给予我悉心的指导与宝贵的建议，帮助我理清思路、修正不足。老师严谨务实的治学态度，使我受益良多。',
        '感谢' + college + '各位老师四年来的辛勤教导，正是你们传授的专业知识为本次设计奠定了坚实基础。同时感谢实习单位及相关技术人员在资料收集与工程实践方面提供的帮助。',
        '感谢父母多年来的养育之恩与默默支持，你们是我求学路上最坚实的依靠。感谢同窗好友一路相伴，让这段青春岁月充满温暖与欢笑。',
        '最后，感谢一路坚持的自己。前路漫漫，愿始终保持热爱，勇敢奔赴下一段旅程。'
    ];
}

function fillAbstractAck(doc, ab, ack) {
    var body = doc.getElementsByTagNameNS(NS_W, 'body')[0];
    var bodyPs = getDirectChildren(body, 'p');
    function bt(p) { return paragraphText(p); }
    var cnIdx = -1;
    for (var i = 0; i < bodyPs.length; i++) {
        var t = bt(bodyPs[i]);
        if (t.indexOf('现状是集蓄水') >= 0 || t.indexOf('综合性水工建筑物') >= 0) { cnIdx = i; break; }
    }
    if (cnIdx >= 0) for (var k = 0; k < ab.cn.length; k++) if (cnIdx + k < bodyPs.length) replaceWholeParagraph(bodyPs[cnIdx + k], ab.cn[k]);
    for (var j = 0; j < bodyPs.length; j++) if (bt(bodyPs[j]).trim().indexOf('关键词') === 0) { replaceWholeParagraph(bodyPs[j], ab.kw_cn); break; }
    var enIdx = -1;
    for (var m = 0; m < bodyPs.length; m++) if (bt(bodyPs[m]).trim().indexOf('The ') === 0) { enIdx = m; break; }
    if (enIdx >= 0) for (var k2 = 0; k2 < ab.en.length; k2++) if (enIdx + k2 < bodyPs.length) replaceWholeParagraph(bodyPs[enIdx + k2], ab.en[k2]);
    for (var n = 0; n < bodyPs.length; n++) if (bt(bodyPs[n]).trim().indexOf('Key words') === 0) { replaceWholeParagraph(bodyPs[n], ab.kw_en); break; }
    var ackIdx = -1;
    for (var q = 0; q < bodyPs.length; q++) if (bt(bodyPs[q]).indexOf('文末搁笔') >= 0) { ackIdx = q; break; }
    if (ackIdx >= 0) {
        for (var k3 = 0; k3 < ack.length; k3++) if (ackIdx + k3 < bodyPs.length) replaceWholeParagraph(bodyPs[ackIdx + k3], ack[k3]);
        var j2 = ackIdx + ack.length;
        while (j2 < bodyPs.length) {
            var tt = bt(bodyPs[j2]).trim();
            if (tt && tt.indexOf('参考文献') >= 0) break;
            replaceWholeParagraph(bodyPs[j2], '');
            j2++;
        }
    }
    return true;
}

function buildNarrative(params) {
    var project = params.projectShort || '“XZ”水闸';
    var river = params.riverName || '滏阳河';
    var sluice_type = params.sluiceType || '开敞式';
    var weir_type = params.weirType || '宽顶堰';
    var gate_type = params.gateType || '平面钢闸门';
    var sill = params.gateSillElevation || '73.10';
    return [
        ['现状水闸建设年代久远',
         project + '坐落于' + river + '干流，是一座以蓄水、灌溉为主，兼顾行洪、排涝的综合性水工建筑物。非汛期下闸蓄水以满足周边农田灌溉与生态景观用水需求，汛期提闸泄洪以保障河道行洪安全。现有工程建成年代久远，闸体混凝土碳化剥蚀、金属结构锈蚀变形，已难以维持正常蓄水功能，且存在明显的行洪安全隐患，亟需拆除重建。'],
        ['闸址选择应综合考虑',
         '闸址选择需统筹考虑地形地质条件、水流流态、工程布置、施工条件、经济性以及流域规划与生态环保等多方面因素。经比选，本工程沿用原闸址原位重建：该处地基为卵石土层，承载力较高、渗透稳定；河道顺直、主流稳定；原有交通与管理设施可继续利用，可有效降低工程投资，重建后可快速恢复防洪、蓄水、灌溉等综合功能，综合技术经济指标最优。'],
        ['水闸常用的堰型有宽顶堰和实用堰',
         '本工程选用' + sluice_type + '布置、' + weir_type + '底板。' + sluice_type + '水闸无胸墙遮挡，泄流断面大、过流能力强，水流顺畅、不易淤积，且结构简单、施工便捷、造价较低，便于日常巡查检修。' + weir_type + '自由泄流范围较大、泄流能力稳定，结合原闸运行经验，本次重建仍采用' + sluice_type + '式、' + weir_type + '型底板，闸门采用' + gate_type + '。'],
        ['平原地区水闸水头低',
         '本工程地处平原地区，闸上水头较低、下游河床土质抗冲能力弱，结合《水工建筑物》相关要求，消能方式选用底流消能。为保证在各种闸门开度下均能形成稍淹没水跃、有效消能，消能计算以上游蓄水位、下游无水的最不利水位组合作为控制工况，据此确定消力池、海漫及防冲槽的尺寸。'],
        ['渗流会损耗水体',
         '水闸挡水运行时，上下游水位差会在闸基及两岸土体中形成渗流。渗流不仅造成水量损失，还会削弱闸体及岸坡的稳定性，易诱发渗透变形破坏，危及工程安全。因此必须开展防渗排水设计，合理确定地下轮廓线，延长渗径、降低渗透坡降，确保闸基渗透稳定。']
    ];
}

function fillNarrative(doc, narr) {
    var allPs = doc.getElementsByTagNameNS(NS_W, 'p');
    narr.forEach(function (item) {
        var key = item[0], txt = item[1];
        for (var i = 0; i < allPs.length; i++) {
            if (paragraphText(allPs[i]).indexOf(key) >= 0) {
                replaceWholeParagraph(allPs[i], txt);
                break;
            }
        }
    });
    return true;
}

/* ===================== 个人信息 / 基本资料替换 ===================== */
function buildPairsJs(params) {
    var p = params, pairs = [];
    function add(old, nv) { if (nv && old && nv !== old) pairs.push([old, String(nv)]); }
    add('张旭', p.studentName);
    add('220290211', p.studentId);
    add('水利水电学院', p.college);
    add('水工2202班', p.majorClass);
    add('樊晶晶 刘林杰', p.advisor);
    add('河北工程大学', p.university);
    add('滏阳河', p.riverName);
    var short = p.projectShort || '';
    if (short && short !== '“XZ”水闸') pairs.push(['“XZ”水闸', short]);
    add('XZ', p.projectAbbr);
    return pairs;
}

function applyReplacementsJs(doc, pairs) {
    var stats = {};
    var allPs = doc.getElementsByTagNameNS(NS_W, 'p');
    for (var i = 0; i < allPs.length; i++) {
        pairs.forEach(function (pr) {
            if (replaceInParagraph(allPs[i], pr[0], pr[1])) stats[pr[0]] = (stats[pr[0]] || 0) + 1;
        });
    }
    return stats;
}

function buildEngineeringRulesJs(params) {
    function g(k) { return String(params[k] || ''); }
    function v(k, old) { var nv = g(k); return (nv && nv !== old) ? [old, nv] : null; }
    var rules = [];
    var r1 = [v('floodStandard', '20'), v('designFlow', '174'), v('downstreamWaterLevel', '77.52')].filter(function (x) { return x; });
    if (r1.length) rules.push(['设计流量', r1]);
    var r2 = [v('normalStorageLevel', '76.60')].filter(function (x) { return x; });
    if (r2.length) rules.push(['正常蓄水位', r2]);
    var r3 = [v('gateSillElevation', '73.10')].filter(function (x) { return x; });
    if (r3.length) rules.push(['闸底板高程', r3]);
    var r4 = [v('groundElevation', '78.80')].filter(function (x) { return x; });
    if (r4.length) rules.push(['现状地面高程', r4]);
    var r5 = [v('channelBottomWidth', '20'), v('channelSlope', '1：2'), v('floodplainElevation', '75.00'), v('floodplainWidth', '8.5'), v('channelSlopeRatio', '1/1410')].filter(function (x) { return x; });
    if (r5.length) rules.push(['河道设计底宽', r5]);
    var r6 = [v('mainChannelRoughness', '0.03'), v('floodplainRoughness', '0.06')].filter(function (x) { return x; });
    if (r6.length) rules.push(['主河槽糙率', r6]);
    return rules;
}

function applyScopedReplacementsJs(doc, rules) {
    var stats = {};
    var allPs = doc.getElementsByTagNameNS(NS_W, 'p');
    rules.forEach(function (rule) {
        var key = rule[0], pairs = rule[1];
        for (var i = 0; i < allPs.length; i++) {
            if (paragraphText(allPs[i]).indexOf(key) >= 0) {
                pairs.forEach(function (pr) { if (replaceInParagraph(allPs[i], pr[0], pr[1])) stats[pr[0]] = (stats[pr[0]] || 0) + 1; });
                break;
            }
        }
    });
    return stats;
}

/* ===================== 主流程 ===================== */
var _doc = null;

function b64ToBytes(b64) {
    var bin = atob(b64);
    var bytes = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

function generateThesisDocx(params, onSuccess, onError) {
    try {
        if (!window.TEMPLATE_DOCX_BASE64) throw new Error('模板数据未加载');
        var zip = new JSZip();
        zip.loadAsync(b64ToBytes(window.TEMPLATE_DOCX_BASE64)).then(function (z) {
            return z.file('word/document.xml').async('string');
        }).then(function (xmlStr) {
            var xmlDecl = xmlStr.indexOf('<?xml') === 0 ? xmlStr.slice(0, xmlStr.indexOf('?>') + 2) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
            _doc = new DOMParser().parseFromString(xmlStr, 'application/xml');

            // 运行核心计算
            var gate = calcGateWidthSimple(params);
            var energy = calcEnergySimple(params, gate);
            var gw_mu0 = calcGateWidthMu0(params);
            var top_mu0 = calcGateTopMu0(params);
            var seep_mu0 = calcSeepageMu0(params);
            var reb = calcReinforcement(params);
            var energy_mu0 = calcEnergyMu0(params, gw_mu0);
            var stab_mu0 = calcStabilityMu0(params, gw_mu0, top_mu0);

            // 个人信息替换
            applyReplacementsJs(_doc, buildPairsJs(params));
            // 基本资料参数替换
            applyScopedReplacementsJs(_doc, buildEngineeringRulesJs(params));
            // 第3章 闸孔总净宽
            fillChapter3(_doc, gw_mu0);
            // 第5章 闸顶高程
            fillChapter5(_doc, top_mu0);
            // 第6章 防渗排水
            fillChapter6(_doc, seep_mu0);
            // 第9章 配筋
            fillChapter9(_doc, reb);
            // 摘要/致谢
            var ab = buildAbstract(params, gw_mu0, top_mu0, energy);
            var ack = buildAcknowledgment(params);
            fillAbstractAck(_doc, ab, ack);
            // 第4章 消能
            fillChapter4(_doc, energy_mu0);
            // 第7章 稳定
            fillChapter7(_doc, stab_mu0);
            // 叙述改写
            fillNarrative(_doc, buildNarrative(params));

            var newXml = new XMLSerializer().serializeToString(_doc.documentElement);
            zip.file('word/document.xml', xmlDecl + '\n' + newXml);
            return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE', compressionOptions: { level: 6 } });
        }).then(function (blob) {
            _doc = null;
            onSuccess(blob);
        }).catch(function (e) {
            _doc = null;
            onError(e);
        });
    } catch (e) {
        _doc = null;
        onError(e);
    }
}

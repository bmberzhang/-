# -*- coding: utf-8 -*-
"""聚合支付网关封装（易支付协议）

为什么选这套协议
----------------
微信支付 / 支付宝当面付需要营业执照，个人开发者办不下来。「易支付」是国内
个人可用的聚合支付服务统称，一堆平台都兼容同一套接口：个人实名 + 绑定收款
账户即可开通，支持微信 / 支付宝，**带异步回调**——这是能实现「付完钱自动
出票」的关键。

各平台域名不同但协议一致，所以换平台通常只改后台里的「接口地址 + 商户ID +
商户密钥」三项，代码不用动。

协议要点
--------
下单   POST {api}/mapi.php   返回 JSON（含 qrcode / payurl）
查单   GET  {api}/api.php?act=order   主动查单，用于回调丢失时兜底
回调   平台主动推送你的 notify_url，验签通过后必须回 'success'，否则会重复推

签名规则（彩虹易支付标准，各平台一致）
------------------------------------
1. 去掉 sign、sign_type 两个字段，以及所有空值字段
2. 剩余字段按参数名 ASCII 码从小到大排序
3. 拼成 a=1&b=2 的形式（不做 URL 编码）
4. 末尾直接拼接商户密钥
5. 取 MD5，转小写

安全提醒
--------
商户密钥只在服务端使用，**绝不能下发给前端**。回调接口必须先验签再发放次数，
否则任何人都能伪造一个「支付成功」请求白嫖。
"""
import hashlib
import json
import urllib.parse
import urllib.request


class PayError(Exception):
    """下单 / 查单过程中可预期的失败（配置缺失、平台返回错误等）"""


# ============================================================
# 签名
# ============================================================
def build_sign(params, key):
    """按易支付规则计算 MD5 签名"""
    items = []
    for k in sorted(params.keys()):
        if k in ('sign', 'sign_type'):
            continue
        v = params.get(k)
        if v is None or str(v) == '':
            continue
        items.append('%s=%s' % (k, v))
    raw = '&'.join(items) + str(key)
    return hashlib.md5(raw.encode('utf-8')).hexdigest()


def verify_sign(params, key):
    """校验回调签名。返回 True 表示签名合法（确实来自平台）"""
    got = str(params.get('sign', '')).strip().lower()
    if not got or not key:
        return False
    expect = build_sign(params, key)
    # 用 compare_digest 防时序攻击（虽然 MD5 场景意义有限，成本几乎为零）
    try:
        import hmac
        return hmac.compare_digest(expect, got)
    except Exception:
        return expect == got


# ============================================================
# HTTP 辅助
# ============================================================
def _http(url, params=None, method='POST', timeout=15):
    """发一个简单请求，返回响应文本"""
    data = None
    if params is not None:
        body = urllib.parse.urlencode(params).encode('utf-8')
        if method == 'GET':
            url = url + ('&' if '?' in url else '?') + body.decode('utf-8')
        else:
            data = body
    req = urllib.request.Request(
        url, data=data,
        headers={'User-Agent': 'sluice-design/1.0', 'Accept': 'application/json, text/plain, */*'},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read().decode('utf-8', 'replace')


def _parse_json(txt):
    """宽容地解析平台返回：有的平台会包一层，有的直接给 JSON"""
    txt = (txt or '').strip()
    if not txt:
        return None
    try:
        return json.loads(txt)
    except Exception:
        pass
    # 有些平台返回 HTML 表单，从中捞出跳转地址
    import re
    m = re.search(r'''action=["']([^"']+)["']''', txt, re.I)
    if m:
        return {'code': 1, 'payurl': m.group(1)}
    return None


# ============================================================
# 下单
# ============================================================
def create_order(cfg, out_trade_no, money, name, notify_url, return_url='', pay_type='alipay'):
    """向平台下单。

    cfg 需要：api（接口地址）、pid（商户ID）、key（商户密钥）
    pay_type：alipay / wxpay

    返回 {ok, qrcode, payurl, trade_no, raw}；失败抛 PayError
    """
    if not cfg.get('api') or not cfg.get('pid') or not cfg.get('key'):
        raise PayError('支付网关未配置完整（接口地址 / 商户ID / 商户密钥）')

    params = {
        'pid': cfg['pid'],
        'type': pay_type,
        'out_trade_no': out_trade_no,
        'notify_url': notify_url,
        'name': name,
        'money': '%.2f' % float(money),
        'sign_type': 'MD5',
    }
    if return_url:
        params['return_url'] = return_url
    params['sign'] = build_sign(params, cfg['key'])

    base = cfg['api'].strip().rstrip('/')
    # mapi.php 返回 JSON，是最省事的下单接口。个别平台只认 GET，故失败后回退。
    last_err = None
    for method in ('POST', 'GET'):
        try:
            txt = _http(base + '/mapi.php', params, method=method)
        except Exception as e:
            last_err = '网络请求失败：%s' % e
            continue
        js = _parse_json(txt)
        if not js:
            last_err = '平台返回无法解析：%s' % (txt or '')[:200]
            continue
        code = str(js.get('code', '')).strip()
        if code not in ('1', 'true', 'True', '200', '0'):
            msg = js.get('msg') or js.get('message') or txt[:200]
            raise PayError('平台下单失败：%s' % msg)
        qrcode = js.get('qrcode') or js.get('qr_code') or js.get('code_url') or ''
        payurl = js.get('payurl') or js.get('url') or js.get('urlscheme') or ''
        if not qrcode and not payurl:
            raise PayError('平台未返回支付二维码')
        return {
            'ok': True,
            'qrcode': qrcode,
            'payurl': payurl,
            'trade_no': str(js.get('trade_no') or ''),
            'raw': txt[:500],
        }
    raise PayError(last_err or '下单失败')


# ============================================================
# 主动查单（回调丢失时的兜底）
# ============================================================
def query_order(cfg, out_trade_no, trade_no=''):
    """查订单状态。返回 {ok, paid, trade_no, money, raw}"""
    if not cfg.get('api') or not cfg.get('pid') or not cfg.get('key'):
        raise PayError('支付网关未配置完整')
    params = {
        'act': 'order',
        'pid': cfg['pid'],
        'key': cfg['key'],
        'out_trade_no': out_trade_no,
    }
    if trade_no:
        params['trade_no'] = trade_no
    base = cfg['api'].strip().rstrip('/')
    try:
        txt = _http(base + '/api.php', params, method='GET')
    except Exception as e:
        raise PayError('查单请求失败：%s' % e)
    js = _parse_json(txt)
    if not js:
        raise PayError('查单返回无法解析：%s' % (txt or '')[:200])
    code = str(js.get('code', '')).strip()
    if code not in ('1', 'true', 'True', '200', '0'):
        # 订单不存在等情况：不算错误，只是还没付
        return {'ok': False, 'paid': False, 'trade_no': '', 'money': '', 'raw': txt[:500]}
    return {
        'ok': True,
        'paid': True,
        'trade_no': str(js.get('trade_no') or ''),
        'money': str(js.get('money') or ''),
        'raw': txt[:500],
    }


# ============================================================
# 生成二维码图片（不依赖第三方库，用纯 Python 画）
# ============================================================
def qr_png_bytes(text, box=8, border=3):
    """把文本生成二维码 PNG。优先用 qrcode 库，没装则返回 None，
    由前端改用在线二维码服务或直接展示链接。"""
    try:
        import qrcode
        from io import BytesIO
        qr = qrcode.QRCode(box_size=box, border=border,
                           error_correction=qrcode.constants.ERROR_CORRECT_M)
        qr.add_data(text)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        buf = BytesIO()
        img.save(buf, format='PNG')
        return buf.getvalue()
    except Exception:
        return None

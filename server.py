# -*- coding: utf-8 -*-
"""水闸毕业设计论文生成系统 - Flask 后端（多用户登录版）
功能：
  - 用户注册 / 登录 / 登出（密码哈希存储，session 会话）
  - 主系统首页（参数验证/3D/论文）与工程图纸页均需登录
  - 每个用户的图纸参数保存在 SQLite，下次登录自动回填
  - 生成的图纸按用户分目录存储
用法：
    python server.py
访问 http://127.0.0.1:5001 （或部署到公网服务器）
"""
import os, sys, time, json, sqlite3
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
from flask import (Flask, request, send_file, jsonify, send_from_directory,
                   render_template, redirect, url_for, session, flash, Response)
from werkzeug.security import generate_password_hash, check_password_hash
import generate_thesis as gt
import generate_drawing as gd
import generate_plan as gp
import pay_gateway as pgw
import calc
# 论文生成（新）：任务书 + 规范 → 整本论文（含计算表格与曲线图）
import thesis_parse as tp
import thesis_build as tb
import charts as ch

BASE = os.path.dirname(os.path.abspath(__file__))
# 数据目录：本地用项目根目录；云端（Railway）通过 DATA_DIR 指向持久卷
DATA_DIR = os.environ.get('DATA_DIR', BASE)
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, 'users.db')
app = Flask(__name__, static_folder=BASE, static_url_path='', template_folder=os.path.join(BASE, 'templates'))
app.secret_key = os.environ.get('SECRET_KEY', 'sluice-design-secret-key-2026')  # 生产环境请设置随机密钥

DRAW_OUT = os.path.join(DATA_DIR, 'output_drawing')
os.makedirs(DRAW_OUT, exist_ok=True)

# 论文生成记录存储目录
THESIS_OUT = os.path.join(DATA_DIR, 'output_thesis')
os.makedirs(THESIS_OUT, exist_ok=True)

# 客户上传的「毕业设计任务书 / 毕业设计规范」存放目录（每个客户一份，规范各不相同）
UPLOAD_DIR = os.path.join(DATA_DIR, 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 任务书 / 规范 上传限制
DOC_KINDS = {'task': '毕业设计任务书', 'spec': '毕业设计规范'}
DOC_EXTS = {'.doc', '.docx'}
DOC_MAX_BYTES = 25 * 1024 * 1024        # 单个文件 25MB


# ============================================================
# 数据库（SQLite）
# ============================================================
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS user_params (
            user_id INTEGER PRIMARY KEY,
            params TEXT NOT NULL DEFAULT '{}',
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );
        CREATE TABLE IF NOT EXISTS payment_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            username TEXT,
            amount TEXT DEFAULT '',
            note TEXT DEFAULT '',
            status TEXT DEFAULT 'pending',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            processed_at TEXT
        );
        -- 在线支付订单：下单 → 扫码付款 → 平台回调 → 自动发放次数
        CREATE TABLE IF NOT EXISTS pay_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            out_trade_no TEXT UNIQUE NOT NULL,
            user_id INTEGER NOT NULL,
            username TEXT,
            kind TEXT NOT NULL,
            num INTEGER NOT NULL DEFAULT 1,
            amount REAL NOT NULL DEFAULT 0,
            pay_type TEXT DEFAULT 'alipay',
            status TEXT DEFAULT 'pending',
            trade_no TEXT DEFAULT '',
            qrcode TEXT DEFAULT '',
            payurl TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            paid_at TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_pay_orders_user ON pay_orders(user_id);
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            username TEXT,
            content TEXT NOT NULL,
            is_admin INTEGER DEFAULT 0,
            to_user_id INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            is_read INTEGER DEFAULT 0,
            type TEXT DEFAULT 'text',
            file_url TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS generation_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            username TEXT,
            type TEXT NOT NULL,
            title TEXT DEFAULT '',
            file_url TEXT DEFAULT '',
            file2_url TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS user_docs (
            user_id INTEGER NOT NULL,
            kind TEXT NOT NULL,
            filename TEXT DEFAULT '',
            stored TEXT DEFAULT '',
            size INTEGER DEFAULT 0,
            parsed TEXT DEFAULT '',
            spec_summary TEXT DEFAULT '',
            updated_at TEXT DEFAULT (datetime('now','localtime')),
            PRIMARY KEY (user_id, kind)
        );
    ''')
    conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('invite_code', ?)",
                 (os.environ.get('INVITE_CODE', 'sluice2026'),))
    # ===== 付费功能：旧库字段迁移 =====
    try:
        ucols = [r[1] for r in conn.execute('PRAGMA table_info(users)').fetchall()]
        if 'drawing_quota' not in ucols:
            conn.execute('ALTER TABLE users ADD COLUMN drawing_quota INTEGER DEFAULT 0')
        if 'thesis_quota' not in ucols:
            conn.execute('ALTER TABLE users ADD COLUMN thesis_quota INTEGER DEFAULT 0')
        # 设计计算工具：新老用户一律送 2 次免费额度
        if 'tool_quota' not in ucols:
            conn.execute('ALTER TABLE users ADD COLUMN tool_quota INTEGER DEFAULT 2')
        # 参数验证：同样送 2 次免费额度
        if 'verify_quota' not in ucols:
            conn.execute('ALTER TABLE users ADD COLUMN verify_quota INTEGER DEFAULT 2')
        # messages 表迁移（旧库可能没有 to_user_id）
        mcols = [r[1] for r in conn.execute('PRAGMA table_info(messages)').fetchall()]
        if 'to_user_id' not in mcols:
            conn.execute('ALTER TABLE messages ADD COLUMN to_user_id INTEGER DEFAULT 0')
        if 'type' not in mcols:
            conn.execute("ALTER TABLE messages ADD COLUMN type TEXT DEFAULT 'text'")
        if 'file_url' not in mcols:
            conn.execute("ALTER TABLE messages ADD COLUMN file_url TEXT DEFAULT ''")
        # 收费配置（settings 表）
        defaults = [
            ('drawing_price', '5'), ('thesis_price', '10'), ('tool_price', '5'), ('verify_price', '5'),
            ('pay_note', '扫码支付后，请联系管理员（微信/QQ 私聊）确认到账，由管理员为您开通对应次数。'),
            ('wechat_qr', ''), ('alipay_qr', ''),
            ('register_drawing_bonus', '0'), ('register_thesis_bonus', '0'),
            ('register_tool_bonus', '2'), ('register_verify_bonus', '2'),
            # 在线支付网关（易支付协议）。pay_gateway 留空或 manual = 走人工收款码流程
            ('pay_gateway', ''), ('epay_api', ''), ('epay_pid', ''), ('epay_key', ''),
            ('pay_site_url', ''),
        ]
        for k, v in defaults:
            conn.execute('INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)', (k, v))
    except Exception as e:
        print('[init_db] 迁移警告:', e)
    conn.commit()
    conn.close()


# 云端首次启动：若数据目录无用户库，则从镜像内种子库恢复
def seed_db_if_needed():
    seed = os.environ.get('SEED_DB', '')
    if not os.path.exists(DB_PATH) and seed and os.path.exists(seed):
        import shutil
        shutil.copy(seed, DB_PATH)
        print(f'[init] 已从种子库恢复用户数据: {seed}')


seed_db_if_needed()
init_db()


def get_invite_code():
    conn = get_db()
    row = conn.execute("SELECT value FROM settings WHERE key='invite_code'").fetchone()
    conn.close()
    return row['value'] if row else 'sluice2026'


# ============================================================
# 付费 / 次数扣费
# ============================================================
def get_setting(key, default=''):
    conn = get_db()
    row = conn.execute('SELECT value FROM settings WHERE key=?', (key,)).fetchone()
    conn.close()
    return row['value'] if row else default


def get_gateway_config():
    """服务端内部使用：聚合支付网关配置（含商户密钥，严禁下发前端）"""
    return {
        'api': get_setting('epay_api', '').strip(),
        'pid': get_setting('epay_pid', '').strip(),
        'key': get_setting('epay_key', '').strip(),
    }


def is_gateway_ready():
    """在线支付是否已配置完成。
    配好后前端走「下单 → 扫码 → 回调自动到账」；没配则回退到人工收款码流程。"""
    if get_setting('pay_gateway', '').strip() != 'epay':
        return False
    c = get_gateway_config()
    return bool(c['api'] and c['pid'] and c['key'])


def site_base_url():
    """站点根地址，用于拼回调 / 跳转地址"""
    url = get_setting('pay_site_url', '').strip().rstrip('/')
    if url:
        return url
    # 没配置就从当前请求推断（Railway 等平台会带 X-Forwarded-* 头）
    try:
        return request.host_url.rstrip('/')
    except Exception:
        return ''


def get_pay_config():
    """返回收费配置（价格 / 收款码 / 说明 / 在线支付开关）。
    注意：商户密钥绝不在这里下发。"""
    return {
        'drawing_price': get_setting('drawing_price', '5'),
        'thesis_price': get_setting('thesis_price', '10'),
        'tool_price': get_setting('tool_price', '5'),
        'verify_price': get_setting('verify_price', '5'),
        'pay_note': get_setting('pay_note', ''),
        'wechat_qr': get_setting('wechat_qr', ''),
        'alipay_qr': get_setting('alipay_qr', ''),
        'online': is_gateway_ready(),
    }


# 次数类型 → users 表字段
QUOTA_COL = {
    'drawing': 'drawing_quota',
    'thesis': 'thesis_quota',
    'tool': 'tool_quota',
    'verify': 'verify_quota',
}
# 次数类型 → 中文名（管理后台提示用）
QUOTA_LABEL = {'drawing': '图纸', 'thesis': '论文', 'tool': '计算工具', 'verify': '参数验证'}


def get_quota(uid):
    conn = get_db()
    u = conn.execute('SELECT drawing_quota, thesis_quota, tool_quota, verify_quota FROM users WHERE id=?', (uid,)).fetchone()
    conn.close()
    if not u:
        return 0, 0, 0, 0
    return u['drawing_quota'], u['thesis_quota'], u['tool_quota'], u['verify_quota']


def consume_quota(uid, kind):
    """扣减一次次数。kind: 'drawing' / 'thesis' / 'tool'。返回 True=扣减成功，False=次数不足"""
    col = QUOTA_COL.get(kind)
    if not col:
        return False
    conn = get_db()
    u = conn.execute(f'SELECT {col} FROM users WHERE id=?', (uid,)).fetchone()
    if not u or u[col] <= 0:
        conn.close()
        return False
    conn.execute(f'UPDATE users SET {col}={col}-1 WHERE id=?', (uid,))
    conn.commit()
    conn.close()
    return True


def add_quota(uid, kind, num=1):
    col = QUOTA_COL.get(kind)
    if not col:
        return
    conn = get_db()
    conn.execute(f'UPDATE users SET {col}={col}+? WHERE id=?', (int(num), uid))
    conn.commit()
    conn.close()


def is_admin(u=None):
    """管理员 = 第一个注册的用户"""
    if u is None:
        u = current_user()
    if not u:
        return False
    conn = get_db()
    first = conn.execute('SELECT id FROM users ORDER BY id ASC LIMIT 1').fetchone()
    conn.close()
    return first is not None and u['id'] == first['id']


def current_user():
    uid = session.get('user_id')
    if not uid:
        return None
    conn = get_db()
    row = conn.execute('SELECT * FROM users WHERE id=?', (uid,)).fetchone()
    conn.close()
    return row


# ============================================================
# 登录保护（静态资源与登录/注册页除外）
# ============================================================
PUBLIC_PATHS = {'/login', '/register', '/static', '/favicon.ico'}


@app.before_request
def require_login():
    path = request.path
    if path in ('/', '/tool', '/verify', '/model', '/thesis', '/drawing', '/generate',
                '/tool/use', '/verify/use', '/drawing/generate', '/drawing/dl'):
        if not current_user():
            if path in ('/drawing/generate', '/drawing/dl', '/tool/use', '/verify/use'):
                return jsonify({"ok": False, "error": "请先登录"}), 401
            return redirect(url_for('login'))
    elif path.startswith('/drawing/dl/'):
        if not current_user():
            return jsonify({"ok": False, "error": "请先登录"}), 401
    # 注意：/api/pay/notify 是支付平台的回调入口，**刻意不要求登录**——
    # 平台服务器不会带用户的 session cookie。它的安全由签名验签保证（见 pay_notify）。


@app.after_request
def add_cors(resp):
    resp.headers['Access-Control-Allow-Origin'] = '*'
    resp.headers['Access-Control-Allow-Methods'] = 'GET,POST,OPTIONS'
    resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return resp


# ============================================================
# 用户认证
# ============================================================
# 注册邀请码：只有知道邀请码的人才能注册（在网页"管理"页面修改）
# 默认值可在环境变量 INVITE_CODE 或数据库 settings 表设置
@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'GET':
        return render_template('register.html')
    username = (request.form.get('username') or '').strip()
    password = request.form.get('password') or ''
    confirm = request.form.get('confirm') or ''
    invite = (request.form.get('invite') or '').strip()
    if not username or not password:
        flash('请填写用户名和密码')
        return redirect(url_for('register'))
    if len(password) < 4:
        flash('密码至少 4 位')
        return redirect(url_for('register'))
    if password != confirm:
        flash('两次输入的密码不一致')
        return redirect(url_for('register'))
    if not invite or invite != get_invite_code():
        flash('邀请码错误，无法注册')
        return redirect(url_for('register'))
    conn = get_db()
    if conn.execute('SELECT id FROM users WHERE username=?', (username,)).fetchone():
        conn.close()
        flash('该用户名已被注册')
        return redirect(url_for('register'))
    conn.execute('INSERT INTO users (username, password_hash, drawing_quota, thesis_quota, tool_quota, verify_quota) VALUES (?, ?, ?, ?, ?, ?)',
                 (username, generate_password_hash(password),
                  int(get_setting('register_drawing_bonus', '0') or 0),
                  int(get_setting('register_thesis_bonus', '0') or 0),
                  int(get_setting('register_tool_bonus', '2') or 0),
                  int(get_setting('register_verify_bonus', '2') or 0)))
    conn.commit()
    conn.close()
    flash('注册成功，请登录')
    return redirect(url_for('login'))


@app.route('/admin', methods=['GET', 'POST'])
def admin():
    """管理页面：修改邀请码、查看用户（仅第一个注册的管理员可用）"""
    u = current_user()
    if not u:
        return redirect(url_for('login'))
    if not is_admin(u):
        flash('只有管理员（第一个注册的用户）能访问管理页面')
        return redirect(url_for('index'))
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'update_invite':
            new_code = (request.form.get('invite_code') or '').strip()
            if new_code:
                conn = get_db()
                conn.execute("UPDATE settings SET value=? WHERE key='invite_code'", (new_code,))
                conn.commit()
                conn.close()
                flash(f'邀请码已更新为：{new_code}')
            else:
                flash('邀请码不能为空')
        elif action == 'delete_user':
            uid = request.form.get('user_id')
            conn = get_db()
            conn.execute('DELETE FROM users WHERE id=?', (uid,))
            conn.execute('DELETE FROM user_params WHERE user_id=?', (uid,))
            conn.commit()
            conn.close()
            flash('用户已删除')
        elif action == 'reset_password':
            uid = request.form.get('user_id')
            new_pw = request.form.get('new_password') or ''
            if len(new_pw) < 4:
                flash('新密码至少 4 位')
            else:
                conn = get_db()
                conn.execute('UPDATE users SET password_hash=? WHERE id=?',
                             (generate_password_hash(new_pw), uid))
                conn.commit()
                conn.close()
                flash('密码已重置')
        elif action == 'add_quota':
            # 管理员确认收款后给用户加次数
            uid = request.form.get('user_id')
            kind = request.form.get('kind')  # drawing / thesis / tool
            num = request.form.get('num') or '1'
            try:
                num = int(num)
            except ValueError:
                num = 1
            if num > 0 and kind in QUOTA_COL:
                add_quota(uid, kind, num)
                flash(f'已为用户增加 {QUOTA_LABEL.get(kind, kind)}次数 × {num}')
        elif action == 'update_pay':
            # 保存价格与付款说明
            conn = get_db()
            for k in ('drawing_price', 'thesis_price', 'tool_price', 'verify_price', 'pay_note'):
                v = request.form.get(k, '')
                if k.endswith('_price'):
                    try:
                        v = str(max(0, int(float(v))))
                    except (ValueError, TypeError):
                        v = {'drawing_price': '5', 'thesis_price': '10',
                             'tool_price': '5', 'verify_price': '5'}[k]
                conn.execute("UPDATE settings SET value=? WHERE key=?", (v, k))
            conn.commit()
            conn.close()
            flash('收费设置已保存')
        elif action == 'update_gateway':
            # 在线支付网关（易支付协议）：配好后即可「付完钱自动出票」
            conn = get_db()
            pairs = {
                'pay_gateway': request.form.get('pay_gateway', '').strip(),
                'epay_api': request.form.get('epay_api', '').strip().rstrip('/'),
                'epay_pid': request.form.get('epay_pid', '').strip(),
                'pay_site_url': request.form.get('pay_site_url', '').strip().rstrip('/'),
            }
            if pairs['pay_gateway'] not in ('', 'epay'):
                pairs['pay_gateway'] = ''
            # 密钥留空表示不改动，避免误清空已保存的密钥
            newkey = request.form.get('epay_key', '').strip()
            if newkey:
                pairs['epay_key'] = newkey
            for k, v in pairs.items():
                conn.execute("UPDATE settings SET value=? WHERE key=?", (v, k))
            conn.commit()
            conn.close()
            if is_gateway_ready():
                flash('在线支付已保存并开启：用户付款后自动到账，无需人工确认')
            else:
                flash('在线支付设置已保存，但自动出票未开启——请检查「接口地址 / 商户ID / 商户密钥」是否填全')
        elif action == 'clear_gateway_key':
            conn = get_db()
            conn.execute("UPDATE settings SET value='' WHERE key='epay_key'")
            conn.commit()
            conn.close()
            flash('商户密钥已清空，自动出票已关闭')
        elif action == 'upload_qr':
            # 上传微信 / 支付宝收款码
            kind = request.form.get('kind')  # wechat / alipay
            f = request.files.get('qr')
            if f and f.filename:
                os.makedirs(os.path.join(BASE, 'static', 'pay'), exist_ok=True)
                ext = os.path.splitext(f.filename)[1].lower() or '.png'
                if ext not in ('.png', '.jpg', '.jpeg', '.webp', '.gif'):
                    ext = '.png'
                name = ('wechat' if kind == 'wechat' else 'alipay') + ext
                f.save(os.path.join(BASE, 'static', 'pay', name))
                conn = get_db()
                conn.execute("UPDATE settings SET value=? WHERE key=?",
                             (f'/static/pay/{name}', 'wechat_qr' if kind == 'wechat' else 'alipay_qr'))
                conn.commit()
                conn.close()
                flash('收款码已上传')
        return redirect(url_for('admin'))
    conn = get_db()
    users = conn.execute('SELECT id, username, created_at, drawing_quota, thesis_quota, tool_quota, verify_quota FROM users ORDER BY id').fetchall()
    orders = conn.execute('SELECT * FROM pay_orders ORDER BY id DESC LIMIT 100').fetchall()
    conn.close()
    _key = get_setting('epay_key', '').strip()
    return render_template(
        'admin.html', invite_code=get_invite_code(), users=users, pay=get_pay_config(),
        orders=orders,
        qlabel=QUOTA_LABEL,
        gw={
            'mode': get_setting('pay_gateway', '').strip(),
            'api': get_setting('epay_api', '').strip(),
            'pid': get_setting('epay_pid', '').strip(),
            'site': get_setting('pay_site_url', '').strip(),
            'key_set': bool(_key),
            'key_tail': _key[-4:] if len(_key) >= 4 else '',   # 只回显末 4 位，不泄露完整密钥
            'ready': is_gateway_ready(),
        })


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return render_template('login.html')
    username = (request.form.get('username') or '').strip()
    password = request.form.get('password') or ''
    conn = get_db()
    row = conn.execute('SELECT * FROM users WHERE username=?', (username,)).fetchone()
    conn.close()
    if row and check_password_hash(row['password_hash'], password):
        session['user_id'] = row['id']
        session['username'] = row['username']
        return redirect(url_for('index'))
    flash('用户名或密码错误')
    return redirect(url_for('login'))


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/change-password', methods=['GET', 'POST'])
def change_password():
    """用户修改自己的密码（需验证旧密码）"""
    u = current_user()
    if not u:
        return redirect(url_for('login'))
    if request.method == 'GET':
        return render_template('change_password.html')
    old = request.form.get('old_password') or ''
    new = request.form.get('new_password') or ''
    confirm = request.form.get('confirm') or ''
    conn = get_db()
    row = conn.execute('SELECT password_hash FROM users WHERE id=?', (u['id'],)).fetchone()
    if not row or not check_password_hash(row['password_hash'], old):
        conn.close()
        flash('旧密码错误')
        return redirect(url_for('change_password'))
    if len(new) < 4:
        conn.close()
        flash('新密码至少 4 位')
        return redirect(url_for('change_password'))
    if new != confirm:
        conn.close()
        flash('两次输入的新密码不一致')
        return redirect(url_for('change_password'))
    conn.execute('UPDATE users SET password_hash=? WHERE id=?',
                 (generate_password_hash(new), u['id']))
    conn.commit()
    conn.close()
    flash('密码修改成功，请重新登录')
    session.clear()
    return redirect(url_for('login'))


@app.route('/api/me')
def api_me():
    u = current_user()
    if not u:
        return jsonify({"logged": False}), 401
    dq, tq, toolq, vq = get_quota(u['id'])
    return jsonify({
        "logged": True,
        "username": u['username'],
        "is_admin": is_admin(u),
        "drawing_quota": dq,
        "thesis_quota": tq,
        "tool_quota": toolq,
        "verify_quota": vq,
        "pay": get_pay_config(),
    })


# ============================================================
# 在线支付（聚合支付网关）：下单 → 扫码付款 → 平台回调 → 自动发放次数
# ============================================================
def gen_out_trade_no():
    """生成本站订单号：SJ + 时间 + 6 位随机。只用字母数字，兼容各平台要求。"""
    import random, string
    return 'SJ' + time.strftime('%Y%m%d%H%M%S') + \
        ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))


def grant_order(out_trade_no, trade_no=''):
    """把订单置为已支付并发放次数。

    幂等：支付平台会重复推送回调直到收到 success，所以这里必须保证
    「同一订单只发一次次数」——用带条件的 UPDATE 抢占，rowcount=0 说明
    已被别的回调发过了，直接返回成功，不重复加次数。

    返回 (ok, msg)
    """
    conn = get_db()
    try:
        row = conn.execute('SELECT * FROM pay_orders WHERE out_trade_no=?', (out_trade_no,)).fetchone()
        if not row:
            return False, '订单不存在'
        if row['status'] == 'paid':
            return True, '订单已发放（重复通知，已忽略）'
        col = QUOTA_COL.get(row['kind'])
        if not col:
            return False, '订单类型无效'
        cur = conn.execute(
            "UPDATE pay_orders SET status='paid', trade_no=?, paid_at=datetime('now','localtime') "
            "WHERE out_trade_no=? AND status<>'paid'",
            (trade_no or row['trade_no'], out_trade_no))
        if cur.rowcount == 0:
            conn.rollback()
            return True, '订单已发放（重复通知，已忽略）'
        # 加次数与改状态在同一事务，避免「状态变了但次数没加」
        conn.execute('UPDATE users SET %s=%s+? WHERE id=?' % (col, col), (row['num'], row['user_id']))
        conn.commit()
        return True, '已自动发放 %s × %d' % (QUOTA_LABEL.get(row['kind'], row['kind']), row['num'])
    except Exception as e:
        conn.rollback()
        return False, '发放失败：%s' % e
    finally:
        conn.close()


@app.route('/api/pay/create', methods=['POST'])
def pay_create():
    """创建支付订单：用户选好要买的次数，向平台下单拿二维码。
    这是自动出票的起点——付完钱由 /api/pay/notify 自动到账。"""
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    if not is_gateway_ready():
        return jsonify({"ok": False, "error": "在线支付未开通", "online": False}), 400
    data = request.get_json(force=True, silent=True) or {}
    kind = str(data.get('kind', '')).strip()
    pay_type = str(data.get('pay_type', 'alipay')).strip()
    if kind not in QUOTA_COL:
        return jsonify({"ok": False, "error": "商品类型无效"}), 400
    if pay_type not in ('alipay', 'wxpay'):
        pay_type = 'alipay'
    try:
        num = int(data.get('num', 1) or 1)
    except (ValueError, TypeError):
        num = 1
    num = max(1, min(num, 999))
    try:
        unit = float(get_setting(kind + '_price', '5') or 5)
    except (ValueError, TypeError):
        unit = 5.0
    amount = round(unit * num, 2)
    if amount <= 0:
        return jsonify({"ok": False, "error": "金额无效，请先在后台设置价格"}), 400

    out_trade_no = gen_out_trade_no()
    base = site_base_url()
    if not base:
        return jsonify({"ok": False, "error": "站点地址未配置，无法生成回调地址"}), 400
    label = QUOTA_LABEL.get(kind, kind)
    try:
        res = pgw.create_order(
            get_gateway_config(), out_trade_no, amount,
            name='%s×%d' % (label, num),
            notify_url=base + '/api/pay/notify',
            return_url=base + '/',
            pay_type=pay_type,
        )
    except pgw.PayError as e:
        return jsonify({"ok": False, "error": str(e)}), 502

    conn = get_db()
    conn.execute(
        'INSERT INTO pay_orders (out_trade_no, user_id, username, kind, num, amount, pay_type, qrcode, payurl) '
        'VALUES (?,?,?,?,?,?,?,?,?)',
        (out_trade_no, u['id'], u['username'], kind, num, amount, pay_type,
         res.get('qrcode', ''), res.get('payurl', '')))
    conn.commit()
    conn.close()
    return jsonify({
        "ok": True, "out_trade_no": out_trade_no, "amount": amount, "num": num,
        "kind": kind, "label": label, "pay_type": pay_type,
        "qrcode": res.get('qrcode', ''), "payurl": res.get('payurl', ''),
    })


@app.route('/api/pay/qr/<out_trade_no>')
def pay_qr(out_trade_no):
    """把订单的支付链接渲染成二维码图片（不把链接丢给第三方接口，本地生成）"""
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    conn = get_db()
    row = conn.execute('SELECT * FROM pay_orders WHERE out_trade_no=?', (out_trade_no,)).fetchone()
    conn.close()
    if not row:
        return jsonify({"ok": False, "error": "订单不存在"}), 404
    if row['user_id'] != u['id'] and not is_admin(u):
        return jsonify({"ok": False, "error": "无权限"}), 403
    target = row['qrcode'] or row['payurl']
    if not target:
        return jsonify({"ok": False, "error": "该订单没有支付链接"}), 404
    data = pgw.qr_png_bytes(target)
    if not data:
        return jsonify({"ok": False, "error": "二维码组件未安装"}), 503
    return Response(data, mimetype='image/png',
                    headers={'Cache-Control': 'no-store, max-age=0'})


@app.route('/api/pay/notify', methods=['GET', 'POST'])
def pay_notify():
    """【支付平台回调入口】用户付款成功后，平台会主动请求这里。

    三道关卡，缺一不可：
      1. 验签——签名不对说明不是平台发的，直接拒（否则别人伪造「支付成功」白嫖）
      2. 金额——回调金额必须与订单金额一致
      3. 幂等——grant_order 内部保证同一订单只发一次次数

    无论成败都返回 success/FAIL 纯文本，平台收到 success 才会停止重推。
    """
    params = {}
    params.update(request.args.to_dict())
    if request.method == 'POST':
        form = request.form.to_dict()
        if form:
            params.update(form)
        else:
            params.update(request.get_json(force=True, silent=True) or {})

    cfg = get_gateway_config()
    if not (cfg['api'] and cfg['pid'] and cfg['key']):
        return 'FAIL', 200
    if not pgw.verify_sign(params, cfg['key']):
        print('[pay_notify] 签名校验失败，已拒绝:', dict(params))
        return 'FAIL', 200

    out_trade_no = str(params.get('out_trade_no', '')).strip()
    status = str(params.get('trade_status', 'TRADE_SUCCESS')).strip().upper()
    if status != 'TRADE_SUCCESS':
        return 'success', 200      # 非成功状态只确认收到，不发货

    conn = get_db()
    row = conn.execute('SELECT * FROM pay_orders WHERE out_trade_no=?', (out_trade_no,)).fetchone()
    conn.close()
    if not row:
        print('[pay_notify] 订单不存在:', out_trade_no)
        return 'FAIL', 200

    try:
        paid = round(float(params.get('money', 0) or 0), 2)
    except (ValueError, TypeError):
        paid = 0.0
    if paid and abs(paid - float(row['amount'])) > 0.01:
        print('[pay_notify] 金额不符，已拒绝:', out_trade_no, paid, row['amount'])
        return 'FAIL', 200

    ok, msg = grant_order(out_trade_no, str(params.get('trade_no') or ''))
    print('[pay_notify]', out_trade_no, '发放' if ok else '失败', '-', msg)
    return ('success' if ok else 'FAIL'), 200


@app.route('/api/pay/status')
def pay_status():
    """前端轮询订单状态。已支付时顺带返回最新次数，页面当场刷新。"""
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    no = request.args.get('out_trade_no', '').strip()
    conn = get_db()
    row = conn.execute('SELECT * FROM pay_orders WHERE out_trade_no=?', (no,)).fetchone()
    conn.close()
    if not row:
        return jsonify({"ok": False, "error": "订单不存在"}), 404
    if row['user_id'] != u['id'] and not is_admin(u):
        return jsonify({"ok": False, "error": "无权限"}), 403
    return jsonify({
        "ok": True, "status": row['status'], "num": row['num'],
        "kind": row['kind'], "amount": row['amount'],
        "quota": get_quota_one(row['user_id'], row['kind']),
    })


@app.route('/api/pay/sync', methods=['POST'])
def pay_sync():
    """用户点「我已支付」时主动向平台查单——回调万一丢了，钱付了也不至于没到账。"""
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    data = request.get_json(force=True, silent=True) or {}
    no = str(data.get('out_trade_no', '')).strip()
    conn = get_db()
    row = conn.execute('SELECT * FROM pay_orders WHERE out_trade_no=?', (no,)).fetchone()
    conn.close()
    if not row:
        return jsonify({"ok": False, "error": "订单不存在"}), 404
    if row['user_id'] != u['id'] and not is_admin(u):
        return jsonify({"ok": False, "error": "无权限"}), 403
    if row['status'] == 'paid':
        return jsonify({"ok": True, "status": "paid", "msg": "订单已到账",
                        "quota": get_quota_one(row['user_id'], row['kind'])})
    try:
        res = pgw.query_order(get_gateway_config(), no, row['trade_no'])
    except pgw.PayError as e:
        return jsonify({"ok": False, "error": str(e)}), 502
    if res.get('paid'):
        ok, msg = grant_order(no, res.get('trade_no', ''))
        return jsonify({"ok": ok, "status": "paid" if ok else "pending", "msg": msg,
                        "quota": get_quota_one(row['user_id'], row['kind'])})
    return jsonify({"ok": True, "status": "pending", "msg": "暂未查到支付记录，请稍候再试"})


@app.route('/api/pay/orders')
def pay_orders_list():
    """管理员：在线支付订单流水"""
    u = current_user()
    if not is_admin(u):
        return jsonify({"ok": False, "error": "无权限"}), 403
    conn = get_db()
    rows = conn.execute('SELECT * FROM pay_orders ORDER BY id DESC LIMIT 200').fetchall()
    conn.close()
    return jsonify({"ok": True, "list": [dict(r) for r in rows]})


# ============================================================
# 充值申请（客户付款后提交，管理员在后台确认开通）
# ============================================================
@app.route('/api/pay/request', methods=['POST'])
def pay_request():
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    data = request.get_json(force=True, silent=True) or {}
    amount = str(data.get('amount', '')).strip()
    note = str(data.get('note', '')).strip()
    if not amount:
        return jsonify({"ok": False, "error": "请填写付款金额"}), 400
    conn = get_db()
    conn.execute('INSERT INTO payment_requests (user_id, username, amount, note) VALUES (?,?,?,?)',
                 (u['id'], u['username'], amount, note))
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "msg": "充值申请已提交，请等待管理员确认开通。"})


@app.route('/api/pay/list')
def pay_list():
    u = current_user()
    if not is_admin(u):
        return jsonify({"ok": False, "error": "无权限"}), 403
    conn = get_db()
    rows = conn.execute("SELECT * FROM payment_requests ORDER BY id DESC LIMIT 100").fetchall()
    conn.close()
    return jsonify({"ok": True, "list": [dict(r) for r in rows]})


@app.route('/api/pay/process', methods=['POST'])
def pay_process():
    """管理员确认收款并开通次数"""
    u = current_user()
    if not is_admin(u):
        return jsonify({"ok": False, "error": "无权限"}), 403
    data = request.get_json(force=True, silent=True) or {}
    rid = data.get('id')
    try:
        counts = {k: max(0, int(data.get(k, 0) or 0)) for k in QUOTA_COL}
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "次数无效"}), 400
    if not any(counts.values()):
        return jsonify({"ok": False, "error": "请填写要开通的次数"}), 400
    conn = get_db()
    row = conn.execute('SELECT * FROM payment_requests WHERE id=? AND status=?', (rid, 'pending')).fetchone()
    if not row:
        conn.close()
        return jsonify({"ok": False, "error": "申请不存在或已处理"}), 400
    for kind, num in counts.items():
        if num:
            add_quota(row['user_id'], kind, num)
    conn.execute("UPDATE payment_requests SET status='done', processed_at=datetime('now','localtime') WHERE id=?", (rid,))
    conn.commit()
    conn.close()
    parts = [f'{QUOTA_LABEL[k]}×{v}' for k, v in counts.items() if v]
    return jsonify({"ok": True, "msg": f"已为用户 {row['username']} 开通 " + "、".join(parts)})


# ============================================================
# 客服消息（站内信：客户留言，管理员回复）
# ============================================================
@app.route('/api/msg/send', methods=['POST'])
def msg_send():
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    data = request.get_json(force=True, silent=True) or {}
    msg_type = str(data.get('type', 'text')) or 'text'
    content = str(data.get('content', '')).strip()
    file_url = ''
    if msg_type == 'image':
        # 图片消息：image_data 为 dataURL（data:image/png;base64,xxxx）
        image_data = str(data.get('image_data', '') or '')
        if not image_data.startswith('data:image'):
            return jsonify({"ok": False, "error": "图片格式不正确"}), 400
        try:
            import base64
            header, b64 = image_data.split(',', 1)
            ext = '.png'
            mime = header.split(';')[0].split('/')[-1].lower()
            if mime in ('jpeg', 'jpg'):
                ext = '.jpg'
            elif mime == 'gif':
                ext = '.gif'
            elif mime == 'webp':
                ext = '.webp'
            raw = base64.b64decode(b64)
            if len(raw) > 8 * 1024 * 1024:
                return jsonify({"ok": False, "error": "图片过大（限 8MB）"}), 400
            img_dir = os.path.join(DATA_DIR, 'msg_imgs')
            os.makedirs(img_dir, exist_ok=True)
            fname = f"msg_{int(time.time()*1000)}_{u['id']}{ext}"
            with open(os.path.join(img_dir, fname), 'wb') as f:
                f.write(raw)
            file_url = f"/api/msg/image/{fname}"
            content = '发来一张图片'
        except Exception as e:
            print('[msg] 图片保存失败:', e)
            return jsonify({"ok": False, "error": "图片保存失败"}), 500
    else:
        if not content:
            return jsonify({"ok": False, "error": "消息不能为空"}), 400
        if len(content) > 2000:
            content = content[:2000]
    is_admin_flag = 1 if is_admin(u) else 0
    # 管理员回复时可指定发给哪个用户（to_user_id），否则发给管理员(0)
    to_user_id = 0
    if is_admin_flag:
        try:
            to_user_id = int(data.get('to_user_id', 0) or 0)
        except (ValueError, TypeError):
            to_user_id = 0
    conn = get_db()
    conn.execute('INSERT INTO messages (user_id, username, content, is_admin, to_user_id, type, file_url) VALUES (?,?,?,?,?,?,?)',
                 (u['id'], u['username'], content, is_admin_flag, to_user_id, msg_type, file_url))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


@app.route('/api/msg/image/<path:filename>')
def msg_image(filename):
    """返回聊天图片（需登录，且属于当前用户或管理员）"""
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    img_dir = os.path.join(DATA_DIR, 'msg_imgs')
    safe = os.path.basename(filename)
    path = os.path.join(img_dir, safe)
    if not os.path.exists(path):
        return jsonify({"ok": False, "error": "图片不存在"}), 404
    return send_from_directory(img_dir, safe)


@app.route('/api/msg/list')
def msg_list():
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    conn = get_db()
    if is_admin(u):
        rows = conn.execute('SELECT * FROM messages ORDER BY id DESC LIMIT 300').fetchall()
    else:
        rows = conn.execute('SELECT * FROM messages WHERE user_id=? OR to_user_id=? ORDER BY id DESC LIMIT 200',
                            (u['id'], u['id'])).fetchall()
    conn.close()
    msgs = [dict(r) for r in rows]
    msgs.reverse()
    return jsonify({"ok": True, "list": msgs, "is_admin": is_admin(u)})


# ============================================================
# 生成记录（客户历史生成，可回看/重新下载）
# ============================================================
@app.route('/api/records/save', methods=['POST'])
def records_save():
    """论文生成后前端上传 docx 并保存记录（图纸在 drawing_generate 里自动记录）"""
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    f = request.files.get('file')
    title = (request.form.get('title') or '毕业设计论文').strip()
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "未收到文件"}), 400
    user_dir = os.path.join(THESIS_OUT, str(u['id']))
    os.makedirs(user_dir, exist_ok=True)
    stamp = int(time.time())
    # 清理文件名中的非法字符
    safe_title = ''.join(c for c in title if c not in '\\/:*?"<>|').strip() or '论文'
    fname = f"{safe_title}_{stamp}.docx"
    fpath = os.path.join(user_dir, fname)
    f.save(fpath)
    conn = get_db()
    conn.execute(
        'INSERT INTO generation_records (user_id, username, type, title, file_url) VALUES (?,?,?,?,?)',
        (u['id'], u['username'], 'thesis', title, f"/thesis/dl/{u['id']}/{fname}"))
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "msg": "论文已保存到你的生成记录"})


@app.route('/api/records')
def records_list():
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    conn = get_db()
    if is_admin(u):
        rows = conn.execute('SELECT * FROM generation_records ORDER BY id DESC LIMIT 500').fetchall()
    else:
        rows = conn.execute('SELECT * FROM generation_records WHERE user_id=? ORDER BY id DESC LIMIT 100',
                            (u['id'],)).fetchall()
    conn.close()
    out = []
    for r in rows:
        d = dict(r)
        # 管理员视角：附上通用下载地址（可下载任意客户文件）
        if is_admin(u):
            _url = d.get('file_url') or ''
            _fn = _url.rsplit('/', 1)[-1] if '/' in _url else _url
            d['admin_dl'] = f"/admin/dl/{d['type']}/{d['user_id']}/{_fn}" if _fn else ''
        out.append(d)
    return jsonify({"ok": True, "list": out})


@app.route('/records/dl/<int:uid>/<path:filename>')
def records_download(uid, filename):
    u = current_user()
    if u is None or (u['id'] != uid and not is_admin(u)):
        return jsonify({"ok": False, "error": "无权限"}), 403
    return send_from_directory(os.path.join(THESIS_OUT, str(uid)), filename)


@app.route('/admin/dl/<kind>/<int:uid>/<path:filename>')
def admin_download(kind, uid, filename):
    """管理员通用下载：kind=thesis|drawing，可下载任意客户的文件"""
    u = current_user()
    if not u or not is_admin(u):
        return jsonify({"ok": False, "error": "无权限"}), 403
    base = THESIS_OUT if kind == 'thesis' else DRAW_OUT
    return send_from_directory(os.path.join(base, str(uid)), filename, as_attachment=True)


@app.route('/records')
def records_page():
    u = current_user()
    if not u:
        return redirect(url_for('login'))
    return render_template('records.html', username=u['username'], is_admin=is_admin(u))


# ============================================================
# 用户参数持久化
# ============================================================
def save_user_params(uid, params):
    conn = get_db()
    conn.execute(
        'INSERT INTO user_params (user_id, params, updated_at) VALUES (?,?,datetime(\'now\',\'localtime\')) '
        'ON CONFLICT(user_id) DO UPDATE SET params=excluded.params, updated_at=excluded.updated_at',
        (uid, json.dumps(params, ensure_ascii=False)))
    conn.commit()
    conn.close()


def load_user_params(uid):
    conn = get_db()
    row = conn.execute('SELECT params FROM user_params WHERE user_id=?', (uid,)).fetchone()
    conn.close()
    if not row:
        return {}
    try:
        return json.loads(row['params'])
    except Exception:
        return {}


# ============================================================
# 页面路由
# ============================================================
@app.route('/', methods=['GET', 'POST', 'OPTIONS'])
def index():
    if request.method == 'OPTIONS':
        return ('', 204)
    return send_from_directory(BASE, 'index.html')


@app.route('/verify')
def page_verify():
    return send_from_directory(BASE, 'index.html')

@app.route('/model')
def page_model():
    return send_from_directory(BASE, 'index.html')


@app.route('/thesis')
def page_thesis():
    return send_from_directory(BASE, 'index.html')


@app.route('/tool')
def page_tool():
    return send_from_directory(BASE, 'index.html')


def get_quota_one(uid, kind):
    """读取某一类次数余额"""
    col = QUOTA_COL.get(kind)
    if not col:
        return 0
    conn = get_db()
    row = conn.execute(f'SELECT {col} FROM users WHERE id=?', (uid,)).fetchone()
    conn.close()
    return row[col] if row else 0


def _quota_take(kind):
    """按次收费的前端功能通用扣次逻辑（设计计算工具 / 参数验证）。
    成功返回 {ok:True, 剩余次数, pay}；次数不足返回 402 + need_pay。"""
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    key = kind + '_quota'
    if not consume_quota(u['id'], kind):
        return jsonify({
            "ok": False,
            "need_pay": True,
            "error": "免费次数已用完，请付费后继续使用。",
            key: 0,
            "pay": get_pay_config(),
        }), 402
    return jsonify({"ok": True, key: get_quota_one(u['id'], kind), "pay": get_pay_config()})


@app.route('/tool/use', methods=['POST'])
def tool_use():
    """设计计算工具：每点一次「开始计算」扣 1 次。免费 2 次用完后需付费开通。"""
    return _quota_take('tool')


@app.route('/verify/use', methods=['POST'])
def verify_use():
    """参数验证：每次点「参数验证与设计计算」扣 1 次。免费 2 次用完后需付费开通。"""
    return _quota_take('verify')


# ============================================================
# 论文生成：任务书 + 规范 → 整本论文
# 每个客户上传自己的任务书与规范（各校规范不同，不能用固定模板），
# 系统从任务书取数据、从规范取格式，正文按实际数据重新撰写。
# ============================================================
def _doc_dir(uid):
    d = os.path.join(UPLOAD_DIR, str(uid))
    os.makedirs(d, exist_ok=True)
    return d


def _doc_row(uid, kind):
    conn = get_db()
    r = conn.execute('SELECT * FROM user_docs WHERE user_id=? AND kind=?', (uid, kind)).fetchone()
    conn.close()
    return r


def _doc_path(uid, kind):
    """返回已上传文件的绝对路径；没有则 None。"""
    r = _doc_row(uid, kind)
    if not r or not r['stored']:
        return None
    p = os.path.join(_doc_dir(uid), r['stored'])
    return p if os.path.exists(p) else None


def _sniff_kind(path):
    """判断文件真实格式：docx(zip) / doc(OLE) / 其它。
    很多客户把 .docx 直接改名成 .doc，所以必须嗅探而不是只看扩展名。"""
    try:
        with open(path, 'rb') as f:
            head = f.read(8)
    except Exception:
        return 'unknown'
    if head[:4] == b'PK\x03\x04':
        return 'docx'
    if head[:4] == b'\xd0\xcf\x11\xe0':
        return 'doc'
    return 'unknown'


def _as_readable_docx(path):
    """把上传的 Word 变成 python-docx 能读的文件，返回可用路径。
    老版 .doc 在 Windows 本地可借 Word COM 转换；云端无法转换时给出明确提示。"""
    k = _sniff_kind(path)
    if k == 'docx':
        return path, None
    if k == 'doc':
        if os.name == 'nt':
            try:
                import win32com.client  # noqa
                out = os.path.splitext(path)[0] + '_conv.docx'
                w = win32com.client.Dispatch('Word.Application')
                w.Visible = False
                try:
                    d = w.Documents.Open(os.path.abspath(path), ReadOnly=True)
                    d.SaveAs2(os.path.abspath(out), FileFormat=16)   # 16 = wdFormatDocumentDefault
                    d.Close(False)
                finally:
                    w.Quit()
                if os.path.exists(out):
                    return out, None
            except Exception as e:
                return None, ('该文件是老版 .doc 格式，服务器无法自动转换（%s）。'
                              '请在 Word 中「另存为 → Word 文档(*.docx)」后重新上传。' % e)
        return None, ('该文件是老版 .doc 格式，服务器无法直接读取。'
                      '请在 Word 中「另存为 → Word 文档(*.docx)」后重新上传。')
    return None, '文件格式无法识别，请上传 Word 文档（.docx 或 .doc）。'


def _parse_docs(uid):
    """读取该用户的任务书与规范。返回 (pack, warnings)。"""
    pack = {'params': {}, 'sources': [], 'sections': [],
            'spec': None, 'task_name': '', 'spec_name': '', 'raw': '',
            'project_name': '', 'has_task': False, 'has_spec': False}
    warns = []

    task_path = _doc_path(uid, 'task')
    if task_path:
        usable, err = _as_readable_docx(task_path)
        if usable:
            try:
                res = tp.parse_task_book(usable)
                pack['params'] = res.get('params') or {}
                pack['sources'] = res.get('sources') or []
                pack['sections'] = res.get('sections') or []
                pack['project_name'] = res.get('project_name') or ''
                pack['has_task'] = True
                # 原文留着：正文撰写引擎用它做随机种子，
                # 也用于判断是否属除险加固工程。
                pack['raw'] = res.get('raw') or ''
                pack['task_meta'] = {'n_tables': res.get('n_tables', 0),
                                     'n_paras': res.get('n_paras', 0)}
                if not pack['params']:
                    warns.append('任务书里没有识别到设计参数，请检查文件是否为任务书正文。')
            except Exception as e:
                warns.append('任务书解析失败：%s' % e)
        else:
            warns.append('任务书：%s' % err)

    spec_path = _doc_path(uid, 'spec')
    if spec_path:
        usable, err = _as_readable_docx(spec_path)
        if usable:
            try:
                sp = tp.parse_spec(usable)
                pack['spec'] = sp
                pack['has_spec'] = True
                if not sp.get('chapters'):
                    warns.append('规范里没有识别到章节标题，论文结构将按通用水闸设计流程组织。')
            except Exception as e:
                warns.append('规范解析失败：%s' % e)
        else:
            warns.append('规范：%s' % err)

    tr, sr = _doc_row(uid, 'task'), _doc_row(uid, 'spec')
    pack['task_name'] = tr['filename'] if tr else ''
    pack['spec_name'] = sr['filename'] if sr else ''
    return pack, warns


def _labeled_rows(pack, limit=18):
    """把抽取到的参数整理成「基本资料一览表」的行。"""
    rows = []
    for s in (pack.get('sources') or [])[:limit]:
        name = s.get('label') or s.get('key')
        unit = s.get('unit') or ''
        rows.append(('%s%s' % (name, ('（%s）' % unit) if unit else ''), s.get('value', '')))
    return rows


def _generate_thesis(uid, params, meta=None):
    """执行完整生成流程，返回 (BytesIO, warnings, pack)。"""
    pack, warns = _parse_docs(uid)
    meta = dict(meta or {})

    P = {}
    P.update(pack['params'])              # 任务书里抽到的
    # 封面/抬头信息也要进参数：thesis_build 的封面与摘要都从 params 里取值
    P.update({k: v for k, v in meta.items()
              if v not in (None, '') and k not in ('labeled', 'warnings')})
    P.update({k: v for k, v in (params or {}).items()
              if v not in (None, '')})    # 客户校对后的值最后覆盖
    # 学校名：客户填的优先，其次从规范标题里认出来的
    if not (P.get('university') or '').strip():
        P['university'] = (pack.get('spec') or {}).get('university', '')
    # 题目：客户填的 > 任务书里的题目 > 按河流名拼一个
    if not (P.get('projectName') or '').strip():
        P['projectName'] = pack.get('project_name') or _proj_name(P)

    meta['labeled'] = _labeled_rows(pack)
    meta.setdefault('taskName', _proj_name(P))
    # 正文撰写引擎的随机种子取自任务书原文：
    # 同一份任务书结果可复现，不同任务书必然得到不同的行文。
    meta['seed_material'] = pack.get('raw') or '\n'.join(pack.get('sections') or [])
    P['_raw'] = meta['seed_material']

    # 任务书缺关键参数要明说：整本论文会按内置示例值（calc.KEY_DEFAULTS）
    # 计算，不提醒的话客户会拿到一本“数据凭空而来”的论文（线上实测踩过）。
    for _k, _label in (
            ('floodStandard', '洪水重现期'),
            ('designFlow', '设计流量'),
            ('checkFlow', '校核流量'),
            ('gateSillElevation', '闸底板顶高程'),
            ('downstreamWaterLevel', '下游设计水位（设计洪水位）'),
            ('normalStorageLevel', '正常蓄水位'),
            ('checkWaterLevel', '校核洪水位'),
            ('groundElevation', '闸址地面高程'),
            ('gateCount', '闸孔数'),
            ('singleGateWidth', '单孔净宽')):
        if P.get(_k) in (None, ''):
            _dv = calc.KEY_DEFAULTS.get(_k)
            warns.append('任务书中未识别到「%s」%s，请在“参数校对”中补充后重新生成'
                         % (_label, ('，本次计算暂按示例值 %s 进行' % _dv)
                            if _dv is not None else ''))

    # 曲线图：与正文同一套计算，图表只插到对应章节
    gw = calc.calc_gate_width_mu0(P)
    top = calc.calc_gate_top_mu0(P)
    figs, fig_warns = ch.render_all(gw, top, calc.calc_seepage_mu0(P),
                                    calc.calc_energy_mu0(P, gw),
                                    calc.calc_stability_mu0(P, gw, top), P)
    warns += fig_warns

    meta['warnings'] = warns
    buf, bw = tb.build(P, spec=pack.get('spec'), task_sections=pack.get('sections'),
                       figures=figs, meta=meta)
    warn_all = list(dict.fromkeys(warns + list(bw or [])))
    return buf, warn_all, pack


def _proj_name(P):
    river = (P.get('riverName') or '').strip()
    return ('%s水闸拆除重建工程' % river) if river else '水闸拆除重建工程'


_UNSAFE_FN = '\\/:*?"<>|\r\n\t'


def _safe_filename(s):
    """去掉文件名里不能用的字符（题目常带引号书名号，直接进 Content-Disposition 会出问题）。"""
    out = ''.join(('' if ch in _UNSAFE_FN else ch) for ch in str(s or ''))
    out = out.replace('“', '').replace('”', '').replace('《', '').replace('》', '')
    out = out.strip().strip('.')
    return out or '水闸设计'


@app.route('/thesis/doc/status')
def thesis_doc_status():
    """查询当前用户已上传的任务书/规范，以及上次解析出的参数。"""
    u = current_user()
    if not u:
        return jsonify({'ok': False, 'error': '请先登录'}), 401
    out = {'ok': True, 'docs': {}}
    for kind, label in DOC_KINDS.items():
        r = _doc_row(u['id'], kind)
        if r and r['stored']:
            out['docs'][kind] = {
                'kind': kind, 'label': label, 'filename': r['filename'],
                'size': r['size'], 'updated_at': r['updated_at'],
                'has_file': _doc_path(u['id'], kind) is not None,
            }
        else:
            out['docs'][kind] = {'kind': kind, 'label': label, 'filename': '', 'has_file': False}
    # 回填上次解析结果，避免重复解析
    tr = _doc_row(u['id'], 'task')
    sr = _doc_row(u['id'], 'spec')
    out['parsed'] = json.loads(tr['parsed']) if (tr and tr['parsed']) else None
    out['spec_summary'] = (sr['spec_summary'] if sr else '') or ''
    out['quota'] = get_quota_one(u['id'], 'thesis')
    out['pay'] = get_pay_config()
    return jsonify(out)


@app.route('/thesis/doc/upload', methods=['POST'])
def thesis_doc_upload():
    """上传任务书或规范。上传与解析均不扣次数，只有生成论文才扣。"""
    u = current_user()
    if not u:
        return jsonify({'ok': False, 'error': '请先登录'}), 401
    # kind 既接受表单字段也接受查询串，前端怎么传都不会 400
    kind = (request.form.get('kind') or request.args.get('kind') or '').strip()
    if kind not in DOC_KINDS:
        return jsonify({'ok': False, 'error': '上传类型不正确（应为 task 或 spec）'}), 400
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify({'ok': False, 'error': '请选择要上传的 Word 文件'}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in DOC_EXTS:
        return jsonify({'ok': False,
                        'error': '只支持 Word 文档（.docx / .doc），当前是「%s」' % (ext or '无扩展名')}), 400

    data = f.read()
    if not data:
        return jsonify({'ok': False, 'error': '文件内容为空'}), 400
    if len(data) > DOC_MAX_BYTES:
        return jsonify({'ok': False,
                        'error': '文件过大（%.1f MB），请控制在 25MB 以内' % (len(data) / 1048576.0)}), 400

    stored = '%s%s' % (kind, ext)
    d = _doc_dir(u['id'])
    # 换扩展名重传时先清掉旧文件，避免残留
    old = _doc_row(u['id'], kind)
    if old and old['stored'] and old['stored'] != stored:
        try:
            os.remove(os.path.join(d, old['stored']))
        except Exception:
            pass
    with open(os.path.join(d, stored), 'wb') as fp:
        fp.write(data)

    conn = get_db()
    conn.execute('''INSERT INTO user_docs (user_id, kind, filename, stored, size, parsed, spec_summary, updated_at)
                    VALUES (?, ?, ?, ?, ?, '', '', datetime('now','localtime'))
                    ON CONFLICT(user_id, kind) DO UPDATE SET
                    filename=excluded.filename, stored=excluded.stored, size=excluded.size,
                    parsed='', spec_summary='', updated_at=datetime('now','localtime')''',
                 (u['id'], kind, f.filename, stored, len(data)))
    conn.commit()
    conn.close()

    # 立刻试解析一次，把问题当场暴露给客户
    try:
        path = os.path.join(d, stored)
        usable, err = _as_readable_docx(path)
        if not usable:
            return jsonify({'ok': True, 'kind': kind, 'filename': f.filename,
                            'warning': err, 'params_n': 0})
        if kind == 'task':
            res = tp.parse_task_book(usable)
            n = len(res.get('params') or {})
            return jsonify({'ok': True, 'kind': kind, 'filename': f.filename,
                            'params_n': n, 'n_tables': res.get('n_tables', 0),
                            'warning': '' if n else '没有从任务书中识别到参数，请确认上传的是设计任务书'})
        sp = tp.parse_spec(usable)
        hint = tp.spec_hint(sp)
        conn = get_db()
        conn.execute('UPDATE user_docs SET spec_summary=? WHERE user_id=? AND kind=?',
                     (hint, u['id'], kind))
        conn.commit()
        conn.close()
        return jsonify({'ok': True, 'kind': kind, 'filename': f.filename,
                        'params_n': len(sp.get('chapters') or []),
                        'spec_summary': hint, 'warning': ''})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'ok': True, 'kind': kind, 'filename': f.filename,
                        'params_n': 0, 'warning': '文件已保存，但解析出错：%s' % e})


@app.route('/thesis/doc/parse', methods=['POST'])
def thesis_doc_parse():
    """解析已上传的任务书与规范，返回可校对的设计参数与格式摘要。不扣次数。"""
    u = current_user()
    if not u:
        return jsonify({'ok': False, 'error': '请先登录'}), 401
    pack, warns = _parse_docs(u['id'])
    if not pack['has_task'] and not pack['has_spec']:
        return jsonify({'ok': False, 'error': '请先上传毕业设计任务书（建议同时上传毕业设计规范）',
                        'warnings': warns}), 400

    hint = tp.spec_hint(pack['spec']) if pack.get('spec') else ''
    conn = get_db()
    conn.execute('''UPDATE user_docs SET parsed=?, updated_at=datetime('now','localtime')
                    WHERE user_id=? AND kind='task' ''',
                 (json.dumps({'params': pack['params'], 'sources': pack['sources']},
                             ensure_ascii=False), u['id']))
    if hint:
        conn.execute('''UPDATE user_docs SET spec_summary=?, updated_at=datetime('now','localtime')
                        WHERE user_id=? AND kind='spec' ''', (hint, u['id']))
    conn.commit()
    conn.close()

    chapters = []
    if pack.get('spec'):
        chapters = [c for c in (pack['spec'].get('chapters') or []) if c.get('level') == 1]
    return jsonify({
        'ok': True,
        'params': pack['params'],
        'sources': pack['sources'],
        'sections': pack['sections'][:6],
        'has_task': pack['has_task'],
        'has_spec': pack['has_spec'],
        'task_name': pack['task_name'],
        'spec_name': pack['spec_name'],
        'spec_summary': hint,
        'university': (pack.get('spec') or {}).get('university', ''),
        'project_name': pack.get('project_name', ''),
        'n_tables': (pack.get('task_meta') or {}).get('n_tables', 0),
        'chapters': chapters,
        'warnings': warns,
        'quota': get_quota_one(u['id'], 'thesis'),
    })


@app.route('/thesis/doc/delete', methods=['POST'])
def thesis_doc_delete():
    """删除已上传的任务书或规范"""
    u = current_user()
    if not u:
        return jsonify({'ok': False, 'error': '请先登录'}), 401
    data = request.get_json(force=True, silent=True) or {}
    kind = (data.get('kind') or '').strip()
    if kind not in DOC_KINDS:
        return jsonify({'ok': False, 'error': '类型不正确'}), 400
    r = _doc_row(u['id'], kind)
    if r and r['stored']:
        try:
            os.remove(os.path.join(_doc_dir(u['id']), r['stored']))
        except Exception:
            pass
    conn = get_db()
    conn.execute('DELETE FROM user_docs WHERE user_id=? AND kind=?', (u['id'], kind))
    conn.commit()
    conn.close()
    return jsonify({'ok': True})


@app.route('/generate', methods=['POST'])
def generate():
    """生成整本毕业设计论文（含计算表格与曲线图）。
    数据来自客户上传的任务书 + 校对后的参数，格式来自客户上传的规范。
    旧的前端直接传参方式仍然兼容（未上传任务书时走原模板）。"""
    u = current_user()
    body = request.get_json(force=True, silent=True) or {}

    # ---- 兼容旧调用：纯参数对象 {'designFlow': ..., ...} ----
    if 'params' not in body and not body.get('use_docs'):
        if u and not _doc_path(u['id'], 'task'):
            try:
                buf = gt.generate(body)
                return send_file(buf, as_attachment=True, download_name='毕业设计论文.docx',
                                 mimetype='application/vnd.openxmlformats-officedocument.'
                                          'wordprocessingml.document')
            except Exception as e:
                import traceback
                traceback.print_exc()
                return jsonify({'error': str(e)}), 500

    if not u:
        return jsonify({'ok': False, 'error': '请先登录'}), 401

    params = body.get('params') or {}
    meta = body.get('meta') or {}

    if not _doc_path(u['id'], 'task'):
        return jsonify({'ok': False, 'need_upload': True,
                        'error': '请先上传毕业设计任务书，系统需要据此取用设计数据'}), 400

    # ---- 计次：只有生成论文才扣 ----
    if not consume_quota(u['id'], 'thesis'):
        return jsonify({'ok': False, 'need_pay': True, 'thesis_quota': 0,
                        'error': '论文生成次数已用完，请付费后继续使用。',
                        'pay': get_pay_config()}), 402

    try:
        buf, warns, pack = _generate_thesis(u['id'], params, meta)
    except Exception as e:
        import traceback
        traceback.print_exc()
        add_quota(u['id'], 'thesis', 1)      # 生成失败，把次数退回
        return jsonify({'ok': False, 'error': '论文生成失败：%s' % e}), 500

    # ---- 落盘存档 + 生成记录 ----
    merged = dict(pack['params'])
    merged.update({k: v for k, v in meta.items() if v not in (None, '')})
    merged.update({k: v for k, v in params.items() if v not in (None, '')})
    if not (merged.get('projectName') or '').strip():
        merged['projectName'] = pack.get('project_name') or _proj_name(merged)
    filename = '%s毕业设计论文.docx' % _safe_filename(merged.get('projectName'))
    stamp = time.strftime('%Y%m%d%H%M%S')
    # 存档名：时间戳 + 正式文件名。filename 本身已带 .docx，这里不能再补扩展名
    save_name = '%s_%s' % (stamp, filename)
    out_dir = os.path.join(THESIS_OUT, str(u['id']))
    os.makedirs(out_dir, exist_ok=True)
    try:
        with open(os.path.join(out_dir, save_name), 'wb') as fp:
            fp.write(buf.getvalue())
    except Exception as e:
        print('[thesis] 存档失败:', e)
    buf.seek(0)

    try:
        conn = get_db()
        conn.execute('''INSERT INTO generation_records (user_id, username, type, title, file_url)
                        VALUES (?, ?, 'thesis', ?, ?)''',
                     (u['id'], u['username'], filename, '/thesis/dl/%d/%s' % (u['id'], save_name)))
        conn.commit()
        conn.close()
    except Exception as e:
        print('[thesis] 记录写入失败:', e)

    resp = send_file(buf, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.'
                              'wordprocessingml.document')
    # 把生成过程中的提示带给前端（供页面展示「哪些章节留白/哪些参数缺失」）
    # 必须用 ASCII 转义：HTTP 头只能 latin-1，直接塞中文会 UnicodeEncodeError，
    # 表现是「生成成功但浏览器收不到响应」。前端 JSON.parse 能正常还原 \uXXXX。
    resp.headers['X-Thesis-Warnings'] = json.dumps(warns[:8])
    resp.headers['Access-Control-Expose-Headers'] = 'X-Thesis-Warnings, Content-Disposition'
    resp.headers['X-Thesis-Quota'] = str(get_quota_one(u['id'], 'thesis'))
    return resp


@app.route('/thesis/dl/<int:uid>/<path:filename>')
def thesis_download(uid, filename):
    """下载历史生成的论文"""
    u = current_user()
    if not u:
        return redirect('/login')
    if u['id'] != uid and not is_admin(u):
        return '无权限', 403
    d = os.path.join(THESIS_OUT, str(uid))
    return send_from_directory(d, filename, as_attachment=True)



# ============================================================
# 工程图纸生成器（水闸纵剖面图）—— 独立页面
# ============================================================
FIELDS = [
    ("pg_len", "铺盖长度", "铺盖", 15.0, "m"), ("pg_h", "铺盖厚度", "铺盖", 0.5, "m"),
    ("pg_cd", "铺盖齿墙深", "铺盖", 0.5, "m"), ("pg_cw", "铺盖齿墙宽", "铺盖", 0.5, "m"),
    ("db_len", "底板长度", "底板", 14.0, "m"), ("db_h", "底板厚度", "底板", 1.2, "m"),
    ("db_cd", "底板齿墙深", "底板", 1.0, "m"), ("db_cw", "底板齿墙宽", "底板", 1.0, "m"),
    ("xl_len", "消力池长度", "消力池", 15.0, "m"), ("xl_h", "消力池底板厚", "消力池", 0.6, "m"),
    ("xl_cd", "消力池齿墙深", "消力池", 0.5, "m"), ("xl_cw", "消力池齿墙宽", "消力池", 0.5, "m"),
    ("fl_gravel", "砾石层厚", "反滤层", 0.2, "m"), ("fl_stone", "碎石层厚", "反滤层", 0.3, "m"),
    ("fl_sand", "粗砂层厚", "反滤层", 0.2, "m"),
    ("hm_total", "海漫总长", "海漫", 20.0, "m"), ("hm_horiz", "水平段长度", "海漫", 10.0, "m"),
    ("hm_stone", "砌石厚度", "海漫", 0.5, "m"), ("hm_cushion", "粗砂垫层厚", "海漫", 0.1, "m"),
    ("hm_slope", "斜坡坡率", "海漫", 0.1, "比率"), ("hm_cd", "海漫齿墙深", "海漫", 0.5, "m"),
    ("hm_cw", "海漫齿墙宽", "海漫", 0.5, "m"),
    ("fcc_d", "防冲槽深度", "防冲槽", 2.85, "m"), ("fcc_bw", "防冲槽底宽", "防冲槽", 5.0, "m"),
    ("fcc_rip", "堆石覆盖厚", "防冲槽", 0.4, "m"), ("fcc_m", "边坡坡率", "防冲槽", 2.0, "比率"),
    ("el_pg", "铺盖顶/底板顶高程", "高程", 73.1, "m"), ("el_bank", "滩地高程", "高程", 75.0, "m"),
    ("el_wl", "正常蓄水位", "高程", 76.6, "m"), ("el_gate_top", "闸顶高程", "高程", 78.8, "m"),
    ("el_trestle", "排架高程", "高程", 84.8, "m"), ("el_bridge", "工作桥高程", "高程", 85.5, "m"),
    ("gate1_x", "检修门距底板左端", "闸门", 0.65, "m"), ("gate1_w", "检修闸门厚", "闸门", 0.3, "m"),
    ("gate2_gap", "工作门距检修门", "闸门", 1.9, "m"), ("gate2_w", "工作闸门厚", "闸门", 0.8, "m"),
    ("pa_w", "排架立柱宽", "排架", 0.4, "m"), ("pa_beam_w", "排架横梁长", "排架", 4.6, "m"),
    ("pa_beam_h", "排架横梁厚", "排架", 0.3, "m"),
    ("br_w", "工作桥总宽", "工作桥", 4.6, "m"), ("br_col_w", "底端腿宽", "工作桥", 0.4, "m"),
    ("br_deck_h", "桥面板厚", "工作桥", 0.2, "m"),
    ("tb_w", "交通桥宽度", "交通桥", 4.6, "m"), ("tb_deck_top", "桥面顶高程", "交通桥", 78.9, "m"),
    ("tb_cushion", "垫层+支座总高", "交通桥", 0.14, "m"), ("tb_slab", "空心板厚", "交通桥", 0.5, "m"),
    ("tb_overlay", "板顶混凝土厚", "交通桥", 0.1, "m"),
    ("house_w", "机房宽度", "启闭机房", 4.6, "m"), ("house_wall_h", "墙体高度", "启闭机房", 2.0, "m"),
    ("house_roof_h", "屋顶高度", "启闭机房", 1.5, "m"), ("house_door_w", "门宽", "启闭机房", 0.9, "m"),
    ("house_door_h", "门高", "启闭机房", 1.6, "m"),
    ("title", "图名", "绘图选项", "水闸纵剖面图", "文本"), ("scale_text", "比例尺", "绘图选项", "1:100", "文本"),
]
RATIO_KEYS = {"hm_slope", "fcc_m"}
TEXT_KEYS = {"title", "scale_text"}
GROUPS = [("铺盖", "结构"), ("底板", "结构"), ("消力池", "结构"), ("反滤层", "结构"), ("海漫", "结构"),
          ("防冲槽", "结构"), ("高程", "高程"), ("闸门", "上部结构"), ("排架", "上部结构"),
          ("工作桥", "上部结构"), ("交通桥", "上部结构"), ("启闭机房", "上部结构"), ("绘图选项", "其他")]


PLAN_PREFIX = 'plan.'   # 平面图参数在同一份存档里加前缀，避免与纵剖面图撞名


@app.route('/drawing', methods=['GET'])
def drawing_page():
    u = current_user()
    # 回填该用户上次保存的参数
    saved = load_user_params(u['id'])
    saved_plan = {k[len(PLAN_PREFIX):]: v for k, v in saved.items() if k.startswith(PLAN_PREFIX)}
    fields = [(k, n, g, saved.get(k, d), un) for k, n, g, d, un in FIELDS]
    plan_fields = [(k, n, g, saved_plan.get(k, d), un) for k, n, g, d, un in gp.FIELDS]
    return render_template('drawing.html', fields=fields, groups=GROUPS,
                           plan_fields=plan_fields, plan_groups=gp.GROUPS,
                           username=u['username'], is_admin=is_admin(u))


def _parse_section_params(data):
    """纵剖面图：表单填米，内部存 mm；比率与文本原样"""
    p = dict(gd.P)
    saved = {}
    for key, *_ in FIELDS:
        if key not in data or data[key] in ("", None):
            continue
        if key in TEXT_KEYS:
            p[key] = str(data[key])
            saved[key] = str(data[key])
        elif key in RATIO_KEYS:
            p[key] = float(data[key])
            saved[key] = float(data[key])
        else:
            p[key] = float(data[key]) * 1000.0
            saved[key] = float(data[key])
    err = None
    for key in ("pg_len", "db_len", "xl_len", "hm_total"):
        if p.get(key, 0) <= 0:
            err = f"{key} 必须大于 0"
            break
    return p, saved, err


def _parse_plan_params(data):
    """平面布置图：表单填米，模块内部再 ×1000；勾选项为 0/1；可留空项保持空串"""
    p = dict(gp.P)
    saved = {}
    for key, name, group, default, unit in gp.FIELDS:
        if key not in data:
            continue
        v = data[key]
        if unit == "勾选":
            val = 1 if v in (1, '1', True, 'true', 'on') else 0
        elif unit == "文本":
            val = str(v)
        elif key in gp.OPT_KEYS and (v == "" or v is None):
            val = ""                     # 留空 = 按公式自动
        elif key == "nBay":
            val = int(float(v))
        else:
            val = float(v)
        p[key] = val
        saved[key] = val
    err = None
    if p["nBay"] < 1:
        err = "闸孔数至少为 1"
    elif p["riverW"] <= 0 or p["bayW"] <= 0 or p["gateL"] <= 0:
        err = "河道宽、单孔净宽、闸室顺流长必须大于 0"
    elif p["slope"] <= 0 or p["hmN"] <= 0 or p["fcN"] <= 0:
        err = "坡比必须大于 0"
    elif p["fcX"] <= 0 or p["fcY"] <= 0:
        err = "防冲槽槽底顺流宽 / 横向长必须大于 0"
    return p, saved, err


@app.route('/drawing/generate', methods=['POST'])
def drawing_generate():
    u = current_user()
    try:
        data = request.get_json(force=True, silent=True) or {}
    except Exception:
        return jsonify({"ok": False, "error": "请求格式错误"}), 400

    # kind: section=纵剖面图 / plan=平面布置图
    kind = (data.get('kind') or 'section').strip()
    if kind not in ('section', 'plan'):
        kind = 'section'

    # 先解析 + 校验，再扣次数（填错参数不该白扣一次）
    try:
        if kind == 'plan':
            p, saved, err = _parse_plan_params(data)
            default_title = '水闸枢纽平面布置图'
        else:
            p, saved, err = _parse_section_params(data)
            default_title = '水闸纵剖面图'
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "存在无效数值，请检查填写内容"}), 400
    if err:
        return jsonify({"ok": False, "error": err}), 400

    # ===== 付费墙：检查图纸剩余次数 =====
    if not consume_quota(u['id'], 'drawing'):
        return jsonify({
            "ok": False,
            "need_pay": True,
            "error": "图纸生成次数不足，请先付费开通。",
            "pay": get_pay_config(),
        }), 402

    # 图纸文件按用户分目录存储
    user_dir = os.path.join(DRAW_OUT, str(u['id']))
    os.makedirs(user_dir, exist_ok=True)
    stamp = int(time.time())
    base = f"sluice_{stamp}"
    dxf_path = os.path.join(user_dir, base + ".dxf")
    svg_path = os.path.join(user_dir, base + ".svg")
    try:
        mod = gp if kind == 'plan' else gd
        mod.generate_dxf(p, dxf_path)
        mod.generate_svg(p, svg_path)
    except Exception as e:
        return jsonify({"ok": False, "error": f"生成失败: {e}"}), 500

    # 保存该用户的参数（下次自动回填）；两张图共用一份存档，各改各的键
    try:
        allp = load_user_params(u['id'])
        if kind == 'plan':
            for k in [k for k in allp if k.startswith(PLAN_PREFIX)]:
                allp.pop(k)
            for k, v in saved.items():
                allp[PLAN_PREFIX + k] = v
        else:
            for k in [k for k in allp if not k.startswith(PLAN_PREFIX)]:
                allp.pop(k)
            allp.update(saved)
        save_user_params(u['id'], allp)
    except Exception as e:
        print('[drawing] 参数保存失败:', e)

    title = str(p.get('title') or default_title).strip() or default_title

    # 自动记录生成历史（供"我的生成记录"回看/下载）
    try:
        conn = get_db()
        conn.execute(
            'INSERT INTO generation_records (user_id, username, type, title, file_url, file2_url) VALUES (?,?,?,?,?,?)',
            (u['id'], u['username'], 'drawing', title,
             f"/drawing/dl/{u['id']}/{os.path.basename(dxf_path)}",
             f"/drawing/dl/{u['id']}/{os.path.basename(svg_path)}"))
        conn.commit()
        conn.close()
    except Exception as e:
        print('[drawing] 记录保存失败:', e)

    return jsonify({
        "ok": True,
        "name": title,
        "dxf_url": f"/drawing/dl/{u['id']}/{os.path.basename(dxf_path)}",
        "svg_url": f"/drawing/dl/{u['id']}/{os.path.basename(svg_path)}",
    })


@app.route('/drawing/dl/<int:uid>/<path:filename>')
def drawing_download(uid, filename):
    u = current_user()
    # 只能下载自己的图纸
    if u is None or u['id'] != uid:
        return jsonify({"ok": False, "error": "无权限"}), 403
    return send_from_directory(os.path.join(DRAW_OUT, str(uid)), filename)


@app.route('/<path:path>')
def static_files(path):
    return send_from_directory(BASE, path)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))  # Railway 会自动注入 PORT
    print('水闸毕业设计论文生成系统（多用户版）已启动: http://0.0.0.0:%d' % port)
    print('登录: /login   注册: /register')
    app.run(host='0.0.0.0', port=port, debug=False)

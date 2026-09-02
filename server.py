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
                   render_template, redirect, url_for, session, flash)
from werkzeug.security import generate_password_hash, check_password_hash
import generate_thesis as gt
import generate_drawing as gd

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
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            username TEXT,
            content TEXT NOT NULL,
            is_admin INTEGER DEFAULT 0,
            to_user_id INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now','localtime')),
            is_read INTEGER DEFAULT 0
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
        # messages 表迁移（旧库可能没有 to_user_id）
        mcols = [r[1] for r in conn.execute('PRAGMA table_info(messages)').fetchall()]
        if 'to_user_id' not in mcols:
            conn.execute('ALTER TABLE messages ADD COLUMN to_user_id INTEGER DEFAULT 0')
        # 收费配置（settings 表）
        defaults = [
            ('drawing_price', '5'), ('thesis_price', '10'),
            ('pay_note', '扫码支付后，请联系管理员（微信/QQ 私聊）确认到账，由管理员为您开通对应次数。'),
            ('wechat_qr', ''), ('alipay_qr', ''),
            ('register_drawing_bonus', '0'), ('register_thesis_bonus', '0'),
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


def get_pay_config():
    """返回收费配置（价格 / 收款码 / 说明）"""
    return {
        'drawing_price': get_setting('drawing_price', '5'),
        'thesis_price': get_setting('thesis_price', '10'),
        'pay_note': get_setting('pay_note', ''),
        'wechat_qr': get_setting('wechat_qr', ''),
        'alipay_qr': get_setting('alipay_qr', ''),
    }


def get_quota(uid):
    conn = get_db()
    u = conn.execute('SELECT drawing_quota, thesis_quota FROM users WHERE id=?', (uid,)).fetchone()
    conn.close()
    if not u:
        return 0, 0
    return u['drawing_quota'], u['thesis_quota']


def consume_quota(uid, kind):
    """扣减一次次数。kind: 'drawing' 或 'thesis'。返回 True=扣减成功，False=次数不足"""
    col = 'drawing_quota' if kind == 'drawing' else 'thesis_quota'
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
    col = 'drawing_quota' if kind == 'drawing' else 'thesis_quota'
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
    if path in ('/', '/tool', '/verify', '/model', '/thesis', '/drawing', '/generate', '/drawing/generate', '/drawing/dl'):
        if not current_user():
            if path == '/drawing/generate' or path == '/drawing/dl':
                return jsonify({"ok": False, "error": "请先登录"}), 401
            return redirect(url_for('login'))
    elif path.startswith('/drawing/dl/'):
        if not current_user():
            return jsonify({"ok": False, "error": "请先登录"}), 401


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
    conn.execute('INSERT INTO users (username, password_hash, drawing_quota, thesis_quota) VALUES (?, ?, ?, ?)',
                 (username, generate_password_hash(password),
                  int(get_setting('register_drawing_bonus', '0') or 0),
                  int(get_setting('register_thesis_bonus', '0') or 0)))
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
            kind = request.form.get('kind')  # drawing / thesis
            num = request.form.get('num') or '1'
            try:
                num = int(num)
            except ValueError:
                num = 1
            if num > 0 and kind in ('drawing', 'thesis'):
                add_quota(uid, kind, num)
                flash(f'已为用户增加 {kind == "drawing" and "图纸" or "论文"}次数 × {num}')
        elif action == 'update_pay':
            # 保存价格与付款说明
            conn = get_db()
            for k in ('drawing_price', 'thesis_price', 'pay_note'):
                v = request.form.get(k, '')
                if k in ('drawing_price', 'thesis_price'):
                    try:
                        v = str(max(0, int(float(v))))
                    except (ValueError, TypeError):
                        v = '5' if k == 'drawing_price' else '10'
                conn.execute("UPDATE settings SET value=? WHERE key=?", (v, k))
            conn.commit()
            conn.close()
            flash('收费设置已保存')
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
    users = conn.execute('SELECT id, username, created_at, drawing_quota, thesis_quota FROM users ORDER BY id').fetchall()
    conn.close()
    return render_template('admin.html', invite_code=get_invite_code(), users=users, pay=get_pay_config())


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
    dq, tq = get_quota(u['id'])
    return jsonify({
        "logged": True,
        "username": u['username'],
        "is_admin": is_admin(u),
        "drawing_quota": dq,
        "thesis_quota": tq,
        "pay": get_pay_config(),
    })


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
        drawing = max(0, int(data.get('drawing', 0) or 0))
        thesis = max(0, int(data.get('thesis', 0) or 0))
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "次数无效"}), 400
    if drawing == 0 and thesis == 0:
        return jsonify({"ok": False, "error": "请填写要开通的次数"}), 400
    conn = get_db()
    row = conn.execute('SELECT * FROM payment_requests WHERE id=? AND status=?', (rid, 'pending')).fetchone()
    if not row:
        conn.close()
        return jsonify({"ok": False, "error": "申请不存在或已处理"}), 400
    if drawing:
        add_quota(row['user_id'], 'drawing', drawing)
    if thesis:
        add_quota(row['user_id'], 'thesis', thesis)
    conn.execute("UPDATE payment_requests SET status='done', processed_at=datetime('now','localtime') WHERE id=?", (rid,))
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "msg": f"已为用户 {row['username']} 开通图纸×{drawing}、论文×{thesis}"})


# ============================================================
# 客服消息（站内信：客户留言，管理员回复）
# ============================================================
@app.route('/api/msg/send', methods=['POST'])
def msg_send():
    u = current_user()
    if not u:
        return jsonify({"ok": False, "error": "请先登录"}), 401
    data = request.get_json(force=True, silent=True) or {}
    content = str(data.get('content', '')).strip()
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
    conn.execute('INSERT INTO messages (user_id, username, content, is_admin, to_user_id) VALUES (?,?,?,?,?)',
                 (u['id'], u['username'], content, is_admin_flag, to_user_id))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


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
        (u['id'], u['username'], 'thesis', title, f"/records/dl/{u['id']}/{fname}"))
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
        rows = conn.execute('SELECT * FROM generation_records ORDER BY id DESC LIMIT 200').fetchall()
    else:
        rows = conn.execute('SELECT * FROM generation_records WHERE user_id=? ORDER BY id DESC LIMIT 100',
                            (u['id'],)).fetchall()
    conn.close()
    return jsonify({"ok": True, "list": [dict(r) for r in rows]})


@app.route('/records/dl/<int:uid>/<path:filename>')
def records_download(uid, filename):
    u = current_user()
    if u is None or u['id'] != uid:
        return jsonify({"ok": False, "error": "无权限"}), 403
    return send_from_directory(os.path.join(THESIS_OUT, str(uid)), filename)


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


@app.route('/generate', methods=['POST'])
def generate():
    params = request.get_json(force=True, silent=True) or {}
    try:
        buf = gt.generate(params)   # 返回 BytesIO（不落盘，避免权限/占用问题）
        return send_file(buf, as_attachment=True, download_name='毕业设计论文.docx',
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


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


@app.route('/drawing', methods=['GET'])
def drawing_page():
    u = current_user()
    # 回填该用户上次保存的参数
    saved = load_user_params(u['id'])
    fields = []
    for key, name, group, default, unit in FIELDS:
        val = saved.get(key, default)
        fields.append((key, name, group, val, unit))
    return render_template('drawing.html', fields=fields, groups=GROUPS, username=u['username'], is_admin=is_admin(u))


@app.route('/drawing/generate', methods=['POST'])
def drawing_generate():
    u = current_user()
    # ===== 付费墙：检查图纸剩余次数 =====
    if not consume_quota(u['id'], 'drawing'):
        return jsonify({
            "ok": False,
            "need_pay": True,
            "error": "图纸生成次数不足，请先付费开通。",
            "pay": get_pay_config(),
        }), 402
    try:
        data = request.get_json(force=True, silent=True) or {}
    except Exception:
        return jsonify({"ok": False, "error": "请求格式错误"}), 400

    p = dict(gd.P)
    saved = {}
    for key, *_ in FIELDS:
        if key not in data or data[key] in ("", None):
            continue
        try:
            if key in TEXT_KEYS:
                p[key] = str(data[key])
                saved[key] = str(data[key])
            elif key in RATIO_KEYS:
                p[key] = float(data[key])
                saved[key] = float(data[key])
            else:
                p[key] = float(data[key]) * 1000.0
                saved[key] = float(data[key])
        except (ValueError, TypeError):
            return jsonify({"ok": False, "error": f"字段 {key} 数值无效"}), 400

    for key in ("pg_len", "db_len", "xl_len", "hm_total"):
        if p.get(key, 0) <= 0:
            return jsonify({"ok": False, "error": f"{key} 必须大于 0"}), 400

    # 图纸文件按用户分目录存储
    user_dir = os.path.join(DRAW_OUT, str(u['id']))
    os.makedirs(user_dir, exist_ok=True)
    stamp = int(time.time())
    base = f"sluice_{stamp}"
    dxf_path = os.path.join(user_dir, base + ".dxf")
    svg_path = os.path.join(user_dir, base + ".svg")
    try:
        gd.generate_dxf(p, dxf_path)
        gd.generate_svg(p, svg_path)
    except Exception as e:
        return jsonify({"ok": False, "error": f"生成失败: {e}"}), 500

    # 保存该用户的参数（下次自动回填）
    save_user_params(u['id'], saved)

    # 自动记录生成历史（供"我的生成记录"回看/下载）
    try:
        conn = get_db()
        conn.execute(
            'INSERT INTO generation_records (user_id, username, type, title, file_url, file2_url) VALUES (?,?,?,?,?,?)',
            (u['id'], u['username'], 'drawing', p.get('title', '水闸纵剖面图'),
             f"/drawing/dl/{u['id']}/{os.path.basename(dxf_path)}",
             f"/drawing/dl/{u['id']}/{os.path.basename(svg_path)}"))
        conn.commit()
        conn.close()
    except Exception as e:
        print('[drawing] 记录保存失败:', e)

    return jsonify({
        "ok": True,
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

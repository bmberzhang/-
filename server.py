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
    ''')
    conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('invite_code', ?)",
                 (os.environ.get('INVITE_CODE', 'sluice2026'),))
    conn.commit()
    conn.close()


init_db()


def get_invite_code():
    conn = get_db()
    row = conn.execute("SELECT value FROM settings WHERE key='invite_code'").fetchone()
    conn.close()
    return row['value'] if row else 'sluice2026'


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
    if path in ('/', '/drawing', '/generate', '/drawing/generate', '/drawing/dl'):
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
    conn.execute('INSERT INTO users (username, password_hash) VALUES (?, ?)',
                 (username, generate_password_hash(password)))
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
        return redirect(url_for('admin'))
    conn = get_db()
    users = conn.execute('SELECT id, username, created_at FROM users ORDER BY id').fetchall()
    conn.close()
    return render_template('admin.html', invite_code=get_invite_code(), users=users)


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


@app.route('/api/me')
def api_me():
    u = current_user()
    if not u:
        return jsonify({"logged": False}), 401
    return jsonify({"logged": True, "username": u['username']})


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
    ("gate1_x", "检修门距底板左端", "闸门", 0.65, "m"), ("gate1_w", "检修闸门宽", "闸门", 0.3, "m"),
    ("gate2_gap", "工作门距检修门", "闸门", 1.9, "m"), ("gate2_w", "工作闸门宽", "闸门", 0.8, "m"),
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
    return render_template('drawing.html', fields=fields, groups=GROUPS, username=u['username'])


@app.route('/drawing/generate', methods=['POST'])
def drawing_generate():
    u = current_user()
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

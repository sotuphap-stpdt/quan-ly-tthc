import time
time.sleep(2)  # Chờ 3 giây trước khi khởi động
import os
import sqlite3
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, jsonify, session
from datetime import datetime
import uuid
from werkzeug.utils import secure_filename
import json
import difflib
import re
import openpyxl
from openpyxl import load_workbook
from functools import wraps

app = Flask(__name__)
app.secret_key = 'sotuphap_dongthap_secret_key_2026'

UPLOAD_FOLDER = 'uploads'
QUYET_DINH_FOLDER = 'uploads/quyet_dinh'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['QUYET_DINH_FOLDER'] = QUYET_DINH_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(QUYET_DINH_FOLDER, exist_ok=True)

# ==================== DECORATOR KIỂM TRA ĐĂNG NHẬP ====================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session or session.get('role') != 'admin':
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

# ==================== HÀM LẤY CƠ QUAN THEO CẤP ====================
def get_co_quan_by_cap(cap_thuc_hien):
    cap_map = {
        'bo': 'Bộ Tư pháp',
        'tinh': 'Sở Tư pháp',
        'xa': 'UBND cấp xã',
        'dung_chung': 'UBND cấp xã, các Phòng công chứng',
        'lien_thong': 'UBND cấp xã, Sở Tư pháp'
    }
    return cap_map.get(cap_thuc_hien, '')

# ==================== KHỞI TẠO DATABASE ====================
def init_db():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    c.execute('''CREATE TABLE IF NOT EXISTS quyet_dinh (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        so_quyet_dinh TEXT UNIQUE,
        ten_quyet_dinh TEXT,
        ngay_ban_hanh DATE,
        loai TEXT,
        mo_ta TEXT,
        file_dinh_kem TEXT,
        ten_file_goc TEXT,
        noi_dung_ocr TEXT
    )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS tthc (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ma_tthc TEXT UNIQUE,
        ten_tthc TEXT,
        linh_vuc TEXT,
        phi TEXT,
        le_phi TEXT,
        lien_thong_cung_cap TEXT,
        lien_thong_02_cap TEXT,
        phi_dia_gioi TEXT,
        dvc_toan_trinh TEXT,
        dvc_mot_phan TEXT,
        dvc_cung_cap_tt TEXT,
        dich_vu_bcci TEXT,
        ghi_chu TEXT,
        cap_thuc_hien TEXT,
        trang_thai_cong_khai TEXT,
        so_quyet_dinh TEXT,
        co_quan_thuc_hien TEXT,
        trang_thai TEXT DEFAULT 'da_cong_bo'
    )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS quy_trinh (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tthc_id INTEGER,
        buoc INTEGER,
        ten_buoc TEXT,
        thoi_gian_xu_ly TEXT,
        co_quan_thuc_hien TEXT
    )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT,
        role TEXT DEFAULT 'user',
        fullname TEXT,
        email TEXT,
        created_at DATE
    )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS user_linh_vuc (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        linh_vuc TEXT,
        FOREIGN KEY(user_id) REFERENCES users(id)
    )''')
    
    import hashlib
    c.execute("SELECT COUNT(*) FROM users WHERE username = 'Admin'")
    if c.fetchone()[0] == 0:
        admin_password = hashlib.sha256('admin@123'.encode()).hexdigest()
        c.execute("INSERT INTO users (username, password, role, fullname, email, created_at) VALUES (?,?,?,?,?,?)",
                  ('Admin', admin_password, 'admin', 'Quản trị viên', 'admin@dongthap.gov.vn', datetime.now().date()))
    
    try:
        c.execute("ALTER TABLE tthc ADD COLUMN trang_thai TEXT DEFAULT 'da_cong_bo'")
    except:
        pass
    
    conn.commit()
    conn.close()

init_db()

# ==================== KIỂM TRA QUYỀN ====================
def check_linh_vuc_access(user_id, linh_vuc):
    if session.get('role') == 'admin':
        return True
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM user_linh_vuc WHERE user_id = ? AND linh_vuc = ?", (user_id, linh_vuc))
    result = c.fetchone()[0] > 0
    conn.close()
    return result

# ==================== ĐĂNG NHẬP ====================
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        
        import hashlib
        hashed_password = hashlib.sha256(password.encode()).hexdigest()
        
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        c.execute("SELECT id, username, role FROM users WHERE username = ? AND password = ?", (username, hashed_password))
        user = c.fetchone()
        conn.close()
        
        if user:
            session['user_id'] = user[0]
            session['username'] = user[1]
            session['role'] = user[2]
            return redirect(url_for('dashboard'))
        else:
            return render_template('login.html', error='Sai tên đăng nhập hoặc mật khẩu!')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ==================== QUẢN LÝ NGƯỜI DÙNG ====================
@app.route('/users')
@admin_required
def danh_sach_users():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT id, username, role, fullname, email, created_at FROM users ORDER BY id")
    users = c.fetchall()
    conn.close()
    return render_template('users.html', users=users)

@app.route('/users/them', methods=['GET', 'POST'])
@admin_required
def them_user():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        role = request.form.get('role', 'user')
        fullname = request.form.get('fullname', '').strip()
        email = request.form.get('email', '').strip()
        linh_vuc_list = request.form.getlist('linh_vuc')
        
        import hashlib
        hashed_password = hashlib.sha256(password.encode()).hexdigest()
        
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        
        try:
            c.execute("INSERT INTO users (username, password, role, fullname, email, created_at) VALUES (?,?,?,?,?,?)",
                      (username, hashed_password, role, fullname, email, datetime.now().date()))
            user_id = c.lastrowid
            
            for lv in linh_vuc_list:
                if lv:
                    c.execute("INSERT INTO user_linh_vuc (user_id, linh_vuc) VALUES (?,?)", (user_id, lv))
            
            conn.commit()
            conn.close()
            return '<script>alert("✅ Thêm người dùng thành công!"); window.location.href="/users";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ Lỗi: {str(e)}"); window.history.back();</script>'
    
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT DISTINCT linh_vuc FROM tthc WHERE linh_vuc IS NOT NULL AND linh_vuc != ''")
    linh_vuc_list = [row[0] for row in c.fetchall()]
    if not linh_vuc_list:
        linh_vuc_list = ['Kinh doanh', 'Hộ tịch', 'Xây dựng', 'Công chứng', 'Đất đai', 'Môi trường', 'Lao động', 'Thông tin']
    conn.close()
    return render_template('user_form.html', user=None, linh_vuc_list=linh_vuc_list)

@app.route('/users/sua/<int:id>', methods=['GET', 'POST'])
@admin_required
def sua_user(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        role = request.form.get('role', 'user')
        fullname = request.form.get('fullname', '').strip()
        email = request.form.get('email', '').strip()
        new_password = request.form.get('new_password', '').strip()
        linh_vuc_list = request.form.getlist('linh_vuc')
        
        import hashlib
        
        if new_password:
            hashed_password = hashlib.sha256(new_password.encode()).hexdigest()
            c.execute("UPDATE users SET username=?, password=?, role=?, fullname=?, email=? WHERE id=?",
                      (username, hashed_password, role, fullname, email, id))
        else:
            c.execute("UPDATE users SET username=?, role=?, fullname=?, email=? WHERE id=?",
                      (username, role, fullname, email, id))
        
        c.execute("DELETE FROM user_linh_vuc WHERE user_id=?", (id,))
        for lv in linh_vuc_list:
            if lv:
                c.execute("INSERT INTO user_linh_vuc (user_id, linh_vuc) VALUES (?,?)", (id, lv))
        
        conn.commit()
        conn.close()
        return '<script>alert("✅ Cập nhật người dùng thành công!"); window.location.href="/users";</script>'
    
    c.execute("SELECT id, username, role, fullname, email FROM users WHERE id=?", (id,))
    user = c.fetchone()
    c.execute("SELECT linh_vuc FROM user_linh_vuc WHERE user_id=?", (id,))
    user_linh_vuc = [row[0] for row in c.fetchall()]
    c.execute("SELECT DISTINCT linh_vuc FROM tthc WHERE linh_vuc IS NOT NULL AND linh_vuc != ''")
    linh_vuc_list = [row[0] for row in c.fetchall()]
    if not linh_vuc_list:
        linh_vuc_list = ['Kinh doanh', 'Hộ tịch', 'Xây dựng', 'Công chứng', 'Đất đai', 'Môi trường', 'Lao động', 'Thông tin']
    conn.close()
    return render_template('user_form.html', user=user, linh_vuc_list=linh_vuc_list, user_linh_vuc=user_linh_vuc)

@app.route('/users/xoa/<int:id>')
@admin_required
def xoa_user(id):
    if id == session.get('user_id'):
        return '<script>alert("❌ Không thể xóa chính mình!"); window.location.href="/users";</script>'
    
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("DELETE FROM user_linh_vuc WHERE user_id=?", (id,))
    c.execute("DELETE FROM users WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return '<script>alert("🗑️ Đã xóa người dùng!"); window.location.href="/users";</script>'

# ==================== TRANG CHỦ ====================
@app.route('/')
@login_required
def dashboard():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    thong_ke = {
        'cap_bo': 0, 'cap_tinh': 0, 'cap_xa': 0, 'cap_dung_chung': 0, 'cap_lien_thong': 0,
        'trang_thai_toan_trinh': 0, 'trang_thai_cong_khai_mot_phan': 0, 'trang_thai_chua_cong_khai': 0,
        'so_da_cong_bo': 0, 'so_bai_bo': 0, 'linh_vuc': []
    }
    
    try:
        for cap in ['bo', 'tinh', 'xa', 'dung_chung', 'lien_thong']:
            c.execute("SELECT COUNT(*) FROM tthc WHERE cap_thuc_hien=?", (cap,))
            result = c.fetchone()
            if result and result[0]:
                thong_ke[f'cap_{cap}'] = result[0]
    except: pass
    
    try:
        c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_toan_trinh = 'X'")
        thong_ke['trang_thai_toan_trinh'] = c.fetchone()[0] or 0
        c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_mot_phan = 'X'")
        thong_ke['trang_thai_cong_khai_mot_phan'] = c.fetchone()[0] or 0
        c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_cung_cap_tt = 'X'")
        thong_ke['trang_thai_chua_cong_khai'] = c.fetchone()[0] or 0
    except: pass
    
    try:
        c.execute("SELECT COUNT(*) FROM tthc WHERE trang_thai = 'da_cong_bo'")
        thong_ke['so_da_cong_bo'] = c.fetchone()[0] or 0
        c.execute("SELECT COUNT(*) FROM tthc WHERE trang_thai = 'bai_bo'")
        thong_ke['so_bai_bo'] = c.fetchone()[0] or 0
    except: pass
    
    try:
        c.execute("SELECT linh_vuc, COUNT(*) FROM tthc WHERE linh_vuc IS NOT NULL AND linh_vuc != '' GROUP BY linh_vuc ORDER BY COUNT(*) DESC")
        raw_data = c.fetchall()
        if raw_data:
            thong_ke['linh_vuc'] = [{'ten': str(row[0]), 'so_luong': row[1]} for row in raw_data]
    except: pass
    
    conn.close()
    return render_template('dashboard.html', thong_ke=thong_ke, username=session.get('username'), role=session.get('role'))

# ==================== API ====================
@app.route('/api/quyet_dinh')
@login_required
def api_quyet_dinh():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT id, so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    data = [{'id': row[0], 'so_quyet_dinh': row[1], 'ten_quyet_dinh': row[2], 'ngay_ban_hanh': row[3], 'loai': row[4]} for row in c.fetchall()]
    conn.close()
    return jsonify(data)

# ==================== TẢI FILE MẪU ====================
@app.route('/tai_file_mau')
@login_required
def tai_file_mau():
    import io
    from flask import send_file
    
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    headers = ['STT', 'MÃ TTHC', 'Tên TTHC', 'Lĩnh vực', 'Phí', 'Lệ Phí', 
               'Cấp thực hiện', 'Cơ quan thực hiện', 'Cùng cấp', '02 cấp', 
               'Phi địa giới', 'Toàn trình', 'Một phần', 'Cung cấp thông tin', 
               'Dịch vụ BCCI', 'Quyết định số', 'TRẠNG THÁI', 'Ghi chú']
    
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)
    
    mau_data = [
        [1, 'TTHC001', 'Đăng ký kinh doanh', 'Kinh doanh', 'X', '', 'tinh', 'Sở Kế hoạch Đầu tư', 'X', '', '', 'X', '', '', 'Có', '123/QĐ-UBND', 'Đã công bố', ''],
        [2, 'TTHC002', 'Chứng thực bản sao', 'Hộ tịch', '', 'X', 'xa', 'UBND cấp xã', '', 'X', '', '', 'X', '', '', '456/QĐ-UBND', 'Đã công bố', ''],
        [3, 'TTHC003', 'Cấp giấy phép xây dựng', 'Xây dựng', 'X', 'X', 'lien_thong', 'UBND cấp xã, Sở Tư pháp', '', '', 'X', 'X', '', '', '', '789/QĐ-UBND', 'Đã công bố', ''],
        [4, 'TTHC004', 'Bổ nhiệm công chứng viên', 'Công chứng', 'X', '', 'bo', 'Bộ Tư pháp', '', '', '', 'X', '', '', 'Có', '111/QĐ-BTP', 'Bãi bỏ', ''],
        [5, 'TTHC005', 'Đăng ký hộ tịch', 'Hộ tịch', '', 'X', 'xa', 'UBND cấp xã', '', '', '', '', 'X', '', '', '222/QĐ-UBND', 'Đã công bố', '']
    ]
    
    for row_idx, row_data in enumerate(mau_data, 2):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Mau_Import_TTHC.xlsx')

# ==================== DANH SÁCH TTHC ====================
@app.route('/tthc')
@login_required
def danh_sach_tthc():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    if session.get('role') == 'admin':
        c.execute('''SELECT id, ma_tthc, ten_tthc, linh_vuc, phi, le_phi, cap_thuc_hien, co_quan_thuc_hien,
                    lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                    dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt,
                    dich_vu_bcci, ghi_chu, so_quyet_dinh, trang_thai
                    FROM tthc ORDER BY id DESC''')
    else:
        c.execute("SELECT linh_vuc FROM user_linh_vuc WHERE user_id=?", (session['user_id'],))
        allowed_linh_vuc = [row[0] for row in c.fetchall()]
        if allowed_linh_vuc:
            placeholders = ','.join(['?'] * len(allowed_linh_vuc))
            c.execute(f'''SELECT id, ma_tthc, ten_tthc, linh_vuc, phi, le_phi, cap_thuc_hien, co_quan_thuc_hien,
                        lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                        dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt,
                        dich_vu_bcci, ghi_chu, so_quyet_dinh, trang_thai
                        FROM tthc WHERE linh_vuc IN ({placeholders}) ORDER BY id DESC''', allowed_linh_vuc)
        else:
            c.execute("SELECT id, ma_tthc, ten_tthc, linh_vuc, phi, le_phi, cap_thuc_hien, co_quan_thuc_hien, lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi, dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, dich_vu_bcci, ghi_chu, so_quyet_dinh, trang_thai FROM tthc WHERE 1=0")
    
    data = c.fetchall()
    conn.close()
    return render_template('tthc_list.html', tthcs=data, username=session.get('username'), role=session.get('role'))

# ==================== THÊM TTHC ====================
@app.route('/tthc/them', methods=['GET', 'POST'])
@login_required
def them_tthc():
    if request.method == 'POST':
        if 'file_excel' in request.files:
            file = request.files['file_excel']
            if file and file.filename != '' and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
                return import_excel_file(file)
        
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        
        ma_tthc = request.form.get('ma_tthc', '').strip()
        ten_tthc = request.form.get('ten_tthc', '').strip()
        linh_vuc = request.form.get('linh_vuc', '').strip()
        phi = 'X' if request.form.get('phi') == 'on' else ''
        le_phi = 'X' if request.form.get('le_phi') == 'on' else ''
        lien_thong_cung_cap = 'X' if request.form.get('lien_thong_cung_cap') == 'on' else ''
        lien_thong_02_cap = 'X' if request.form.get('lien_thong_02_cap') == 'on' else ''
        phi_dia_gioi = 'X' if request.form.get('phi_dia_gioi') == 'on' else ''
        dvc_toan_trinh = 'X' if request.form.get('dvc_toan_trinh') == 'on' else ''
        dvc_mot_phan = 'X' if request.form.get('dvc_mot_phan') == 'on' else ''
        dvc_cung_cap_tt = 'X' if request.form.get('dvc_cung_cap_tt') == 'on' else ''
        dich_vu_bcci = request.form.get('dich_vu_bcci', '').strip()
        ghi_chu = request.form.get('ghi_chu', '').strip()
        cap_thuc_hien = request.form.get('cap_thuc_hien', '').strip()
        trang_thai_cong_khai = request.form.get('trang_thai_cong_khai', '').strip()
        so_qd = request.form.get('so_quyet_dinh', '').strip()
        co_quan = request.form.get('co_quan_thuc_hien', '').strip()
        trang_thai_tt = request.form.get('trang_thai_tt', 'da_cong_bo')
        
        if not co_quan and cap_thuc_hien:
            co_quan = get_co_quan_by_cap(cap_thuc_hien)
        
        try:
            c.execute("SELECT id FROM tthc WHERE ma_tthc = ?", (ma_tthc,))
            existing = c.fetchone()
            
            if existing:
                c.execute('''UPDATE tthc SET ten_tthc=?, linh_vuc=?, phi=?, le_phi=?, 
                    lien_thong_cung_cap=?, lien_thong_02_cap=?, phi_dia_gioi=?,
                    dvc_toan_trinh=?, dvc_mot_phan=?, dvc_cung_cap_tt=?, 
                    dich_vu_bcci=?, ghi_chu=?, cap_thuc_hien=?, trang_thai_cong_khai=?, 
                    so_quyet_dinh=?, co_quan_thuc_hien=?, trang_thai=? WHERE ma_tthc = ?''',
                    (ten_tthc, linh_vuc, phi, le_phi, lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                     dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, dich_vu_bcci, ghi_chu,
                     cap_thuc_hien, trang_thai_cong_khai, so_qd, co_quan, trang_thai_tt, ma_tthc))
            else:
                c.execute('''INSERT INTO tthc (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, 
                    lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                    dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, dich_vu_bcci, ghi_chu,
                    cap_thuc_hien, trang_thai_cong_khai, so_quyet_dinh, co_quan_thuc_hien, trang_thai)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                    (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                     dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, dich_vu_bcci, ghi_chu,
                     cap_thuc_hien, trang_thai_cong_khai, so_qd, co_quan, trang_thai_tt))
            
            conn.commit()
            conn.close()
            return '<script>alert("✅ THÊM THÀNH CÔNG!"); window.location.href="/tthc";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ LỖI: {str(e)}"); window.history.back();</script>'
    
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT id, so_quyet_dinh FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    quyet_dinh_list = c.fetchall()
    conn.close()
    return render_template('tthc_form.html', tthc=None, quyet_dinh_list=quyet_dinh_list, username=session.get('username'), role=session.get('role'))

def import_excel_file(file):
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        wb = load_workbook(filepath, data_only=True)
        ws = wb.active
        
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        
        c.execute("SELECT COUNT(*) FROM quyet_dinh")
        if c.fetchone()[0] == 0:
            c.execute("INSERT INTO quyet_dinh (so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai, mo_ta) VALUES (?,?,?,?,?)",
                      ("QD001", "Quyết định công bố mặc định", datetime.now().date(), "cong_bo", "Tự động tạo khi import"))
        
        headers = [str(cell.value).strip() if cell.value else '' for cell in ws[1]]
        so_luong_thanh_cong = 0
        danh_sach_loi = []
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or not any(cell for cell in row):
                continue
            
            try:
                row_dict = {}
                for i, header in enumerate(headers):
                    if i < len(row) and row[i] is not None:
                        row_dict[header] = str(row[i]).strip()
                    else:
                        row_dict[header] = ''
                
                ma_tthc = row_dict.get('MÃ TTHC', '').strip()
                if not ma_tthc or ma_tthc == 'None':
                    danh_sach_loi.append(f"Dòng {row_idx}: Thiếu MÃ TTHC")
                    continue
                
                ten_tthc = row_dict.get('Tên TTHC', '').strip()
                linh_vuc = row_dict.get('Lĩnh vực', '').strip()
                phi = 'X' if row_dict.get('Phí', '').upper() == 'X' else ''
                le_phi = 'X' if row_dict.get('Lệ Phí', '').upper() == 'X' else ''
                cap_thuc_hien = row_dict.get('Cấp thực hiện', '').strip()
                co_quan = row_dict.get('Cơ quan thực hiện', '').strip()
                lien_thong_cung_cap = 'X' if row_dict.get('Cùng cấp', '').upper() == 'X' else ''
                lien_thong_02_cap = 'X' if row_dict.get('02 cấp', '').upper() == 'X' else ''
                phi_dia_gioi = 'X' if row_dict.get('Phi địa giới', '').upper() == 'X' else ''
                dvc_toan_trinh = 'X' if row_dict.get('Toàn trình', '').upper() == 'X' else ''
                dvc_mot_phan = 'X' if row_dict.get('Một phần', '').upper() == 'X' else ''
                dvc_cung_cap_tt = 'X' if row_dict.get('Cung cấp thông tin', '').upper() == 'X' else ''
                dich_vu_bcci = row_dict.get('Dịch vụ BCCI', '').strip()
                so_qd = row_dict.get('Quyết định số', '').strip()
                ghi_chu = row_dict.get('Ghi chú', '').strip()
                trang_thai_excel = row_dict.get('TRẠNG THÁI', '').strip()
                trang_thai_tt = 'bai_bo' if trang_thai_excel == 'Bãi bỏ' else 'da_cong_bo'
                
                trang_thai_cong_khai_val = ''
                if dvc_toan_trinh == 'X': trang_thai_cong_khai_val = 'toan_trinh'
                elif dvc_mot_phan == 'X': trang_thai_cong_khai_val = 'cong_khai_mot_phan'
                elif dvc_cung_cap_tt == 'X': trang_thai_cong_khai_val = 'chua_cong_khai'
                
                if cap_thuc_hien and cap_thuc_hien not in ['bo', 'tinh', 'xa', 'dung_chung', 'lien_thong', '']:
                    danh_sach_loi.append(f"Dòng {row_idx}: Cấp thực hiện '{cap_thuc_hien}' không hợp lệ")
                    continue
                
                if not co_quan and cap_thuc_hien:
                    co_quan = get_co_quan_by_cap(cap_thuc_hien)
                
                c.execute("SELECT id FROM tthc WHERE ma_tthc = ?", (ma_tthc,))
                existing = c.fetchone()
                
                if existing:
                    c.execute('''UPDATE tthc SET ten_tthc=?, linh_vuc=?, phi=?, le_phi=?, 
                        lien_thong_cung_cap=?, lien_thong_02_cap=?, phi_dia_gioi=?,
                        dvc_toan_trinh=?, dvc_mot_phan=?, dvc_cung_cap_tt=?, 
                        dich_vu_bcci=?, ghi_chu=?, cap_thuc_hien=?, trang_thai_cong_khai=?, 
                        so_quyet_dinh=?, co_quan_thuc_hien=?, trang_thai=? WHERE ma_tthc = ?''',
                        (ten_tthc, linh_vuc, phi, le_phi, lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                         dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, dich_vu_bcci, ghi_chu,
                         cap_thuc_hien, trang_thai_cong_khai_val, so_qd, co_quan, trang_thai_tt, ma_tthc))
                else:
                    c.execute('''INSERT INTO tthc (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, 
                        lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                        dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, dich_vu_bcci, ghi_chu,
                        cap_thuc_hien, trang_thai_cong_khai, so_quyet_dinh, co_quan_thuc_hien, trang_thai)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                        (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                         dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, dich_vu_bcci, ghi_chu,
                         cap_thuc_hien, trang_thai_cong_khai_val, so_qd, co_quan, trang_thai_tt))
                
                so_luong_thanh_cong += 1
                
            except Exception as e:
                danh_sach_loi.append(f"Dòng {row_idx}: {str(e)}")
        
        conn.commit()
        conn.close()
        try: os.remove(filepath)
        except: pass
        
        if so_luong_thanh_cong > 0:
            thong_bao = f"✅ ĐÃ THÊM DỮ LIỆU THÀNH CÔNG!\n\nSố thủ tục đã thêm: {so_luong_thanh_cong}"
            if danh_sach_loi: thong_bao += f"\n\n⚠️ CẢNH BÁO ({len(danh_sach_loi)} lỗi):\n" + "\n".join(danh_sach_loi[:5])
            thong_bao_js = thong_bao.replace("'", "\\'").replace('"', '\\"').replace("\n", "\\n")
            return f'''<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Import thành công</title></head><body><script>alert("{thong_bao_js}"); window.location.href="/tthc";</script></body></html>'''
        else:
            loi_text = "\n".join(danh_sach_loi[:5]) if danh_sach_loi else "Không tìm thấy dữ liệu hợp lệ."
            loi_text_js = loi_text.replace("'", "\\'").replace('"', '\\"').replace("\n", "\\n")
            return f'''<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Import thất bại</title></head><body><script>alert("❌ IMPORT THẤT BẠI!\n\n{loi_text_js}"); window.location.href="/tthc/them";</script></body></html>'''
            
    except Exception as e:
        loi_text = str(e).replace("'", "\\'").replace('"', '\\"').replace("\n", "\\n")
        return f'''<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Lỗi hệ thống</title></head><body><script>alert("❌ LỖI HỆ THỐNG: {loi_text}"); window.location.href="/tthc/them";</script></body></html>'''

# ==================== SỬA TTHC ====================
@app.route('/tthc/sua/<int:id>', methods=['GET', 'POST'])
@login_required
def sua_tthc(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    c.execute("SELECT linh_vuc FROM tthc WHERE id=?", (id,))
    result = c.fetchone()
    if result and session.get('role') != 'admin':
        if not check_linh_vuc_access(session['user_id'], result[0]):
            conn.close()
            return '<script>alert("Bạn không có quyền sửa thủ tục này!"); window.location.href="/tthc";</script>'
    
    if request.method == 'POST':
        ma_tthc = request.form.get('ma_tthc', '').strip()
        ten_tthc = request.form.get('ten_tthc', '').strip()
        linh_vuc = request.form.get('linh_vuc', '').strip()
        phi = 'X' if request.form.get('phi') == 'on' else ''
        le_phi = 'X' if request.form.get('le_phi') == 'on' else ''
        lien_thong_cung_cap = 'X' if request.form.get('lien_thong_cung_cap') == 'on' else ''
        lien_thong_02_cap = 'X' if request.form.get('lien_thong_02_cap') == 'on' else ''
        phi_dia_gioi = 'X' if request.form.get('phi_dia_gioi') == 'on' else ''
        dvc_toan_trinh = 'X' if request.form.get('dvc_toan_trinh') == 'on' else ''
        dvc_mot_phan = 'X' if request.form.get('dvc_mot_phan') == 'on' else ''
        dvc_cung_cap_tt = 'X' if request.form.get('dvc_cung_cap_tt') == 'on' else ''
        dich_vu_bcci = request.form.get('dich_vu_bcci', '').strip()
        ghi_chu = request.form.get('ghi_chu', '').strip()
        cap_thuc_hien = request.form.get('cap_thuc_hien', '').strip()
        trang_thai_cong_khai = request.form.get('trang_thai_cong_khai', '').strip()
        so_qd = request.form.get('so_quyet_dinh', '').strip()
        co_quan = request.form.get('co_quan_thuc_hien', '').strip()
        trang_thai_tt = request.form.get('trang_thai_tt', 'da_cong_bo')
        
        if not co_quan and cap_thuc_hien:
            co_quan = get_co_quan_by_cap(cap_thuc_hien)
        
        try:
            c.execute('''UPDATE tthc SET ma_tthc=?, ten_tthc=?, linh_vuc=?, phi=?, le_phi=?, 
                lien_thong_cung_cap=?, lien_thong_02_cap=?, phi_dia_gioi=?,
                dvc_toan_trinh=?, dvc_mot_phan=?, dvc_cung_cap_tt=?, 
                dich_vu_bcci=?, ghi_chu=?, cap_thuc_hien=?, trang_thai_cong_khai=?, 
                so_quyet_dinh=?, co_quan_thuc_hien=?, trang_thai=? WHERE id=?''',
                (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                 dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, dich_vu_bcci, ghi_chu,
                 cap_thuc_hien, trang_thai_cong_khai, so_qd, co_quan, trang_thai_tt, id))
            conn.commit()
            conn.close()
            return '<script>alert("✅ CẬP NHẬT THÀNH CÔNG!"); window.location.href="/tthc";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ LỖI: {str(e)}"); window.history.back();</script>'
    
    c.execute("SELECT * FROM tthc WHERE id=?", (id,))
    tthc = c.fetchone()
    c.execute("SELECT id, so_quyet_dinh FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    quyet_dinh_list = c.fetchall()
    conn.close()
    return render_template('tthc_form.html', tthc=tthc, quyet_dinh_list=quyet_dinh_list, username=session.get('username'), role=session.get('role'))

# ==================== XÓA TTHC ====================
@app.route('/tthc/xoa/<int:id>')
@login_required
def xoa_tthc(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    c.execute("SELECT linh_vuc FROM tthc WHERE id=?", (id,))
    result = c.fetchone()
    if result and session.get('role') != 'admin':
        if not check_linh_vuc_access(session['user_id'], result[0]):
            conn.close()
            return '<script>alert("Bạn không có quyền xóa thủ tục này!"); window.location.href="/tthc";</script>'
    
    c.execute("DELETE FROM quy_trinh WHERE tthc_id=?", (id,))
    c.execute("DELETE FROM tthc WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return '<script>alert("🗑️ Đã xóa thủ tục!"); window.location.href="/tthc";</script>'

# ==================== CHI TIẾT TTHC ====================
@app.route('/tthc/chi_tiet/<int:id>')
@login_required
def chi_tiet_tthc(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT * FROM tthc WHERE id=?", (id,))
    tthc = c.fetchone()
    conn.close()
    return render_template('tthc_detail.html', tthc=tthc, username=session.get('username'), role=session.get('role'))

# ==================== QUẢN LÝ QUYẾT ĐỊNH ====================
@app.route('/quyet_dinh')
@login_required
def danh_sach_quyet_dinh():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT * FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    data = c.fetchall()
    conn.close()
    return render_template('quyet_dinh_list.html', quyet_dinhs=data, username=session.get('username'), role=session.get('role'))

@app.route('/quyet_dinh/them', methods=['GET', 'POST'])
@admin_required
def them_quyet_dinh():
    if request.method == 'POST':
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        so_qd = request.form['so_quyet_dinh']
        ten_qd = request.form['ten_quyet_dinh']
        ngay_bh = request.form['ngay_ban_hanh']
        loai = request.form['loai']
        mo_ta = request.form.get('mo_ta', '')
        
        file_dinh_kem = None
        ten_file_goc = None
        
        if 'file_dinh_kem' in request.files:
            file = request.files['file_dinh_kem']
            if file and file.filename != '' and file.filename.endswith('.pdf'):
                ten_file_goc = secure_filename(file.filename)
                unique_name = f"{uuid.uuid4().hex}_{ten_file_goc}"
                filepath = os.path.join(app.config['QUYET_DINH_FOLDER'], unique_name)
                file.save(filepath)
                file_dinh_kem = unique_name
        
        try:
            c.execute("INSERT INTO quyet_dinh (so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai, mo_ta, file_dinh_kem, ten_file_goc) VALUES (?,?,?,?,?,?,?)",
                      (so_qd, ten_qd, ngay_bh, loai, mo_ta, file_dinh_kem, ten_file_goc))
            conn.commit()
            conn.close()
            return '<script>alert("✅ Thêm quyết định thành công!"); window.location.href="/quyet_dinh";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ Lỗi: {str(e)}"); window.history.back();</script>'
    
    return render_template('quyet_dinh_form.html', quyet_dinh=None)

@app.route('/quyet_dinh/sua/<int:id>', methods=['GET', 'POST'])
@admin_required
def sua_quyet_dinh(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    if request.method == 'POST':
        so_qd = request.form['so_quyet_dinh']
        ten_qd = request.form['ten_quyet_dinh']
        ngay_bh = request.form['ngay_ban_hanh']
        loai = request.form['loai']
        mo_ta = request.form.get('mo_ta', '')
        
        c.execute("SELECT file_dinh_kem FROM quyet_dinh WHERE id=?", (id,))
        old_file = c.fetchone()
        old_file_name = old_file[0] if old_file else None
        
        file_dinh_kem = old_file_name
        ten_file_goc = None
        
        if 'file_dinh_kem' in request.files:
            file = request.files['file_dinh_kem']
            if file and file.filename != '' and file.filename.endswith('.pdf'):
                if old_file_name:
                    old_path = os.path.join(app.config['QUYET_DINH_FOLDER'], old_file_name)
                    if os.path.exists(old_path):
                        os.remove(old_path)
                ten_file_goc = secure_filename(file.filename)
                unique_name = f"{uuid.uuid4().hex}_{ten_file_goc}"
                filepath = os.path.join(app.config['QUYET_DINH_FOLDER'], unique_name)
                file.save(filepath)
                file_dinh_kem = unique_name
        
        try:
            c.execute("UPDATE quyet_dinh SET so_quyet_dinh=?, ten_quyet_dinh=?, ngay_ban_hanh=?, loai=?, mo_ta=?, file_dinh_kem=?, ten_file_goc=? WHERE id=?",
                      (so_qd, ten_qd, ngay_bh, loai, mo_ta, file_dinh_kem, ten_file_goc, id))
            conn.commit()
            conn.close()
            return '<script>alert("✅ Cập nhật quyết định thành công!"); window.location.href="/quyet_dinh";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ Lỗi: {str(e)}"); window.history.back();</script>'
    
    c.execute("SELECT * FROM quyet_dinh WHERE id=?", (id,))
    quyet_dinh = c.fetchone()
    conn.close()
    return render_template('quyet_dinh_form.html', quyet_dinh=quyet_dinh)

@app.route('/quyet_dinh/xoa/<int:id>')
@admin_required
def xoa_quyet_dinh(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT file_dinh_kem FROM quyet_dinh WHERE id=?", (id,))
    result = c.fetchone()
    if result and result[0]:
        filepath = os.path.join(app.config['QUYET_DINH_FOLDER'], result[0])
        if os.path.exists(filepath):
            os.remove(filepath)
    c.execute("DELETE FROM quyet_dinh WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return '<script>alert("🗑️ Đã xóa quyết định!"); window.location.href="/quyet_dinh";</script>'

# ==================== XEM PDF ====================
@app.route('/xem_pdf/<int:id>')
@login_required
def xem_pdf(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT file_dinh_kem, ten_file_goc FROM quyet_dinh WHERE id=?", (id,))
    result = c.fetchone()
    conn.close()
    if result and result[0]:
        return render_template('xem_pdf.html', file_name=result[0], ten_file=result[1], quyet_dinh_id=id)
    return '<script>alert("Không có file đính kèm!"); window.history.back();</script>'

@app.route('/uploads/quyet_dinh/<filename>')
def download_pdf(filename):
    return send_from_directory(app.config['QUYET_DINH_FOLDER'], filename)

# ==================== SO SÁNH QUYẾT ĐỊNH ====================
@app.route('/so_sanh_quyet_dinh')
@login_required
def so_sanh_qd():
    return render_template('so_sanh_quyet_dinh.html', username=session.get('username'), role=session.get('role'))

@app.route('/api/so_sanh_quyet_dinh')
@login_required
def api_so_sanh_quyet_dinh():
    qd_ids = request.args.get('qd_ids', '')
    if not qd_ids:
        return jsonify({'error': 'Chưa chọn quyết định'})
    ids = [int(x) for x in qd_ids.split(',')]
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    quyet_dinhs = []
    for qd_id in ids:
        c.execute("SELECT id, so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai FROM quyet_dinh WHERE id=?", (qd_id,))
        qd = c.fetchone()
        if qd:
            quyet_dinhs.append({'id': qd[0], 'so_quyet_dinh': qd[1], 'ten_quyet_dinh': qd[2], 'ngay_ban_hanh': qd[3], 'loai': qd[4]})
    tthc_theo_qd = {}
    for qd_id in ids:
        qd = next((q for q in quyet_dinhs if q['id'] == qd_id), None)
        if qd:
            c.execute("SELECT ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc WHERE so_quyet_dinh=?", (qd['so_quyet_dinh'],))
            tthc_theo_qd[qd_id] = c.fetchall()
    conn.close()
    return jsonify({'quyet_dinhs': quyet_dinhs, 'tthc_theo_qd': tthc_theo_qd})

# ==================== BÁO CÁO ====================
@app.route('/bao_cao')
@login_required
def bao_cao():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    c.execute("SELECT id, so_quyet_dinh FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    danh_sach_quyet_dinh = c.fetchall()
    
    if session.get('role') == 'admin':
        c.execute("SELECT DISTINCT linh_vuc FROM tthc WHERE linh_vuc IS NOT NULL AND linh_vuc != '' ORDER BY linh_vuc")
        danh_sach_linh_vuc = c.fetchall()
        c.execute("SELECT DISTINCT cap_thuc_hien FROM tthc WHERE cap_thuc_hien IS NOT NULL AND cap_thuc_hien != '' ORDER BY cap_thuc_hien")
        danh_sach_cap = c.fetchall()
        c.execute("SELECT DISTINCT co_quan_thuc_hien FROM tthc WHERE co_quan_thuc_hien IS NOT NULL AND co_quan_thuc_hien != '' ORDER BY co_quan_thuc_hien")
        danh_sach_co_quan = c.fetchall()
        c.execute("SELECT DISTINCT dich_vu_bcci FROM tthc WHERE dich_vu_bcci IS NOT NULL AND dich_vu_bcci != '' ORDER BY dich_vu_bcci")
        danh_sach_bcci = c.fetchall()
    else:
        c.execute("SELECT linh_vuc FROM user_linh_vuc WHERE user_id=?", (session['user_id'],))
        allowed = [row[0] for row in c.fetchall()]
        if allowed:
            placeholders = ','.join(['?'] * len(allowed))
            c.execute(f"SELECT DISTINCT linh_vuc FROM tthc WHERE linh_vuc IN ({placeholders}) AND linh_vuc IS NOT NULL AND linh_vuc != '' ORDER BY linh_vuc", allowed)
            danh_sach_linh_vuc = c.fetchall()
            c.execute(f"SELECT DISTINCT cap_thuc_hien FROM tthc WHERE linh_vuc IN ({placeholders}) AND cap_thuc_hien IS NOT NULL AND cap_thuc_hien != '' ORDER BY cap_thuc_hien", allowed)
            danh_sach_cap = c.fetchall()
            c.execute(f"SELECT DISTINCT co_quan_thuc_hien FROM tthc WHERE linh_vuc IN ({placeholders}) AND co_quan_thuc_hien IS NOT NULL AND co_quan_thuc_hien != '' ORDER BY co_quan_thuc_hien", allowed)
            danh_sach_co_quan = c.fetchall()
            c.execute(f"SELECT DISTINCT dich_vu_bcci FROM tthc WHERE linh_vuc IN ({placeholders}) AND dich_vu_bcci IS NOT NULL AND dich_vu_bcci != '' ORDER BY dich_vu_bcci", allowed)
            danh_sach_bcci = c.fetchall()
        else:
            danh_sach_linh_vuc = danh_sach_cap = danh_sach_co_quan = danh_sach_bcci = []
    
    conn.close()
    return render_template('bao_cao.html', danh_sach_quyet_dinh=danh_sach_quyet_dinh, danh_sach_linh_vuc=danh_sach_linh_vuc,
                         danh_sach_cap=danh_sach_cap, danh_sach_co_quan=danh_sach_co_quan, danh_sach_bcci=danh_sach_bcci,
                         username=session.get('username'), role=session.get('role'))

@app.route('/api/bao_cao_data')
@login_required
def api_bao_cao_data():
    so_quyet_dinh = request.args.get('so_quyet_dinh', '')
    linh_vuc = request.args.get('linh_vuc', '')
    cap_thuc_hien = request.args.get('cap_thuc_hien', '')
    trang_thai = request.args.get('trang_thai', '')
    phi_le_phi = request.args.get('phi_le_phi', '')
    dich_vu_bcci = request.args.get('dich_vu_bcci', '')
    co_quan_thuc_hien = request.args.get('co_quan_thuc_hien', '')
    trang_thai_tt = request.args.get('trang_thai_tt', '')
    
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    sql = """SELECT t.ma_tthc, t.ten_tthc, t.linh_vuc, t.phi, t.le_phi, 
                    t.cap_thuc_hien, t.co_quan_thuc_hien, t.dvc_toan_trinh, t.dvc_mot_phan, t.dvc_cung_cap_tt,
                    t.dich_vu_bcci, t.ghi_chu, t.so_quyet_dinh, t.trang_thai
             FROM tthc t WHERE 1=1"""
    params = []
    
    if session.get('role') != 'admin':
        c.execute("SELECT linh_vuc FROM user_linh_vuc WHERE user_id=?", (session['user_id'],))
        allowed = [row[0] for row in c.fetchall()]
        if allowed:
            placeholders = ','.join(['?'] * len(allowed))
            sql += f" AND t.linh_vuc IN ({placeholders})"
            params.extend(allowed)
        else:
            sql += " AND 1=0"
    
    if so_quyet_dinh and so_quyet_dinh != '':
        sql += " AND t.so_quyet_dinh = ?"
        params.append(so_quyet_dinh)
    if linh_vuc and linh_vuc != '':
        sql += " AND t.linh_vuc = ?"
        params.append(linh_vuc)
    if cap_thuc_hien and cap_thuc_hien != '':
        sql += " AND t.cap_thuc_hien = ?"
        params.append(cap_thuc_hien)
    if trang_thai and trang_thai != '':
        if trang_thai == 'toan_trinh': sql += " AND t.dvc_toan_trinh = 'X'"
        elif trang_thai == 'mot_phan': sql += " AND t.dvc_mot_phan = 'X'"
        elif trang_thai == 'cung_cap_tt': sql += " AND t.dvc_cung_cap_tt = 'X'"
    if trang_thai_tt and trang_thai_tt != '':
        if trang_thai_tt == 'da_cong_bo': sql += " AND t.trang_thai = 'da_cong_bo'"
        elif trang_thai_tt == 'bai_bo': sql += " AND t.trang_thai = 'bai_bo'"
    if phi_le_phi and phi_le_phi != '':
        if phi_le_phi == 'co_phi': sql += " AND t.phi = 'X'"
        elif phi_le_phi == 'co_le_phi': sql += " AND t.le_phi = 'X'"
        elif phi_le_phi == 'mien_thu': sql += " AND (t.phi IS NULL OR t.phi = '') AND (t.le_phi IS NULL OR t.le_phi = '')"
    if dich_vu_bcci and dich_vu_bcci != '':
        if dich_vu_bcci == 'co': sql += " AND t.dich_vu_bcci IS NOT NULL AND t.dich_vu_bcci != ''"
        elif dich_vu_bcci == 'khong': sql += " AND (t.dich_vu_bcci IS NULL OR t.dich_vu_bcci = '')"
        else:
            sql += " AND t.dich_vu_bcci = ?"
            params.append(dich_vu_bcci)
    if co_quan_thuc_hien and co_quan_thuc_hien != '':
        sql += " AND t.co_quan_thuc_hien = ?"
        params.append(co_quan_thuc_hien)
    sql += " ORDER BY t.ma_tthc"
    
    c.execute(sql, params)
    data = c.fetchall()
    
    c.execute("SELECT COUNT(*) FROM tthc")
    tong_so = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_toan_trinh = 'X'")
    so_toan_trinh = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_mot_phan = 'X'")
    so_mot_phan = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_cung_cap_tt = 'X'")
    so_cung_cap_tt = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE trang_thai = 'da_cong_bo'")
    so_da_cong_bo = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE trang_thai = 'bai_bo'")
    so_bai_bo = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE phi = 'X'")
    so_co_phi = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE le_phi = 'X'")
    so_co_le_phi = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE (phi IS NULL OR phi = '') AND (le_phi IS NULL OR le_phi = '')")
    so_mien_thu = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE dich_vu_bcci IS NOT NULL AND dich_vu_bcci != ''")
    so_co_bcci = c.fetchone()[0]
    conn.close()
    
    ket_qua = []
    for row in data:
        ket_qua.append({
            'ma_tthc': row[0], 'ten_tthc': row[1], 'linh_vuc': row[2] or '',
            'phi': 'Có' if row[3] == 'X' else 'Không', 'le_phi': 'Có' if row[4] == 'X' else 'Không',
            'cap_thuc_hien': row[5] or '', 'co_quan_thuc_hien': row[6] or '',
            'toan_trinh': 'X' if row[7] == 'X' else '', 'mot_phan': 'X' if row[8] == 'X' else '',
            'cung_cap_tt': 'X' if row[9] == 'X' else '', 'dich_vu_bcci': row[10] or '',
            'ghi_chu': row[11] or '', 'so_quyet_dinh': row[12] or '', 'trang_thai': row[13] or 'da_cong_bo'
        })
    
    return jsonify({'data': ket_qua, 'tong_so': tong_so, 'so_toan_trinh': so_toan_trinh, 'so_mot_phan': so_mot_phan,
                    'so_cung_cap_tt': so_cung_cap_tt, 'so_da_cong_bo': so_da_cong_bo, 'so_bai_bo': so_bai_bo,
                    'so_co_phi': so_co_phi, 'so_co_le_phi': so_co_le_phi, 'so_mien_thu': so_mien_thu, 'so_co_bcci': so_co_bcci})

@app.route('/export_bao_cao_excel')
@login_required
def export_bao_cao_excel():
    import io
    from flask import send_file
    from openpyxl import Workbook
    
    so_quyet_dinh = request.args.get('so_quyet_dinh', '')
    linh_vuc = request.args.get('linh_vuc', '')
    cap_thuc_hien = request.args.get('cap_thuc_hien', '')
    trang_thai = request.args.get('trang_thai', '')
    phi_le_phi = request.args.get('phi_le_phi', '')
    dich_vu_bcci = request.args.get('dich_vu_bcci', '')
    co_quan_thuc_hien = request.args.get('co_quan_thuc_hien', '')
    trang_thai_tt = request.args.get('trang_thai_tt', '')
    
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    sql = """SELECT t.ma_tthc, t.ten_tthc, t.linh_vuc, t.phi, t.le_phi, 
                    t.cap_thuc_hien, t.co_quan_thuc_hien, t.dvc_toan_trinh, t.dvc_mot_phan, t.dvc_cung_cap_tt,
                    t.dich_vu_bcci, t.ghi_chu, t.so_quyet_dinh, t.trang_thai
             FROM tthc t WHERE 1=1"""
    params = []
    
    if session.get('role') != 'admin':
        c.execute("SELECT linh_vuc FROM user_linh_vuc WHERE user_id=?", (session['user_id'],))
        allowed = [row[0] for row in c.fetchall()]
        if allowed:
            placeholders = ','.join(['?'] * len(allowed))
            sql += f" AND t.linh_vuc IN ({placeholders})"
            params.extend(allowed)
    
    if so_quyet_dinh and so_quyet_dinh != '':
        sql += " AND t.so_quyet_dinh = ?"
        params.append(so_quyet_dinh)
    if linh_vuc and linh_vuc != '':
        sql += " AND t.linh_vuc = ?"
        params.append(linh_vuc)
    if cap_thuc_hien and cap_thuc_hien != '':
        sql += " AND t.cap_thuc_hien = ?"
        params.append(cap_thuc_hien)
    if trang_thai and trang_thai != '':
        if trang_thai == 'toan_trinh': sql += " AND t.dvc_toan_trinh = 'X'"
        elif trang_thai == 'mot_phan': sql += " AND t.dvc_mot_phan = 'X'"
        elif trang_thai == 'cung_cap_tt': sql += " AND t.dvc_cung_cap_tt = 'X'"
    if trang_thai_tt and trang_thai_tt != '':
        if trang_thai_tt == 'da_cong_bo': sql += " AND t.trang_thai = 'da_cong_bo'"
        elif trang_thai_tt == 'bai_bo': sql += " AND t.trang_thai = 'bai_bo'"
    if phi_le_phi and phi_le_phi != '':
        if phi_le_phi == 'co_phi': sql += " AND t.phi = 'X'"
        elif phi_le_phi == 'co_le_phi': sql += " AND t.le_phi = 'X'"
        elif phi_le_phi == 'mien_thu': sql += " AND (t.phi IS NULL OR t.phi = '') AND (t.le_phi IS NULL OR t.le_phi = '')"
    if dich_vu_bcci and dich_vu_bcci != '':
        if dich_vu_bcci == 'co': sql += " AND t.dich_vu_bcci IS NOT NULL AND t.dich_vu_bcci != ''"
        elif dich_vu_bcci == 'khong': sql += " AND (t.dich_vu_bcci IS NULL OR t.dich_vu_bcci = '')"
        else:
            sql += " AND t.dich_vu_bcci = ?"
            params.append(dich_vu_bcci)
    if co_quan_thuc_hien and co_quan_thuc_hien != '':
        sql += " AND t.co_quan_thuc_hien = ?"
        params.append(co_quan_thuc_hien)
    sql += " ORDER BY t.ma_tthc"
    
    c.execute(sql, params)
    data = c.fetchall()
    conn.close()
    
    wb = Workbook()
    ws = wb.active
    ws.title = "BaoCaoTTHC"
    
    headers = ['MÃ TTHC', 'TÊN TTHC', 'LĨNH VỰC', 'PHÍ', 'LỆ PHÍ',
               'CẤP THỰC HIỆN', 'CƠ QUAN THỰC HIỆN', 'TOÀN TRÌNH', 'MỘT PHẦN', 'CUNG CẤP TT',
               'DỊCH VỤ BCCI', 'GHI CHÚ', 'QUYẾT ĐỊNH SỐ', 'TRẠNG THÁI']
    
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)
    
    cap_map = {'bo': 'Cấp Bộ', 'tinh': 'Cấp tỉnh', 'xa': 'Cấp xã', 'dung_chung': 'Dùng chung', 'lien_thong': 'Liên thông'}
    
    for row_idx, row in enumerate(data, 2):
        ws.cell(row=row_idx, column=1, value=row[0])
        ws.cell(row=row_idx, column=2, value=row[1])
        ws.cell(row=row_idx, column=3, value=row[2] or '')
        ws.cell(row=row_idx, column=4, value='X' if row[3] == 'X' else '')
        ws.cell(row=row_idx, column=5, value='X' if row[4] == 'X' else '')
        ws.cell(row=row_idx, column=6, value=cap_map.get(row[5], row[5]))
        ws.cell(row=row_idx, column=7, value=row[6] or '')
        ws.cell(row=row_idx, column=8, value='X' if row[7] == 'X' else '')
        ws.cell(row=row_idx, column=9, value='X' if row[8] == 'X' else '')
        ws.cell(row=row_idx, column=10, value='X' if row[9] == 'X' else '')
        ws.cell(row=row_idx, column=11, value=row[10] or '')
        ws.cell(row=row_idx, column=12, value=row[11] or '')
        ws.cell(row=row_idx, column=13, value=row[12] or '')
        ws.cell(row=row_idx, column=14, value='Đã công bố' if row[13] == 'da_cong_bo' else 'Bãi bỏ')
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    ten_file = f'BaoCao_TTHC_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=ten_file)

# ==================== TRA CỨU ====================
@app.route('/tim_kiem')
@login_required
def tim_kiem():
    tu_khoa = request.args.get('q', '')
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    if session.get('role') == 'admin':
        if tu_khoa:
            c.execute("SELECT id, ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc WHERE ten_tthc LIKE ? OR ma_tthc LIKE ?", (f'%{tu_khoa}%', f'%{tu_khoa}%'))
        else:
            c.execute("SELECT id, ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc LIMIT 20")
    else:
        c.execute("SELECT linh_vuc FROM user_linh_vuc WHERE user_id=?", (session['user_id'],))
        allowed = [row[0] for row in c.fetchall()]
        if allowed:
            placeholders = ','.join(['?'] * len(allowed))
            if tu_khoa:
                c.execute(f"SELECT id, ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc WHERE (ten_tthc LIKE ? OR ma_tthc LIKE ?) AND linh_vuc IN ({placeholders})", (f'%{tu_khoa}%', f'%{tu_khoa}%', *allowed))
            else:
                c.execute(f"SELECT id, ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc WHERE linh_vuc IN ({placeholders}) LIMIT 20", allowed)
        else:
            c.execute("SELECT id, ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc WHERE 1=0")
    
    ket_qua = c.fetchall()
    conn.close()
    return render_template('tim_kiem.html', ket_qua=ket_qua, tu_khoa=tu_khoa, username=session.get('username'), role=session.get('role'))

if __name__ == '__main__':
    app.run(debug=True)

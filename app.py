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
import shutil
import hashlib

app = Flask(__name__)
app.secret_key = 'sotuphap_dongthap_secret_key_2026'

UPLOAD_FOLDER = 'uploads'
QUYET_DINH_FOLDER = 'uploads/quyet_dinh'
BACKUP_FOLDER = 'backups'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['QUYET_DINH_FOLDER'] = QUYET_DINH_FOLDER
app.config['BACKUP_FOLDER'] = BACKUP_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(QUYET_DINH_FOLDER, exist_ok=True)
os.makedirs(BACKUP_FOLDER, exist_ok=True)

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

# ==================== HÀM BACKUP ====================
def backup_database():
    """Tạo bản sao lưu database"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(BACKUP_FOLDER, f"tthc_backup_{timestamp}.db")
    shutil.copy2('tthc.db', backup_file)
    
    # Xóa các bản backup cũ hơn 30 ngày
    for f in os.listdir(BACKUP_FOLDER):
        if f.startswith('tthc_backup_') and f.endswith('.db'):
            file_path = os.path.join(BACKUP_FOLDER, f)
            if os.path.getmtime(file_path) < (datetime.now().timestamp() - 30 * 24 * 3600):
                os.remove(file_path)
    return backup_file

def restore_database(backup_file):
    """Khôi phục database từ bản backup"""
    if os.path.exists(backup_file):
        shutil.copy2(backup_file, 'tthc.db')
        return True
    return False

def get_backup_list():
    """Lấy danh sách các file backup"""
    backups = []
    for f in os.listdir(BACKUP_FOLDER):
        if f.startswith('tthc_backup_') and f.endswith('.db'):
            file_path = os.path.join(BACKUP_FOLDER, f)
            stat = os.stat(file_path)
            backups.append({
                'name': f,
                'size': stat.st_size,
                'date': datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                'path': file_path
            })
    backups.sort(key=lambda x: x['date'], reverse=True)
    return backups

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
    
    admin_password = hashlib.sha256('admin@123'.encode()).hexdigest()
    c.execute("SELECT COUNT(*) FROM users WHERE username = 'Admin'")
    if c.fetchone()[0] == 0:
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
        
        hashed_password = hashlib.sha256(password.encode()).hexdigest()
        
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        c.execute("SELECT id, username, role, fullname FROM users WHERE username = ? AND password = ?", (username, hashed_password))
        user = c.fetchone()
        conn.close()
        
        if user:
            session['user_id'] = user[0]
            session['username'] = user[1]
            session['role'] = user[2]
            session['fullname'] = user[3] or user[1]
            return redirect(url_for('dashboard'))
        else:
            return render_template('login.html', error='Sai tên đăng nhập hoặc mật khẩu!')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ==================== ĐỔI MẬT KHẨU ====================
@app.route('/doi_mat_khau', methods=['GET', 'POST'])
@login_required
def doi_mat_khau():
    if request.method == 'POST':
        old_password = request.form.get('old_password', '').strip()
        new_password = request.form.get('new_password', '').strip()
        confirm_password = request.form.get('confirm_password', '').strip()
        
        if new_password != confirm_password:
            return '<script>alert("Mật khẩu mới không khớp!"); window.history.back();</script>'
        
        if len(new_password) < 6:
            return '<script>alert("Mật khẩu mới phải có ít nhất 6 ký tự!"); window.history.back();</script>'
        
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        old_hashed = hashlib.sha256(old_password.encode()).hexdigest()
        c.execute("SELECT password FROM users WHERE id = ?", (session['user_id'],))
        current_password = c.fetchone()[0]
        
        if current_password != old_hashed:
            conn.close()
            return '<script>alert("Mật khẩu cũ không đúng!"); window.history.back();</script>'
        
        new_hashed = hashlib.sha256(new_password.encode()).hexdigest()
        c.execute("UPDATE users SET password = ? WHERE id = ?", (new_hashed, session['user_id']))
        conn.commit()
        conn.close()
        
        return '<script>alert("✅ Đổi mật khẩu thành công!"); window.location.href="/";</script>'
    
    return render_template('doi_mat_khau.html', username=session.get('username'), role=session.get('role'))

# ==================== BACKUP VÀ RESTORE ====================
@app.route('/backup')
@admin_required
def backup_list():
    backups = get_backup_list()
    return render_template('backup.html', backups=backups, username=session.get('username'), role=session.get('role'))

@app.route('/backup/create')
@admin_required
def create_backup():
    try:
        backup_file = backup_database()
        return f'<script>alert("✅ Đã tạo bản sao lưu thành công!"); window.location.href="/backup";</script>'
    except Exception as e:
        return f'<script>alert("❌ Lỗi tạo backup: {str(e)}"); window.location.href="/backup";</script>'

@app.route('/backup/restore/<filename>')
@admin_required
def restore_backup(filename):
    try:
        backup_path = os.path.join(BACKUP_FOLDER, filename)
        if restore_database(backup_path):
            return '<script>alert("✅ Khôi phục dữ liệu thành công! Vui lòng đăng nhập lại."); window.location.href="/login";</script>'
        else:
            return '<script>alert("❌ Khôi phục thất bại!"); window.location.href="/backup";</script>'
    except Exception as e:
        return f'<script>alert("❌ Lỗi khôi phục: {str(e)}"); window.location.href="/backup";</script>'

@app.route('/backup/download/<filename>')
@admin_required
def download_backup(filename):
    return send_from_directory(BACKUP_FOLDER, filename, as_attachment=True)

@app.route('/backup/delete/<filename>')
@admin_required
def delete_backup(filename):
    try:
        file_path = os.path.join(BACKUP_FOLDER, filename)
        if os.path.exists(file_path):
            os.remove(file_path)
        return '<script>alert("🗑️ Đã xóa bản sao lưu!"); window.location.href="/backup";</script>'
    except Exception as e:
        return f'<script>alert("❌ Lỗi xóa: {str(e)}"); window.location.href="/backup";</script>'

# ==================== QUẢN LÝ NGƯỜI DÙNG ====================
@app.route('/users')
@admin_required
def danh_sach_users():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT id, username, role, fullname, email, created_at FROM users ORDER BY id")
    users = c.fetchall()
    conn.close()
    return render_template('users.html', users=users, username=session.get('username'), role=session.get('role'))

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
    return render_template('user_form.html', user=None, linh_vuc_list=linh_vuc_list, username=session.get('username'), role=session.get('role'))

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
    return render_template('user_form.html', user=user, linh_vuc_list=linh_vuc_list, user_linh_vuc=user_linh_vuc, username=session.get('username'), role=session.get('role'))

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
    return render_template('dashboard.html', thong_ke=thong_ke, username=session.get('username'), role=session.get('role'), fullname=session.get('fullname'))

# ==================== CÁC ROUTE KHÁC (GIỮ NGUYÊN TỪ CÁC PHẦN TRƯỚC) ====================
# ... (các route /tthc, /quyet_dinh, /bao_cao, /tim_kiem, /so_sanh_quyet_dinh, /api/... giữ nguyên)
# Để tránh quá dài, tôi sẽ giữ nguyên các route đã có từ các phiên bản trước

if __name__ == '__main__':
    app.run(debug=True)

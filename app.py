import time
time.sleep(2)  # Chờ 2 giây trước khi khởi động
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
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(BACKUP_FOLDER, f"tthc_backup_{timestamp}.db")
    shutil.copy2('tthc.db', backup_file)
    
    for f in os.listdir(BACKUP_FOLDER):
        if f.startswith('tthc_backup_') and f.endswith('.db'):
            file_path = os.path.join(BACKUP_FOLDER, f)
            if os.path.getmtime(file_path) < (datetime.now().timestamp() - 30 * 24 * 3600):
                os.remove(file_path)
    return backup_file

def restore_database(backup_file):
    if os.path.exists(backup_file):
        shutil.copy2(backup_file, 'tthc.db')
        return True
    return False

def get_backup_list():
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
    
    # Bảng quyet_dinh
    c.execute('''CREATE TABLE IF NOT EXISTS quyet_dinh (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        so_quyet_dinh TEXT UNIQUE,
        ten_quyet_dinh TEXT,
        ngay_ban_hanh DATE,
        loai TEXT,
        mo_ta TEXT,
        file_dinh_kem TEXT,
        ten_file_goc TEXT,
        noi_dung_ocr TEXT,
        ngay_hieu_luc DATE,
        tthc_thay_the TEXT
    )''')
    
    # Bảng tthc
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
        trang_thai TEXT DEFAULT 'da_cong_bo',
        thanh_phan_ho_so TEXT,
        thoi_gian_giai_quyet TEXT,
        can_cu_phap_ly TEXT,
        so_luong_da_xu_ly INTEGER DEFAULT 0,
        so_luong_dang_xu_ly INTEGER DEFAULT 0,
        created_at DATE
    )''')
    
    # Bảng quy_trinh
    c.execute('''CREATE TABLE IF NOT EXISTS quy_trinh (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tthc_id INTEGER,
        buoc INTEGER,
        ten_buoc TEXT,
        thoi_gian_xu_ly TEXT,
        co_quan_thuc_hien TEXT
    )''')
    
    # Bảng users
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT,
        role TEXT DEFAULT 'user',
        fullname TEXT,
        email TEXT,
        created_at DATE
    )''')
    
    # Bảng user_linh_vuc
    c.execute('''CREATE TABLE IF NOT EXISTS user_linh_vuc (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        linh_vuc TEXT,
        FOREIGN KEY(user_id) REFERENCES users(id)
    )''')
    
    # Bảng lich_su_thay_doi
    c.execute('''CREATE TABLE IF NOT EXISTS lich_su_thay_doi (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tthc_id INTEGER,
        quyet_dinh_id INTEGER,
        loai_thay_doi TEXT,
        ghi_chu TEXT,
        ngay_ap_dung DATE,
        created_at DATE
    )''')
    
    # Tạo tài khoản admin mặc định
    admin_password = hashlib.sha256('admin@123'.encode()).hexdigest()
    c.execute("SELECT COUNT(*) FROM users WHERE username = 'Admin'")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO users (username, password, role, fullname, email, created_at) VALUES (?,?,?,?,?,?)",
                  ('Admin', admin_password, 'admin', 'Quản trị viên', 'admin@dongthap.gov.vn', datetime.now().date()))
    
    # Tạo dữ liệu mẫu nếu bảng tthc trống
    c.execute("SELECT COUNT(*) FROM tthc")
    if c.fetchone()[0] == 0:
        sample_tthc = [
            ('TTHC001', 'Đăng ký khai sinh', 'Hộ tịch', 'Miễn phí', '0', 'X', '', '', 'X', '', '', '', '', 'xa', 'Đã công khai', '', 'UBND cấp xã', 'da_cong_bo', 
             'Giấy khai sinh, Giấy chứng sinh', '03 ngày', 'Luật Hộ tịch 2014', 120, 15, datetime.now().date()),
            ('TTHC002', 'Cấp Căn cước công dân', 'Căn cước', '50.000 VNĐ', '0', '', '', '', 'X', '', '', '', '', 'tinh', 'Đã công khai', '', 'Công an tỉnh', 'da_cong_bo',
             'Tờ khai CCCD, Ảnh thẻ', '07 ngày', 'Luật CCCD 2014', 250, 30, datetime.now().date()),
            ('TTHC003', 'Đăng ký kinh doanh', 'Kinh doanh', '100.000 VNĐ', '0', 'X', 'X', '', 'X', '', '', '', '', 'tinh', 'Đã công khai', '', 'Sở KH&ĐT', 'da_cong_bo',
             'Đơn đăng ký, Điều lệ', '03 ngày', 'Luật Doanh nghiệp 2020', 89, 12, datetime.now().date()),
            ('TTHC004', 'Cấp giấy phép xây dựng', 'Xây dựng', '200.000 VNĐ', '0', '', '', '', 'X', '', '', '', '', 'xa', 'Đã công khai', '', 'UBND huyện', 'da_cong_bo',
             'Đơn xin phép, Bản vẽ', '15 ngày', 'Luật Xây dựng 2014', 45, 8, datetime.now().date()),
            ('TTHC005', 'Công chứng hợp đồng', 'Công chứng', '0.1% giá trị', '0', '', '', '', '', 'X', 'X', '', '', 'dung_chung', 'Đã công khai', '', 'Phòng công chứng', 'da_cong_bo',
             'Hợp đồng, CMND', '01 ngày', 'Luật Công chứng', 380, 42, datetime.now().date()),
        ]
        for item in sample_tthc:
            c.execute('''INSERT INTO tthc (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, lien_thong_cung_cap, 
                      lien_thong_02_cap, phi_dia_gioi, dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, 
                      dich_vu_bcci, ghi_chu, cap_thuc_hien, trang_thai_cong_khai, so_quyet_dinh, 
                      co_quan_thuc_hien, trang_thai, thanh_phan_ho_so, thoi_gian_giai_quyet, 
                      can_cu_phap_ly, so_luong_da_xu_ly, so_luong_dang_xu_ly, created_at) 
                      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', item)
    
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

# ==================== TRANG CHỦ DASHBOARD ====================
@app.route('/')
@login_required
def dashboard():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    thong_ke = {
        'cap_bo': 0, 'cap_tinh': 0, 'cap_xa': 0, 'cap_dung_chung': 0, 'cap_lien_thong': 0,
        'trang_thai_toan_trinh': 0, 'trang_thai_cong_khai_mot_phan': 0, 'trang_thai_chua_cong_khai': 0,
        'so_da_cong_bo': 0, 'so_bai_bo': 0, 'linh_vuc': [],
        'tong_so_tthc': 0, 'tthc_da_xu_ly': 0, 'tthc_dang_xu_ly': 0
    }
    
    try:
        # Thống kê theo cấp thực hiện
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
        c.execute("SELECT SUM(so_luong_da_xu_ly) FROM tthc")
        thong_ke['tthc_da_xu_ly'] = c.fetchone()[0] or 0
        c.execute("SELECT SUM(so_luong_dang_xu_ly) FROM tthc")
        thong_ke['tthc_dang_xu_ly'] = c.fetchone()[0] or 0
        c.execute("SELECT COUNT(*) FROM tthc")
        thong_ke['tong_so_tthc'] = c.fetchone()[0] or 0
    except: pass
    
    try:
        c.execute("SELECT linh_vuc, COUNT(*) FROM tthc WHERE linh_vuc IS NOT NULL AND linh_vuc != '' GROUP BY linh_vuc ORDER BY COUNT(*) DESC")
        raw_data = c.fetchall()
        if raw_data:
            thong_ke['linh_vuc'] = [{'ten': str(row[0]), 'so_luong': row[1]} for row in raw_data]
    except: pass
    
    # Lấy TTHC mới nhất
    c.execute("SELECT id, ma_tthc, ten_tthc, linh_vuc, trang_thai FROM tthc ORDER BY id DESC LIMIT 5")
    recent_tthc = c.fetchall()
    
    # Lấy quyết định mới nhất
    c.execute("SELECT id, so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh FROM quyet_dinh ORDER BY id DESC LIMIT 5")
    recent_qd = c.fetchall()
    
    conn.close()
    return render_template('dashboard.html', thong_ke=thong_ke, recent_tthc=recent_tthc, recent_qd=recent_qd,
                          username=session.get('username'), role=session.get('role'), fullname=session.get('fullname'))

# ==================== QUẢN LÝ TTHC ====================
@app.route('/tthc')
@login_required
def tthc_list():
    page = request.args.get('page', 1, type=int)
    per_page = 10
    offset = (page - 1) * per_page
    
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    # Đếm tổng số
    c.execute("SELECT COUNT(*) FROM tthc WHERE trang_thai = 'da_cong_bo'")
    total = c.fetchone()[0]
    
    # Lấy danh sách
    c.execute('''SELECT id, ma_tthc, ten_tthc, linh_vuc, cap_thuc_hien, 
                 co_quan_thuc_hien, trang_thai, so_luong_da_xu_ly, so_luong_dang_xu_ly 
                 FROM tthc WHERE trang_thai = 'da_cong_bo' 
                 ORDER BY id DESC LIMIT ? OFFSET ?''', (per_page, offset))
    tthc_list = c.fetchall()
    conn.close()
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('tthc_list.html', tthc_list=tthc_list, page=page, total_pages=total_pages,
                          username=session.get('username'), role=session.get('role'))

@app.route('/tthc/them', methods=['GET', 'POST'])
@login_required
def tthc_create():
    if request.method == 'POST':
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        
        try:
            ma_tthc = request.form.get('ma_tthc', '').strip()
            c.execute("SELECT COUNT(*) FROM tthc WHERE ma_tthc = ?", (ma_tthc,))
            if c.fetchone()[0] > 0:
                conn.close()
                return '<script>alert("❌ Mã TTHC đã tồn tại!"); window.history.back();</script>'
            
            c.execute('''INSERT INTO tthc (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, cap_thuc_hien, 
                      co_quan_thuc_hien, thanh_phan_ho_so, thoi_gian_giai_quyet, can_cu_phap_ly,
                      dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, trang_thai, created_at)
                      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                      (ma_tthc, request.form.get('ten_tthc'), request.form.get('linh_vuc'),
                       request.form.get('phi'), request.form.get('le_phi'), request.form.get('cap_thuc_hien'),
                       request.form.get('co_quan_thuc_hien'), request.form.get('thanh_phan_ho_so'),
                       request.form.get('thoi_gian_giai_quyet'), request.form.get('can_cu_phap_ly'),
                       request.form.get('dvc_toan_trinh', ''), request.form.get('dvc_mot_phan', ''),
                       request.form.get('dvc_cung_cap_tt', ''), 'da_cong_bo', datetime.now().date()))
            conn.commit()
            conn.close()
            return '<script>alert("✅ Thêm TTHC thành công!"); window.location.href="/tthc";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ Lỗi: {str(e)}"); window.history.back();</script>'
    
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT DISTINCT linh_vuc FROM tthc WHERE linh_vuc IS NOT NULL AND linh_vuc != ''")
    linh_vuc_list = [row[0] for row in c.fetchall()]
    conn.close()
    
    return render_template('tthc_form.html', tthc=None, linh_vuc_list=linh_vuc_list,
                          username=session.get('username'), role=session.get('role'))

@app.route('/tthc/sua/<int:id>', methods=['GET', 'POST'])
@login_required
def tthc_edit(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    if request.method == 'POST':
        try:
            c.execute('''UPDATE tthc SET ma_tthc=?, ten_tthc=?, linh_vuc=?, phi=?, le_phi=?, 
                      cap_thuc_hien=?, co_quan_thuc_hien=?, thanh_phan_ho_so=?, thoi_gian_giai_quyet=?, 
                      can_cu_phap_ly=?, dvc_toan_trinh=?, dvc_mot_phan=?, dvc_cung_cap_tt=?
                      WHERE id=?''',
                      (request.form.get('ma_tthc'), request.form.get('ten_tthc'), request.form.get('linh_vuc'),
                       request.form.get('phi'), request.form.get('le_phi'), request.form.get('cap_thuc_hien'),
                       request.form.get('co_quan_thuc_hien'), request.form.get('thanh_phan_ho_so'),
                       request.form.get('thoi_gian_giai_quyet'), request.form.get('can_cu_phap_ly'),
                       request.form.get('dvc_toan_trinh', ''), request.form.get('dvc_mot_phan', ''),
                       request.form.get('dvc_cung_cap_tt', ''), id))
            conn.commit()
            conn.close()
            return '<script>alert("✅ Cập nhật TTHC thành công!"); window.location.href="/tthc";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ Lỗi: {str(e)}"); window.history.back();</script>'
    
    c.execute("SELECT * FROM tthc WHERE id=?", (id,))
    tthc = c.fetchone()
    c.execute("SELECT DISTINCT linh_vuc FROM tthc WHERE linh_vuc IS NOT NULL AND linh_vuc != ''")
    linh_vuc_list = [row[0] for row in c.fetchall()]
    conn.close()
    
    return render_template('tthc_form.html', tthc=tthc, linh_vuc_list=linh_vuc_list,
                          username=session.get('username'), role=session.get('role'))

@app.route('/tthc/xoa/<int:id>')
@admin_required
def tthc_delete(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("UPDATE tthc SET trang_thai = 'bai_bo' WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return '<script>alert("🗑️ Đã xóa TTHC!"); window.location.href="/tthc";</script>'

@app.route('/tthc/chitiet/<int:id>')
@login_required
def tthc_detail(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT * FROM tthc WHERE id=?", (id,))
    tthc = c.fetchone()
    conn.close()
    return render_template('tthc_detail.html', tthc=tthc, username=session.get('username'), role=session.get('role'))

# ==================== QUẢN LÝ QUYẾT ĐỊNH ====================
@app.route('/quyet_dinh')
@login_required
def quyet_dinh_list():
    page = request.args.get('page', 1, type=int)
    per_page = 10
    offset = (page - 1) * per_page
    
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    c.execute("SELECT COUNT(*) FROM quyet_dinh")
    total = c.fetchone()[0]
    
    c.execute("SELECT id, so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai, file_dinh_kem FROM quyet_dinh ORDER BY id DESC LIMIT ? OFFSET ?", (per_page, offset))
    quyet_dinh_list = c.fetchall()
    conn.close()
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('quyet_dinh_list.html', quyet_dinh_list=quyet_dinh_list, page=page, total_pages=total_pages,
                          username=session.get('username'), role=session.get('role'))

@app.route('/quyet_dinh/them', methods=['GET', 'POST'])
@login_required
def quyet_dinh_create():
    if request.method == 'POST':
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        
        try:
            so_quyet_dinh = request.form.get('so_quyet_dinh', '').strip()
            c.execute("SELECT COUNT(*) FROM quyet_dinh WHERE so_quyet_dinh = ?", (so_quyet_dinh,))
            if c.fetchone()[0] > 0:
                conn.close()
                return '<script>alert("❌ Số quyết định đã tồn tại!"); window.history.back();</script>'
            
            file_dinh_kem = ''
            if 'file_dinh_kem' in request.files:
                file = request.files['file_dinh_kem']
                if file and file.filename:
                    filename = secure_filename(file.filename)
                    filepath = os.path.join(app.config['QUYET_DINH_FOLDER'], filename)
                    file.save(filepath)
                    file_dinh_kem = filename
            
            c.execute('''INSERT INTO quyet_dinh (so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai, mo_ta, file_dinh_kem, ten_file_goc)
                      VALUES (?,?,?,?,?,?,?)''',
                      (so_quyet_dinh, request.form.get('ten_quyet_dinh'), request.form.get('ngay_ban_hanh'),
                       request.form.get('loai'), request.form.get('mo_ta'), file_dinh_kem, file_dinh_kem))
            conn.commit()
            conn.close()
            return '<script>alert("✅ Thêm quyết định thành công!"); window.location.href="/quyet_dinh";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ Lỗi: {str(e)}"); window.history.back();</script>'
    
    return render_template('quyet_dinh_form.html', quyet_dinh=None,
                          username=session.get('username'), role=session.get('role'))

@app.route('/quyet_dinh/sua/<int:id>', methods=['GET', 'POST'])
@login_required
def quyet_dinh_edit(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    if request.method == 'POST':
        try:
            file_dinh_kem = request.form.get('existing_file', '')
            if 'file_dinh_kem' in request.files:
                file = request.files['file_dinh_kem']
                if file and file.filename:
                    filename = secure_filename(file.filename)
                    filepath = os.path.join(app.config['QUYET_DINH_FOLDER'], filename)
                    file.save(filepath)
                    file_dinh_kem = filename
            
            c.execute('''UPDATE quyet_dinh SET so_quyet_dinh=?, ten_quyet_dinh=?, ngay_ban_hanh=?, 
                      loai=?, mo_ta=?, file_dinh_kem=? WHERE id=?''',
                      (request.form.get('so_quyet_dinh'), request.form.get('ten_quyet_dinh'),
                       request.form.get('ngay_ban_hanh'), request.form.get('loai'),
                       request.form.get('mo_ta'), file_dinh_kem, id))
            conn.commit()
            conn.close()
            return '<script>alert("✅ Cập nhật quyết định thành công!"); window.location.href="/quyet_dinh";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ Lỗi: {str(e)}"); window.history.back();</script>'
    
    c.execute("SELECT * FROM quyet_dinh WHERE id=?", (id,))
    quyet_dinh = c.fetchone()
    conn.close()
    
    return render_template('quyet_dinh_form.html', quyet_dinh=quyet_dinh,
                          username=session.get('username'), role=session.get('role'))

@app.route('/quyet_dinh/xoa/<int:id>')
@admin_required
def quyet_dinh_delete(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("DELETE FROM quyet_dinh WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return '<script>alert("🗑️ Đã xóa quyết định!"); window.location.href="/quyet_dinh";</script>'

@app.route('/quyet_dinh/download/<filename>')
@login_required
def download_quyet_dinh(filename):
    return send_from_directory(app.config['QUYET_DINH_FOLDER'], filename, as_attachment=True)

# ==================== BÁO CÁO ====================
@app.route('/bao_cao')
@login_required
def bao_cao():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    # Thống kê theo lĩnh vực
    c.execute("SELECT linh_vuc, COUNT(*) FROM tthc WHERE linh_vuc IS NOT NULL GROUP BY linh_vuc")
    linh_vuc_stats = c.fetchall()
    
    # Thống kê theo cấp thực hiện
    c.execute("SELECT cap_thuc_hien, COUNT(*) FROM tthc GROUP BY cap_thuc_hien")
    cap_stats = c.fetchall()
    
    # Thống kê DVC
    c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_toan_trinh = 'X'")
    toan_trinh = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_mot_phan = 'X'")
    mot_phan = c.fetchone()[0]
    
    conn.close()
    
    return render_template('bao_cao.html', linh_vuc_stats=linh_vuc_stats, cap_stats=cap_stats,
                          toan_trinh=toan_trinh, mot_phan=mot_phan,
                          username=session.get('username'), role=session.get('role'))

# ==================== TÌM KIẾM ====================
@app.route('/tim_kiem', methods=['GET', 'POST'])
@login_required
def tim_kiem():
    results = []
    keyword = ''
    
    if request.method == 'POST':
        keyword = request.form.get('keyword', '').strip()
        
        if keyword:
            conn = sqlite3.connect('tthc.db')
            c = conn.cursor()
            c.execute('''SELECT id, ma_tthc, ten_tthc, linh_vuc, cap_thuc_hien, co_quan_thuc_hien 
                      FROM tthc WHERE ten_tthc LIKE ? OR ma_tthc LIKE ? OR linh_vuc LIKE ?
                      ORDER BY ten_tthc''', (f'%{keyword}%', f'%{keyword}%', f'%{keyword}%'))
            results = c.fetchall()
            conn.close()
    
    return render_template('tim_kiem.html', results=results, keyword=keyword,
                          username=session.get('username'), role=session.get('role'))

# ==================== SO SÁNH QUYẾT ĐỊNH ====================
@app.route('/so_sanh_quyet_dinh', methods=['GET', 'POST'])
@login_required
def so_sanh_quyet_dinh():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT id, so_quyet_dinh, ten_quyet_dinh FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    quyet_dinh_list = c.fetchall()
    conn.close()
    
    result = None
    qd1 = None
    qd2 = None
    
    if request.method == 'POST':
        qd1_id = request.form.get('qd1_id', type=int)
        qd2_id = request.form.get('qd2_id', type=int)
        
        if qd1_id and qd2_id:
            conn = sqlite3.connect('tthc.db')
            c = conn.cursor()
            c.execute("SELECT * FROM quyet_dinh WHERE id=?", (qd1_id,))
            qd1 = c.fetchone()
            c.execute("SELECT * FROM quyet_dinh WHERE id=?", (qd2_id,))
            qd2 = c.fetchone()
            conn.close()
            
            result = {
                'so_quyet_dinh': [qd1[1] if qd1 else '', qd2[1] if qd2 else ''],
                'ten_quyet_dinh': [qd1[2] if qd1 else '', qd2[2] if qd2 else ''],
                'ngay_ban_hanh': [qd1[3] if qd1 else '', qd2[3] if qd2 else ''],
                'loai': [qd1[4] if qd1 else '', qd2[4] if qd2 else '']
            }
    
    return render_template('so_sanh_quyet_dinh.html', quyet_dinh_list=quyet_dinh_list,
                          result=result, qd1=qd1, qd2=qd2,
                          username=session.get('username'), role=session.get('role'))

# ==================== XEM PDF ====================
@app.route('/uploads/quyet_dinh/<filename>')
@login_required
def xem_pdf(filename):
    return send_from_directory(app.config['QUYET_DINH_FOLDER'], filename)

# ==================== API ====================
@app.route('/api/thong_ke')
@login_required
def api_thong_ke():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM tthc")
    tong = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_toan_trinh = 'X'")
    toan_trinh = c.fetchone()[0]
    conn.close()
    return jsonify({'tong_tthc': tong, 'toan_trinh': toan_trinh})

# ==================== IMPORT EXCEL ====================
@app.route('/import_excel', methods=['POST'])
@admin_required
def import_excel():
    if 'file' not in request.files:
        return '<script>alert("❌ Không có file được chọn!"); window.history.back();</script>'
    
    file = request.files['file']
    if file.filename == '':
        return '<script>alert("❌ Chưa chọn file!"); window.history.back();</script>'
    
    if file and file.filename.endswith(('.xlsx', '.xls')):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            wb = load_workbook(filepath)
            ws = wb.active
            conn = sqlite3.connect('tthc.db')
            c = conn.cursor()
            success = 0
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:
                    try:
                        c.execute('''INSERT OR IGNORE INTO tthc (ma_tthc, ten_tthc, linh_vuc, cap_thuc_hien, created_at)
                                  VALUES (?,?,?,?,?)''', (str(row[0]), str(row[1]), str(row[2]) if row[2] else '', 
                                  str(row[3]) if row[3] else '', datetime.now().date()))
                        success += 1
                    except: pass
            
            conn.commit()
            conn.close()
            os.remove(filepath)
            return f'<script>alert("✅ Import thành công {success} dòng dữ liệu!"); window.location.href="/tthc";</script>'
        except Exception as e:
            return f'<script>alert("❌ Lỗi import: {str(e)}"); window.history.back();</script>'
    
    return '<script>alert("❌ File không hợp lệ!"); window.history.back();</script>'

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

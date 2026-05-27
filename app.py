import time
time.sleep(3)  # Chờ 3 giây trước khi khởi động
import os
import sqlite3
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, jsonify
from datetime import datetime
import uuid
from werkzeug.utils import secure_filename
import pandas as pd
import json
import difflib
import re

# OCR imports
try:
    from PIL import Image
    import pytesseract
    from pdf2image import convert_from_path
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
QUYET_DINH_FOLDER = 'uploads/quyet_dinh'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['QUYET_DINH_FOLDER'] = QUYET_DINH_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(QUYET_DINH_FOLDER, exist_ok=True)

if OCR_AVAILABLE:
    try:
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    except:
        pass

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
    
    conn.commit()
    
    # Thêm cột trang_thai nếu chưa có
    try:
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        c.execute("ALTER TABLE tthc ADD COLUMN trang_thai TEXT DEFAULT 'da_cong_bo'")
        conn.commit()
        conn.close()
    except:
        pass
    
    conn.close()
    print("Đã khởi tạo database với cấu trúc bảng trống!")

init_db()

# ==================== HÀM OCR ====================
def extract_text_from_pdf(filepath):
    try:
        import PyPDF2
        with open(filepath, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
            return text if text.strip() else ""
    except:
        return ""

def extract_text_from_image_pdf(filepath):
    if not OCR_AVAILABLE:
        return ""
    try:
        images = convert_from_path(filepath)
        text = ""
        for image in images:
            image = image.convert('L')
            text += pytesseract.image_to_string(image, lang='vie+eng') + "\n"
        return text
    except:
        return ""

# ==================== TRANG CHỦ ====================
@app.route('/')
def dashboard():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    thong_ke = {
        'cap_bo': 0,
        'cap_tinh': 0,
        'cap_xa': 0,
        'cap_dung_chung': 0,
        'cap_lien_thong': 0,
        'trang_thai_toan_trinh': 0,
        'trang_thai_cong_khai_mot_phan': 0,
        'trang_thai_chua_cong_khai': 0,
        'so_da_cong_bo': 0,
        'so_bai_bo': 0,
        'linh_vuc': []
    }
    
    try:
        for cap in ['bo', 'tinh', 'xa', 'dung_chung', 'lien_thong']:
            c.execute("SELECT COUNT(*) FROM tthc WHERE cap_thuc_hien=?", (cap,))
            result = c.fetchone()
            if result and result[0]:
                thong_ke[f'cap_{cap}'] = result[0]
    except Exception as e:
        print(f"Lỗi đếm cấp: {e}")
    
    try:
        c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_toan_trinh = 'X'")
        thong_ke['trang_thai_toan_trinh'] = c.fetchone()[0] or 0
        
        c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_mot_phan = 'X'")
        thong_ke['trang_thai_cong_khai_mot_phan'] = c.fetchone()[0] or 0
        
        c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_cung_cap_tt = 'X'")
        thong_ke['trang_thai_chua_cong_khai'] = c.fetchone()[0] or 0
    except Exception as e:
        print(f"Lỗi đếm trạng thái công khai: {e}")
    
    try:
        c.execute("SELECT COUNT(*) FROM tthc WHERE trang_thai = 'da_cong_bo'")
        thong_ke['so_da_cong_bo'] = c.fetchone()[0] or 0
        
        c.execute("SELECT COUNT(*) FROM tthc WHERE trang_thai = 'bai_bo'")
        thong_ke['so_bai_bo'] = c.fetchone()[0] or 0
    except Exception as e:
        print(f"Lỗi đếm trạng thái: {e}")
    
    try:
        c.execute("SELECT linh_vuc, COUNT(*) FROM tthc WHERE linh_vuc IS NOT NULL AND linh_vuc != '' GROUP BY linh_vuc ORDER BY COUNT(*) DESC")
        raw_data = c.fetchall()
        if raw_data:
            thong_ke['linh_vuc'] = [{'ten': str(row[0]), 'so_luong': row[1]} for row in raw_data]
    except Exception as e:
        print(f"Lỗi thống kê lĩnh vực: {e}")
    
    conn.close()
    
    print("=" * 50)
    print("DASHBOARD THỐNG KÊ:")
    print(f"Cấp Bộ: {thong_ke['cap_bo']}")
    print(f"Cấp tỉnh: {thong_ke['cap_tinh']}")
    print(f"Cấp xã: {thong_ke['cap_xa']}")
    print(f"Dùng chung: {thong_ke['cap_dung_chung']}")
    print(f"Liên thông: {thong_ke['cap_lien_thong']}")
    print(f"Toàn trình: {thong_ke['trang_thai_toan_trinh']}")
    print(f"Một phần: {thong_ke['trang_thai_cong_khai_mot_phan']}")
    print(f"Cung cấp TT: {thong_ke['trang_thai_chua_cong_khai']}")
    print(f"Đã công bố: {thong_ke['so_da_cong_bo']}")
    print(f"Bãi bỏ: {thong_ke['so_bai_bo']}")
    print("=" * 50)
    
    return render_template('dashboard.html', thong_ke=thong_ke)

# ==================== API JSON ====================
@app.route('/api/quyet_dinh')
def api_quyet_dinh():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT id, so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    data = []
    for row in c.fetchall():
        data.append({
            'id': row[0],
            'so_quyet_dinh': row[1],
            'ten_quyet_dinh': row[2],
            'ngay_ban_hanh': row[3],
            'loai': row[4]
        })
    conn.close()
    return jsonify(data)

# ==================== TẢI FILE MẪU EXCEL ====================
@app.route('/tai_file_mau')
def tai_file_mau():
    import io
    from flask import send_file
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        mau_data = pd.DataFrame({
            'STT': [1, 2, 3, 4, 5],
            'MÃ TTHC': ['TTHC001', 'TTHC002', 'TTHC003', 'TTHC004', 'TTHC005'],
            'Tên TTHC': ['Đăng ký kinh doanh', 'Chứng thực bản sao', 'Cấp giấy phép xây dựng', 'Bổ nhiệm công chứng viên', 'Đăng ký hộ tịch'],
            'Lĩnh vực': ['Kinh doanh', 'Hộ tịch', 'Xây dựng', 'Công chứng', 'Hộ tịch'],
            'Phí': ['X', '', 'X', 'X', ''],
            'Lệ Phí': ['', 'X', '', '', 'X'],
            'Cấp thực hiện': ['tinh', 'xa', 'lien_thong', 'bo', 'xa'],
            'Cơ quan thực hiện': ['Sở Kế hoạch Đầu tư', 'UBND cấp xã', 'UBND cấp xã, Sở Tư pháp', 'Bộ Tư pháp', 'UBND cấp xã'],
            'Cùng cấp': ['X', '', '', '', ''],
            '02 cấp': ['', 'X', '', '', ''],
            'Phi địa giới': ['', '', 'X', '', ''],
            'Toàn trình': ['X', '', '', 'X', ''],
            'Một phần': ['', 'X', '', '', 'X'],
            'Cung cấp thông tin': ['', '', 'X', '', ''],
            'Dịch vụ BCCI': ['Có', '', 'Không', 'Có', ''],
            'Quyết định số': ['123/QĐ-UBND', '456/QĐ-UBND', '789/QĐ-UBND', '111/QĐ-BTP', '222/QĐ-UBND'],
            'TRẠNG THÁI': ['Đã công bố', 'Đã công bố', 'Đã công bố', 'Bãi bỏ', 'Đã công bố'],
            'Ghi chú': ['Áp dụng từ 01/2024', 'Có phí công chứng', 'Liên thông 3 cấp', 'Theo Nghị định mới', '']
        })
        mau_data.to_excel(writer, sheet_name='Sheet1', index=False)
    
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Mau_Import_TTHC.xlsx'
    )

# ==================== DANH SÁCH TTHC ====================
@app.route('/tthc')
def danh_sach_tthc():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute('''SELECT id, ma_tthc, ten_tthc, linh_vuc, phi, le_phi, cap_thuc_hien, co_quan_thuc_hien,
                lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt,
                dich_vu_bcci, ghi_chu, so_quyet_dinh, trang_thai
                FROM tthc ORDER BY id DESC''')
    data = c.fetchall()
    conn.close()
    return render_template('tthc_list.html', tthcs=data)

# ==================== THÊM TTHC ====================
@app.route('/tthc/them', methods=['GET', 'POST'])
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
                c.execute('''UPDATE tthc SET 
                    ten_tthc=?, linh_vuc=?, phi=?, le_phi=?, 
                    lien_thong_cung_cap=?, lien_thong_02_cap=?, phi_dia_gioi=?,
                    dvc_toan_trinh=?, dvc_mot_phan=?, dvc_cung_cap_tt=?, 
                    dich_vu_bcci=?, ghi_chu=?, cap_thuc_hien=?, trang_thai_cong_khai=?, 
                    so_quyet_dinh=?, co_quan_thuc_hien=?, trang_thai=?
                    WHERE ma_tthc = ?''',
                    (ten_tthc, linh_vuc, phi, le_phi, 
                     lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                     dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, 
                     dich_vu_bcci, ghi_chu, cap_thuc_hien, trang_thai_cong_khai, 
                     so_qd, co_quan, trang_thai_tt, ma_tthc))
            else:
                c.execute('''INSERT INTO tthc 
                    (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, 
                     lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                     dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, 
                     dich_vu_bcci, ghi_chu, cap_thuc_hien, trang_thai_cong_khai, 
                     so_quyet_dinh, co_quan_thuc_hien, trang_thai)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                    (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, 
                     lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                     dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, 
                     dich_vu_bcci, ghi_chu, cap_thuc_hien, trang_thai_cong_khai, 
                     so_qd, co_quan, trang_thai_tt))
            
            conn.commit()
            conn.close()
            return '<script>alert("✅ THEM THANH CONG!"); window.location.href="/tthc";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ LOI: {str(e)}"); window.history.back();</script>'
    
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT id, so_quyet_dinh FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    quyet_dinh_list = c.fetchall()
    conn.close()
    return render_template('tthc_form.html', tthc=None, quyet_dinh_list=quyet_dinh_list)

def import_excel_file(file):
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        df = pd.read_excel(filepath, header=0)
        conn = sqlite3.connect('tthc.db')
        c = conn.cursor()
        
        # Đảm bảo có quyết định mặc định
        c.execute("SELECT COUNT(*) FROM quyet_dinh")
        if c.fetchone()[0] == 0:
            c.execute("INSERT INTO quyet_dinh (so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai, mo_ta) VALUES (?,?,?,?,?)",
                      ("QD001", "Quyết định công bố mặc định", datetime.now().date(), "cong_bo", "Tự động tạo khi import"))
        
        so_luong_thanh_cong = 0
        danh_sach_loi = []
        danh_sach_thanh_cong = []
        
        for idx, row in df.iterrows():
            try:
                ma_tthc = str(row.get('MÃ TTHC', '')).strip() if pd.notna(row.get('MÃ TTHC')) else ''
                ten_tthc = str(row.get('Tên TTHC', '')).strip() if pd.notna(row.get('Tên TTHC')) else ''
                
                if not ma_tthc or ma_tthc == 'nan' or ma_tthc == '':
                    danh_sach_loi.append(f"Dong {idx+2}: Thieu MA TTHC")
                    continue
                
                linh_vuc = str(row.get('Lĩnh vực', '')) if pd.notna(row.get('Lĩnh vực')) else ''
                phi = 'X' if pd.notna(row.get('Phí')) and str(row.get('Phí')).upper() == 'X' else ''
                le_phi = 'X' if pd.notna(row.get('Lệ Phí')) and str(row.get('Lệ Phí')).upper() == 'X' else ''
                cap_thuc_hien = str(row.get('Cấp thực hiện', '')) if pd.notna(row.get('Cấp thực hiện')) else ''
                co_quan = str(row.get('Cơ quan thực hiện', '')) if pd.notna(row.get('Cơ quan thực hiện')) else ''
                lien_thong_cung_cap = 'X' if pd.notna(row.get('Cùng cấp')) and str(row.get('Cùng cấp')).upper() == 'X' else ''
                lien_thong_02_cap = 'X' if pd.notna(row.get('02 cấp')) and str(row.get('02 cấp')).upper() == 'X' else ''
                phi_dia_gioi = 'X' if pd.notna(row.get('Phi địa giới')) and str(row.get('Phi địa giới')).upper() == 'X' else ''
                dvc_toan_trinh = 'X' if pd.notna(row.get('Toàn trình')) and str(row.get('Toàn trình')).upper() == 'X' else ''
                dvc_mot_phan = 'X' if pd.notna(row.get('Một phần')) and str(row.get('Một phần')).upper() == 'X' else ''
                dvc_cung_cap_tt = 'X' if pd.notna(row.get('Cung cấp thông tin')) and str(row.get('Cung cấp thông tin')).upper() == 'X' else ''
                dich_vu_bcci = str(row.get('Dịch vụ BCCI', '')) if pd.notna(row.get('Dịch vụ BCCI')) else ''
                so_qd = str(row.get('Quyết định số', '')) if pd.notna(row.get('Quyết định số')) else ''
                ghi_chu = str(row.get('Ghi chú', '')) if pd.notna(row.get('Ghi chú')) else ''
                
                if not so_qd or so_qd == 'nan':
                    c.execute("SELECT so_quyet_dinh FROM quyet_dinh LIMIT 1")
                    default_qd = c.fetchone()
                    if default_qd:
                        so_qd = default_qd[0]
                
                trang_thai_excel = str(row.get('TRẠNG THÁI', '')).strip() if pd.notna(row.get('TRẠNG THÁI')) else ''
                trang_thai_tt = 'bai_bo' if trang_thai_excel == 'Bãi bỏ' else 'da_cong_bo'
                
                trang_thai_cong_khai = ''
                if dvc_toan_trinh == 'X':
                    trang_thai_cong_khai = 'toan_trinh'
                elif dvc_mot_phan == 'X':
                    trang_thai_cong_khai = 'cong_khai_mot_phan'
                elif dvc_cung_cap_tt == 'X':
                    trang_thai_cong_khai = 'chua_cong_khai'
                
                if cap_thuc_hien and cap_thuc_hien not in ['bo', 'tinh', 'xa', 'dung_chung', 'lien_thong', '']:
                    danh_sach_loi.append(f"Dong {idx+2}: Cap thuc hien '{cap_thuc_hien}' khong hop le")
                    continue
                
                if not co_quan and cap_thuc_hien:
                    co_quan = get_co_quan_by_cap(cap_thuc_hien)
                
                c.execute("SELECT id FROM tthc WHERE ma_tthc = ?", (ma_tthc,))
                existing = c.fetchone()
                
                if existing:
                    c.execute('''UPDATE tthc SET 
                        ten_tthc=?, linh_vuc=?, phi=?, le_phi=?, 
                        lien_thong_cung_cap=?, lien_thong_02_cap=?, phi_dia_gioi=?,
                        dvc_toan_trinh=?, dvc_mot_phan=?, dvc_cung_cap_tt=?, 
                        dich_vu_bcci=?, ghi_chu=?, cap_thuc_hien=?, trang_thai_cong_khai=?, 
                        so_quyet_dinh=?, co_quan_thuc_hien=?, trang_thai=?
                        WHERE ma_tthc = ?''',
                        (ten_tthc, linh_vuc, phi, le_phi, 
                         lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                         dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, 
                         dich_vu_bcci, ghi_chu, cap_thuc_hien, trang_thai_cong_khai, 
                         so_qd, co_quan, trang_thai_tt, ma_tthc))
                else:
                    c.execute('''INSERT INTO tthc 
                        (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, 
                         lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                         dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, 
                         dich_vu_bcci, ghi_chu, cap_thuc_hien, trang_thai_cong_khai, 
                         so_quyet_dinh, co_quan_thuc_hien, trang_thai)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                        (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, 
                         lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                         dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, 
                         dich_vu_bcci, ghi_chu, cap_thuc_hien, trang_thai_cong_khai, 
                         so_qd, co_quan, trang_thai_tt))
                
                so_luong_thanh_cong += 1
                danh_sach_thanh_cong.append(f"{ma_tthc} - {ten_tthc}")
                
            except Exception as e:
                danh_sach_loi.append(f"Dong {idx+2}: {str(e)}")
        
        conn.commit()
        conn.close()
        
        try:
            os.remove(filepath)
        except:
            pass
        
        if so_luong_thanh_cong > 0:
            thong_bao = f"✅ ĐÃ THÊM DỮ LIỆU THÀNH CÔNG!\n\nSố thủ tục đã thêm: {so_luong_thanh_cong}"
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
def sua_tthc(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
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
            c.execute('''UPDATE tthc SET 
                ma_tthc=?, ten_tthc=?, linh_vuc=?, phi=?, le_phi=?, 
                lien_thong_cung_cap=?, lien_thong_02_cap=?, phi_dia_gioi=?,
                dvc_toan_trinh=?, dvc_mot_phan=?, dvc_cung_cap_tt=?, 
                dich_vu_bcci=?, ghi_chu=?, cap_thuc_hien=?, trang_thai_cong_khai=?, 
                so_quyet_dinh=?, co_quan_thuc_hien=?, trang_thai=?
                WHERE id=?''',
                (ma_tthc, ten_tthc, linh_vuc, phi, le_phi, 
                 lien_thong_cung_cap, lien_thong_02_cap, phi_dia_gioi,
                 dvc_toan_trinh, dvc_mot_phan, dvc_cung_cap_tt, 
                 dich_vu_bcci, ghi_chu, cap_thuc_hien, trang_thai_cong_khai, 
                 so_qd, co_quan, trang_thai_tt, id))
            
            conn.commit()
            conn.close()
            return '<script>alert("✅ CAP NHAT THANH CONG!"); window.location.href="/tthc";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ LOI: {str(e)}"); window.history.back();</script>'
    
    c.execute("SELECT * FROM tthc WHERE id=?", (id,))
    tthc = c.fetchone()
    c.execute("SELECT id, so_quyet_dinh FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    quyet_dinh_list = c.fetchall()
    conn.close()
    return render_template('tthc_form.html', tthc=tthc, quyet_dinh_list=quyet_dinh_list)

# ==================== XÓA TTHC ====================
@app.route('/tthc/xoa/<int:id>')
def xoa_tthc(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("DELETE FROM quy_trinh WHERE tthc_id=?", (id,))
    c.execute("DELETE FROM tthc WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return '<script>alert("🗑️ Da xoa thu tuc!"); window.location.href="/tthc";</script>'

# ==================== CHI TIẾT TTHC ====================
@app.route('/tthc/chi_tiet/<int:id>')
def chi_tiet_tthc(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT * FROM tthc WHERE id=?", (id,))
    tthc = c.fetchone()
    conn.close()
    return render_template('tthc_detail.html', tthc=tthc)

# ==================== QUẢN LÝ QUYẾT ĐỊNH ====================
@app.route('/quyet_dinh')
def danh_sach_quyet_dinh():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT * FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    data = c.fetchall()
    conn.close()
    return render_template('quyet_dinh_list.html', quyet_dinhs=data)

@app.route('/quyet_dinh/them', methods=['GET', 'POST'])
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
        noi_dung_ocr = None
        
        if 'file_dinh_kem' in request.files:
            file = request.files['file_dinh_kem']
            if file and file.filename != '' and file.filename.endswith('.pdf'):
                ten_file_goc = secure_filename(file.filename)
                unique_name = f"{uuid.uuid4().hex}_{ten_file_goc}"
                filepath = os.path.join(app.config['QUYET_DINH_FOLDER'], unique_name)
                file.save(filepath)
                file_dinh_kem = unique_name
                noi_dung_ocr = extract_text_from_pdf(filepath)
                if not noi_dung_ocr:
                    noi_dung_ocr = extract_text_from_image_pdf(filepath)
        
        try:
            c.execute("INSERT INTO quyet_dinh (so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai, mo_ta, file_dinh_kem, ten_file_goc, noi_dung_ocr) VALUES (?,?,?,?,?,?,?,?)",
                      (so_qd, ten_qd, ngay_bh, loai, mo_ta, file_dinh_kem, ten_file_goc, noi_dung_ocr))
            conn.commit()
            conn.close()
            return '<script>alert("✅ Them quyet dinh thanh cong!"); window.location.href="/quyet_dinh";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ Loi: {str(e)}"); window.history.back();</script>'
    
    return render_template('quyet_dinh_form.html', quyet_dinh=None)

@app.route('/quyet_dinh/sua/<int:id>', methods=['GET', 'POST'])
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
        noi_dung_ocr = None
        
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
                noi_dung_ocr = extract_text_from_pdf(filepath)
                if not noi_dung_ocr:
                    noi_dung_ocr = extract_text_from_image_pdf(filepath)
        
        try:
            c.execute("UPDATE quyet_dinh SET so_quyet_dinh=?, ten_quyet_dinh=?, ngay_ban_hanh=?, loai=?, mo_ta=?, file_dinh_kem=?, ten_file_goc=?, noi_dung_ocr=? WHERE id=?",
                      (so_qd, ten_qd, ngay_bh, loai, mo_ta, file_dinh_kem, ten_file_goc, noi_dung_ocr, id))
            conn.commit()
            conn.close()
            return '<script>alert("✅ Cap nhat quyet dinh thanh cong!"); window.location.href="/quyet_dinh";</script>'
        except Exception as e:
            conn.close()
            return f'<script>alert("❌ Loi: {str(e)}"); window.history.back();</script>'
    
    c.execute("SELECT * FROM quyet_dinh WHERE id=?", (id,))
    quyet_dinh = c.fetchone()
    conn.close()
    return render_template('quyet_dinh_form.html', quyet_dinh=quyet_dinh)

@app.route('/quyet_dinh/xoa/<int:id>')
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
    return '<script>alert("🗑️ Da xoa quyet dinh!"); window.location.href="/quyet_dinh";</script>'

# ==================== XEM PDF ====================
@app.route('/xem_pdf/<int:id>')
def xem_pdf(id):
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    c.execute("SELECT file_dinh_kem, ten_file_goc, noi_dung_ocr FROM quyet_dinh WHERE id=?", (id,))
    result = c.fetchone()
    conn.close()
    if result and result[0]:
        return render_template('xem_pdf.html', file_name=result[0], ten_file=result[1], noi_dung_ocr=result[2], quyet_dinh_id=id)
    return '<script>alert("Khong co file dinh kem!"); window.history.back();</script>'

@app.route('/uploads/quyet_dinh/<filename>')
def download_pdf(filename):
    return send_from_directory(app.config['QUYET_DINH_FOLDER'], filename)

# ==================== SO SÁNH QUYẾT ĐỊNH ====================
@app.route('/so_sanh_quyet_dinh')
def so_sanh_qd():
    return render_template('so_sanh_quyet_dinh.html')

@app.route('/api/so_sanh_quyet_dinh')
def api_so_sanh_quyet_dinh():
    qd_ids = request.args.get('qd_ids', '')
    if not qd_ids:
        return jsonify({'error': 'Chua chon quyet dinh'})
    ids = [int(x) for x in qd_ids.split(',')]
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    quyet_dinhs = []
    noi_dung = {}
    for qd_id in ids:
        c.execute("SELECT id, so_quyet_dinh, ten_quyet_dinh, ngay_ban_hanh, loai, noi_dung_ocr FROM quyet_dinh WHERE id=?", (qd_id,))
        qd = c.fetchone()
        if qd:
            quyet_dinhs.append({'id': qd[0], 'so_quyet_dinh': qd[1], 'ten_quyet_dinh': qd[2], 'ngay_ban_hanh': qd[3], 'loai': qd[4], 'noi_dung': qd[5] or ''})
            noi_dung[qd_id] = qd[5] or ''
    tthc_theo_qd = {}
    for qd_id in ids:
        qd = next((q for q in quyet_dinhs if q['id'] == qd_id), None)
        if qd:
            c.execute("SELECT ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc WHERE so_quyet_dinh=?", (qd['so_quyet_dinh'],))
            tthc_theo_qd[qd_id] = c.fetchall()
    conn.close()
    so_sanh_van_ban = []
    if len(ids) >= 2:
        for i in range(len(ids)):
            for j in range(i+1, len(ids)):
                if noi_dung.get(ids[i]) and noi_dung.get(ids[j]):
                    diff = difflib.HtmlDiff(wrapcolumn=80).make_file(noi_dung[ids[i]].splitlines(), noi_dung[ids[j]].splitlines(), fromdesc=f"QĐ {quyet_dinhs[i]['so_quyet_dinh']}", todesc=f"QĐ {quyet_dinhs[j]['so_quyet_dinh']}")
                    so_sanh_van_ban.append({'qd1': quyet_dinhs[i]['so_quyet_dinh'], 'qd2': quyet_dinhs[j]['so_quyet_dinh'], 'diff_html': diff})
    return jsonify({'quyet_dinhs': quyet_dinhs, 'tthc_theo_qd': tthc_theo_qd, 'so_sanh_van_ban': so_sanh_van_ban})

# ==================== BÁO CÁO THỐNG KÊ NÂNG CAO ====================
@app.route('/bao_cao')
def bao_cao():
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    
    c.execute("SELECT id, so_quyet_dinh FROM quyet_dinh ORDER BY ngay_ban_hanh DESC")
    danh_sach_quyet_dinh = c.fetchall()
    
    c.execute("SELECT DISTINCT linh_vuc FROM tthc WHERE linh_vuc IS NOT NULL AND linh_vuc != '' ORDER BY linh_vuc")
    danh_sach_linh_vuc = c.fetchall()
    
    c.execute("SELECT DISTINCT cap_thuc_hien FROM tthc WHERE cap_thuc_hien IS NOT NULL AND cap_thuc_hien != '' ORDER BY cap_thuc_hien")
    danh_sach_cap = c.fetchall()
    
    c.execute("SELECT DISTINCT co_quan_thuc_hien FROM tthc WHERE co_quan_thuc_hien IS NOT NULL AND co_quan_thuc_hien != '' ORDER BY co_quan_thuc_hien")
    danh_sach_co_quan = c.fetchall()
    
    c.execute("SELECT DISTINCT dich_vu_bcci FROM tthc WHERE dich_vu_bcci IS NOT NULL AND dich_vu_bcci != '' ORDER BY dich_vu_bcci")
    danh_sach_bcci = c.fetchall()
    
    conn.close()
    return render_template('bao_cao.html', 
                         danh_sach_quyet_dinh=danh_sach_quyet_dinh,
                         danh_sach_linh_vuc=danh_sach_linh_vuc,
                         danh_sach_cap=danh_sach_cap,
                         danh_sach_co_quan=danh_sach_co_quan,
                         danh_sach_bcci=danh_sach_bcci)

@app.route('/api/bao_cao_data')
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
        if trang_thai == 'toan_trinh':
            sql += " AND t.dvc_toan_trinh = 'X'"
        elif trang_thai == 'mot_phan':
            sql += " AND t.dvc_mot_phan = 'X'"
        elif trang_thai == 'cung_cap_tt':
            sql += " AND t.dvc_cung_cap_tt = 'X'"
    
    if trang_thai_tt and trang_thai_tt != '':
        if trang_thai_tt == 'da_cong_bo':
            sql += " AND t.trang_thai = 'da_cong_bo'"
        elif trang_thai_tt == 'bai_bo':
            sql += " AND t.trang_thai = 'bai_bo'"
    
    if phi_le_phi and phi_le_phi != '':
        if phi_le_phi == 'co_phi':
            sql += " AND t.phi = 'X'"
        elif phi_le_phi == 'co_le_phi':
            sql += " AND t.le_phi = 'X'"
        elif phi_le_phi == 'mien_thu':
            sql += " AND (t.phi IS NULL OR t.phi = '') AND (t.le_phi IS NULL OR t.le_phi = '')"
    
    if dich_vu_bcci and dich_vu_bcci != '':
        if dich_vu_bcci == 'co':
            sql += " AND t.dich_vu_bcci IS NOT NULL AND t.dich_vu_bcci != ''"
        elif dich_vu_bcci == 'khong':
            sql += " AND (t.dich_vu_bcci IS NULL OR t.dich_vu_bcci = '')"
        else:
            sql += " AND t.dich_vu_bcci = ?"
            params.append(dich_vu_bcci)
    
    if co_quan_thuc_hien and co_quan_thuc_hien != '':
        sql += " AND t.co_quan_thuc_hien = ?"
        params.append(co_quan_thuc_hien)
    
    sql += " ORDER BY t.ma_tthc"
    
    c.execute(sql, params)
    data = c.fetchall()
    
    # Tổng số
    c.execute("SELECT COUNT(*) FROM tthc")
    tong_so = c.fetchone()[0]
    
    # Thống kê DVC trực tuyến
    c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_toan_trinh = 'X'")
    so_toan_trinh = c.fetchone()[0]
    
    c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_mot_phan = 'X'")
    so_mot_phan = c.fetchone()[0]
    
    c.execute("SELECT COUNT(*) FROM tthc WHERE dvc_cung_cap_tt = 'X'")
    so_cung_cap_tt = c.fetchone()[0]
    
    # Thống kê trạng thái
    c.execute("SELECT COUNT(*) FROM tthc WHERE trang_thai = 'da_cong_bo'")
    so_da_cong_bo = c.fetchone()[0]
    
    c.execute("SELECT COUNT(*) FROM tthc WHERE trang_thai = 'bai_bo'")
    so_bai_bo = c.fetchone()[0]
    
    # Thống kê Phí, Lệ phí, Miễn thu, BCCI
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
            'ma_tthc': row[0],
            'ten_tthc': row[1],
            'linh_vuc': row[2] or '',
            'phi': 'Có' if row[3] == 'X' else 'Không',
            'le_phi': 'Có' if row[4] == 'X' else 'Không',
            'cap_thuc_hien': row[5] or '',
            'co_quan_thuc_hien': row[6] or '',
            'toan_trinh': 'X' if row[7] == 'X' else '',
            'mot_phan': 'X' if row[8] == 'X' else '',
            'cung_cap_tt': 'X' if row[9] == 'X' else '',
            'dich_vu_bcci': row[10] or '',
            'ghi_chu': row[11] or '',
            'so_quyet_dinh': row[12] or '',
            'trang_thai': row[13] or 'da_cong_bo'
        })
    
    return jsonify({
        'data': ket_qua,
        'tong_so': tong_so,
        'so_toan_trinh': so_toan_trinh,
        'so_mot_phan': so_mot_phan,
        'so_cung_cap_tt': so_cung_cap_tt,
        'so_da_cong_bo': so_da_cong_bo,
        'so_bai_bo': so_bai_bo,
        'so_co_phi': so_co_phi,
        'so_co_le_phi': so_co_le_phi,
        'so_mien_thu': so_mien_thu,
        'so_co_bcci': so_co_bcci
    })

@app.route('/export_bao_cao_excel')
def export_bao_cao_excel():
    import io
    from flask import send_file
    
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
        if trang_thai == 'toan_trinh':
            sql += " AND t.dvc_toan_trinh = 'X'"
        elif trang_thai == 'mot_phan':
            sql += " AND t.dvc_mot_phan = 'X'"
        elif trang_thai == 'cung_cap_tt':
            sql += " AND t.dvc_cung_cap_tt = 'X'"
    
    if trang_thai_tt and trang_thai_tt != '':
        if trang_thai_tt == 'da_cong_bo':
            sql += " AND t.trang_thai = 'da_cong_bo'"
        elif trang_thai_tt == 'bai_bo':
            sql += " AND t.trang_thai = 'bai_bo'"
    
    if phi_le_phi and phi_le_phi != '':
        if phi_le_phi == 'co_phi':
            sql += " AND t.phi = 'X'"
        elif phi_le_phi == 'co_le_phi':
            sql += " AND t.le_phi = 'X'"
        elif phi_le_phi == 'mien_thu':
            sql += " AND (t.phi IS NULL OR t.phi = '') AND (t.le_phi IS NULL OR t.le_phi = '')"
    
    if dich_vu_bcci and dich_vu_bcci != '':
        if dich_vu_bcci == 'co':
            sql += " AND t.dich_vu_bcci IS NOT NULL AND t.dich_vu_bcci != ''"
        elif dich_vu_bcci == 'khong':
            sql += " AND (t.dich_vu_bcci IS NULL OR t.dich_vu_bcci = '')"
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
    
    df = pd.DataFrame(data, columns=[
        'MÃ TTHC', 'TÊN TTHC', 'LĨNH VỰC', 'PHÍ', 'LỆ PHÍ',
        'CẤP THỰC HIỆN', 'CƠ QUAN THỰC HIỆN', 'TOÀN TRÌNH', 'MỘT PHẦN', 'CUNG CẤP TT',
        'DỊCH VỤ BCCI', 'GHI CHÚ', 'QUYẾT ĐỊNH SỐ', 'TRẠNG THÁI'
    ])
    
    df['PHÍ'] = df['PHÍ'].apply(lambda x: 'X' if x == 'X' else '')
    df['LỆ PHÍ'] = df['LỆ PHÍ'].apply(lambda x: 'X' if x == 'X' else '')
    df['TOÀN TRÌNH'] = df['TOÀN TRÌNH'].apply(lambda x: 'X' if x == 'X' else '')
    df['MỘT PHẦN'] = df['MỘT PHẦN'].apply(lambda x: 'X' if x == 'X' else '')
    df['CUNG CẤP TT'] = df['CUNG CẤP TT'].apply(lambda x: 'X' if x == 'X' else '')
    df['TRẠNG THÁI'] = df['TRẠNG THÁI'].apply(lambda x: 'Đã công bố' if x == 'da_cong_bo' else 'Bãi bỏ')
    
    cap_map = {'bo': 'Cấp Bộ', 'tinh': 'Cấp tỉnh', 'xa': 'Cấp xã', 'dung_chung': 'Dùng chung', 'lien_thong': 'Liên thông'}
    df['CẤP THỰC HIỆN'] = df['CẤP THỰC HIỆN'].apply(lambda x: cap_map.get(x, x))
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='BaoCaoTTHC', index=False)
        worksheet = writer.sheets['BaoCaoTTHC']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    ten_file = f'BaoCao_TTHC_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=ten_file
    )

# ==================== TRA CỨU ====================
@app.route('/tim_kiem')
def tim_kiem():
    tu_khoa = request.args.get('q', '')
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    if tu_khoa:
        c.execute("SELECT id, ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc WHERE ten_tthc LIKE ? OR ma_tthc LIKE ?", (f'%{tu_khoa}%', f'%{tu_khoa}%'))
    else:
        c.execute("SELECT id, ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc LIMIT 20")
    ket_qua = c.fetchall()
    conn.close()
    return render_template('tim_kiem.html', ket_qua=ket_qua, tu_khoa=tu_khoa)

@app.route('/api/tim_kiem')
def api_tim_kiem():
    tu_khoa = request.args.get('q', '')
    conn = sqlite3.connect('tthc.db')
    c = conn.cursor()
    if tu_khoa:
        c.execute("SELECT id, ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc WHERE ten_tthc LIKE ? OR ma_tthc LIKE ?", (f'%{tu_khoa}%', f'%{tu_khoa}%'))
    else:
        c.execute("SELECT id, ma_tthc, ten_tthc, cap_thuc_hien, trang_thai_cong_khai, trang_thai FROM tthc LIMIT 20")
    data = c.fetchall()
    conn.close()
    
    ket_qua = []
    for row in data:
        ket_qua.append({
            'id': row[0],
            'ma_tthc': row[1],
            'ten_tthc': row[2],
            'cap_thuc_hien': row[3],
            'trang_thai_cong_khai': row[4],
            'trang_thai': row[5]
        })
    
    return jsonify({'data': ket_qua})
if __name__ == '__main__':
    app.run(debug=True)
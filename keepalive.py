# keepalive.py
import urllib.request
import time
import os
import sys
from datetime import datetime

# Cấu hình
CHECK_INTERVAL = 300  # 5 phút
RENDER_URL = os.environ.get('RENDER_EXTERNAL_URL', 'https://your-app.onrender.com')

def ping_url(url):
    """Ping URL để giữ ứng dụng hoạt động"""
    try:
        full_url = f"{url}/ping"
        req = urllib.request.Request(full_url, method='GET')
        with urllib.request.urlopen(req, timeout=30) as response:
            status = response.status
            print(f"[{datetime.now().isoformat()}] Ping thành công: {status} - {full_url}")
            return True
    except Exception as e:
        print(f"[{datetime.now().isoformat()}] Ping thất bại: {e}")
        return False

def main():
    print(f"[{datetime.now().isoformat()}] Khởi động KeepAlive service")
    print(f"Target URL: {RENDER_URL}")
    print(f"Check interval: {CHECK_INTERVAL} giây")
    
    while True:
        ping_url(RENDER_URL)
        time.sleep(CHECK_INTERVAL)

if __name__ == '__main__':
    # Chỉ chạy nếu có RENDER_URL
    if RENDER_URL and RENDER_URL != 'https://your-app.onrender.com':
        main()
    else:
        print("Vui lòng set biến môi trường RENDER_EXTERNAL_URL")
        print("Ví dụ: RENDER_EXTERNAL_URL=https://your-app.onrender.com")
        sys.exit(1)

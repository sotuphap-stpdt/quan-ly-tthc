#!/bin/bash

# start.sh - Script khởi động cho Render

echo "=========================================="
echo "Khởi động Hệ thống TTHC Sở Tư pháp Đồng Tháp"
echo "=========================================="

# Hiển thị thông tin
echo "Port: $PORT"
echo "Render URL: $RENDER_EXTERNAL_URL"
echo "Python version: $(python --version)"

# Khởi tạo database nếu cần
if [ ! -f "tthc.db" ]; then
    echo "Khởi tạo database mới..."
    python -c "from app import init_db; init_db()"
fi

# Chạy ứng dụng với Gunicorn (production)
echo "Khởi động Gunicorn server..."
exec gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 2 --timeout 120

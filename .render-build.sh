#!/bin/bash
# .render-build.sh

echo "Bắt đầu build trên Render..."

# Cài đặt dependencies
pip install --upgrade pip
pip install -r requirements.txt

# Tạo thư mục cần thiết
mkdir -p uploads
mkdir -p uploads/quyet_dinh
mkdir -p backups

# Khởi tạo database
python -c "from app import init_db; init_db()" || echo "Database đã tồn tại"

echo "Build hoàn tất!"

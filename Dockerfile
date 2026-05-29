# Dockerfile
FROM python:3.11-slim

WORKDIR /app

# Cài đặt dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code
COPY . .

# Tạo thư mục cần thiết
RUN mkdir -p uploads uploads/quyet_dinh backups

# Khởi tạo database
RUN python -c "from app import init_db; init_db()"

# Expose port
EXPOSE 5000

# Chạy ứng dụng
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:5000", "--workers", "2"]

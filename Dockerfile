FROM python:3.11-slim

WORKDIR /app

# Thiết lập biến môi trường
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080

# Cài đặt thư viện phụ thuộc
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Sao chép toàn bộ mã nguồn ứng dụng
COPY . .

# Mở cổng dịch vụ
EXPOSE 8080

# Khởi chạy server
CMD ["python", "server.py"]

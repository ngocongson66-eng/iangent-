# HƯỚNG DẪN TRIỂN KHAI (DEPLOY) WEB XỬ LÝ VÀ PHÂN TÍCH DỮ LIỆU

Hệ thống đã được cấu hình sẵn sàng để triển khai lên **Vercel**, **Render**, **Railway** hoặc bằng **Docker**.

---

## ⚡ CÁCH 1: TRIỂN KHAI LÊN VERCEL (ĐÃ FIX LỖI PYTHON_ENTRYPOINT_NOT_FOUND)

Lỗi của bạn xảy ra vì Vercel là nền tảng Serverless, nó không biết file nào trong các file `server.py`, `main.py` là điểm chạy. 
Dự án hiện đã được bổ sung đầy đủ:
- `api/index.py`: Serverless handler chuẩn cho Vercel.
- `vercel.json`: Cấu hình builder `@vercel/python` và định tuyến toàn bộ URL về `api/index.py`.
- `pyproject.toml`: Chỉ định `entrypoint = "api/index.py"`.

### Các bước deploy lại trên Vercel:
1. Mở terminal và đẩy các file mới cập nhật lên GitHub repo `github.com/ngocongson66-eng/ai301`:
   ```bash
   git add .
   git commit -m "Fix Vercel Python Entrypoint"
   git push origin main
   ```
2. Trên giao diện Vercel, bấm nút **Redeploy** (hoặc Vercel sẽ tự động build lại commit mới).
3. Đợi khoảng 1 phút, dự án sẽ hoàn tất (Ready) và có link truy cập `xxx.vercel.app`.

---

## CÁCH 2: TRIỂN KHAI LÊN RENDER.COM (KHUYÊN DÙNG CHO PYTHON WEB DÀI HẠN)

1. Truy cập **[Render.com](https://render.com/)**, đăng nhập bằng GitHub.
2. Chọn **New +** ➔ Chọn **Web Service**.
3. Chọn repo `ai301` bạn vừa đẩy lên.
4. Render sẽ tự động nhận diện:
   * **Runtime:** `Python 3`
   * **Build Command:** `pip install -r requirements.txt`
   * **Start Command:** `python server.py`
5. Bấm **Create Web Service**.

---

## CÁCH 3: TRIỂN KHAI LÊN RAILWAY.APP

1. Truy cập **[Railway.app](https://railway.app/)**, đăng nhập bằng GitHub.
2. Chọn **New Project** ➔ **Deploy from GitHub repo** ➔ chọn `ai301`.
3. Railway tự động nhận file `railway.json` và build ngay.

---

## CÁCH 4: TRIỂN KHAI BẰNG DOCKER (CHO MÁY CHỦ VPS)

```bash
docker build -t data-analytics-web .
docker run -d -p 8080:8080 --name data-web data-analytics-web
```

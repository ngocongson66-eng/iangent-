import os
import sys
import json
import urllib.parse
import traceback
from http.server import BaseHTTPRequestHandler

# Đảm bảo thư mục gốc dự án nằm trong sys.path để import server và src
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)
try:
    os.chdir(root_dir)
except Exception:
    pass

from server import RequestHandler

class handler(RequestHandler):
    """
    Entrypoint Serverless Function chuẩn mực dành riêng cho Vercel (@vercel/python).
    Kế thừa RequestHandler của server.py và hỗ trợ rewrite URL thông minh 100% chuẩn xác.
    """
    def get_real_path(self):
        # 1. Ưu tiên kiểm tra query param __vercel_path từ rewrite của vercel.json
        parsed = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(parsed.query)
        if "__vercel_path" in qs and qs["__vercel_path"]:
            real = qs["__vercel_path"][0]
            other_params = {k: v for k, v in qs.items() if k != "__vercel_path"}
            if other_params:
                new_q = urllib.parse.urlencode(other_params, doseq=True)
                return f"{real}?{new_q}"
            return real

        # 2. Kiểm tra các header của Vercel / Reverse Proxy
        for header_name in ["x-forwarded-uri", "x-matched-path", "x-vercel-matched-path", "x-now-route-matches"]:
            h_val = self.headers.get(header_name)
            if h_val and h_val not in ["/api/index.py", "/api/index", "/api"]:
                return h_val

        # 3. Fallback
        if parsed.path in ["/api/index.py", "/api/index", "/api"]:
            return "/"
        return self.path

    def do_GET(self):
        try:
            self.path = self.get_real_path()
            return super().do_GET()
        except Exception as e:
            traceback.print_exc()
            self.send_response(500)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write(f"500 Internal Server Error:\n{traceback.format_exc()}".encode("utf-8"))

    def do_POST(self):
        try:
            self.path = self.get_real_path()
            return super().do_POST()
        except Exception as e:
            traceback.print_exc()
            self.send_response(500)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            err_json = json.dumps({"error": str(e), "traceback": traceback.format_exc()})
            self.wfile.write(err_json.encode("utf-8"))

    def do_OPTIONS(self):
        return super().do_OPTIONS()

    def do_HEAD(self):
        try:
            self.path = self.get_real_path()
            return super().do_HEAD()
        except Exception:
            return super().do_HEAD()

# Export alias tương thích cả hai kiểu nhận diện của Vercel
app = handler

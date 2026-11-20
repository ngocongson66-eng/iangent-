import io
import sys
import os
import urllib.parse
from http.server import BaseHTTPRequestHandler

# Test get_real_path logic
def get_real_path(path, headers={}):
    parsed = urllib.parse.urlparse(path)
    qs = urllib.parse.parse_qs(parsed.query)
    if "__vercel_path" in qs and qs["__vercel_path"]:
        real = qs["__vercel_path"][0]
        other_params = {k: v for k, v in qs.items() if k != "__vercel_path"}
        if other_params:
            new_q = urllib.parse.urlencode(other_params, doseq=True)
            return f"{real}?{new_q}"
        return real

    for header_name in ["x-forwarded-uri", "x-matched-path", "x-vercel-matched-path", "x-now-route-matches"]:
        h_val = headers.get(header_name)
        if h_val and h_val not in ["/api/index.py", "/api/index", "/api"]:
            return h_val

    if parsed.path in ["/api/index.py", "/api/index", "/api"]:
        return "/"
    return path

# Test cases
assert get_real_path("/api/index?__vercel_path=/api/batch_process") == "/api/batch_process"
assert get_real_path("/api/index?__vercel_path=/api/upload_file&filename=test.xlsx") == "/api/upload_file?filename=test.xlsx"
assert get_real_path("/api/index?__vercel_path=/") == "/"
assert get_real_path("/api/index") == "/"
assert get_real_path("/") == "/"
assert get_real_path("/api/load_hung_phat_sample", {"x-forwarded-uri": "/api/load_hung_phat_sample"}) == "/api/load_hung_phat_sample"

print("All get_real_path assertions passed successfully!")

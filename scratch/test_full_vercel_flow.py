import os
import sys
import io
import json

# Test updating api/index.py and vercel.json
sys.path.insert(0, os.path.abspath('.'))

from api.index import handler

class MockSocket:
    def __init__(self, data=b""):
        self.rfile = io.BytesIO(data)
        self.wfile = io.BytesIO()
    def makefile(self, *args, **kwargs):
        if 'r' in args[0] or 'rb' in args[0]:
            return self.rfile
        return self.wfile
    def sendall(self, b):
        self.wfile.write(b)
    def close(self):
        pass

# Test 1: GET /api/index?__vercel_path=/
mock_req1 = b"GET /api/index?__vercel_path=/ HTTP/1.1\r\nHost: myapp.vercel.app\r\n\r\n"
sock1 = MockSocket(mock_req1)
h1 = handler(sock1, ('127.0.0.1', 8080), None)
res1 = sock1.wfile.getvalue()
assert b"200 OK" in res1
print("Test 1 (GET / via __vercel_path): PASS")

# Test 2: GET /api/index?__vercel_path=/api/load_hung_phat_sample
mock_req2 = b"GET /api/index?__vercel_path=/api/load_hung_phat_sample HTTP/1.1\r\nHost: myapp.vercel.app\r\n\r\n"
sock2 = MockSocket(mock_req2)
h2 = handler(sock2, ('127.0.0.1', 8080), None)
res2 = sock2.wfile.getvalue()
assert b"200 OK" in res2
assert b"application/json" in res2
print("Test 2 (GET /api/load_hung_phat_sample): PASS")

# Test 3: POST /api/index?__vercel_path=/api/batch_process
post_body = b'[{"opportunity_id":"TEST-01","raw_request_text":"Ban lam viec"}]'
mock_req3 = f"POST /api/index?__vercel_path=/api/batch_process HTTP/1.1\r\nHost: myapp.vercel.app\r\nContent-Length: {len(post_body)}\r\nContent-Type: application/json\r\n\r\n".encode('utf-8') + post_body
sock3 = MockSocket(mock_req3)
h3 = handler(sock3, ('127.0.0.1', 8080), None)
res3 = sock3.wfile.getvalue()
assert b"200 OK" in res3
assert b"application/json" in res3
assert b"dossiers" in res3
print("Test 3 (POST /api/batch_process via __vercel_path): PASS")

print("ALL TESTS PASSED SUCCESSFULLY!")

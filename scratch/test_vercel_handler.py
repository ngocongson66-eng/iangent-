import io
import sys
import os

# Test importing api.index
sys.path.insert(0, os.path.abspath('.'))
from api.index import handler

print("Successfully imported handler from api.index!")

# Test handler initialization and do_GET
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

mock_req = b"GET / HTTP/1.1\r\nHost: localhost\r\n\r\n"
sock = MockSocket(mock_req)
try:
    h = handler(sock, ('127.0.0.1', 8080), None)
    output = sock.wfile.getvalue()
    print("do_GET / output length:", len(output))
    print("Status in output:", output[:100])
except Exception as e:
    import traceback
    traceback.print_exc()

# Test POST /api/batch_process
post_body = b'{"items": [{"id": 1, "raw_text": "C\\u1ea7n g\\u1ea5p 5 B\\u00e0n l\\u00e0m vi\\u1ec7c g\\u1ed7 s\\u1ed3i giao H\\u00e0 N\\u1ed9i deadline ng\\u00e0y mai ng\\u00e2n s\\u00e1ch 10tr"}]}'
mock_post = f"POST /api/batch_process HTTP/1.1\r\nHost: localhost\r\nContent-Length: {len(post_body)}\r\nContent-Type: application/json\r\nx-forwarded-uri: /api/batch_process\r\n\r\n".encode('utf-8') + post_body
sock_post = MockSocket(mock_post)
try:
    h_post = handler(sock_post, ('127.0.0.1', 8080), None)
    out_post = sock_post.wfile.getvalue()
    print("\nPOST /api/batch_process output length:", len(out_post))
    print("POST output preview:", out_post[:300].decode('utf-8', errors='ignore'))
except Exception as e:
    import traceback
    traceback.print_exc()

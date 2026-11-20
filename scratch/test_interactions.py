import subprocess
import time
import json
import urllib.request
import socket
import base64
import os
import struct

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
user_dir = r"C:\Users\Son\.gemini\antigravity\brain\beab687d-e87c-4454-8cb8-a2b689101e7b\scratch\cdp_user2"
os.makedirs(user_dir, exist_ok=True)

port = 9333
proc = subprocess.Popen([
    chrome_path,
    "--headless=new",
    "--disable-gpu",
    "--no-sandbox",
    f"--remote-debugging-port={port}",
    f"--user-data-dir={user_dir}",
    "about:blank"
])

time.sleep(2)

try:
    res = urllib.request.urlopen(f"http://127.0.0.1:{port}/json", timeout=5)
    pages = json.loads(res.read().decode())
    page_info = [p for p in pages if p.get("type") == "page"][0]
    ws_url = page_info["webSocketDebuggerUrl"]

    parts = ws_url.replace("ws://", "").split("/", 1)
    host_port = parts[0].split(":")
    host = host_port[0]
    p = int(host_port[1])
    path = "/" + parts[1]

    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.connect((host, p))
    key = base64.b64encode(os.urandom(16)).decode()
    handshake = (
        f"GET {path} HTTP/1.1\r\n"
        f"Host: {host}:{p}\r\n"
        f"Upgrade: websocket\r\n"
        f"Connection: Upgrade\r\n"
        f"Sec-WebSocket-Key: {key}\r\n"
        f"Sec-WebSocket-Version: 13\r\n\r\n"
    )
    s.sendall(handshake.encode())
    resp = s.recv(4096).decode(errors='ignore')
    if "101" not in resp:
        raise RuntimeError("Handshake failed")

    msg_id = 0
    def send_cdp(method, params=None):
        global msg_id
        msg_id += 1
        payload = {"id": msg_id, "method": method, "params": params or {}}
        raw = json.dumps(payload).encode()
        frame = bytearray([0x81, 0x80 | len(raw)])
        mask = os.urandom(4)
        frame.extend(mask)
        masked = bytearray(b ^ mask[i % 4] for i, b in enumerate(raw))
        frame.extend(masked)
        s.sendall(frame)
        return msg_id

    def read_frames(duration=2):
        s.settimeout(0.5)
        end_time = time.time() + duration
        msgs = []
        buf = bytearray()
        while time.time() < end_time:
            try:
                data = s.recv(8192)
                if not data:
                    break
                buf.extend(data)
                while len(buf) >= 2:
                    payload_len = buf[1] & 0x7F
                    offset = 2
                    if payload_len == 126:
                        if len(buf) < 4: break
                        payload_len = struct.unpack("!H", buf[2:4])[0]
                        offset = 4
                    elif payload_len == 127:
                        if len(buf) < 10: break
                        payload_len = struct.unpack("!Q", buf[2:10])[0]
                        offset = 10
                    if len(buf) < offset + payload_len:
                        break
                    frame_data = buf[offset:offset+payload_len]
                    buf = buf[offset+payload_len:]
                    try:
                        msgs.append(json.loads(frame_data.decode('utf-8')))
                    except Exception:
                        pass
            except socket.timeout:
                continue
        return msgs

    send_cdp("Runtime.enable")
    send_cdp("Page.navigate", {"url": "http://localhost:8080"})
    time.sleep(2)
    read_frames(1)

    def eval_js(expr):
        cid = send_cdp("Runtime.evaluate", {"expression": expr, "returnByValue": True})
        frames = read_frames(1)
        for f in frames:
            if f.get("id") == cid:
                res = f.get("result", {}).get("result", {})
                return res.get("value")
        return None

    # Test 1: switchTask(2)
    print("Test 1 - Switch to Task 2:", eval_js("switchTask(2); document.getElementById('panel-task-2').classList.contains('active')"))
    # Test 2: switchTask(3)
    print("Test 2 - Switch to Task 3:", eval_js("switchTask(3); document.getElementById('panel-task-3').classList.contains('active')"))
    # Test 3: switchTask(1)
    print("Test 3 - Switch back to Task 1:", eval_js("switchTask(1); document.getElementById('panel-task-1').classList.contains('active')"))
    # Test 4: loadSampleBusinessQuestions
    print("Test 4 - Load Questions:", eval_js("loadSampleBusinessQuestions(); document.getElementById('ctx-questions').value.length > 10"))
    # Test 5: Confirm Business Context
    print("Test 5 - Confirm Context:", eval_js("confirmBusinessContext(); document.getElementById('ctx-status-badge').innerText"))
    # Test 6: Click Retail Sample (loadBo1Sample)
    print("Test 6 - Load Sample Data:", eval_js("loadBo1Sample(); rawItems ? rawItems.length : 0"))
    time.sleep(1)
    print("Test 6b - Raw items count after load:", eval_js("rawItems ? rawItems.length : 0"))
    print("Test 7 - Task 2 clean rows count:", eval_js("document.querySelectorAll('#clean-tbody tr').length"))

except Exception as err:
    print("Script Error:", err)
finally:
    try: s.close()
    except: pass
    proc.terminate()

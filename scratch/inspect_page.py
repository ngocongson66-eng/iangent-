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

# 1. Start Chrome
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
    # 2. Get target page ws url
    res = urllib.request.urlopen(f"http://127.0.0.1:{port}/json", timeout=5)
    pages = json.loads(res.read().decode())
    page_info = [p for p in pages if p.get("type") == "page"][0]
    ws_url = page_info["webSocketDebuggerUrl"]
    print("Page target created:", page_info["id"])

    # Parse ws url: ws://127.0.0.1:9333/devtools/page/...
    parts = ws_url.replace("ws://", "").split("/", 1)
    host_port = parts[0].split(":")
    host = host_port[0]
    p = int(host_port[1])
    path = "/" + parts[1]

    # Simple RFC 6455 WebSocket Handshake
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
        print("WS Handshake failed:", resp)
        raise RuntimeError("Handshake failed")
    print("WS Connected!")

    def send_cdp(method, params=None, msg_id=1):
        payload = {"id": msg_id, "method": method, "params": params or {}}
        raw = json.dumps(payload).encode()
        frame = bytearray([0x81, 0x80 | len(raw)])
        mask = os.urandom(4)
        frame.extend(mask)
        masked = bytearray(b ^ mask[i % 4] for i, b in enumerate(raw))
        frame.extend(masked)
        s.sendall(frame)

    def read_frames(duration=4):
        s.settimeout(1.0)
        end_time = time.time() + duration
        msgs = []
        buf = bytearray()
        while time.time() < end_time:
            try:
                data = s.recv(4096)
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

    # Enable domains
    send_cdp("Runtime.enable", msg_id=1)
    send_cdp("Log.enable", msg_id=2)
    send_cdp("Page.enable", msg_id=3)
    read_frames(0.5)

    # Navigate to localhost:8080
    print("Navigating to http://localhost:8080 ...")
    send_cdp("Page.navigate", {"url": "http://localhost:8080"}, msg_id=4)
    
    events = read_frames(3)
    print(f"\n--- CAPTURED {len(events)} CDP EVENTS ---")
    for ev in events:
        method = ev.get("method", "")
        if "exception" in method.lower() or "error" in method.lower() or "console" in method.lower():
            print("\n*** EXCEPTION / CONSOLE ***", json.dumps(ev, indent=2, ensure_ascii=False))
        elif method == "Runtime.exceptionThrown":
            print("\n🚨 RUNTIME EXCEPTION:", json.dumps(ev["params"]["exceptionDetails"], indent=2, ensure_ascii=False))
        elif method == "Runtime.consoleAPICalled":
            args = [a.get("value", a.get("description", "")) for a in ev["params"].get("args", [])]
            print(f"[{ev['params'].get('type', 'log').upper()}]", *args)

except Exception as err:
    print("Script Error:", err)
finally:
    try: s.close()
    except: pass
    proc.terminate()

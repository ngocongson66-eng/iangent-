import subprocess
import time
import json
import urllib.request
import socket
import base64
import os
import struct

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
user_dir = r"C:\Users\Son\.gemini\antigravity\brain\beab687d-e87c-4454-8cb8-a2b689101e7b\scratch\cdp_user_interact"
os.makedirs(user_dir, exist_ok=True)

port = 9334
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
    assert "101" in resp

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
                data = s.recv(4096)
                if not data: break
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
                    if len(buf) < offset + payload_len: break
                    frame_data = buf[offset:offset+payload_len]
                    buf = buf[offset+payload_len:]
                    try:
                        msgs.append(json.loads(frame_data.decode('utf-8')))
                    except Exception: pass
            except socket.timeout:
                continue
        return msgs

    send_cdp("Runtime.enable")
    send_cdp("Log.enable")
    send_cdp("Page.enable")
    read_frames(0.5)

    print("Navigating to http://localhost:8080 ...")
    send_cdp("Page.navigate", {"url": "http://localhost:8080"})
    events = read_frames(3)
    for ev in events:
        if ev.get("method") == "Runtime.exceptionThrown":
            print("🚨 EXCEPTION:", ev["params"]["exceptionDetails"])
        elif ev.get("method") == "Runtime.consoleAPICalled":
            args = [a.get("value", a.get("description", "")) for a in ev["params"].get("args", [])]
            print(f"[{ev['params'].get('type', 'log').upper()}]", *args)

    def eval_js(expr):
        req_id = send_cdp("Runtime.evaluate", {"expression": expr, "returnByValue": True, "awaitPromise": True})
        responses = read_frames(2)
        for r in responses:
            if r.get("id") == req_id:
                return r.get("result", {}).get("result", {})
        return None

    print("\n--- TEST 1: Check typeof switchTask ---")
    res = eval_js("typeof switchTask")
    print("typeof switchTask:", res.get("value"))

    print("\n--- TEST 2: Check active task before & after switchTask(2) ---")
    eval_js("switchTask(2)")
    active_panel = eval_js("document.querySelector('.task-panel.active') ? document.querySelector('.task-panel.active').id : 'none'")
    print("Active panel after switchTask(2):", active_panel.get("value"))

    print("\n--- TEST 3: Click 'loadSampleBusinessQuestions()' ---")
    eval_js("loadSampleBusinessQuestions()")
    q_val = eval_js("document.getElementById('ctx-questions') ? document.getElementById('ctx-questions').value : ''")
    print("Questions textarea length:", len(str(q_val.get("value", ""))))

    print("\n--- TEST 4: Click 'applyContextAndAnalyzeHungPhat()' ---")
    eval_js("applyContextAndAnalyzeHungPhat()")
    time.sleep(1.5)
    events2 = read_frames(2)
    for ev in events2:
        if ev.get("method") == "Runtime.exceptionThrown":
            print("🚨 EXCEPTION in HungPhat:", ev["params"]["exceptionDetails"])
        elif ev.get("method") == "Runtime.consoleAPICalled":
            args = [a.get("value", a.get("description", "")) for a in ev["params"].get("args", [])]
            print(f"[{ev['params'].get('type', 'log').upper()}]", *args)

    items_len = eval_js("rawItems.length")
    dossiers_len = eval_js("allDossiers.length")
    print("rawItems.length:", items_len.get("value"))
    print("allDossiers.length:", dossiers_len.get("value"))

finally:
    try: s.close()
    except: pass
    proc.terminate()

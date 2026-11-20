import subprocess
import time
import json
import urllib.request
import asyncio
import sys

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
user_dir = r"C:\Users\Son\.gemini\antigravity\brain\beab687d-e87c-4454-8cb8-a2b689101e7b\scratch\cdp_user"

# Start chrome with remote debugging
proc = subprocess.Popen([
    chrome_path,
    "--headless=new",
    "--disable-gpu",
    "--no-sandbox",
    f"--remote-debugging-port=9222",
    f"--user-data-dir={user_dir}",
    "about:blank"
])

print("Started chrome, waiting for CDP...")
time.sleep(2)

try:
    resp = urllib.request.urlopen("http://127.0.0.1:9222/json", timeout=5)
    pages = json.loads(resp.read().decode())
    print("CDP pages:", pages)
    ws_url = pages[0]["webSocketDebuggerUrl"]
    print("WebSocket URL:", ws_url)
except Exception as e:
    print("Failed to connect to CDP HTTP:", e)
finally:
    proc.terminate()

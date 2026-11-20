import subprocess
import os
import shutil

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
url = r"file:///C:/New folder (2)/scratch/test_inline.html"
ud = r"C:\Users\Son\.gemini\antigravity\brain\beab687d-e87c-4454-8cb8-a2b689101e7b\scratch\chrome_test_profile2"
os.makedirs(ud, exist_ok=True)

try:
    res = subprocess.run([
        chrome_path,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        f"--user-data-dir={ud}",
        "--dump-dom",
        url
    ], capture_output=True, text=True, timeout=12)

    for line in res.stdout.splitlines():
        if "test-output" in line or "EXACT_ERROR" in line or "SUCCESS" in line:
            print(line)
except Exception as e:
    print("Error:", e)
finally:
    shutil.rmtree(ud, ignore_errors=True)

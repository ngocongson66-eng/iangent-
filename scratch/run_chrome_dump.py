import subprocess

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
url = r"file:///C:/New folder (2)/scratch/test_syntax.html"

res = subprocess.run([
    chrome_path,
    "--headless",
    "--disable-gpu",
    "--no-sandbox",
    "--dump-dom",
    url
], capture_output=True, text=True, timeout=10)

print("Return code:", res.returncode)
print("STDOUT length:", len(res.stdout))
print("STDERR length:", len(res.stderr))
if "SUCCESS" in res.stdout:
    print("MATCH: SUCCESS found in stdout!")
if "ERROR" in res.stdout:
    print("MATCH: ERROR found in stdout!")

for line in res.stdout.splitlines():
    if "test-output" in line or "ERRORS" in line or "SUCCESS" in line:
        print("DOM line:", line)

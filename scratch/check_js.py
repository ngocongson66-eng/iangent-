import re

with open('server.py', 'r', encoding='utf-8') as f:
    content = f.read()

scripts = re.findall(r'<script>(.*?)</script>', content, re.DOTALL)
print(f"Total <script> blocks found: {len(scripts)}")
js_code = scripts[0]
with open('scratch/app.js', 'w', encoding='utf-8') as f:
    f.write(js_code)

print(f"app.js written, size: {len(js_code)} characters, {len(js_code.splitlines())} lines")

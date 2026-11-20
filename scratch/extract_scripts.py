import re

with open('server.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Extract HTML_TEMPLATE
m = re.search(r'HTML_TEMPLATE\s*=\s*"""(.*?)"""', code, re.DOTALL)
if not m:
    print("Could not find HTML_TEMPLATE")
    exit(1)

html = m.group(1)
scripts = re.findall(r'<script>(.*?)</script>', html, re.DOTALL)
print(f"Found {len(scripts)} scripts in HTML_TEMPLATE.")
for i, s in enumerate(scripts):
    fname = f"scratch/script_{i}.js"
    with open(fname, "w", encoding="utf-8") as f:
        f.write(s)
    print(f"Saved {fname} ({len(s)} bytes)")

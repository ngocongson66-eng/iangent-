import re

with open('server.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Find all regex literals in server.py
# Regex literals look like /.../flags
regexes = re.findall(r'/[^/\n\r]+/[gimsuy]*', code)
print(f"Found {len(regexes)} regex-like patterns.")

# Check each one in node/python
suspicious = []
for r in regexes:
    if '\\(' in r or '\\)' in r or '\\\\' in r:
        suspicious.append(r)

print(f"Suspicious regex patterns ({len(suspicious)}):")
for s in suspicious:
    print("  ", s)

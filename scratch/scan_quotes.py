import re

with open('server.py', 'r', encoding='utf-8') as f:
    html = f.read()

# Extract script
script = re.search(r'<script>(.*?)</script>', html, re.DOTALL).group(1)

# Check for unescaped newlines in single- or double-quoted strings
lines = script.splitlines()
in_backtick = False
issues = []

for idx, line in enumerate(lines, 1):
    # count backticks
    bt_count = 0
    escaped = False
    for ch in line:
        if ch == '\\':
            escaped = not escaped
            continue
        if ch == '`' and not escaped:
            bt_count += 1
        escaped = False
    if bt_count % 2 != 0:
        in_backtick = not in_backtick
    
    if not in_backtick:
        # Check single quotes and double quotes parity
        sq = 0
        dq = 0
        esc = False
        for ch in line:
            if ch == '\\':
                esc = not esc
                continue
            if ch == "'" and not esc:
                sq += 1
            elif ch == '"' and not esc:
                dq += 1
            esc = False
        if sq % 2 != 0 or dq % 2 != 0:
            issues.append((idx, line.strip(), sq % 2, dq % 2))

print(f"Total potential quote parity issues outside backticks: {len(issues)}")
for line_no, text, sq_odd, dq_odd in issues[:10]:
    print(f"Line {line_no}: sq_odd={sq_odd}, dq_odd={dq_odd} -> {text[:100]}")

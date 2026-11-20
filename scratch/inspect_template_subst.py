with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'HTML_PAGE = HTML_TEMPLATE' in l:
        for j in range(i, min(len(lines), i+10)):
            print(f"{j+1}: {lines[j].strip()[:100]}")
        break

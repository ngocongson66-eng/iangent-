with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'function renderTask4Data' in l or 'async function renderTask4Data' in l:
        for j in range(i, min(i+60, len(lines))):
            print(f'{j+1}: {lines[j].rstrip()[:110]}')
        break

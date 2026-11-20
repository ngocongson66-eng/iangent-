with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'load_hung_phat_sample' in l and ('elif' in l or 'if' in l) and 'parsed.path' in l:
        for j in range(i, min(i+50, len(lines))):
            print(f'{j+1}: {lines[j].rstrip()[:110]}')
        break

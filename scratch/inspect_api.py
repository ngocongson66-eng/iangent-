with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'quickLoadHungPhat' in l or 'Nạp 6 Ca Đề Bài' in l or 'fetch(' in l:
        if any(k in l for k in ['function', 'fetch', '/api/']):
            print(f'{i+1}: {l.strip()[:110]}')

with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'Task 2: Làm Sạch Dữ Liệu' in l:
        print(f'{i+1}: {l.strip()[:100]}')

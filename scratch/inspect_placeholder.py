with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'Bấm Duyệt luồng để bắt đầu' in l or 'Chưa có kết quả xử lý' in l:
        print(f'{i+1}: {l.strip()[:110]}')

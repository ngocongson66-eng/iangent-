with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'Phê Duyệt Luồng & Mở Báo Cáo Ngay' in l:
        for j in range(max(0, i-5), min(i+15, len(lines))):
            print(f'{j+1}: {lines[j].rstrip()[:110]}')

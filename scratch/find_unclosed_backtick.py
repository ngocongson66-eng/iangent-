with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'Tổng chi phí chiết khấu đã giải ngân trong tập dữ liệu là' in l:
        for j in range(max(0, i-5), min(len(lines), i+25)):
            print(f"Line {j+1}: {repr(lines[j])}")
        break

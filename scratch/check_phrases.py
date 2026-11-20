import os

phrases = [
    'ĐÃ HOÀN TẤT LÀM SẠCH',
    'Xuất File Excel Đã Làm Sạch',
    'Phê Duyệt & Sang Task 3',
    'Tải Báo Cáo Word Dạng HTML',
    'step-transition-banner'
]

found = []
for root, dirs, files in os.walk('.'):
    if any(x in root for x in ['.git', 'scratch', '__pycache__']):
        continue
    for file in files:
        if file.endswith(('.py', '.html', '.js', '.json')):
            path = os.path.join(root, file)
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                for phrase in phrases:
                    if phrase.lower() in content.lower():
                        found.append((phrase, path))

for p, path in found:
    print(f'Found "{p}" in {path}')

if not found:
    print('NONE FOUND in codebase!')

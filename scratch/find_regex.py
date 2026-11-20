with open('scratch/script_0.js', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'Khảo sát' in l:
        print(f'{i+1}: {l.strip()[:100]}')

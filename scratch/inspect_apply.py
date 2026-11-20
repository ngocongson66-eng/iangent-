with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'function applyContextAndAnalyzeHungPhat' in l or 'async function applyContextAndAnalyzeHungPhat' in l or 'function quickLoadHungPhat' in l:
        for j in range(i, min(i+50, len(lines))):
            print(f'{j+1}: {lines[j].rstrip()[:110]}')
        print("="*60)

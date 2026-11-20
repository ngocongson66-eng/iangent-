with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'async function approveAndRunPipeline(' in l or 'function approveAndRunPipeline(' in l:
        for j in range(i, min(i+70, len(lines))):
            print(f'{j+1}: {lines[j].rstrip()[:110]}')
        break

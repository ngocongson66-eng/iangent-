with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'approveAndRunPipelineFast' in l and '<button' in l:
        print(f'{i+1}: {l.strip()[:110]}')

with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if any(k in l for k in ['open(', 'os.path.exists(', 'pd.ExcelFile(', 'pd.read_']):
        if 'get_rel_path' not in l and not l.strip().startswith('#'):
            print(f'{i+1}: {l.strip()[:110]}')

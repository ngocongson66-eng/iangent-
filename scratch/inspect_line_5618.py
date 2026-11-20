with open('scratch/script_0.js', 'r', encoding='utf-8') as f:
    lines = f.readlines()

line_num = 5618
for i in range(max(0, line_num - 15), min(len(lines), line_num + 15)):
    print(f"{i+1}: {lines[i].rstrip()[:110]}")

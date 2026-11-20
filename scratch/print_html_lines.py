with open('scratch/test_inline.html', 'r', encoding='utf-8') as f:
    lines = f.readlines()

line_num = 5618
for i in range(max(0, line_num - 10), min(len(lines), line_num + 10)):
    print(f"HTML Line {i+1}: {repr(lines[i])}")

with open('scratch/script_0.js', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(5610, 5625):
    print(f"Line {i+1}: {repr(lines[i])}")

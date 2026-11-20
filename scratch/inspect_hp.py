with open('server.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, l in enumerate(lines):
    if 'Hung_Phat_Test_Data_5_Steps.xlsx' in l and ('open(' in l or 'exists(' in l or 'ExcelFile(' in l):
        print(f'{i+1}: {l.strip()[:110]}')

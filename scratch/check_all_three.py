import re

with open('server.py', 'r', encoding='utf-8') as f:
    content = f.read()

print("--- 1. BANNERS ---")
matches = re.findall(r'.*step-transition-banner.*', content)
print(f"Occurrences of step-transition-banner: {len(matches)}")
for m in matches:
    print("  ", m.strip()[:100])

matches_clean_banner = re.findall(r'.*ĐÃ HOÀN TẤT LÀM SẠCH & THẨM ĐỊNH KHO.*', content)
print(f"Occurrences of 'ĐÃ HOÀN TẤT LÀM SẠCH & THẨM ĐỊNH KHO': {len(matches_clean_banner)}")

print("\n--- 2. TASK 1 BỐI CẢNH NGHIỆP VỤ ---")
has_scope = 'ctx-scope-info' in content
has_btn_confirm = 'btn-confirm-context' in content
print(f"Has ctx-scope-info (Phạm vi & Thông tin bổ sung): {has_scope}")
print(f"Has btn-confirm-context (Nút Thiết Lập): {has_btn_confirm}")

print("\n--- 3. TASK 4 REPORT BUTTONS ---")
matches_buttons = re.findall(r'.*(?:Xuất File Word|Tải Báo Cáo|Báo Cáo Word Dạng HTML|Báo Cáo Dữ Liệu).*', content)
for m in matches_buttons:
    if any(k in m for k in ['<button', '<span>', 'class="btn']):
        print("  ", m.strip()[:100])

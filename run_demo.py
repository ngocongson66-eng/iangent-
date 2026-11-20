import time
import json
from src.agent import MinhAnSalesAgent
from src.tools import DocumentGenerationTool

def print_separator(char="=", length=85):
    print(char * length)

def run_interactive_demo():
    print_separator("=")
    print("      DEMO HỆ THỐNG AI AGENT XỬ LÝ CƠ HỘI BÁN HÀNG B2B - CÔNG TY MINH AN")
    print("      Kịch bản kiểm thử: Thông thường - Thiếu dữ liệu - Ranh giới quy tắc")
    print_separator("=")

    agent = MinhAnSalesAgent(
        catalog_path="data/catalog.json",
        inventory_path="data/inventory.json",
        policies_path="data/policies.json"
    )

    with open("data/sample_requests.json", "r", encoding="utf-8") as f:
        samples = json.load(f)

    print(f"\n[+] Khởi tạo thành công AI Agent cùng 5 kỹ năng chuyên trách:")
    print("    1. InformationExtractionTool     (Phân tích NLP, trích xuất thực thể, kiểm tra đầy đủ)")
    print("    2. InventoryLeadTimeTool          (Thẩm định tồn kho khả dụng, tính toán Lead time)")
    print("    3. PricingDiscountTool            (Tính bảng giá đề xuất, áp dụng chính sách chiết khấu)")
    print("    4. ApprovalRoutingTool            (Kiểm tra thẩm quyền phê duyệt, phát hiện vượt quyền)")
    print("    5. DocumentGenerationTool         (Tạo Hồ sơ cơ hội, Báo giá, Email và Phiếu duyệt)")

    print(f"\n[+] Đang tiến hành kiểm thử thực tế trên {len(samples)} yêu cầu khách hàng đa dạng...")
    time.sleep(0.5)

    dossiers = []
    for idx, req in enumerate(samples, 1):
        print("\n" + "-" * 85)
        print(f"▶ CA KIỂM THỬ {idx:02d}: [{req['id']}] - {req['customer_name']}")
        print(f"  Kịch bản: {req['scenario_type']}")
        print(f"  Văn bản tiếp nhận: \"{req['raw_request_text'][:95]}...\"")
        
        d = agent.process_request(req)
        dossiers.append(d)

        # In vết suy luận của Agent
        print(f"  [*] Trạng thái phân loại: >> {d.status} <<")
        print(f"  [*] Chuyên viên phụ trách: {d.assigned_sales_rep}")
        if d.status == "CẦN BỔ SUNG THÔNG TIN":
            print(f"  [!] Thông tin thiếu phát hiện: {', '.join(d.extracted_data.missing_fields)}")
        elif d.status == "CHƯA ĐÁP ỨNG":
            print(f"  [!] Lý do từ chối/hạn chế: {d.feasibility.rejection_reason}")
        elif d.status == "CHỜ DUYỆT":
            print(f"  [!] Cần chuyển duyệt: {d.approval.approver_role}")
            print(f"  [!] Căn cứ: {'; '.join(d.approval.reasons)}")
            print(f"  [!] Giá trị đề xuất: {d.pricing.final_proposal_value:,.0f} VNĐ (Chiết khấu: {d.pricing.applied_discount_rate*100:.1f}%)")
        else: # CÓ THỂ TIẾP TỤC
            print(f"  [✓] Đầy đủ thông tin, tồn kho đáp ứng, trong quyền NVKD.")
            print(f"  [✓] Giá trị đề xuất: {d.pricing.final_proposal_value:,.0f} VNĐ")

        print(f"  [*] Hành động tiếp theo: {d.next_action}")

    # Bảng tổng hợp
    status_counts = {"CÓ THỂ TIẾP TỤC": 0, "CẦN BỔ SUNG THÔNG TIN": 0, "CHƯA ĐÁP ỨNG": 0, "CHỜ DUYỆT": 0}
    total_val = 0.0
    for d in dossiers:
        status_counts[d.status] += 1
        if d.status in ["CÓ THỂ TIẾP TỤC", "CHỜ DUYỆT"] and d.pricing:
            total_val += d.pricing.final_proposal_value

    print("\n" + "=" * 85)
    print("                    BẢNG TỔNG HỢP THEO TIÊU CHUẨN ĐỀ BÀI")
    print("=" * 85)
    print(f"1. Số cơ hội CÓ THỂ TIẾP TỤC:       {status_counts['CÓ THỂ TIẾP TỤC']:>2} cơ hội")
    print(f"2. Số cơ hội CẦN BỔ SUNG THÔNG TIN: {status_counts['CẦN BỔ SUNG THÔNG TIN']:>2} cơ hội")
    print(f"3. Số cơ hội CHƯA ĐÁP ỨNG:          {status_counts['CHƯA ĐÁP ỨNG']:>2} cơ hội")
    print(f"4. Số cơ hội CHỜ DUYỆT:             {status_counts['CHỜ DUYỆT']:>2} cơ hội")
    print("-" * 55)
    print(f"   TỔNG SỐ YÊU CẦU ĐÃ XỬ LÝ:        {len(dossiers):>2} cơ hội")
    print(f"\n=> TỔNG GIÁ TRỊ CÁC CƠ HỘI ĐỦ ĐIỀU KIỆN TIẾP TỤC HOẶC CHỜ DUYỆT:")
    print(f"   ★ {total_val:,.0f} VNĐ ★")
    print("=" * 85)

if __name__ == "__main__":
    run_interactive_demo()

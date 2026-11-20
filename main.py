import os
import sys
import json
import argparse
from src.agent import MinhAnSalesAgent
from src.tools import DocumentGenerationTool

def main():
    parser = argparse.ArgumentParser(description="Hệ thống AI Agent Xử lý Cơ hội Bán hàng B2B - Công ty Minh An")
    parser.add_argument("--catalog", default="data/catalog.json", help="Đường dẫn file danh mục sản phẩm và bảng giá")
    parser.add_argument("--inventory", default="data/inventory.json", help="Đường dẫn file dữ liệu tồn kho khả dụng")
    parser.add_argument("--policies", default="data/policies.json", help="Đường dẫn file chính sách chiết khấu và phân quyền")
    parser.add_argument("--data", default="data/sample_requests.json", help="Đường dẫn file dữ liệu yêu cầu khách hàng")
    parser.add_argument("--output-dir", default="outputs", help="Thư mục xuất kết quả")
    args = parser.parse_args()

    print("=" * 80)
    print("     HỆ THỐNG AI AGENT XỬ LÝ CƠ HỘI BÁN HÀNG B2B - CÔNG TY MINH AN")
    print("     Học phần: AI301 | Đồ án thực thi quy trình tự động hóa")
    print("=" * 80)

    # 1. Khởi tạo Agent
    print(f"\n[1] Đang nạp cấu hình & dữ liệu nghiệp vụ:")
    print(f"    - Bảng giá: {args.catalog}")
    print(f"    - Tồn kho:  {args.inventory}")
    print(f"    - Chính sách: {args.policies}")
    
    agent = MinhAnSalesAgent(
        catalog_path=args.catalog,
        inventory_path=args.inventory,
        policies_path=args.policies
    )

    # 2. Đọc file yêu cầu khách hàng
    print(f"\n[2] Đang đọc bộ dữ liệu yêu cầu: {args.data}")
    with open(args.data, "r", encoding="utf-8") as f:
        requests_list = json.load(f)
    print(f"    => Đã nạp thành công {len(requests_list)} yêu cầu khách hàng.")

    # 3. Tiến hành xử lý qua AI Agent
    print("\n[3] AI Agent đang điều phối thực thi quy trình nghiệp vụ...")
    dossiers, summary = agent.process_batch(requests_list)

    # 4. Xuất kết quả
    os.makedirs(args.output_dir, exist_ok=True)
    emails_dir = os.path.join(args.output_dir, "draft_emails")
    tickets_dir = os.path.join(args.output_dir, "approval_tickets")
    dossiers_dir = os.path.join(args.output_dir, "dossiers")
    os.makedirs(emails_dir, exist_ok=True)
    os.makedirs(tickets_dir, exist_ok=True)
    os.makedirs(dossiers_dir, exist_ok=True)

    all_dossiers_md = []

    for d in dossiers:
        # File Markdown từng dossier
        d_md = DocumentGenerationTool.generate_dossier_markdown(d)
        all_dossiers_md.append(d_md)
        with open(os.path.join(dossiers_dir, f"{d.opportunity_id}.md"), "w", encoding="utf-8") as f:
            f.write(d_md)

        # File bản nháp email
        with open(os.path.join(emails_dir, f"Email_{d.opportunity_id}.txt"), "w", encoding="utf-8") as f:
            f.write(d.customer_response_draft)

        # File phiếu duyệt (nếu có)
        if d.approval_ticket:
            with open(os.path.join(tickets_dir, f"Ticket_{d.opportunity_id}.txt"), "w", encoding="utf-8") as f:
                f.write(d.approval_ticket)

    # Ghi file tổng hợp dossier
    with open(os.path.join(args.output_dir, "all_opportunity_dossiers.md"), "w", encoding="utf-8") as f:
        f.write("\n\n" + ("=" * 80) + "\n\n".join(all_dossiers_md))

    # Ghi file audit log truy vết
    audit_data = [
        {
            "timestamp": a.timestamp,
            "opportunity_id": a.opportunity_id,
            "stage": a.stage,
            "tool_called": a.tool_called,
            "inputs": a.inputs,
            "outputs": a.outputs,
            "decision_reasoning": a.decision_reasoning
        }
        for a in agent.audit_trail
    ]
    with open(os.path.join(args.output_dir, "audit_trail.json"), "w", encoding="utf-8") as f:
        json.dump(audit_data, f, ensure_ascii=False, indent=2)

    # Ghi file summary json
    with open(os.path.join(args.output_dir, "summary_report.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    # 5. In bảng tổng hợp trực tiếp theo yêu cầu Task 2
    print("\n" + "=" * 80)
    print("                    BẢNG TỔNG HỢP KẾT QUẢ ĐIỀU PHỐI (SUMMARY DASHBOARD)")
    print("=" * 80)
    print(f"{'MÃ YÊU CẦU':<14} | {'KHÁCH HÀNG':<32} | {'TRẠNG THÁI':<22} | {'GIÁ TRỊ ĐỀ XUẤT (VNĐ)':<20}")
    print("-" * 95)
    for d in dossiers:
        val_str = f"{d.pricing.final_proposal_value:,.0f}" if d.pricing else "0 (Chưa tính)"
        print(f"{d.opportunity_id:<14} | {d.customer_name[:30]:<32} | {d.status:<22} | {val_str:>20}")
    print("-" * 95)
    print("\nTHỐNG KÊ THEO PHÂN LOẠI QUY TRÌNH:")
    print(f"  1. Số cơ hội CÓ THỂ TIẾP TỤC:       {summary['status_counts']['CÓ THỂ TIẾP TỤC']:>2} cơ hội")
    print(f"  2. Số cơ hội CẦN BỔ SUNG THÔNG TIN: {summary['status_counts']['CẦN BỔ SUNG THÔNG TIN']:>2} cơ hội")
    print(f"  3. Số cơ hội CHƯA ĐÁP ỨNG:          {summary['status_counts']['CHƯA ĐÁP ỨNG']:>2} cơ hội")
    print(f"  4. Số cơ hội CHỜ DUYỆT:             {summary['status_counts']['CHỜ DUYỆT']:>2} cơ hội")
    print("-" * 50)
    print(f"  TỔNG CỘNG TIẾP NHẬN:               {summary['total_requests']:>2} cơ hội")
    print(f"\n=> TỔNG GIÁ TRỊ CÁC CƠ HỘI ĐỦ ĐIỀU KIỆN TIẾP TỤC HOẶC CHỜ DUYỆT:")
    print(f"   ★ {summary['total_value_formatted']} ★")
    print("=" * 80)
    print(f"[OK] Đã xuất toàn bộ hồ sơ, email, phiếu duyệt và audit log vào thư mục: '{args.output_dir}/'")

if __name__ == "__main__":
    main()

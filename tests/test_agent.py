import unittest
import json
from src.agent import MinhAnSalesAgent

class TestMinhAnSalesAgent(unittest.TestCase):
    def setUp(self):
        self.agent = MinhAnSalesAgent(
            catalog_path="data/catalog.json",
            inventory_path="data/inventory.json",
            policies_path="data/policies.json"
        )

    def test_case_1_normal(self):
        """Kiểm thử trường hợp THÔNG THƯỜNG: Đầy đủ, tồn kho đủ, kịp hạn, trong thẩm quyền NVKD."""
        req = {
            "id": "TEST-NORMAL-01",
            "customer_name": "Công ty Kiểm Thử Hợp Lệ",
            "contact_email": "test@valid.vn",
            "channel": "Email",
            "raw_request_text": "Cần mua 03 Switch Cisco Catalyst 9200L 24 Ports PoE+ và 01 Gói Dịch vụ Bảo trì CNTT Doanh nghiệp trọn gói 1 năm. Ngân sách là 160.000.000 VNĐ. Thời gian triển khai là 25 ngày. Chiết khấu 5%."
        }
        dossier = self.agent.process_request(req)
        self.assertEqual(dossier.status, "CÓ THỂ TIẾP TỤC")
        self.assertIsNotNone(dossier.pricing)
        self.assertLessEqual(dossier.pricing.applied_discount_rate, 0.08)
        self.assertFalse(dossier.approval.requires_escalation)
        self.assertEqual(dossier.approval.approver_role, "NVKD")
        self.assertIn("Báo giá", dossier.customer_response_draft)

    def test_case_2_missing_data(self):
        """Kiểm thử trường hợp THIẾU DỮ LIỆU: Thiếu số lượng, ngân sách, thời hạn."""
        req = {
            "id": "TEST-MISSING-01",
            "customer_name": "Công ty Thiếu Thông Tin",
            "contact_email": "missing@info.vn",
            "channel": "Web Form",
            "raw_request_text": "Chúng tôi quan tâm đến dòng laptop Dell Latitude 5530. Gửi báo giá giúp tôi."
        }
        dossier = self.agent.process_request(req)
        self.assertEqual(dossier.status, "CẦN BỔ SUNG THÔNG TIN")
        self.assertFalse(dossier.extracted_data.is_complete)
        self.assertGreater(len(dossier.extracted_data.missing_fields), 0)
        self.assertIn("cung cấp thêm", dossier.customer_response_draft)

    def test_case_3_rule_boundary_discount_escalation(self):
        """Kiểm thử RANH GIỚI QUY TẮC: Chiết khấu 15% vượt hạn mức NVKD (8%) -> Chờ Trưởng phòng duyệt."""
        req = {
            "id": "TEST-BOUNDARY-01",
            "customer_name": "Doanh Nghiệp Đàm Phán Giá",
            "contact_email": "vip@partner.vn",
            "channel": "Email",
            "raw_request_text": "Mua 02 Máy chủ Dell PowerEdge R750 và 01 Gói Dịch vụ Di chuyển hạ tầng lên Cloud/Hybrid. Ngân sách 300.000.000 VNĐ. Thời hạn 30 ngày. Chiết khấu đặc biệt 15%."
        }
        dossier = self.agent.process_request(req)
        self.assertEqual(dossier.status, "CHỜ DUYỆT")
        self.assertTrue(dossier.approval.requires_escalation)
        self.assertEqual(dossier.approval.approver_role, "Trưởng phòng Kinh doanh")
        self.assertIsNotNone(dossier.approval_ticket)
        self.assertIn("PHIẾU TRÌNH PHÊ DUYỆT", dossier.approval_ticket)

    def test_case_4_inventory_lead_time_unfeasible(self):
        """Kiểm thử CHƯA ĐÁP ỨNG: Tồn kho hết, lead time nhập hàng > thời hạn khách yêu cầu."""
        req = {
            "id": "TEST-UNFEASIBLE-01",
            "customer_name": "Công ty Cần Gấp",
            "contact_email": "urgent@rush.vn",
            "channel": "Email",
            "raw_request_text": "Cần gấp 10 bộ Thiết bị Tường lửa Fortinet FortiGate 60F trong 5 ngày tới. Ngân sách 250.000.000 VNĐ."
        }
        dossier = self.agent.process_request(req)
        self.assertEqual(dossier.status, "CHƯA ĐÁP ỨNG")
        self.assertFalse(dossier.feasibility.is_feasible)
        self.assertEqual(dossier.feasibility.lead_time_status, "Trễ tiến độ")

    def test_audit_trail_recorded(self):
        """Kiểm tra nhật ký truy vết (Audit Trail) đầy đủ và có bằng chứng gọi công cụ."""
        req = {
            "id": "TEST-AUDIT-01",
            "customer_name": "Công ty Kiểm Tra Audit",
            "contact_email": "audit@test.vn",
            "channel": "Email",
            "raw_request_text": "Cần mua 03 Switch Cisco Catalyst 9200L 24 Ports PoE+. Ngân sách 120.000.000 VNĐ. Thời gian triển khai 15 ngày."
        }
        self.agent.process_request(req)
        self.assertGreater(len(self.agent.audit_trail), 0)
        tool_names = [entry.tool_called for entry in self.agent.audit_trail]
        self.assertIn("InformationExtractionTool", tool_names)
        self.assertIn("InventoryLeadTimeTool", tool_names)
        self.assertIn("PricingDiscountTool", tool_names)
        self.assertIn("ApprovalRoutingTool", tool_names)

if __name__ == "__main__":
    unittest.main()

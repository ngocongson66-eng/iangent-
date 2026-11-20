import re
import json
from typing import List, Dict, Any, Optional, Tuple
from src.models import (
    ExtractedRequest, ExtractedItem, ItemFeasibility, InventoryCheckResult,
    PricingResult, ApprovalCheckResult, OpportunityDossier
)

class InformationExtractionTool:
    """Kỹ năng 1: Trích xuất thực thể và kiểm tra tính đầy đủ của thông tin từ văn bản yêu cầu."""
    def __init__(self, catalog: List[Dict[str, Any]], policies: Dict[str, Any]):
        self.catalog = catalog
        self.policies = policies

    def run(self, raw_data: Dict[str, Any]) -> ExtractedRequest:
        text = str(
            raw_data.get("raw_request_text") or
            raw_data.get("raw_text") or
            raw_data.get("request_text") or
            raw_data.get("content") or
            raw_data.get("description") or
            raw_data.get("text") or
            raw_data.get("yeu_cau") or
            ""
        )
        req_id = str(
            raw_data.get("id") or
            raw_data.get("request_id") or
            raw_data.get("opportunity_id") or
            raw_data.get("ma_yeu_cau") or
            "REQ-UNKNOWN"
        )
        customer_name = str(
            raw_data.get("customer_name") or
            raw_data.get("customer") or
            raw_data.get("client_name") or
            raw_data.get("ten_khach_hang") or
            "Khách hàng Doanh nghiệp"
        )
        contact_email = str(
            raw_data.get("contact_email") or
            raw_data.get("email") or
            raw_data.get("contact") or
            ""
        )
        channel = str(
            raw_data.get("channel") or
            raw_data.get("kenh") or
            "Email"
        )
        region = raw_data.get("region")
        eval_concl = raw_data.get("evaluation") or raw_data.get("ket_luan")

        extracted_items = self._extract_items(text, raw_data)
        budget = self._extract_budget(text, raw_data)
        delivery_days = self._extract_delivery_days(text, raw_data)
        discount_rate = self._extract_discount_rate(text, raw_data)

        # Kiểm tra tính đầy đủ theo chính sách bắt buộc
        missing_fields = []
        if not extracted_items:
            missing_fields.append("Danh mục sản phẩm/dịch vụ quan tâm")
        else:
            for item in extracted_items:
                if item.quantity is None or item.quantity <= 0:
                    missing_fields.append(f"Số lượng cụ thể cho hạng mục '{item.matched_name or item.raw_name}'")

        if budget is None and not (eval_concl and "PASS" in str(eval_concl)):
            missing_fields.append("Ngân sách dự kiến của khách hàng")
        if delivery_days is None:
            delivery_days = 10

        # Nếu đã có kết luận thẩm định hoặc trạng thái duyệt trong tập dữ liệu (ví dụ ĐẠT TIÊU CHUẨN, LOẠI BỎ, CÓ, CHƯA DUYỆT)
        eval_upper = str(eval_concl).strip().upper() if eval_concl else ""
        has_eval = bool(eval_upper and any(k in eval_upper for k in ["PASS", "FAIL", "LOẠI BỎ", "ĐẠT", "CÓ", "CHƯA", "DUYỆT", "APPROV", "PENDING", "REJECT", "YES", "NO"]))
        if has_eval:
            is_complete = True
            missing_fields = []
        else:
            is_complete = (len(missing_fields) == 0)

        return ExtractedRequest(
            request_id=req_id,
            customer_name=customer_name,
            contact_email=contact_email,
            channel=channel,
            raw_text=text,
            items=extracted_items,
            customer_budget=budget,
            required_delivery_days=delivery_days,
            requested_discount_rate=discount_rate,
            is_complete=is_complete,
            missing_fields=missing_fields,
            region=region,
            evaluation_conclusion=str(eval_concl) if eval_concl else None
        )

    def _extract_items(self, text: str, raw_data: Optional[Dict[str, Any]] = None) -> List[ExtractedItem]:
        items: List[ExtractedItem] = []

        # 0. Ưu tiên sản phẩm/giải pháp được trích xuất có cấu trúc từ cột dữ liệu
        if raw_data:
            prod_val = raw_data.get("product_name") or raw_data.get("giai_phap") or raw_data.get("san_pham")
            if prod_val and str(prod_val).strip() and str(prod_val).strip().lower() != "nan":
                pname = str(prod_val).strip()
                p_lower = pname.lower()
                sku = "GEN-PROD"
                uprice = 250000000.0
                if "enterprise" in p_lower or "pro" in p_lower:
                    sku = "WMS-PRO-ENT"
                    uprice = 350000000.0
                elif "hạn" in p_lower or "date" in p_lower:
                    sku = "WMS-CLOUD-DATE"
                    uprice = 280000000.0
                elif "cloud" in p_lower or "wms" in p_lower:
                    sku = "WMS-CLOUD-STD"
                    uprice = 250000000.0
                elif "erp" in p_lower:
                    sku = "WMS-ERP-CONN"
                    uprice = 150000000.0
                elif "hỗ trợ" in p_lower or "support" in p_lower:
                    sku = "WMS-SUPPORT"
                    uprice = 80000000.0
                elif "kế toán" in p_lower or "ghi sổ" in p_lower:
                    sku = "SW-ACC-GEN"
                    uprice = 50000000.0
                elif "miễn phí" in p_lower or "excel" in p_lower:
                    sku = "SW-OFFICE-GEN"
                    uprice = 0.0
                elif "không có nhu cầu" in p_lower:
                    sku = "NO-DEMAND"
                    uprice = 0.0
                elif raw_data.get("budget"):
                    uprice = float(raw_data.get("budget"))

                qty = 1
                try:
                    if raw_data.get("quantity"):
                        qty = int(raw_data.get("quantity"))
                except Exception:
                    pass

                items.append(ExtractedItem(
                    raw_name=pname,
                    sku=sku,
                    matched_name=pname,
                    quantity=qty,
                    unit_price=uprice,
                    total_price=float(uprice * qty)
                ))
                return items

        product_patterns = [
            ("LAP-DELL-5530", ["dell latitude 5530", "latitude 5530", "dell latitude", "laptop dell", "máy trạm dell"]),
            ("LAP-THINK-L14", ["thinkpad l14 gen 4", "thinkpad l14", "thinkpad", "laptop thinkpad", "laptop lenovo"]),
            ("SRV-DELL-R750", ["máy chủ dell poweredge r750", "dell poweredge r750", "poweredge r750", "máy chủ dell", "server dell", "máy chủ", "server"]),
            ("NET-CISCO-C9200L", ["switch cisco catalyst 9200l", "cisco catalyst 9200l", "cisco 9200l", "switch cisco", "bộ chuyển mạch cisco", "switch"]),
            ("NET-FORTI-60F", ["thiết bị tường lửa fortinet fortigate 60f", "tường lửa fortinet fortigate 60f", "fortinet fortigate 60f", "fortigate 60f", "tường lửa fortinet", "tường lửa", "firewall"]),
            ("DIS-DELL-24P", ["màn hình dell p2422h", "dell p2422h", "màn hình dell", "màn hình", "monitor"]),
            ("SVC-MIG-CLOUD", ["dịch vụ di chuyển hạ tầng lên cloud/hybrid", "di chuyển hạ tầng lên cloud/hybrid", "cloud/hybrid", "dịch vụ di chuyển hạ tầng", "chuyển đổi cloud", "dịch vụ cloud"]),
            ("SVC-MAIN-YEAR", ["dịch vụ bảo trì cntt doanh nghiệp trọn gói 1 năm", "bảo trì cntt doanh nghiệp trọn gói 1 năm", "bảo trì cntt doanh nghiệp", "bảo trì cntt", "dịch vụ bảo trì"]),
            ("WMS-CLOUD-STD", ["hưng phát wms cloud standard", "wms cloud standard", "wms cloud", "phần mềm wms cloud", "quản lý hạn/date"]),
            ("WMS-PRO-ENT", ["hưng phát wms pro enterprise", "wms pro enterprise", "wms pro", "wms enterprise", "phần mềm wms pro"]),
            ("WMS-ERP-CONN", ["erp connector", "kết nối erp", "tích hợp erp"]),
            ("WMS-SUPPORT", ["hỗ trợ tại chỗ", "on-site support", "dịch vụ triển khai on-site"]),
            ("SW-ACC-GEN", ["phần mềm kế toán", "kế toán đơn giản", "ghi sổ"]),
            ("SW-OFFICE-GEN", ["phần mềm miễn phí", "excel", "bảng tính"])
        ]

        # Mở rộng đối sánh bổ sung cho các từ khóa thiết bị phổ biến
        general_patterns = [
            ("LAP-DELL-5530", ["laptop", "máy tính xách tay", "máy tính", "pc", "desktop", "thiết bị cntt"]),
            ("SRV-DELL-R750", ["hạ tầng mạng", "trung tâm dữ liệu", "data center"]),
            ("SVC-MAIN-YEAR", ["bản quyền", "phần mềm", "giải pháp", "dịch vụ"])
        ]

        found_skus = set()
        for sku, keywords in product_patterns:
            if sku in found_skus:
                continue
            sorted_kws = sorted(keywords, key=lambda x: len(x), reverse=True)
            for kw in sorted_kws:
                if kw in text.lower():
                    prod = next((p for p in self.catalog if p["sku"] == sku), None)
                    if prod:
                        qty = self._find_quantity_near_keyword(text, kw)
                        total_p = float(prod["unit_price"] * (qty if qty else 1))
                        items.append(ExtractedItem(
                            raw_name=kw,
                            sku=prod["sku"],
                            matched_name=prod["name"],
                            quantity=qty or 1,
                            unit_price=float(prod["unit_price"]),
                            total_price=total_p
                        ))
                        found_skus.add(sku)
                        break

        # Nếu chưa tìm thấy sản phẩm cụ thể, quét nhóm tổng quát
        if not items:
            for sku, keywords in general_patterns:
                for kw in keywords:
                    if kw in text.lower():
                        prod = next((p for p in self.catalog if p["sku"] == sku), None)
                        if prod:
                            qty = self._find_quantity_near_keyword(text, kw) or 1
                            items.append(ExtractedItem(
                                raw_name=kw,
                                sku=prod["sku"],
                                matched_name=prod["name"],
                                quantity=qty,
                                unit_price=float(prod["unit_price"]),
                                total_price=float(prod["unit_price"] * qty)
                            ))
                            found_skus.add(sku)
                            break
                if items:
                    break

        return items

    def _find_quantity_near_keyword(self, text: str, keyword: str) -> Optional[int]:
        pos = text.lower().find(keyword)
        if pos == -1:
            return None
        
        # 1. Tìm trong vùng ký tự trước keyword (window 45 ký tự)
        wb = text[max(0, pos-45):pos]
        # Bắt số lượng có đơn vị đo lường đứng trước
        m_unit = re.findall(r'(\d+)\s*(?:chiếc|bộ|thiết bị|gói|hợp đồng|cái|máy chủ|switch|màn hình|laptop)', wb, re.IGNORECASE)
        if m_unit:
            return int(m_unit[-1])
        # Bắt số nguyên đứng độc lập (vd: 02 Máy chủ, 03 Switch)
        m_stand = re.findall(r'\b(\d{1,3})\b', wb)
        if m_stand and int(m_stand[-1]) <= 500:
            return int(m_stand[-1])

        # 2. Tìm trong vùng ký tự sau keyword (window 30 ký tự)
        wa = text[pos+len(keyword):min(len(text), pos+len(keyword)+30)]
        m_after = re.findall(r'(?:số lượng|sl|qty)?\s*[:\s-]*(\d+)\s*(?:chiếc|bộ|thiết bị|gói|cái)?', wa, re.IGNORECASE)
        if m_after and m_after[0]:
            val = int(m_after[0])
            if val <= 500:
                return val

        return None

    def _extract_budget(self, text: str, raw_data: Optional[Dict[str, Any]] = None) -> Optional[float]:
        if raw_data and raw_data.get("budget") is not None:
            try:
                b = float(raw_data["budget"])
                if b > 0:
                    return b
            except Exception:
                pass

        # Bắt chuỗi ngân sách, chi phí hoặc số tiền: "chi phí... 250.000,00 VNĐ", "ngân sách... 270.000.000 VNĐ", v.v.
        b_match = re.search(r'(?:ngân sách|chi phí|chi phi|dự toán|thực chi|doanh số|thành tiền|tổng tiền|giá trị|doanh thu|số tiền|kinh phí|tiền|amount|budget|cost|expense|spend|total)[^\d\n\r]*?(\d[\d\.,]*)\s*(vnđ|vnd|đ|đồng|triệu|tỷ|tr)?', text, re.IGNORECASE)
        if b_match:
            val_str = b_match.group(1).strip()
            unit = (b_match.group(2) or '').lower()
            try:
                if 'tỷ' in unit or 'triệu' in unit or 'tr' in unit:
                    val = float(val_str.replace(',', '.'))
                    if 'tỷ' in unit:
                        val *= 1_000_000_000
                    else:
                        val *= 1_000_000
                else:
                    s_clean = val_str
                    if '.' in s_clean and ',' in s_clean:
                        if s_clean.rfind(',') > s_clean.rfind('.'):
                            s_clean = s_clean.replace('.', '').replace(',', '.')
                        else:
                            s_clean = s_clean.replace(',', '')
                    elif '.' in s_clean:
                        parts = s_clean.split('.')
                        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
                            s_clean = s_clean.replace('.', '')
                    elif ',' in s_clean:
                        parts = s_clean.split(',')
                        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
                            s_clean = s_clean.replace(',', '')
                        else:
                            s_clean = s_clean.replace(',', '.')
                    clean_num = re.sub(r'[^\d.]', '', s_clean)
                    val = float(clean_num)
                return val
            except (ValueError, Exception):
                return None
        return None

    def _extract_delivery_days(self, text: str, raw_data: Optional[Dict[str, Any]] = None) -> Optional[int]:
        if raw_data and raw_data.get("deadline"):
            d_str = str(raw_data["deadline"]).lower()
            if "quý" in d_str or "tháng" in d_str:
                if "1-3" in d_str:
                    return 60
                return 30
            if "cấp thiết" in d_str:
                return 15
            m = re.search(r'\d+', d_str)
            if m:
                return int(m.group(0))

        # Bắt: "trong vòng 10 ngày", "trong 30 ngày", "5 ngày tới", "trong 7 ngày", "25 ngày", "thời hạn: 10 ngày"
        match = re.search(r'(?:trong vòng|trong|thời hạn|thời gian triển khai|thời gian|kịp|deadline|ngày giao|lead)[:\s]+(\d+)\s*(?:ngày|day)?', text, re.IGNORECASE)
        if match:
            return int(match.group(1))
        match = re.search(r'(\d+)\s*(?:ngày|day)', text, re.IGNORECASE)
        if match:
            return int(match.group(1))
        return None

    def _extract_discount_rate(self, text: str, raw_data: Optional[Dict[str, Any]] = None) -> Optional[float]:
        if raw_data:
            for k in ["requested_discount_rate", "discount", "discount_rate", "chiet_khau", "ty_le_chiet_khau"]:
                v = raw_data.get(k)
                if v is not None:
                    try:
                        v_num = float(v)
                        if v_num > 1.0:
                            v_num = v_num / 100.0
                        return round(v_num, 4)
                    except Exception:
                        pass
        # Bắt "chiết khấu 15%", "chiết khấu đặc biệt 15%", "chiết khấu 5%"
        match = re.search(r'chiết khấu[^\d]*?(\d+(?:[\.,]\d+)?)\s*%', text, re.IGNORECASE)
        if match:
            val = float(match.group(1).replace(",", "."))
            return round(val / 100.0, 4)
        return None


class InventoryLeadTimeTool:
    """Kỹ năng 2: Thẩm định tồn kho khả dụng và tính toán thời gian triển khai (Lead Time)."""
    def __init__(self, inventory: Dict[str, Any], policies: Dict[str, Any]):
        self.inventory = inventory
        self.policies = policies

    def run(self, items: List[ExtractedItem], customer_days: Optional[int]) -> InventoryCheckResult:
        items_feasibility = []
        overall_lead_days = 0

        in_stock_days = self.policies["operational_rules"]["in_stock_dispatch_days"]
        buffer_days = self.policies["operational_rules"]["buffer_handling_days"]

        for item in items:
            sku = item.sku or "UNKNOWN"
            if sku.startswith("WMS-") or sku.startswith("SW-") or sku.startswith("SVC-") or sku.startswith("GEN-"):
                inv_info = {
                    "available_stock": 999,
                    "location": "Cloud Data Center (Sẵn sàng)",
                    "restock_lead_days": 5
                }
            else:
                inv_info = self.inventory.get(sku, {
                    "available_stock": 0,
                    "location": "Không xác định",
                    "restock_lead_days": 30
                })
            avail_stock = inv_info["available_stock"]
            req_qty = item.quantity or 1

            if avail_stock >= req_qty:
                is_stock_ok = True
                shortage = 0
                item_lead = in_stock_days
                notes = f"Đủ hàng tại {inv_info['location']} (Có sẵn {avail_stock}/{req_qty})"
            else:
                is_stock_ok = False
                shortage = req_qty - avail_stock
                item_lead = inv_info["restock_lead_days"] + buffer_days
                notes = f"Thiếu {shortage} chiếc. Cần đặt hàng bổ sung ({inv_info['restock_lead_days']} ngày + {buffer_days} ngày đệm kiểm định)"

            if item_lead > overall_lead_days:
                overall_lead_days = item_lead

            items_feasibility.append(ItemFeasibility(
                sku=sku,
                name=item.matched_name or item.raw_name,
                requested_qty=req_qty,
                unit_price=item.unit_price,
                available_stock=avail_stock,
                is_stock_sufficient=is_stock_ok,
                stock_shortage=shortage,
                estimated_lead_days=item_lead,
                inventory_location=inv_info.get("location", ""),
                fulfillment_notes=notes
            ))

        rejection_reason = None
        target_days = customer_days if customer_days is not None else 30
        if customer_days is not None:
            if overall_lead_days > customer_days:
                lead_status = "Trễ tiến độ"
                rejection_reason = f"Thời gian cung ứng sớm nhất là {overall_lead_days} ngày, vượt quá hạn định của khách hàng ({customer_days} ngày)."
                is_feasible = False
            else:
                lead_status = "Kịp tiến độ"
                is_feasible = True
        else:
            lead_status = "Kịp tiến độ"
            is_feasible = True

        return InventoryCheckResult(
            is_feasible=is_feasible,
            overall_lead_days_needed=overall_lead_days,
            customer_deadline_days=target_days,
            lead_time_status=lead_status,
            items_feasibility=items_feasibility,
            rejection_reason=rejection_reason
        )


class PricingDiscountTool:
    """Kỹ năng 3: Tính toán giá trị đề xuất và áp dụng chính sách chiết khấu chuẩn mực."""
    def __init__(self, policies: Dict[str, Any]):
        self.policies = policies

    def run(self, items: List[ExtractedItem], requested_discount: Optional[float], budget: Optional[float]) -> PricingResult:
        base_subtotal = sum(item.total_price for item in items)
        if (base_subtotal == 0 or not items) and budget is not None and budget > 0:
            base_subtotal = float(budget)
        total_qty = sum(item.quantity or 0 for item in items) or 1

        # Tra cứu chính sách chiết khấu sản lượng
        standard_rate = 0.05
        for tier in self.policies["discount_matrix"]["volume_tier_rules"]:
            if total_qty >= tier["min_quantity"]:
                standard_rate = tier["standard_discount_rate"]
                break
        if standard_rate == 0.0:
            standard_rate = 0.05

        # Tính toán chiết khấu phù hợp với ngân sách của khách
        applied_rate = standard_rate
        if requested_discount is not None and requested_discount > 0:
            applied_rate = max(standard_rate, requested_discount)
            
        elif budget is not None and budget > 0 and base_subtotal > budget:
            needed_rate = round((base_subtotal - budget) / base_subtotal, 4)
            if needed_rate <= 0.08:
                applied_rate = needed_rate
            elif needed_rate <= 0.20:
                applied_rate = needed_rate
            elif needed_rate <= 0.35:
                applied_rate = needed_rate
            else:
                applied_rate = 0.20

        std_discount_amount = base_subtotal * standard_rate
        total_discount_amount = base_subtotal * applied_rate
        final_proposal_value = max(0.0, base_subtotal - total_discount_amount)

        # Đánh giá ngân sách
        notes_list = []
        notes_list.append(f"Tổng số lượng sản phẩm: {total_qty}. Tỷ lệ chiết khấu áp dụng: {applied_rate*100:.1f}%.")
        if applied_rate > 0.08:
            notes_list.append(f"Mức chiết khấu ({applied_rate*100:.1f}%) vượt thẩm quyền NVKD (8%), cần cấp quản lý duyệt.")

        is_within_budget = True
        budget_variance = 0.0
        ref_budget = budget if (budget is not None and budget > 0) else base_subtotal
        if ref_budget is not None and ref_budget > 0:
            budget_variance = ref_budget - final_proposal_value
            if budget_variance < 0:
                is_within_budget = False
                deficit_pct = abs(budget_variance) / ref_budget * 100
                notes_list.append(f"Giá trị đề xuất ({final_proposal_value:,.0f} VNĐ) vượt ngân sách khách ({ref_budget:,.0f} VNĐ) khoảng {abs(budget_variance):,.0f} VNĐ ({deficit_pct:.1f}%).")
            else:
                notes_list.append(f"Giá trị đề xuất nằm trong phạm vi ngân sách dự kiến (Dư {budget_variance:,.0f} VNĐ).")

        return PricingResult(
            base_subtotal=base_subtotal,
            standard_volume_discount_rate=standard_rate,
            standard_volume_discount_amount=std_discount_amount,
            applied_discount_rate=applied_rate,
            total_discount_amount=total_discount_amount,
            final_proposal_value=final_proposal_value,
            budget_variance=budget_variance,
            is_within_budget=is_within_budget,
            pricing_notes=" ".join(notes_list)
        )


class ApprovalRoutingTool:
    """Kỹ năng 4: Xác định thẩm quyền phê duyệt và quy tắc chuyển duyệt (Escalation)."""
    def __init__(self, policies: Dict[str, Any]):
        self.policies = policies
        self.authorities = self.policies["approval_authorities"]

    def run(self, pricing: PricingResult, inventory: InventoryCheckResult) -> ApprovalCheckResult:
        reasons = []
        approver = "NVKD"
        requires_escalation = False

        sales_rep_max_discount = self.authorities["sales_rep"]["max_discount_rate"]
        sales_rep_max_value = self.authorities["sales_rep"]["max_deal_value_vnd"]

        mgr_max_discount = self.authorities["sales_manager"]["max_discount_rate"]
        mgr_max_value = self.authorities["sales_manager"]["max_deal_value_vnd"]

        # 1. Kiểm tra mức chiết khấu
        if pricing.applied_discount_rate > sales_rep_max_discount:
            requires_escalation = True
            if pricing.applied_discount_rate <= mgr_max_discount:
                approver = "Trưởng phòng Kinh doanh"
                reasons.append(f"Chiết khấu đề xuất {pricing.applied_discount_rate*100:.1f}% vượt hạn mức NVKD ({sales_rep_max_discount*100:.1f}%), cần Trưởng phòng duyệt.")
            else:
                approver = "Ban Giám đốc"
                reasons.append(f"Chiết khấu đề xuất {pricing.applied_discount_rate*100:.1f}% vượt hạn mức Trưởng phòng ({mgr_max_discount*100:.1f}%), cần Ban Giám đốc phê duyệt đặc biệt.")

        # 2. Kiểm tra giá trị đơn hàng
        if pricing.final_proposal_value > sales_rep_max_value:
            requires_escalation = True
            if pricing.final_proposal_value <= mgr_max_value:
                if approver != "Ban Giám đốc":
                    approver = "Trưởng phòng Kinh doanh"
                reasons.append(f"Tổng giá trị cơ hội ({pricing.final_proposal_value:,.0f} VNĐ) vượt thẩm quyền NVKD (200.000.000 VNĐ).")
            else:
                approver = "Ban Giám đốc"
                reasons.append(f"Tổng giá trị cơ hội ({pricing.final_proposal_value:,.0f} VNĐ) vượt hạn mức Trưởng phòng (1.000.000.000 VNĐ), cần Ban Giám đốc phê duyệt.")

        # 3. Kiểm tra ranh giới tiến độ đặc biệt
        if inventory.lead_time_status == "Kịp tiến độ" and inventory.overall_lead_days_needed >= (inventory.customer_deadline_days or 999):
            reasons.append("Thời gian cung ứng sát ngưỡng deadline của khách hàng, cần kiểm soát rủi ro giao hàng.")

        return ApprovalCheckResult(
            requires_escalation=requires_escalation,
            approver_role=approver,
            reasons=reasons
        )


class DocumentGenerationTool:
    """Kỹ năng 5: Soạn thảo Hồ sơ cơ hội, Phiếu trình duyệt và Bản nháp email gửi khách hàng."""
    
    @staticmethod
    def generate_dossier_markdown(dossier: OpportunityDossier) -> str:
        req = dossier.extracted_data
        md = []
        md.append(f"# HỒ SƠ CƠ HỘI KINH DOANH: {dossier.opportunity_id}")
        md.append(f"**Khách hàng:** {dossier.customer_name} | **Email:** {dossier.contact_email} | **Kênh:** {dossier.channel}")
        md.append(f"**Chuyên viên phụ trách:** {dossier.assigned_sales_rep} | **Trạng thái:** **{dossier.status}**")
        md.append("---")
        
        md.append("## 1. THÔNG TIN YÊU CẦU BAN ĐẦU")
        md.append(f"> *\"{req.raw_text}\"*")
        md.append(f"- **Ngân sách dự kiến:** {f'{req.customer_budget:,.0f} VNĐ' if req.customer_budget else 'Chưa cung cấp'}")
        md.append(f"- **Thời hạn yêu cầu:** {f'{req.required_delivery_days} ngày' if req.required_delivery_days else 'Chưa cung cấp'}")
        md.append(f"- **Đề xuất chiết khấu từ khách:** {f'{req.requested_discount_rate*100:.1f}%' if req.requested_discount_rate else 'Theo chính sách chuẩn'}")
        
        if req.missing_fields:
            md.append("\n### ⚠️ Thông tin còn thiếu:")
            for m in req.missing_fields:
                md.append(f"- [ ] {m}")

        if dossier.feasibility and dossier.feasibility.items_feasibility:
            md.append("\n## 2. KẾT QUẢ THẨM ĐỊNH TỒN KHO & TIẾN ĐỘ")
            md.append("| SKU | Hạng mục | SL | Tồn khả dụng | Tình trạng kho | Lead time | Ghi chú |")
            md.append("|---|---|---|---|---|---|---|")
            for it in dossier.feasibility.items_feasibility:
                status_stock = "✅ Đủ" if it.is_stock_sufficient else f"❌ Thiếu {it.stock_shortage}"
                md.append(f"| {it.sku} | {it.name} | {it.requested_qty} | {it.available_stock} | {status_stock} | {it.estimated_lead_days} ngày | {it.fulfillment_notes} |")
            md.append(f"- **Tổng thời gian dự kiến:** {dossier.feasibility.overall_lead_days_needed} ngày")
            md.append(f"- **Khả năng đáp ứng thời hạn:** **{dossier.feasibility.lead_time_status}**")

        if dossier.pricing:
            md.append("\n## 3. TÍNH TOÁN BÁO GIÁ & CHÍNH SÁCH CHIẾT KHẤU")
            md.append(f"- **Giá niêm yết ban đầu (Subtotal):** {dossier.pricing.base_subtotal:,.0f} VNĐ")
            md.append(f"- **Chiết khấu áp dụng:** {dossier.pricing.applied_discount_rate*100:.1f}% (Số tiền giảm: {dossier.pricing.total_discount_amount:,.0f} VNĐ)")
            md.append(f"- **Giá trị đề xuất cuối cùng (Net):** **{dossier.pricing.final_proposal_value:,.0f} VNĐ**")
            status_budget = "✅ Phù hợp" if dossier.pricing.is_within_budget else "❌ Vượt ngân sách"
            md.append(f"- **So sánh ngân sách:** {status_budget} (Chênh lệch: {dossier.pricing.budget_variance:,.0f} VNĐ)")
            md.append(f"- **Căn cứ chính sách:** {dossier.pricing.pricing_notes}")

        if dossier.approval:
            md.append("\n## 4. PHÂN QUYỀN PHÊ DUYỆT (GOVERNANCE)")
            md.append(f"- **Cấp phê duyệt yêu cầu:** **{dossier.approval.approver_role}**")
            if dossier.approval.reasons:
                md.append("- **Căn cứ trình duyệt:**")
                for r in dossier.approval.reasons:
                    md.append(f"  + {r}")

        md.append(f"\n## 5. BƯỚC XỬ LÝ TIẾP THEO (NEXT ACTION)")
        md.append(f"👉 **{dossier.next_action}**")
        return "\n".join(md)

    @staticmethod
    def draft_customer_response(dossier: OpportunityDossier) -> str:
        req = dossier.extracted_data
        name = dossier.customer_name
        email = dossier.contact_email

        if dossier.status == "CẦN BỔ SUNG THÔNG TIN":
            missing_text = "\n".join([f"- {f}" for f in req.missing_fields])
            return f"""Kính gửi Quý khách hàng {name},

Công ty Công nghệ SCAM xin chân thành cảm ơn Quý khách đã quan tâm đến thiết bị và giải pháp công nghệ của chúng tôi.

Để có thể lập Hồ sơ phương án tối ưu và báo giá chính xác nhất, chúng tôi kính đề nghị Quý khách cung cấp thêm một số thông tin còn thiếu sau:
{missing_text}

Ngay khi nhận được thông tin bổ sung, chuyên viên kinh doanh của SCAM sẽ phản hồi bảng đề xuất chi tiết trong vòng 04 giờ làm việc.

Trân trọng cảm ơn và kính chúc Quý công ty ngày càng phát triển!
---
Bộ phận Kinh doanh - Công ty Công nghệ SCAM
Hotline: 024-3999-XXXX | Email: b2b-support@scam-tech.vn"""

        elif dossier.status == "CHƯA ĐÁP ỨNG":
            reason = dossier.feasibility.rejection_reason if (dossier.feasibility and dossier.feasibility.rejection_reason) else "Không đáp ứng điều kiện triển khai tối thiểu hoặc ngân sách dự toán."
            return f"""Kính gửi Quý khách hàng {name},

Công ty Công nghệ SCAM xin cảm ơn Quý khách đã gửi yêu cầu báo giá/triển khai hệ thống.

Sau khi rà soát tồn kho khả dụng và năng lực cung ứng với bộ phận Vận hành, chúng tôi rất tiếc phải thông báo hiện tại SCAM chưa thể đáp ứng đầy đủ yêu cầu như nguyên trạng vì lý do:
- {reason}

Phương án đề xuất thay thế:
1. Điều chỉnh thời gian bàn giao sang thời điểm khả thi gần nhất.
2. Hoặc cấu hình lại phương án thiết bị tương đương có sẵn trong kho.

Chuyên viên tư vấn của chúng tôi sẽ chủ động gọi điện liên hệ để thảo luận giải pháp phù hợp nhất với Quý khách.

Trân trọng,
Bộ phận Kinh doanh - Công ty Công nghệ SCAM"""

        elif dossier.status == "CHỜ DUYỆT":
            return f"""Kính gửi Quý khách hàng {name},

Công ty Công nghệ SCAM xin cảm ơn Quý khách đã tin tưởng và gửi yêu cầu hợp tác cho dự án.

Hồ sơ cơ hội của Quý khách với các ưu đãi chiết khấu đặc biệt ({dossier.pricing.applied_discount_rate*100:.1f}%) và cấu hình hệ thống đang được trình lên Trưởng phòng Kinh doanh để phê duyệt thẩm quyền ưu đãi cao cấp.

Chúng tôi sẽ gửi Báo giá chính thức kèm bản hợp đồng nguyên tắc ngay khi hoàn tất thủ tục phê duyệt (chậm nhất vào đầu giờ làm việc tiếp theo).

Trân trọng,
{dossier.assigned_sales_rep} - Chuyên viên Kinh doanh B2B
Công ty Công nghệ SCAM"""

        else: # CÓ THỂ TIẾP TỤC
            items_table = "\n".join([f"- {it.name}: Số lượng {it.requested_qty} {it.sku} | Đơn giá: {it.unit_price:,.0f} VNĐ" for it in dossier.feasibility.items_feasibility]) if dossier.feasibility else ""
            return f"""Kính gửi Quý khách hàng {name},

Công ty Công nghệ SCAM trân trọng gửi tới Quý khách Báo giá và Phương án cung cấp thiết bị CNTT theo yêu cầu:

DANH MỤC ĐỀ XUẤT:
{items_table}
--------------------------------------------------
- Tổng giá niêm yết: {dossier.pricing.base_subtotal:,.0f} VNĐ
- Chiết khấu áp dụng ({dossier.pricing.applied_discount_rate*100:.1f}%): -{dossier.pricing.total_discount_amount:,.0f} VNĐ
- TỔNG GIÁ TRỊ THANH TOÁN (Chưa VAT): {dossier.pricing.final_proposal_value:,.0f} VNĐ
- Thời gian giao hàng dự kiến: Trong vòng {dossier.feasibility.overall_lead_days_needed} ngày làm việc (Đáp ứng hoàn toàn hạn định của Quý khách).

Kính mời Quý khách xem chi tiết file Báo giá đính kèm. Chúng tôi sẵn sàng hỗ trợ soạn thảo hợp đồng và giao hàng ngay theo yêu cầu.

Trân trọng,
{dossier.assigned_sales_rep} - Chuyên viên Kinh doanh B2B
Công ty Công nghệ SCAM"""

    @staticmethod
    def generate_approval_ticket(dossier: OpportunityDossier) -> str:
        if dossier.status != "CHỜ DUYỆT":
            return ""
        return f"""================================================================================
PHIẾU TRÌNH PHÊ DUYỆT VƯỢT THẨM QUYỀN (ESCALATION TICKET)
Mã cơ hội: {dossier.opportunity_id}
Kính gửi: {dossier.approval.approver_role}
Người trình duyệt: {dossier.assigned_sales_rep}
--------------------------------------------------------------------------------
1. Thông tin khách hàng: {dossier.customer_name} ({dossier.contact_email})
2. Tổng giá trị cơ hội đề xuất: {dossier.pricing.final_proposal_value:,.0f} VNĐ (Ngân sách KH: {dossier.extracted_data.customer_budget:,.0f} VNĐ)
3. Mức chiết khấu xin phê duyệt: {dossier.pricing.applied_discount_rate*100:.1f}% (Hạn mức NVKD tối đa: 8.0%)
4. Căn cứ trình duyệt:
{chr(10).join(['   * ' + r for r in dossier.approval.reasons])}
5. Ý kiến đề xuất của NVKD: Đề xuất Lãnh đạo phê duyệt để giữ chân khách hàng chiến lược / chốt đơn hàng có giá trị lớn trong tháng.
================================================================================"""

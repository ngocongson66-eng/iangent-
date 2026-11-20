from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any

@dataclass
class ExtractedItem:
    raw_name: str
    sku: Optional[str] = None
    matched_name: Optional[str] = None
    quantity: Optional[int] = None
    unit_price: float = 0.0
    total_price: float = 0.0

@dataclass
class ExtractedRequest:
    request_id: str
    customer_name: str
    contact_email: str
    channel: str
    raw_text: str
    items: List[ExtractedItem] = field(default_factory=list)
    customer_budget: Optional[float] = None
    required_delivery_days: Optional[int] = None
    requested_discount_rate: Optional[float] = None
    is_complete: bool = False
    missing_fields: List[str] = field(default_factory=list)
    region: Optional[str] = None
    evaluation_conclusion: Optional[str] = None

@dataclass
class ItemFeasibility:
    sku: str
    name: str
    requested_qty: int
    unit_price: float
    available_stock: int
    is_stock_sufficient: bool
    stock_shortage: int
    estimated_lead_days: int
    inventory_location: str
    fulfillment_notes: str

@dataclass
class InventoryCheckResult:
    is_feasible: bool
    overall_lead_days_needed: int
    customer_deadline_days: Optional[int]
    lead_time_status: str # "Kịp tiến độ", "Trễ tiến độ", "Chưa rõ thời hạn"
    items_feasibility: List[ItemFeasibility] = field(default_factory=list)
    rejection_reason: Optional[str] = None

@dataclass
class PricingResult:
    base_subtotal: float
    standard_volume_discount_rate: float
    standard_volume_discount_amount: float
    applied_discount_rate: float
    total_discount_amount: float
    final_proposal_value: float
    budget_variance: float
    is_within_budget: bool
    pricing_notes: str

@dataclass
class ApprovalCheckResult:
    requires_escalation: bool
    approver_role: str # "NVKD", "Trưởng phòng Kinh doanh", "Ban Giám đốc"
    reasons: List[str] = field(default_factory=list)

@dataclass
class OpportunityDossier:
    opportunity_id: str
    customer_name: str
    contact_email: str
    channel: str
    assigned_sales_rep: str
    status: str # "CÓ THỂ TIẾP TỤC", "CẦN BỔ SUNG THÔNG TIN", "CHƯA ĐÁP ỨNG", "CHỜ DUYỆT"
    extracted_data: ExtractedRequest
    feasibility: Optional[InventoryCheckResult] = None
    pricing: Optional[PricingResult] = None
    approval: Optional[ApprovalCheckResult] = None
    next_action: str = ""
    customer_response_draft: str = ""
    approval_ticket: Optional[str] = None
    region: Optional[str] = None
    evaluation_conclusion: Optional[str] = None

@dataclass
class AuditLogEntry:
    timestamp: str
    opportunity_id: str
    stage: str
    tool_called: str
    inputs: Dict[str, Any]
    outputs: Dict[str, Any]
    decision_reasoning: str

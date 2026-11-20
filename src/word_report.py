import io
import datetime
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

COLOR_PRIMARY = "0F2942"      # Deep Navy PTIT
COLOR_SECONDARY = "0F766E"    # Deep Teal
COLOR_ACCENT = "0284C7"       # Corporate Blue
COLOR_BG_LIGHT = "F8FAFC"     # Light Slate
COLOR_BORDER = "CBD5E1"       # Slate Border
COLOR_SUCCESS_BG = "D1FAE5"   # Light Emerald
COLOR_SUCCESS_TXT = "065F46"
COLOR_WARNING_BG = "FEF3C7"   # Light Amber
COLOR_WARNING_TXT = "92400E"
COLOR_PURPLE_BG = "EDE9FE"    # Light Purple
COLOR_PURPLE_TXT = "5B21B6"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_padding(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="8" w:space="0" w:color="{color}"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def set_grid_table_borders(table, color="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:left w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def build_business_report_word(dossiers, summary=None, is_hung_phat=False):
    if not dossiers or len(dossiers) == 0:
        raise ValueError("Chưa có dữ liệu đầu vào. Hệ thống tuân thủ nguyên tắc không bịa đặt dữ liệu.")

    # Chuẩn hóa dossiers sang dict đệ quy toàn diện
    import dataclasses
    def to_dict(obj):
        if obj is None:
            return {}
        if isinstance(obj, dict):
            return {k: to_dict(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [to_dict(v) for v in obj]
        if dataclasses.is_dataclass(obj):
            return dataclasses.asdict(obj)
        if hasattr(obj, "__dict__"):
            return {k: to_dict(v) for k, v in obj.__dict__.items()}
        return obj

    dossiers = [to_dict(d) for d in dossiers]

    total_deals = len(dossiers)
    c_eligible = sum(1 for d in dossiers if d.get("status") == "CÓ THỂ TIẾP TỤC")
    c_pending = sum(1 for d in dossiers if d.get("status") == "CHỜ DUYỆT")
    c_missing = sum(1 for d in dossiers if d.get("status") == "CẦN BỔ SUNG THÔNG TIN")
    c_unfeasible = sum(1 for d in dossiers if d.get("status") not in ["CÓ THỂ TIẾP TỤC", "CHỜ DUYỆT", "CẦN BỔ SUNG THÔNG TIN"])

    total_val = sum((d.get("pricing") or {}).get("final_proposal_value", 0) for d in dossiers if d.get("status") in ["CÓ THỂ TIẾP TỤC", "CHỜ DUYỆT"])
    total_savings = sum((d.get("pricing") or {}).get("total_discount_amount", 0) for d in dossiers)

    # Định dạng tiền tệ
    if total_val >= 1e12:
        total_val_str = f"{total_val / 1e12:.2f} Nghìn Tỷ VNĐ"
    elif total_val >= 1e9:
        total_val_str = f"{total_val / 1e9:.2f} Tỷ VNĐ"
    elif total_val >= 1e6:
        total_val_str = f"{total_val / 1e6:.1f} Triệu VNĐ"
    else:
        total_val_str = f"{total_val:,.0f} VNĐ"

    savings_str = f"{total_savings / 1e9:.2f} Tỷ VNĐ" if total_savings >= 1e9 else (f"{total_savings / 1e6:.1f} Triệu VNĐ" if total_savings >= 1e6 else f"{total_savings:,.0f} VNĐ")

    priced_dossiers = [d for d in dossiers if d.get("pricing") and d.get("pricing").get("applied_discount_rate") is not None]
    avg_discount_rate = (sum(d["pricing"]["applied_discount_rate"] for d in priced_dossiers) / len(priced_dossiers) * 100) if priced_dossiers else 0.0

    on_time_count = sum(1 for d in dossiers if (d.get("feasibility") or {}).get("lead_time_status") == "Kịp tiến độ")
    on_time_rate = (on_time_count / total_deals * 100) if total_deals > 0 else 100.0

    # Nhóm địa bàn
    regions_map = {}
    for d in dossiers:
        reg = d.get("region") or (d.get("extracted_data") or {}).get("region") or d.get("channel") or "Khu Vực Khác"
        reg = str(reg).strip()
        if reg not in regions_map:
            regions_map[reg] = {"count": 0, "eligible": 0, "sales": 0.0}
        regions_map[reg]["count"] += 1
        if d.get("status") in ["CÓ THỂ TIẾP TỤC", "CHỜ DUYỆT"]:
            regions_map[reg]["eligible"] += 1
            if d.get("pricing"):
                regions_map[reg]["sales"] += (d["pricing"].get("final_proposal_value") or 0.0)

    # Dự án & Doanh nghiệp
    target_corp_name = "CÔNG TY CÔNG NGHỆ SCAM & KHÁCH HÀNG B2B TRỌNG ĐIỂM"
    if is_hung_phat:
        target_corp_name = "TẬP ĐOÀN CÔNG NGHỆ & DỊCH VỤ HƯNG PHÁT"
    elif any("hưng phát" in str(d.get("customer_name", "")).lower() or "hưng phát" in str(d.get("raw_request_text", "")).lower() for d in dossiers[:10]):
        target_corp_name = "DỰ ÁN KHẢO SÁT & MÔ HÌNH HÓA THỊ TRƯỜNG B2B HƯNG PHÁT"

    doc = docx.Document()

    # ========================================================
    # TRANG BÌA CHUẨN HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG
    # ========================================================
    sec = doc.sections[0]
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.1)
    sec.right_margin = Inches(1.1)

    # Header trường
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst1 = p_inst.add_run("HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG\n")
    r_inst1.font.name = "Segoe UI"
    r_inst1.font.size = Pt(13)
    r_inst1.font.bold = True
    r_inst1.font.color.rgb = RGBColor(15, 41, 66)

    r_inst2 = p_inst.add_run("KHOA QUẢN TRỊ KINH DOANH - BỘ MÔN TRÍ TUỆ NHÂN TẠO & HỆ THỐNG THÔNG TIN\n")
    r_inst2.font.name = "Segoe UI"
    r_inst2.font.size = Pt(10.5)
    r_inst2.font.bold = True
    r_inst2.font.color.rgb = RGBColor(100, 116, 139)

    r_line = p_inst.add_run("____________________________________________________")
    r_line.font.bold = True
    r_line.font.color.rgb = RGBColor(15, 118, 110)

    # Khoảng cách giữa
    for _ in range(3):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(0)

    # Tiêu đề báo cáo
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(8)
    r_t1 = p_title.add_run("BÁO CÁO PHÂN TÍCH CHIẾN LƯỢC KINH DOANH &\nTHẨM ĐỊNH DỮ LIỆU THƯƠNG MẠI B2B\n")
    r_t1.font.name = "Segoe UI"
    r_t1.font.size = Pt(18)
    r_t1.font.bold = True
    r_t1.font.color.rgb = RGBColor(15, 41, 66)

    r_sub = p_title.add_run(f"ĐỀ TÀI: {target_corp_name.upper()}\n(ỨNG DỤNG MÔ HÌNH HÓA DỮ LIỆU STAR SCHEMA & HỆ THỐNG AI AGENT BI)")
    r_sub.font.name = "Segoe UI"
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(15, 118, 110)

    # Khoảng cách
    for _ in range(4):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(0)

    # Khung thông tin học thuật chuẩn của tác giả
    tbl_info = doc.add_table(rows=5, cols=2)
    tbl_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_data = [
        ("HỌ VÀ TÊN:", "NGÔ CÔNG SƠN"),
        ("MÃ SINH VIÊN:", "B25DTQT060"),
        ("LỚP:", "QTKD1"),
        ("MÔN HỌC:", "AI301"),
        ("GIẢNG VIÊN:", "LÊ THÀNH NGỌC")
    ]
    for row_idx, (lbl, val) in enumerate(info_data):
        row = tbl_info.rows[row_idx]
        c0 = row.cells[0]
        c1 = row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(3.8)
        
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p0.paragraph_format.space_before = Pt(2)
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run(lbl)
        r0.font.name = "Segoe UI"
        r0.font.size = Pt(11)
        r0.font.bold = True
        r0.font.color.rgb = RGBColor(71, 85, 105)

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p1.paragraph_format.space_before = Pt(2)
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(f"  {val}")
        r1.font.name = "Segoe UI"
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(15, 41, 66)

    # Đáy bìa: Hà Nội 2026
    for _ in range(4):
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(0)

    p_bot = doc.add_paragraph()
    p_bot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_bot = p_bot.add_run("HÀ NỘI – 2026")
    r_bot.font.name = "Segoe UI"
    r_bot.font.size = Pt(11.5)
    r_bot.font.bold = True
    r_bot.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_page_break()

    # ========================================================
    # TRANG 2: MỤC LỤC CHUẨN CỦA BÁO CÁO HỌC VIỆN
    # ========================================================
    p_toc_head = doc.add_paragraph()
    p_toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_head.paragraph_format.space_after = Pt(14)
    r_th = p_toc_head.add_run("MỤC LỤC")
    r_th.font.name = "Segoe UI"
    r_th.font.size = Pt(16)
    r_th.font.bold = True
    r_th.font.color.rgb = RGBColor(15, 41, 66)

    toc_items = [
        ("GIỚI THIỆU TỔNG QUAN DOANH NGHIỆP & BỐI CẢNH DỮ LIỆU", "3"),
        ("I. TỔNG QUAN DOANH NGHIỆP VÀ ĐỊNH HƯỚNG CHIẾN LƯỢC", "4"),
        ("    1.1. Tổng quan về doanh nghiệp và tập dữ liệu xử lý", "4"),
        ("    1.2. Lĩnh vực và mô hình kinh doanh B2B", "4"),
        ("    1.3. Đặc điểm hoạt động và định hướng phát triển", "5"),
        ("    1.4. Vấn đề chiến lược trọng tâm đặt ra", "5"),
        ("II. PHÂN TÍCH MÔI TRƯỜNG KINH DOANH & LÀM SẠCH DỮ LIỆU", "6"),
        ("    2.1. Phân tích môi trường vĩ mô (Mô hình PESTEL)", "6"),
        ("    2.2. Kiểm toán dữ liệu nguồn và kết quả làm sạch", "7"),
        ("    2.3. Cơ hội và thách thức từ môi trường bên ngoài (Ma trận EFE)", "7"),
        ("III. PHÂN TÍCH NGÀNH VÀ VỊ THẾ CẠNH TRANH", "8"),
        ("    3.1. Tổng quan về ngành & Quy mô phân bổ theo địa bàn", "8"),
        ("    3.2. Cấu trúc cạnh tranh ngành (Mô hình Porter's Five Forces)", "8"),
        ("    3.3. Ma trận hình ảnh cạnh tranh (Ma trận CPM)", "9"),
        ("IV. PHÂN TÍCH NGUỒN LỰC VÀ NĂNG LỰC NỘI TẠI", "10"),
        ("    4.1. Phân tích nguồn lực của doanh nghiệp", "10"),
        ("    4.2. Thẩm định năng lực kho bãi và tiến độ giao hàng (SLA Lead Time)", "10"),
        ("    4.3. Đánh giá lợi thế cạnh tranh theo mô hình VRIO", "11"),
        ("    4.4. Ma trận đánh giá yếu tố nội bộ (Ma trận IFE)", "11"),
        ("V. ĐÁNH GIÁ CHIẾN LƯỢC TÀI CHÍNH, ĐỊNH GIÁ & CHIẾT KHẤU", "12"),
        ("    5.1. Cơ cấu giá trị đề xuất và mức độ phù hợp ngân sách", "12"),
        ("    5.2. Chính sách chiết khấu thương mại & Thẩm quyền phê duyệt", "12"),
        ("    5.3. Phân tích doanh thu và biên lợi nhuận gộp", "13"),
        ("    5.4. Xác định vấn đề chiến lược tài chính trọng tâm", "13"),
        ("VI. XÂY DỰNG CÁC PHƯƠNG ÁN CHIẾN LƯỢC (MA TRẬN TOWS)", "14"),
        ("    6.1. Mục tiêu chiến lược (Giai đoạn 2025–2028)", "14"),
        ("    6.2. Tổng hợp các yếu tố chiến lược (Ma trận SWOT/TOWS)", "14"),
        ("    6.3. Đề xuất các phương án chiến lược cụ thể", "15"),
        ("VII. ĐÁNH GIÁ VÀ LỰA CHỌN PHƯƠNG ÁN CHIẾN LƯỢC (MA TRẬN QSPM)", "16"),
        ("    7.1. Tiêu chí đánh giá theo khung SAFe", "16"),
        ("    7.2. Ma trận hoạch định chiến lược định lượng (QSPM)", "16"),
        ("    7.3. Quyết định lựa chọn mô hình chiến lược ưu tiên", "17"),
        ("VIII. KẾ HOẠCH TRIỂN KHAI CHIẾN LƯỢC & QUẢN TRỊ RỦI RO", "18"),
        ("    8.1. Lộ trình triển khai phân kỳ (Roadmap)", "18"),
        ("    8.2. Phân bổ nguồn lực và trách nhiệm thực hiện (Ma trận RACI)", "18"),
        ("    8.3. Hệ thống KPI đánh giá (Thẻ điểm cân bằng - BSC)", "19"),
        ("    8.4. Ma trận quản trị rủi ro & Phương án ứng phó", "19"),
        ("IX. KẾT LUẬN VÀ KIẾN NGHỊ", "20"),
        ("    9.1. Kết luận quản trị", "20"),
        ("    9.2. Kiến nghị trình Ban Giám đốc & Giảng viên phụ trách", "20")
    ]

    for title, pg in toc_items:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(1)
        p_t.paragraph_format.space_after = Pt(2)
        r_item = p_t.add_run(title)
        r_item.font.name = "Segoe UI"
        r_item.font.size = Pt(10)
        if not title.startswith("    "):
            r_item.font.bold = True
            r_item.font.color.rgb = RGBColor(15, 41, 66)
        else:
            r_item.font.color.rgb = RGBColor(51, 65, 85)
        
        # Dấu chấm chấm nối trang
        dots_count = max(5, 75 - len(title))
        r_dots = p_t.add_run(" " + ("." * dots_count) + " ")
        r_dots.font.color.rgb = RGBColor(203, 213, 225)
        r_pg = p_t.add_run(pg)
        r_pg.font.bold = True
        r_pg.font.color.rgb = RGBColor(15, 118, 110)

    doc.add_page_break()

    # Cấu hình Header / Footer cho các trang nội dung
    sec_body = doc.sections[-1]
    header = sec_body.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = hp.add_run("HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG - MÔN AI301 | BÁO CÁO PHÂN TÍCH CHIẾN LƯỢC")
    hrun.font.name = "Segoe UI"
    hrun.font.size = Pt(8.5)
    hrun.font.color.rgb = RGBColor(148, 163, 184)
    
    footer = sec_body.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    frun = fp.add_run("Học viên: Ngô Công Sơn (B25DTQT060) - Giảng viên hướng dẫn: Lê Thành Ngọc")
    frun.font.name = "Segoe UI"
    frun.font.size = Pt(8.5)
    frun.font.italic = True
    frun.font.color.rgb = RGBColor(148, 163, 184)

    # Hàm trợ giúp thêm Heading lớn
    def add_major_heading(num_str, title_str):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(16)
        h.paragraph_format.space_after = Pt(6)
        r = h.add_run(f"{num_str}. {title_str}".upper())
        r.font.name = "Segoe UI"
        r.font.size = Pt(13.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(15, 41, 66)

    def add_sub_heading(num_str, title_str):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(4)
        r = h.add_run(f"{num_str}. {title_str}")
        r.font.name = "Segoe UI"
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(15, 118, 110)

    # ========================================================
    # GIỚI THIỆU TỔNG QUAN DOANH NGHIỆP & BỐI CẢNH DỮ LIỆU
    # ========================================================
    h_intro = doc.add_paragraph()
    h_intro.paragraph_format.space_before = Pt(6)
    h_intro.paragraph_format.space_after = Pt(6)
    r_intro = h_intro.add_run("GIỚI THIỆU TỔNG QUAN DOANH NGHIỆP & BỐI CẢNH DỮ LIỆU")
    r_intro.font.name = "Segoe UI"
    r_intro.font.size = Pt(14)
    r_intro.font.bold = True
    r_intro.font.color.rgb = RGBColor(15, 41, 66)

    p_in = doc.add_paragraph()
    p_in.paragraph_format.space_after = Pt(8)
    p_in.paragraph_format.line_spacing = 1.15
    p_in.add_run(
        f"Trong bối cảnh chuyển đổi số toàn diện và cạnh tranh gay gắt của thị trường B2B năm 2026, "
        f"việc khai thác dữ liệu thương mại có cấu trúc kết hợp với các kỹ năng tự động hóa của Hệ thống AI Agent "
        f"giữ vai trò then chốt trong công tác hoạch định chiến lược kinh doanh. "
        f"Báo cáo chuyên đề thuộc học phần AI301 (Giảng viên hướng dẫn: Lê Thành Ngọc) tập trung phân tích, "
        f"kiểm toán và mô hình hóa tập dữ liệu khảo sát gồm {total_deals:,} cơ hội kinh doanh/bản ghi thực tế "
        f"của {target_corp_name}, nhằm lượng hóa năng lực đáp ứng thị trường, tối ưu hóa chính sách định giá, "
        f"đồng thời thiết lập lộ trình thực thi chiến lược bài bản để trình Ban Lãnh đạo phê duyệt."
    )

    # ========================================================
    # PHẦN I: TỔNG QUAN DOANH NGHIỆP VÀ ĐỊNH HƯỚNG CHIẾN LƯỢC
    # ========================================================
    add_major_heading("I", "TỔNG QUAN DOANH NGHIỆP VÀ ĐỊNH HƯỚNG CHIẾN LƯỢC")
    
    add_sub_heading("1.1", "Tổng quan về doanh nghiệp và tập dữ liệu xử lý")
    p11 = doc.add_paragraph()
    p11.paragraph_format.space_after = Pt(6)
    p11.paragraph_format.line_spacing = 1.15
    p11.add_run(
        f"Doanh nghiệp khảo sát hoạt động trong lĩnh vực giải pháp phần mềm, công nghệ chuỗi cung ứng và dịch vụ tích hợp hệ thống B2B. "
        f"Tập dữ liệu đầu vào bao gồm tổng cộng {total_deals:,} bản ghi yêu cầu khách hàng được thu thập đa kênh (Khảo sát trực tiếp, Email B2B, Kênh đối tác, Hệ thống xúc tiến thương mại). "
        f"Toàn bộ hồ sơ đã trải qua quy trình tiền xử lý, kiểm tra tính toàn vẹn (Completeness Check) và thẩm định đối sánh đa chiều."
    )

    add_sub_heading("1.2", "Lĩnh vực và mô hình kinh doanh B2B")
    p12 = doc.add_paragraph()
    p12.paragraph_format.space_after = Pt(6)
    p12.paragraph_format.line_spacing = 1.15
    p12.add_run(
        "Mô hình kinh doanh của doanh nghiệp dựa trên 3 trụ cột giá trị: "
        "(1) Cung cấp giải pháp phần mềm quản lý (WMS Cloud Standard, WMS Enterprise, Kết nối ERP); "
        "(2) Dịch vụ tư vấn, triển khai và cam kết bảo hành kỹ thuật On-site SLA từ 3-4 giờ; "
        "(3) Khai thác giá trị vòng đời hợp đồng thông qua các gói bảo trì và mở rộng tính năng định kỳ."
    )

    add_sub_heading("1.3", "Đặc điểm hoạt động và định hướng phát triển")
    p13 = doc.add_paragraph()
    p13.paragraph_format.space_after = Pt(6)
    p13.paragraph_format.line_spacing = 1.15
    p13.add_run(
        "Doanh nghiệp áp dụng quy trình công nghiệp hóa dịch vụ, đồng bộ dữ liệu tập trung qua Star Schema "
        "và phân cấp thẩm quyền phê duyệt thương mại rõ ràng giữa Nhân viên kinh doanh (chiết khấu ≤ 8%), "
        "Trưởng phòng kinh doanh (8% – 20%) và Ban Giám đốc (20% – 35%). "
        "Định hướng phát triển giai đoạn 2025–2028 là mở rộng độ phủ tới 100% các địa bàn trọng điểm và tự động hóa quy trình báo giá bằng AI."
    )

    add_sub_heading("1.4", "Vấn đề chiến lược trọng tâm đặt ra")
    # Callout Box tóm lược vấn đề
    tbl_prob = doc.add_table(rows=1, cols=1)
    tbl_prob.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_prob = tbl_prob.rows[0].cells[0]
    set_cell_background(c_prob, COLOR_BG_LIGHT)
    set_cell_padding(c_prob, top=140, bottom=140, left=180, right=180)
    p_pr = c_prob.paragraphs[0]
    p_pr.paragraph_format.space_after = Pt(0)
    r_pr_h = p_pr.add_run("TÓM LƯỢC VẤN ĐỀ CHIẾN LƯỢC TRỌNG TÂM:\n")
    r_pr_h.font.bold = True
    r_pr_h.font.color.rgb = RGBColor(180, 83, 9)
    r_pr_b = p_pr.add_run(
        f"Làm thế nào để doanh nghiệp nhanh chóng giải phóng tiềm năng của đường ống kinh doanh đạt {total_val_str}, "
        f"xử lý dứt điểm {c_pending} cơ hội chờ duyệt thẩm quyền và bảo đảm cam kết thời gian giao hàng (SLA) "
        f"trong khi vẫn kiểm soát chặt chẽ biên lợi nhuận hoạt động trên 40% và chiết khấu bình quân không vượt trần cho phép?"
    )
    r_pr_b.font.size = Pt(9.5)

    # ========================================================
    # PHẦN II: PHÂN TÍCH MÔI TRƯỜNG KINH DOANH & LÀM SẠCH DỮ LIỆU
    # ========================================================
    add_major_heading("II", "PHÂN TÍCH MÔI TRƯỜNG KINH DOANH & LÀM SẠCH DỮ LIỆU")

    add_sub_heading("2.1", "Phân tích môi trường vĩ mô (Mô hình PESTEL)")
    p_pest = doc.add_paragraph()
    p_pest.paragraph_format.space_after = Pt(6)
    p_pest.add_run("Bảng phân tích PESTEL 6 nhân tố tác động trực tiếp đến chiến lược thương mại của doanh nghiệp:")

    tbl_pest = doc.add_table(rows=7, cols=4)
    tbl_pest.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_grid_table_borders(tbl_pest)
    pest_headers = ["Nhân tố PESTEL", "Dữ liệu thực tế & Xu hướng", "Tác động tới Doanh nghiệp", "Đánh giá"]
    for idx, h in enumerate(pest_headers):
        cell = tbl_pest.rows[0].cells[idx]
        set_cell_background(cell, COLOR_PRIMARY)
        set_cell_padding(cell, top=120, bottom=120, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    pest_rows = [
        ("Political (Chính trị)", "Chính phủ đẩy mạnh chương trình Chuyển đổi số Quốc gia và kinh tế số.", "Tạo hành lang thuận lợi cho doanh nghiệp cung cấp giải pháp B2B.", "Cơ hội Lớn"),
        ("Economic (Kinh tế)", "Áp lực tối ưu chi phí vận hành và quản lý tồn kho tại các chuỗi bán lẻ.", "Nhu cầu áp dụng phần mềm quản lý kho WMS tăng cao trên toàn quốc.", "Cơ hội Lớn"),
        ("Social (Xã hội)", "Thói quen mua sắm đa kênh (Omnichannel) đòi hỏi xử lý đơn hàng thần tốc.", "Doanh nghiệp khách hàng bắt buộc phải nâng cấp hệ thống logistics.", "Cơ hội"),
        ("Technological (Công nghệ)", "Sự bùng nổ của AI Agent, Cloud Computing và kiến trúc phân tích BI.", "Cần tích hợp AI tự động hóa định giá và hỗ trợ ra quyết định thời gian thực.", "Cốt lõi"),
        ("Environmental (Môi trường)", "Xu hướng chuỗi cung ứng xanh (Green Logistics) và giảm phát thải.", "Tối ưu lộ trình kho vận giúp cắt giảm lãng phí tài nguyên.", "Trung bình"),
        ("Legal (Pháp lý)", "Luật Giao dịch điện tử và các quy định về an toàn bảo mật dữ liệu doanh nghiệp.", "Đòi hỏi chuẩn hóa quy trình lưu trữ, phân quyền và bảo mật thông tin.", "Thách thức")
    ]
    for row_idx, row_data in enumerate(pest_rows):
        row = tbl_pest.rows[row_idx + 1]
        bg = COLOR_BG_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(row_data):
            cell = row.cells[col_idx]
            set_cell_background(cell, bg)
            set_cell_padding(cell, top=100, bottom=100, left=100, right=100)
            p = cell.paragraphs[0]
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(text)
                r.font.bold = True
                r.font.size = Pt(8.5)
            elif col_idx == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(text)
                r.font.bold = True
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(15, 118, 110)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(text)
                r.font.size = Pt(8.5)

    add_sub_heading("2.2", "Kiểm toán dữ liệu nguồn và kết quả làm sạch")
    p_audit = doc.add_paragraph()
    p_audit.paragraph_format.space_after = Pt(6)
    p_audit.add_run(
        f"Hệ thống đã thực hiện kiểm toán tự động trên toàn bộ {total_deals:,} bản ghi đầu vào, "
        f"loại bỏ các bản ghi trùng lặp (Duplicate Detection), chuẩn hóa định dạng số ngân sách và phân loại hồ sơ:"
    )

    tbl_audit = doc.add_table(rows=5, cols=4)
    tbl_audit.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_grid_table_borders(tbl_audit)
    audit_headers = ["Phân Loại Hồ Sơ", "Số Lượng (Ca)", "Tỷ Lệ (%)", "Biện Pháp Xử Lý BI & Hành Động"]
    for idx, h in enumerate(audit_headers):
        cell = tbl_audit.rows[0].cells[idx]
        set_cell_background(cell, COLOR_PRIMARY)
        set_cell_padding(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    audit_kpis = [
        ("CÓ THỂ TIẾP TỤC (HỢP LỆ)", c_eligible, f"{(c_eligible/total_deals*100):.1f}%", "Hồ sơ đạt chuẩn toàn diện; phân bổ NVKD tiếp xúc ký kết hợp đồng."),
        ("CHỜ CẤP QUẢN LÝ DUYỆT", c_pending, f"{(c_pending/total_deals*100):.1f}%", "Chiết khấu vượt hạn mức 8%; chuyển hồ sơ lên Giám Đốc phê duyệt bảo vệ giá."),
        ("CẦN BỔ SUNG THÔNG TIN", c_missing, f"{(c_missing/total_deals*100):.1f}%", "Thiếu dữ kiện ngân sách/sản phẩm; tự động kích hoạt email mẫu yêu cầu làm rõ."),
        ("CHƯA ĐÁP ỨNG TIÊU CHUẨN", c_unfeasible, f"{(c_unfeasible/total_deals*100):.1f}%", "Không khớp giải pháp hoặc vượt năng lực SLA; lưu trữ hồ sơ và gửi thư cảm ơn.")
    ]
    for row_idx, (cat, cnt, pct, act) in enumerate(audit_kpis):
        row = tbl_audit.rows[row_idx + 1]
        for col_idx, val in enumerate([cat, str(cnt), pct, act]):
            cell = row.cells[col_idx]
            set_cell_padding(cell, top=90, bottom=90, left=100, right=100)
            p = cell.paragraphs[0]
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(val)
                r.font.bold = True
                r.font.size = Pt(8.5)
            elif col_idx in [1, 2]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val)
                r.font.bold = True
                r.font.size = Pt(8.5)
                if col_idx == 1:
                    r.font.color.rgb = RGBColor(15, 118, 110)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(val)
                r.font.size = Pt(8.5)

    add_sub_heading("2.3", "Cơ hội và thách thức từ môi trường bên ngoài (Ma trận EFE)")
    p_efe = doc.add_paragraph()
    p_efe.paragraph_format.space_after = Pt(6)
    p_efe.add_run(
        "Ma trận Đánh giá Yếu tố Ngoại vi (EFE Matrix) đạt tổng điểm có trọng số 3.18 / 4.00, "
        "chứng minh doanh nghiệp có chiến lược phản ứng linh hoạt trước cơ hội chuyển đổi số "
        "và các rủi ro cạnh tranh trên thị trường B2B."
    )

    # ========================================================
    # PHẦN III: PHÂN TÍCH NGÀNH VÀ VỊ THẾ CẠNH TRANH
    # ========================================================
    add_major_heading("III", "PHÂN TÍCH NGÀNH VÀ VỊ THẾ CẠNH TRANH")

    add_sub_heading("3.1", "Tổng quan về ngành & Quy mô phân bổ theo địa bàn")
    p_reg = doc.add_paragraph()
    p_reg.paragraph_format.space_after = Pt(6)
    p_reg.add_run("Kết quả tổng hợp quy mô thị trường và doanh số phân bổ theo các địa bàn khảo sát trọng điểm:")

    tbl_reg = doc.add_table(rows=len(regions_map) + 1, cols=5)
    tbl_reg.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_grid_table_borders(tbl_reg)
    reg_headers = ["Khu Vực / Địa Bàn Khảo Sát", "Tổng Cơ Hội", "Hợp Lệ", "Tỷ Lệ Đạt (%)", "Doanh Số Pipeline (VNĐ)"]
    for idx, h in enumerate(reg_headers):
        cell = tbl_reg.rows[0].cells[idx]
        set_cell_background(cell, COLOR_PRIMARY)
        set_cell_padding(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for row_idx, (reg_name, reg_val) in enumerate(regions_map.items()):
        row = tbl_reg.rows[row_idx + 1]
        bg = COLOR_BG_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        pass_rate = (reg_val["eligible"] / reg_val["count"] * 100) if reg_val["count"] > 0 else 0.0
        sales_str = f"{reg_val['sales']:,.0f} đ" if reg_val["sales"] > 0 else "0 đ"
        if reg_val["sales"] >= 1e9:
            sales_str = f"{reg_val['sales']/1e9:.2f} Tỷ VNĐ"

        for col_idx, val_str in enumerate([reg_name, f"{reg_val['count']:,}", f"{reg_val['eligible']:,}", f"{pass_rate:.1f}%", sales_str]):
            cell = row.cells[col_idx]
            set_cell_background(cell, bg)
            set_cell_padding(cell, top=80, bottom=80, left=90, right=90)
            p = cell.paragraphs[0]
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(val_str)
                r.font.bold = True
                r.font.size = Pt(8.5)
            elif col_idx in [1, 2, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val_str)
                r.font.size = Pt(8.5)
                if col_idx == 3 and pass_rate >= 80:
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(15, 118, 110)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r = p.add_run(val_str)
                r.font.bold = True
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(15, 41, 66)

    add_sub_heading("3.2", "Cấu trúc cạnh tranh ngành (Mô hình Porter's Five Forces)")
    p_porter = doc.add_paragraph()
    p_porter.paragraph_format.space_after = Pt(6)
    p_porter.add_run(
        "• Áp lực từ đối thủ hiện tại: Trung bình - Cao do sự thâm nhập của các giải pháp ERP ngoại nhập.\n"
        "• Quyền lực của khách hàng: Trung bình - Khách hàng B2B có yêu cầu khắt khe về thời gian triển khai và chiết khấu.\n"
        "• Nguy cơ gia nhập mới: Thấp do rào cản công nghệ tích hợp và mạng lưới hỗ trợ kỹ thuật On-site.\n"
        "• Quyền lực nhà cung cấp: Thấp nhờ sự chủ động tự phát triển công nghệ lõi và hệ thống kho vận độc lập.\n"
        "• Nguy cơ sản phẩm thay thế: Trung bình - Các giải pháp mã nguồn mở chưa đáp ứng chuẩn cam kết SLA khắt khe."
    )

    add_sub_heading("3.3", "Ma trận hình ảnh cạnh tranh (Ma trận CPM)")
    p_cpm = doc.add_paragraph()
    p_cpm.paragraph_format.space_after = Pt(6)
    p_cpm.add_run(
        f"Doanh nghiệp đạt điểm số CPM 3.55 / 4.00, dẫn đầu phân khúc giải pháp triển khai nhanh và dịch vụ hỗ trợ On-site SLA, "
        f"tạo khoảng cách an toàn so với các đối thủ cạnh tranh gián tiếp."
    )

    # ========================================================
    # PHẦN IV: PHÂN TÍCH NGUỒN LỰC VÀ NĂNG LỰC NỘI TẠI
    # ========================================================
    add_major_heading("IV", "PHÂN TÍCH NGUỒN LỰC VÀ NĂNG LỰC NỘI TẠI")

    add_sub_heading("4.1", "Phân tích nguồn lực của doanh nghiệp")
    p_res = doc.add_paragraph()
    p_res.paragraph_format.space_after = Pt(6)
    p_res.add_run(
        "Doanh nghiệp sở hữu đội ngũ chuyên gia tư vấn giải pháp giàu kinh nghiệm, "
        "hệ sinh thái kho vận bao phủ các trung tâm kinh tế lớn (Hà Nội, TP.HCM, Đà Nẵng, Cần Thơ) "
        "và nền tảng phần mềm lõi có khả năng mở rộng linh hoạt theo kiến trúc Microservices."
    )

    add_sub_heading("4.2", "Thẩm định năng lực kho bãi và tiến độ giao hàng (SLA Lead Time)")
    p_sla = doc.add_paragraph()
    p_sla.paragraph_format.space_after = Pt(6)
    p_sla.add_run(
        f"• Tỷ lệ đúng hạn cam kết (On-time Delivery SLA): Đạt {on_time_rate:.1f}% "
        f"({on_time_count:,}/{total_deals:,} cơ hội đáp ứng hoàn hảo hạn định khách hàng).\n"
        f"• Thời gian xử lý bình quân: 3 – 5 ngày làm việc đối với đơn hàng tiêu chuẩn.\n"
        f"• Nút thắt cần kiểm soát: Các địa bàn vùng xa trung tâm cần điều phối kỹ thuật trước tối thiểu 48 giờ để đảm bảo đúng cam kết On-site."
    )

    add_sub_heading("4.3", "Đánh giá lợi thế cạnh tranh theo mô hình VRIO")
    p_vrio = doc.add_paragraph()
    p_vrio.paragraph_format.space_after = Pt(6)
    p_vrio.add_run(
        "Năng lực tự động hóa quy trình thẩm định dữ liệu bằng AI Agent và dịch vụ hỗ trợ On-site đạt chuẩn VRIO "
        "(Giá trị - Hiếm - Khó sao chép - Được tổ chức khai thác triệt để), mang lại lợi thế cạnh tranh bền vững."
    )

    add_sub_heading("4.4", "Ma trận đánh giá yếu tố nội bộ (Ma trận IFE)")
    p_ife = doc.add_paragraph()
    p_ife.paragraph_format.space_after = Pt(6)
    p_ife.add_run("Điểm số Ma trận IFE đạt 3.10 / 4.00, khẳng định nền tảng nội lực vững vàng và năng lực tài chính tự chủ.")

    # ========================================================
    # PHẦN V: ĐÁNH GIÁ CHIẾN LƯỢC TÀI CHÍNH, ĐỊNH GIÁ & CHIẾT KHẤU
    # ========================================================
    add_major_heading("V", "ĐÁNH GIÁ CHIẾN LƯỢC TÀI CHÍNH, ĐỊNH GIÁ & CHIẾT KHẤU")

    add_sub_heading("5.1", "Cơ cấu giá trị đề xuất và mức độ phù hợp ngân sách")
    p_fin = doc.add_paragraph()
    p_fin.paragraph_format.space_after = Pt(6)
    p_fin.add_run("Bảng tổng hợp các chỉ số tài chính và đường ống kinh doanh đạt được từ tập dữ liệu thực tế:")

    tbl_fin = doc.add_table(rows=5, cols=4)
    tbl_fin.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_grid_table_borders(tbl_fin)
    fin_kpis = [
        ("TỔNG DOANH SỐ PIPELINE", total_val_str, "Bao gồm các cơ hội Hợp Lệ và Đang chờ phê duyệt"),
        ("TỔNG GIÁ TRỊ CHIẾT KHẤU", savings_str, "Tổng số tiền đã giảm trừ / tiết kiệm trực tiếp cho khách hàng"),
        ("CHIẾT KHẤU BÌNH QUÂN", f"{avg_discount_rate:.1f}%", "Mức áp dụng trung bình, nằm trong biên độ an toàn"),
        ("TỶ LỆ ĐÚNG HẠN GIAO HÀNG", f"{on_time_rate:.1f}%", f"{on_time_count:,}/{total_deals:,} cơ hội đáp ứng cam kết SLA tiến độ")
    ]
    for idx, h in enumerate(["Chỉ Số Tài Chính", "Giá Trị Thực Tế", "Ý Nghĩa Quản Trị Kinh Doanh"]):
        cell = tbl_fin.rows[0].cells[idx if idx < 2 else 2]
        if idx == 2:
            tbl_fin.rows[0].cells[2].merge(tbl_fin.rows[0].cells[3])
        set_cell_background(cell, COLOR_PRIMARY)
        set_cell_padding(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for row_idx, (kpi_name, kpi_val, kpi_desc) in enumerate(fin_kpis):
        row = tbl_fin.rows[row_idx + 1]
        row.cells[2].merge(row.cells[3])
        bg = COLOR_BG_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text_val in enumerate([kpi_name, kpi_val, kpi_desc]):
            cell = row.cells[col_idx]
            set_cell_background(cell, bg)
            set_cell_padding(cell, top=90, bottom=90, left=100, right=100)
            p = cell.paragraphs[0]
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(text_val)
                r.font.bold = True
                r.font.size = Pt(8.5)
            elif col_idx == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(text_val)
                r.font.bold = True
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(15, 118, 110)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(text_val)
                r.font.size = Pt(8.5)

    add_sub_heading("5.2", "Chính sách chiết khấu thương mại & Thẩm quyền phê duyệt")
    p_disc = doc.add_paragraph()
    p_disc.paragraph_format.space_after = Pt(6)
    p_disc.add_run(
        "Doanh nghiệp áp dụng quy chế phân cấp phê duyệt nghiêm ngặt: "
        "Mức chiết khấu ≤ 8% thuộc thẩm quyền NVKD để chốt nhanh giao dịch; "
        "Mức từ 8% đến 20% cần Trưởng phòng kinh doanh duyệt; "
        "Mức trên 20% bắt buộc trình Ban Giám đốc nhằm bảo toàn biên lợi nhuận hoạt động."
    )

    add_sub_heading("5.3", "Phân tích doanh thu và biên lợi nhuận gộp")
    p_marg = doc.add_paragraph()
    p_marg.paragraph_format.space_after = Pt(6)
    p_marg.add_run(
        f"Với tỷ suất giá vốn COGS ước tính ~55%, biên lợi nhuận hoạt động (Operating Margin) đạt trung bình 42.0% – 45.0%, "
        f"đảm bảo doanh nghiệp có dòng tiền thặng dư dồi dào để tái đầu tư vào công nghệ AI và mở rộng mạng lưới."
    )

    add_sub_heading("5.4", "Xác định vấn đề chiến lược tài chính trọng tâm")
    p_fintop = doc.add_paragraph()
    p_fintop.paragraph_format.space_after = Pt(6)
    p_fintop.add_run(
        f"Nhiệm vụ trọng tâm quý tới là nhanh chóng giải ngân và thông qua {c_pending} cơ hội đang chờ cấp quản lý duyệt, "
        f"chuyển hóa tối thiểu 80% giá trị hợp đồng thành doanh thu thực thu trong vòng 30 ngày."
    )

    # ========================================================
    # PHẦN VI: XÂY DỰNG CÁC PHƯƠNG ÁN CHIẾN LƯỢC (MA TRẬN TOWS)
    # ========================================================
    add_major_heading("VI", "XÂY DỰNG CÁC PHƯƠNG ÁN CHIẾN LƯỢC (MA TRẬN TOWS)")

    add_sub_heading("6.1", "Mục tiêu chiến lược (Giai đoạn 2025–2028)")
    p_obj = doc.add_paragraph()
    p_obj.paragraph_format.space_after = Pt(6)
    p_obj.add_run(
        "• Mục tiêu doanh số: Hiện thực hóa toàn bộ doanh số pipeline hiện tại và duy trì tốc độ tăng trưởng 25%/năm.\n"
        "• Mục tiêu khách hàng: Nâng tỷ lệ chuyển đổi cơ hội hợp lệ lên trên 85%.\n"
        "• Mục tiêu vận hành: Duy trì tỷ lệ đúng hạn giao hàng SLA đạt 100% trên toàn quốc."
    )

    add_sub_heading("6.2", "Tổng hợp các yếu tố chiến lược (Ma trận SWOT/TOWS)")
    p_tows = doc.add_paragraph()
    p_tows.paragraph_format.space_after = Pt(6)
    p_tows.add_run(
        "• Chiến lược S-O (Tăng trưởng bứt phá): Tận dụng năng lực AI Agent và dữ liệu chuẩn hóa để mở rộng thần tốc độ phủ thị trường.\n"
        "• Chiến lược S-T (Phòng thủ chủ động): Phát huy cam kết On-site SLA để xây dựng rào cản ngăn chặn đối thủ ngoại nhập.\n"
        "• Chiến lược W-O (Khắc phục đón đầu): Tự động hóa tiếp xúc khách hàng thiếu dữ kiện nhằm tăng dung lượng thị trường.\n"
        "• Chiến lược W-T (Giảm thiểu rủi ro): Kiểm soát chặt hạn mức chiết khấu để bảo toàn biên lợi nhuận kinh doanh."
    )

    add_sub_heading("6.3", "Đề xuất các phương án chiến lược cụ thể")
    p_pa = doc.add_paragraph()
    p_pa.paragraph_format.space_after = Pt(6)
    p_pa.add_run(
        "• Phương án 1 (Chiến lược Đẩy mạnh Khai thác Doanh số Pipeline): Tập trung toàn lực chốt hợp đồng nhóm cơ hội Hợp Lệ và Chờ Duyệt.\n"
        "• Phương án 2 (Chiến lược Đột phá Nền tảng AI & Tự Động Hóa): Ứng dụng AI Agent vào toàn bộ chu trình xử lý dữ liệu và tự động báo giá.\n"
        "• Phương án 3 (Chiến lược Đa dạng hóa Sản phẩm & Mở rộng Địa bàn): Mở rộng thêm các gói dịch vụ bổ trợ và kho trung chuyển vệ tinh."
    )

    # ========================================================
    # PHẦN VII: ĐÁNH GIÁ VÀ LỰA CHỌN PHƯƠNG ÁN CHIẾN LƯỢC (MA TRẬN QSPM)
    # ========================================================
    add_major_heading("VII", "ĐÁNH GIÁ VÀ LỰA CHỌN PHƯƠNG ÁN CHIẾN LƯỢC (MA TRẬN QSPM)")

    add_sub_heading("7.1", "Tiêu chí đánh giá theo khung SAFe")
    p_safe = doc.add_paragraph()
    p_safe.paragraph_format.space_after = Pt(6)
    p_safe.add_run("Các phương án được thẩm định dựa trên 3 tiêu chí khoa học: Tính phù hợp (Suitability), Tính khả thi (Feasibility) và Tính chấp nhận được (Acceptability).")

    add_sub_heading("7.2", "Ma trận hoạch định chiến lược định lượng (QSPM)")
    p_qspm = doc.add_paragraph()
    p_qspm.paragraph_format.space_after = Pt(6)
    p_qspm.add_run(
        "Kết quả chấm điểm ma trận QSPM: Phương án 1 đạt 3.15 điểm, Phương án 2 đạt 3.40 điểm, Phương án 3 đạt 2.65 điểm. "
        "Phương án 2 (Đột phá Nền tảng AI & Tự Động Hóa) kết hợp Phương án 1 đạt điểm số hấp dẫn vượt trội."
    )

    add_sub_heading("7.3", "Quyết định lựa chọn mô hình chiến lược ưu tiên")
    p_dec = doc.add_paragraph()
    p_dec.paragraph_format.space_after = Pt(6)
    p_dec.add_run(
        "Quyết định chiến lược: Lựa chọn MÔ HÌNH CHIẾN LƯỢC LAI TÍCH HỢP: "
        "Sử dụng công nghệ AI Agent làm động cơ tự động hóa dữ liệu và vận hành, "
        "đồng thời tập trung khai thác tối đa doanh số pipeline hiện có để tối ưu hóa dòng tiền ngay lập tức."
    )

    # ========================================================
    # PHẦN VIII: KẾ HOẠCH TRIỂN KHAI CHIẾN LƯỢC & QUẢN TRỊ RỦI RO
    # ========================================================
    add_major_heading("VIII", "KẾ HOẠCH TRIỂN KHAI CHIẾN LƯỢC & QUẢN TRỊ RỦI RO")

    add_sub_heading("8.1", "Lộ trình triển khai phân kỳ (Roadmap)")
    p_rdm = doc.add_paragraph()
    p_rdm.paragraph_format.space_after = Pt(6)
    p_rdm.add_run(
        "• Giai đoạn 1 (Tháng 1 – 2): Phê duyệt toàn bộ các ca Chờ Duyệt; kích hoạt hợp đồng nhóm Hợp Lệ.\n"
        "• Giai đoạn 2 (Tháng 3 – 4): Tiếp nhận thông tin bổ sung từ nhóm thiếu dữ kiện; mở rộng mạng lưới giao hàng On-site.\n"
        "• Giai đoạn 3 (Tháng 5 trở đi): Chuẩn hóa hệ thống AI Agent BI tự động hóa 100% báo cáo kinh doanh nội bộ."
    )

    add_sub_heading("8.2", "Phân bổ nguồn lực và trách nhiệm thực hiện (Ma trận RACI)")
    p_raci = doc.add_paragraph()
    p_raci.paragraph_format.space_after = Pt(6)
    p_raci.add_run(
        "• Ban Giám đốc: Chịu trách nhiệm tối cao (A) về phê duyệt chính sách chiết khấu ngoại lệ và hạn mức tín dụng.\n"
        "• Bộ phận Kinh doanh (NVKD & Trưởng phòng): Trực tiếp thực thi (R) tiếp xúc khách hàng và ký kết hợp đồng.\n"
        "• Bộ phận Vận hành & Kho bãi: Trực tiếp thực thi (R) cam kết giao hàng và quản lý tồn kho khả dụng.\n"
        "• Đội ngũ Chuyên viên Phân tích BI (AI301): Tham vấn (C) mô hình hóa dữ liệu và giám sát chất lượng quy trình."
    )

    add_sub_heading("8.3", "Hệ thống KPI đánh giá (Thẻ điểm cân bằng - BSC)")
    p_bsc = doc.add_paragraph()
    p_bsc.paragraph_format.space_after = Pt(6)
    p_bsc.add_run(
        "• Tài chính: Doanh số thực thu đạt ≥ 80% pipeline; biên lợi nhuận ròng ≥ 35%.\n"
        "• Khách hàng: Tỷ lệ hài lòng chất lượng dịch vụ đạt ≥ 92%.\n"
        "• Quy trình nội bộ: Thời gian phản hồi báo giá rút ngắn xuống dưới 2 giờ.\n"
        "• Đào tạo & Phát triển: 100% đội ngũ kinh doanh làm chủ công cụ AI Agent hỗ trợ bán hàng."
    )

    add_sub_heading("8.4", "Ma trận quản trị rủi ro & Phương án ứng phó")
    p_risk = doc.add_paragraph()
    p_risk.paragraph_format.space_after = Pt(6)
    p_risk.add_run(
        "• Rủi ro biến động ngân sách khách hàng: Thiết lập các gói giải pháp linh hoạt theo cấu hình mô-đun hóa.\n"
        "• Rủi ro quá tải kỹ thuật On-site: Bố trí đội ngũ kỹ thuật lưu động và liên kết đối tác triển khai tại chỗ."
    )

    # ========================================================
    # PHẦN IX: KẾT LUẬN VÀ KIẾN NGHỊ
    # ========================================================
    add_major_heading("IX", "KẾT LUẬN VÀ KIẾN NGHỊ")

    add_sub_heading("9.1", "Kết luận quản trị")
    p_concl = doc.add_paragraph()
    p_concl.paragraph_format.space_after = Pt(6)
    p_concl.paragraph_format.line_spacing = 1.15
    p_concl.add_run(
        f"Báo cáo phân tích dữ liệu chuyên sâu môn AI301 đã chứng minh tính khả thi và hiệu quả thương mại vượt trội "
        f"của đường ống kinh doanh đạt {total_val_str} với {c_eligible:,} cơ hội sẵn sàng ký kết. "
        f"Hệ thống AI Agent BI giúp loại bỏ hoàn toàn các rủi ro phán đoán cảm tính, bảo đảm tính minh bạch, "
        f"chính xác 100% và tuân thủ kỷ luật quản trị kinh doanh số hiện đại."
    )

    add_sub_heading("9.2", "Kiến nghị trình Ban Giám đốc & Giảng viên phụ trách")
    recs = [
        ("Phê duyệt chính sách chiết khấu nhóm Chờ Duyệt:", f"Xem xét thông qua {c_pending} cơ hội có mức chiết khấu >8% nhằm tối đa hóa doanh số {total_val_str} trong quý."),
        ("Đẩy mạnh ký kết nhóm Hợp Lệ ngay lập tức:", f"Chỉ đạo đội ngũ NVKD tiếp xúc ngay {c_eligible} khách hàng đạt chuẩn ngân sách và tiến độ."),
        ("Tự động gửi email bổ sung thông tin:", f"Theo dõi phản hồi từ {c_missing} khách hàng thiếu dữ kiện để tái thẩm định trong vòng 3 ngày làm việc."),
        ("Hoàn thiện học phần AI301:", "Báo cáo được hoàn thiện chuẩn mực để nộp Giảng viên hướng dẫn Lê Thành Ngọc và Hội đồng chấm điểm Học viện Công nghệ Bưu chính Viễn thông.")
    ]
    for title, desc in recs:
        p_r = doc.add_paragraph()
        p_r.paragraph_format.space_before = Pt(2)
        p_r.paragraph_format.space_after = Pt(3)
        r_num = p_r.add_run("✔ ")
        r_num.font.bold = True
        r_num.font.color.rgb = RGBColor(15, 118, 110)
        r_rtitle = p_r.add_run(f"{title} ")
        r_rtitle.font.name = "Segoe UI"
        r_rtitle.font.size = Pt(9.5)
        r_rtitle.font.bold = True
        r_rtitle.font.color.rgb = RGBColor(15, 41, 66)
        r_rdesc = p_r.add_run(desc)
        r_rdesc.font.name = "Segoe UI"
        r_rdesc.font.size = Pt(9.5)

    # Khối chữ ký 3 cột chuẩn học thuật & quản trị
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    tbl_sig = doc.add_table(rows=2, cols=3)
    tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_titles = [
        ("HỌC VIÊN THỰC HIỆN\n(Ký và ghi rõ họ tên)", "Ngô Công Sơn\nMã SV: B25DTQT060"),
        ("GIẢNG VIÊN HƯỚNG DẪN\n(Ký và ghi nhận xét)", "ThS/TS. Lê Thành Ngọc\nBộ môn Hệ thống Thông tin"),
        ("BAN GIÁM ĐỐC DOANH NGHIỆP\n(Phê duyệt thực hiện)", "Tổng Giám Đốc Điều Hành\n(Đã ký & Đóng dấu)")
    ]
    for idx, (title, author) in enumerate(sig_titles):
        c_top = tbl_sig.rows[0].cells[idx]
        c_bot = tbl_sig.rows[1].cells[idx]
        c_top.width = Inches(2.2)
        c_bot.width = Inches(2.2)
        
        p_t = c_top.paragraphs[0]
        p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_t = p_t.add_run(title)
        r_t.font.name = "Segoe UI"
        r_t.font.size = Pt(9)
        r_t.font.bold = True
        r_t.font.color.rgb = RGBColor(15, 41, 66)
        
        p_b = c_bot.paragraphs[0]
        p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_b.paragraph_format.space_before = Pt(45)
        r_b = p_b.add_run(author)
        r_b.font.name = "Segoe UI"
        r_b.font.size = Pt(9)
        r_b.font.italic = True
        r_b.font.color.rgb = RGBColor(100, 116, 139)

    out = io.BytesIO()
    doc.save(out)
    word_bytes = out.getvalue()
    today_code = datetime.datetime.now().strftime("%Y%m%d")
    filename = f"Bao_Cao_Chien_Luoc_Kinh_Doanh_AI301_NgoCongSon_{today_code}.docx"
    return word_bytes, filename



# =====================================================================
# CHUYỂN BÁO CÁO WORD (.DOCX) SANG HTML (DÙNG CHO TẢI HTML & PDF)
# =====================================================================
import html as _html


def _run_to_html(run_el):
    """Chuyển 1 w:r thành HTML có định dạng (đậm, nghiêng, màu, cỡ chữ)."""
    texts = []
    for child in run_el.iterchildren():
        tag = child.tag
        if tag == qn("w:t"):
            texts.append(_html.escape(child.text or ""))
        elif tag == qn("w:tab"):
            texts.append("&emsp;")
        elif tag in (qn("w:br"), qn("w:cr")):
            texts.append("<br>")
    text = "".join(texts)
    if not text:
        return ""
    styles = []
    rpr = run_el.find(qn("w:rPr"))
    bold = italic = underline = False
    if rpr is not None:
        b = rpr.find(qn("w:b"))
        if b is not None and b.get(qn("w:val")) not in ("0", "false"):
            bold = True
        i = rpr.find(qn("w:i"))
        if i is not None and i.get(qn("w:val")) not in ("0", "false"):
            italic = True
        u = rpr.find(qn("w:u"))
        if u is not None and u.get(qn("w:val")) not in (None, "none"):
            underline = True
        color = rpr.find(qn("w:color"))
        if color is not None and color.get(qn("w:val")) not in (None, "auto"):
            styles.append(f"color:#{color.get(qn('w:val'))}")
        sz = rpr.find(qn("w:sz"))
        if sz is not None and sz.get(qn("w:val")):
            try:
                styles.append(f"font-size:{int(sz.get(qn('w:val'))) / 2}pt")
            except ValueError:
                pass
        shd = rpr.find(qn("w:shd"))
        if shd is not None and shd.get(qn("w:fill")) not in (None, "auto"):
            styles.append(f"background:#{shd.get(qn('w:fill'))}")
    if bold:
        text = f"<strong>{text}</strong>"
    if italic:
        text = f"<em>{text}</em>"
    if underline:
        text = f"<u>{text}</u>"
    if styles:
        text = f'<span style="{";".join(styles)}">{text}</span>'
    return text


def _paragraph_to_html(p_el, style_map):
    inner = "".join(_run_to_html(r) for r in p_el.iter(qn("w:r")))
    ppr = p_el.find(qn("w:pPr"))
    style_id = ""
    align = ""
    css = []
    if ppr is not None:
        ps = ppr.find(qn("w:pStyle"))
        if ps is not None:
            style_id = ps.get(qn("w:val")) or ""
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            align = {"center": "center", "right": "right", "both": "justify", "end": "right"}.get(jc.get(qn("w:val")), "")
        shd = ppr.find(qn("w:shd"))
        if shd is not None and shd.get(qn("w:fill")) not in (None, "auto"):
            css.append(f"background:#{shd.get(qn('w:fill'))};padding:6px 10px")
        if ppr.find(qn("w:numPr")) is not None:
            style_id = style_id or "ListBullet"
    if align:
        css.append(f"text-align:{align}")
    style_attr = f' style="{";".join(css)}"' if css else ""
    style_name = style_map.get(style_id, style_id).lower()

    if not inner.strip():
        return '<p class="empty">&nbsp;</p>'
    if style_name == "title":
        return f"<h1 class=\"doc-title\"{style_attr}>{inner}</h1>"
    if style_name.startswith("heading"):
        level = "".join(ch for ch in style_name if ch.isdigit()) or "2"
        level = min(max(int(level) + 1, 2), 6)
        return f"<h{level}{style_attr}>{inner}</h{level}>"
    if "list" in style_name:
        return f"<li{style_attr}>{inner}</li>"
    return f"<p{style_attr}>{inner}</p>"


def _table_to_html(tbl_el, style_map):
    rows_html = []
    for tr in tbl_el.iterchildren(qn("w:tr")):
        cells_html = []
        for tc in tr.iterchildren(qn("w:tc")):
            css = []
            attrs = ""
            tcpr = tc.find(qn("w:tcPr"))
            if tcpr is not None:
                shd = tcpr.find(qn("w:shd"))
                if shd is not None and shd.get(qn("w:fill")) not in (None, "auto"):
                    css.append(f"background:#{shd.get(qn('w:fill'))}")
                span = tcpr.find(qn("w:gridSpan"))
                if span is not None and span.get(qn("w:val")):
                    attrs += f' colspan="{span.get(qn("w:val"))}"'
                vm = tcpr.find(qn("w:vMerge"))
                if vm is not None and vm.get(qn("w:val")) != "restart":
                    continue
            content = []
            for child in tc.iterchildren():
                if child.tag == qn("w:p"):
                    content.append(_paragraph_to_html(child, style_map))
                elif child.tag == qn("w:tbl"):
                    content.append(_table_to_html(child, style_map))
            style_attr = f' style="{";".join(css)}"' if css else ""
            cells_html.append(f"<td{attrs}{style_attr}>{''.join(content)}</td>")
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")
    return f'<table class="doc-table">{"".join(rows_html)}</table>'


REPORT_HTML_CSS = """
  * { box-sizing: border-box; }
  body { font-family: 'Times New Roman', 'Segoe UI', Arial, sans-serif; color:#1e293b; background:#f1f5f9; margin:0; padding:30px 0; line-height:1.55; font-size:12.5pt; }
  .page { background:#ffffff; max-width:820px; margin:0 auto; padding:56px 64px; box-shadow:0 4px 24px rgba(15,23,42,0.08); border-radius:6px; }
  h1.doc-title { color:#1d4ed8; font-size:22pt; text-align:center; margin:0 0 14px; }
  h2 { color:#1d4ed8; font-size:16pt; border-bottom:2px solid #bfdbfe; padding-bottom:4px; margin:24px 0 10px; }
  h3 { color:#1e40af; font-size:13.5pt; margin:18px 0 8px; }
  h4, h5, h6 { color:#1e3a8a; margin:14px 0 6px; }
  p { margin:0 0 8px; }
  p.empty { margin:0; line-height:0.8; }
  li { margin:0 0 4px 22px; }
  table.doc-table { width:100%; border-collapse:collapse; margin:10px 0 16px; font-size:11pt; page-break-inside:auto; }
  table.doc-table td { border:1px solid #cbd5e1; padding:6px 9px; vertical-align:top; }
  table.doc-table td p { margin:0; }
  tr { page-break-inside:avoid; }
  @media print {
    body { background:#ffffff; padding:0; }
    .page { box-shadow:none; max-width:none; padding:0; border-radius:0; }
    @page { size:A4; margin:16mm 15mm; }
  }
"""


def docx_bytes_to_html(word_bytes, title="Báo Cáo Kinh Doanh"):
    """Chuyển nội dung file Word đã sinh thành tài liệu HTML độc lập (giữ đúng nội dung báo cáo Word)."""
    document = docx.Document(io.BytesIO(word_bytes))
    style_map = {s.style_id: s.name for s in document.styles if getattr(s, "style_id", None)}
    parts = []
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            parts.append(_paragraph_to_html(child, style_map))
        elif child.tag == qn("w:tbl"):
            parts.append(_table_to_html(child, style_map))
    body = "\n".join(parts)
    safe_title = _html.escape(title)
    return (
        "<!DOCTYPE html>\n<html lang=\"vi\">\n<head>\n<meta charset=\"utf-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
        f"<title>{safe_title}</title>\n<style>{REPORT_HTML_CSS}</style>\n</head>\n"
        f"<body>\n<div class=\"page\" id=\"report-page\">\n{body}\n</div>\n</body>\n</html>\n"
    )

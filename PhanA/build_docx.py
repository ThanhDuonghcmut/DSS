# -*- coding: utf-8 -*-
"""Dựng báo cáo Word Phần A (Data-Driven DSS) từ kpi_summary.json + star_schema.png."""
import json
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))
S = json.load(open(os.path.join(HERE, "tableau_data", "kpi_summary.json")))
NAVY = RGBColor(0x1F, 0x4E, 0x79)
GREY = RGBColor(0x59, 0x59, 0x59)


def money(x):
    return "${:,.0f}".format(x)


def vn(x):
    """Số nguyên kiểu Việt Nam: 172.765"""
    return f"{x:,.0f}".replace(",", ".")


def vn2(x, nd=2):
    """Số thập phân kiểu Việt Nam: 35.214.429,65"""
    return f"{x:,.{nd}f}".replace(",", "X").replace(".", ",").replace("X", ".")


doc = Document()
# font mặc định
st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(11)


def _shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def h(text, level=1):
    p = doc.add_heading(text, level=level)
    for r in p.runs:
        r.font.color.rgb = NAVY
    return p


def para(text, italic=False, bold=False, size=11, color=None, align=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.italic = italic; r.bold = bold; r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    if align:
        p.alignment = align
    return p


def bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r = p.add_run(bold_prefix); r.bold = True
    p.add_run(text)
    return p


def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, hd in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ""
        run = c.paragraphs[0].add_run(hd); run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); run.font.size = Pt(10)
        _shade(c, "1F4E79")
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(str(val)); run.font.size = Pt(9.5)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    return t

# ============================================================ TRANG BÌA
p = para("CAO HỌC K34 – HỆ HỖ TRỢ RA QUYẾT ĐỊNH", bold=True, size=13, color=GREY,
         align=WD_ALIGN_PARAGRAPH.CENTER)
para("ĐỒ ÁN DSS – PHẦN A: DATA-DRIVEN DSS", bold=True, size=20, color=NAVY,
     align=WD_ALIGN_PARAGRAPH.CENTER)
para("Hệ hỗ trợ ra quyết định phân tích Bán hàng & Giao hàng Chuỗi cung ứng toàn cầu "
     "(Bộ dữ liệu DataCo Supply Chain)", italic=True, size=12,
     align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_paragraph()
table(["Hạng mục", "Nội dung"],
      [["Môn học", "Hệ hỗ trợ ra quyết định (DSS)"],
       ["Đề tài", "Data-Driven DSS cho chuỗi cung ứng DataCo"],
       ["Công cụ Phần A", "Tableau (trực quan hoá) + Python (ETL/làm sạch dữ liệu)"],
       ["Mã nhóm", "<MaNhom> — điền trước khi nộp"],
       ["Thành viên & phân công", "<MSHV1 – …>, <MSHV2 – …>, <MSHV3 – …>"]],
      widths=[1.8, 4.5])
para("Ghi chú: các trường <…> là chỗ dành cho nhóm điền thông tin.", italic=True,
     size=9, color=GREY)
doc.add_page_break()

# ============================================================ 1. CHỦ ĐỀ
h("1. Giới thiệu chủ đề & nhu cầu phân tích", 1)
para("Chủ đề nhóm lựa chọn là quản trị chuỗi cung ứng – thương mại điện tử toàn cầu, "
     "dựa trên bộ dữ liệu giao dịch của công ty DataCo Global. Doanh nghiệp bán sản phẩm "
     "thể thao – thời trang – công nghệ qua nhiều thị trường (Market) trên thế giới và "
     "đối mặt hai nhóm vấn đề cốt lõi: (1) hiệu quả bán hàng & lợi nhuận, (2) chất lượng "
     "giao hàng (đúng hẹn / trễ hẹn). Đây là bài toán điển hình mà một Hệ hỗ trợ ra quyết "
     "định hướng dữ liệu (Data-Driven DSS) có thể giải quyết thông qua phân tích mô tả và "
     "trực quan hoá.")
para("Các nhu cầu phân tích đặt ra:")
bullet("doanh thu, lợi nhuận và biên lợi nhuận đang ở mức nào, theo thị trường/ngành hàng/thời gian?")
bullet("tỷ lệ giao hàng trễ bao nhiêu, tập trung ở phương thức vận chuyển hay khu vực nào?")
bullet("những đơn hàng/ngành hàng nào đang bị lỗ, nguyên nhân có phải do chiết khấu?")
bullet("nhóm khách hàng và thị trường nào tạo ra giá trị lớn nhất để ưu tiên nguồn lực?")

# ============================================================ 2. DATASET
h("2. Giới thiệu dataset", 1)
bullet("Kaggle – DataCo Smart Supply Chain for Big Data Analysis.", "Nguồn: ")
bullet("DataCoSupplyChainDataset.csv (mã hoá latin-1).", "Tệp: ")
bullet(f"{S['raw_rows']:,} dòng × {S['raw_cols']} cột ở dạng thô.", "Quy mô thô: ")
bullet("mỗi dòng là một mặt hàng trong đơn hàng (order item line-item).", "Mức chi tiết (grain): ")
bullet(f"{S['date_min']} đến {S['date_max']} (hơn 3 năm giao dịch).", "Khoảng thời gian: ")
para("Các trường dữ liệu chính (vượt yêu cầu tối thiểu 6 trường):")
table(["Nhóm", "Trường tiêu biểu"],
      [["Giao dịch/đo lường", "Sales, Order Item Total, Order Profit Per Order, Order Item Quantity, Order Item Discount Rate"],
       ["Khách hàng", "Customer Id, Customer Segment, Customer City/State/Country"],
       ["Sản phẩm", "Product Name, Category Name, Department Name, Product Price"],
       ["Thời gian", "order date (DateOrders), shipping date (DateOrders)"],
       ["Địa lý đơn hàng", "Market, Order Region, Order Country/State/City"],
       ["Giao hàng", "Shipping Mode, Delivery Status, Days for shipping (real/scheduled), Late_delivery_risk"]],
      widths=[1.8, 4.5])

# ============================================================ 3. LÀM SẠCH
h("3. Xử lý & làm sạch dữ liệu", 1)
para("Toàn bộ quy trình được tự động hoá bằng script Python (pandas) trong tệp "
     "etl_phanA.py. Các bước làm sạch:")
para("3.1. Loại bỏ cột không mang thông tin phân tích", bold=True)
table(["Cột bị loại", "Lý do"],
      [["Product Description", "trống 100%"],
       ["Order Zipcode", "trống 86,2%"],
       ["Customer Email, Customer Password", "chỉ 1 giá trị giả (ẩn danh), vô nghĩa"],
       ["Product Status", "toàn bộ = 0 (hằng số)"],
       ["Product Image", "đường dẫn ảnh, không dùng phân tích"],
       ["Benefit per order", "trùng lặp hoàn toàn với Order Profit Per Order"]],
      widths=[2.6, 3.7])
para("3.2. Xử lý giá trị trống", bold=True)
bullet("Customer Lname: 8 dòng trống → điền \"Unknown\".")
bullet("Customer Zipcode: 3 dòng trống → điền 0.")
para("3.3. Chuẩn hoá định dạng & trùng lặp", bold=True)
bullet(f"Kiểm tra trùng lặp toàn dòng: {S['dup_removed']} dòng trùng (dữ liệu sạch sẵn).")
bullet("Chuyển order date / shipping date (chuỗi M/D/YYYY) sang kiểu datetime chuẩn.")
bullet("Trim khoảng trắng thừa ở các cột văn bản (ví dụ \"Health and Beauty \").")
para("3.4. Tạo trường dẫn xuất", bold=True)
bullet("shipping_delay = Days for shipping (real) − Days for shipment (scheduled): số ngày "
       "giao trễ so với kế hoạch, phục vụ KPI độ trễ.")
para("3.5. Kiểm tra tính nhất quán theo thời gian", bold=True)
bullet("Dữ liệu từ khoảng 10/2017 có cấu trúc khác phần còn lại: dữ liệu thô có đúng 68–69 đơn mỗi ngày (còn 60–69 sau khi lọc), "
       "mỗi đơn chỉ một mặt hàng (trước đó ~3 mặt hàng/đơn), số sản phẩm có bán giảm từ "
       "42–53 xuống 8–20 và các ngành Fan Shop, Golf, Footwear gần như không còn đơn; "
       "Technology chỉ xuất hiện từ giai đoạn này.")
bullet("Nhóm giữ nguyên các dòng này (không xoá vì không có căn cứ chứng minh là lỗi) nhưng "
       "ghi chú trên S1 và trong phần insight để người xem đọc xu hướng cuối kỳ thận trọng.")
para("3.6. Lọc đơn huỷ và đơn nghi gian lận (quyết định nghiệp vụ)", bold=True)
X = S["excluded"]; B0 = X["before"]
bullet("LOẠI khỏi mô hình mọi đơn có Order Status = CANCELED (đã huỷ) hoặc SUSPECTED_FRAUD "
       "(nghi gian lận) ngay trong ETL (hằng số EXCLUDED_STATUS ở etl_phanA.py). Các trạng thái còn "
       "lại (COMPLETE, CLOSED, PENDING, PENDING_PAYMENT, PROCESSING, ON_HOLD, PAYMENT_REVIEW) được "
       "giữ nguyên.", "Quyết định: ")
bullet(f"{vn(X['rows_by_status']['CANCELED'])} dòng CANCELED + {vn(X['rows_by_status']['SUSPECTED_FRAUD'])} "
       f"dòng SUSPECTED_FRAUD = {vn(X['rows'])} dòng ({vn(X['orders'])} đơn; "
       f"{vn2(X['rows'] / B0['rows'] * 100, 1)}% số dòng, {vn2(X['sales'] / B0['total_sales'] * 100, 1)}% "
       "doanh thu). Không đơn nào mang hai trạng thái khác nhau (ETL kiểm tra bằng assert) nên lọc "
       "theo dòng không làm cụt đơn.", "Phạm vi: ")
bullet("(1) đơn huỷ và đơn nghi gian lận không tạo doanh thu hợp lệ nhưng nếu giữ lại sẽ bị cộng vào "
       f"Σ sales ({money(X['sales'])}) và Σ profit ({money(X['profit'])}); (2) toàn bộ "
       f"{vn(X['shipping_canceled_rows'])} dòng này chính là các dòng « Shipping canceled » (đơn không "
       "được giao). Nguồn dữ liệu gán Late_delivery_risk = 0 cho chúng nên chúng sẽ bị tính là ĐÚNG HẠN, "
       "làm tỷ lệ giao trễ thấp giả tạo — ví dụ phương thức First Class chỉ hiện 95,3% trễ nếu giữ các "
       "đơn này, trong khi thực tế 100% đơn giao trễ.", "Lý do: ")
para("Tác động của việc lọc lên các chỉ số chính:", size=10)
table(["Chỉ tiêu", "Trước khi lọc (CSV thô)", "Sau khi lọc (mô hình)"],
      [["Số dòng", vn(B0['rows']), vn(S['clean_rows'])],
       ["Số đơn", vn(B0['orders']), vn(S['n_orders'])],
       ["Số khách hàng", vn(B0['customers']), vn(S['n_customers'])],
       ["Tổng doanh thu", money(B0['total_sales']), money(S['total_sales'])],
       ["Tổng lợi nhuận", money(B0['total_profit']), money(S['total_profit'])],
       ["Biên lợi nhuận", f"{vn2(B0['profit_margin_pct'])}%", f"{vn2(S['profit_margin_pct'])}%"],
       ["Tỷ lệ giao trễ (cấp đơn)", f"{vn2(B0['late_delivery_rate_order_pct'])}%",
        f"{vn2(S['late_delivery_rate_order_pct'])}%"],
       ["Tỷ lệ đúng hạn (cấp đơn)", f"{vn2(100 - B0['late_delivery_rate_order_pct'])}%",
        f"{vn2(S['on_time_rate_order_pct'])}%"],
       ["Giá trị đơn TB (AOV)", f"${vn2(B0['aov'])}", f"${vn2(S['aov'])}"],
       ["Số tổ hợp dim_delivery", vn(B0['dim_delivery_rows']), vn(S['tables']['dim_delivery'][0])]],
      widths=[2.2, 2.0, 2.0])
para("")
bullet("loại 4,3% số dòng làm doanh thu giảm "
       f"{vn2(X['sales'] / B0['total_sales'] * 100, 1)}% và lợi nhuận giảm "
       f"{vn2(X['profit'] / B0['total_profit'] * 100, 1)}%, nên biên lợi nhuận gần như không đổi; "
       f"ngược lại tỷ lệ giao trễ tăng từ {vn2(B0['late_delivery_rate_order_pct'])}% lên "
       f"{vn2(S['late_delivery_rate_order_pct'])}% vì các đơn không được giao không còn bị tính là "
       "đúng hạn. Mọi số liệu ở các mục sau đều tính trên dữ liệu SAU khi lọc.", "Nhận xét: ")
bullet("mô hình không còn dữ liệu để phân tích riêng đơn huỷ/gian lận (cần quay về CSV thô). Nếu "
       "muốn giữ lại, bỏ trạng thái tương ứng khỏi EXCLUDED_STATUS rồi chạy lại etl_phanA.py và "
       "verify_star_schema.py.", "Đánh đổi: ")
para(f"Kết quả: sau làm sạch và lọc đơn huỷ/nghi gian lận, dữ liệu còn {vn(S['clean_rows'])} dòng "
     f"({vn(S['n_orders'])} đơn), đồng nhất và sẵn sàng đưa vào mô hình phân tích.")

# ============================================================ 4. MỤC TIÊU DASHBOARD
h("4. Mục tiêu dashboard & đối tượng sử dụng", 1)
bullet("giám sát đồng thời hiệu quả BÁN HÀNG (doanh thu, lợi nhuận, biên lợi nhuận) và "
       "chất lượng GIAO HÀNG (tỷ lệ trễ, độ trễ) của chuỗi cung ứng toàn cầu, giúp phát "
       "hiện nhanh khu vực/ngành hàng có vấn đề và ra quyết định điều chỉnh.", "Mục tiêu dashboard: ")
bullet("Giám đốc Vận hành / Chuỗi cung ứng (theo dõi giao hàng), Giám đốc Kinh doanh và "
       "quản lý vùng (theo dõi doanh thu – lợi nhuận theo thị trường).", "Đối tượng sử dụng: ")

bullet("Dashboard 1 — Hiệu quả kinh doanh (doanh thu, lợi nhuận); Dashboard 2 — Giao hàng "
       "& đơn bán lỗ; Dashboard 3 — Khách hàng, sản phẩm & chi tiết (thiết kế chi tiết ở "
       "Mục 8.3).", "Cấu trúc: ")

# ============================================================ 5. KPI
h("5. Các KPI chính", 1)
para("Dashboard thể hiện 7 KPI (vượt mức tối thiểu 5). Cột \"Calculated field (Tableau)\" "
     "là công thức để tạo trực tiếp trong Tableau.")
kpi_rows = [
    ["1. Tổng doanh thu", "Tổng giá trị bán ra", "SUM([sales])",
     money(S['total_sales']), "Quy mô kinh doanh; nền để so sánh theo chiều."],
    ["2. Tổng lợi nhuận", "Tổng lãi/lỗ thực tế", "SUM([profit])",
     money(S['total_profit']), "Hiệu quả tài chính thực sự của hoạt động bán hàng."],
    ["3. Biên lợi nhuận %", "Lợi nhuận / Doanh thu", "SUM([profit]) / SUM([sales])",
     f"{S['profit_margin_pct']}%", "Cảnh báo ngành hàng/thị trường bào mòn lợi nhuận."],
    ["4. Tỷ lệ giao trễ %", "% đơn hàng giao trễ", "SUM([late_delivery_flag])/COUNT([order_id])",
     f"{S['late_delivery_rate_order_pct']}%", "KPI vận hành cốt lõi về chất lượng giao hàng."],
    ["5. Độ trễ giao TB (ngày)", "Ngày thực − ngày kế hoạch (cấp đơn)", "AVG([shipping_delay])",
     f"{S['avg_shipping_delay']} ngày", "Mức độ chậm so với cam kết để cải thiện SLA."],
    ["6. Tỷ lệ đơn lỗ %", "% dòng đơn có profit < 0",
     "SUM([Loss Flag])/COUNT([order_item_id]); Loss Flag = IF [profit]<0 THEN 1 ELSE 0 END",
     f"{S['loss_order_line_pct']}%", "Khoanh vùng giao dịch lỗ để xử lý chiết khấu/giá."],
    ["7. Giá trị đơn TB (AOV)", "Doanh thu ròng / số đơn",
     "SUM([net_sales])/COUNTD([order_id])",
     money(S['aov']), "Đo sức mua mỗi đơn; định hướng up-sell/cross-sell."],
]
table(["KPI", "Cách tính", "Calculated field (Tableau)", "Giá trị", "Ý nghĩa ra quyết định"],
      kpi_rows, widths=[1.1, 1.1, 1.6, 0.8, 1.7])

# ============================================================ 6. STAR SCHEMA
h("6. Thiết kế lược đồ sao (Star Schema)", 1)
para("Mục này trình bày thiết kế kho dữ liệu theo quy trình chuẩn của Kimball: xác định "
     "business process → grain → dimensions → facts/measures → loại fact table → vẽ lược "
     "đồ sao → kiểm chứng. Thiết kế gồm đúng hai bảng fact.")

# ---- 6.1 Business Process ----
h("6.1. Xác định Business Process", 2)
para("Từ đặc tả nghiệp vụ, dữ liệu DataCo chứa hai quy trình nghiệp vụ (business process) "
     "có mức chi tiết khác nhau, tương ứng hai bảng fact:")
bullet("phân tích bán hàng – lợi nhuận ở mức TỪNG MẶT HÀNG trong đơn (doanh thu, chiết "
       "khấu, lợi nhuận theo sản phẩm/khách hàng/thị trường).", "FACT_SALES — ")
bullet("phân tích hiệu quả giao hàng ở mức TỪNG ĐƠN HÀNG/LẦN VẬN CHUYỂN (thời gian giao, "
       "trễ hẹn, đúng hạn theo phương thức/khu vực).", "FACT_ORDER_DELIVERY — ")
para("Vì sao hai quy trình có grain khác nhau và KHÔNG gộp chung một fact table?", bold=True)
bullet("Bán hàng phát sinh ở mức dòng mặt hàng (một đơn nhiều sản phẩm, mỗi sản phẩm một "
       "doanh thu/lợi nhuận riêng). Giao hàng lại là thuộc tính của CẢ ĐƠN: ngày giao, số "
       "ngày vận chuyển, trạng thái trễ là như nhau cho mọi mặt hàng trong đơn.")
bullet(f"Nếu nhét chỉ số giao hàng vào fact mức mặt hàng, chúng sẽ bị lặp lại trên mọi dòng "
       f"(trung bình {S['avg_items_per_order']} dòng/đơn) → khi tính 'tỷ lệ giao trễ' sẽ "
       "ĐẾM TRÙNG theo số mặt hàng, thiên lệch về các đơn lớn. Tách fact giúp mỗi đơn chỉ "
       "được tính một lần.")
bullet("Trộn hai grain trong một bảng vi phạm nguyên tắc 'một grain cho một fact' và làm "
       "các phép tổng hợp (SUM, AVG) không còn chính xác.")

# ---- 6.2 Grain ----
h("6.2. Xác định Grain", 2)
bullet("“Một dòng trong FACT_SALES biểu diễn một mặt hàng (line-item) thuộc một đơn "
       "hàng.”", "")
bullet("“Một dòng trong FACT_ORDER_DELIVERY biểu diễn một đơn hàng được vận chuyển "
       "tới người nhận.”", "")
para("Kiểm tra grain có đáp ứng yêu cầu phân tích:")
table(["Yêu cầu", "Grain hỗ trợ?", "Giải thích"],
      [["KPI doanh thu/lợi nhuận theo sản phẩm", "Có", "Grain mặt hàng của FACT_SALES đủ chi tiết để roll-up lên category, department."],
       ["KPI tỷ lệ giao trễ / đúng hạn", "Có", "Grain đơn hàng của FACT_ORDER_DELIVERY tính đúng mỗi đơn một lần."],
       ["Drill-down Year→Quarter→Month→Date", "Có", "Cả hai fact đều có khoá ngày (order_date_key, ship_date_key)."],
       ["Drill-down Market→Region→Country→State→City", "Có", "Cả hai fact đều nối DIM_LOCATION."]],
      widths=[2.3, 0.9, 3.0])

# ---- 6.3 Dimensions ----
h("6.3. Xác định Dimensions", 2)
para("Các dimension cùng business key (BK), surrogate key (SK), thuộc tính mô tả và phân cấp:")
table(["Dimension", "Business key", "Surrogate key", "Thuộc tính mô tả", "Hierarchy", "Số dòng"],
      [["DIM_DATE", "full_date", "date_key (YYYYMMDD)", "year, quarter, month, weekday, is_weekend", "Year→Quarter→Month→Date", f"{S['tables']['dim_date'][0]:,}"],
       ["DIM_CUSTOMER", "Customer Id", "customer_key", "segment, city, state, country", "Country→State→City", f"{S['tables']['dim_customer'][0]:,}"],
       ["DIM_PRODUCT", "Product Card Id", "product_key", "product_name, category, department, product_price", "Department→Category→Product", f"{S['tables']['dim_product'][0]:,}"],
       ["DIM_LOCATION", "(market..city)", "location_key", "market, region, country, state, city", "Market→Region→Country→State→City", f"{S['tables']['dim_location'][0]:,}"],
       ["DIM_DELIVERY", "(mode..status)", "delivery_key", "shipping_mode, delivery_status, payment_type, order_status", "—", f"{S['tables']['dim_delivery'][0]:,}"]],
      widths=[1.2, 1.0, 1.1, 1.6, 1.4, 0.6])
para("Conformed dimensions (chia sẻ giữa hai fact): ", bold=True)
bullet("DIM_DATE, DIM_CUSTOMER, DIM_LOCATION, DIM_DELIVERY dùng chung cho cả FACT_SALES và "
       "FACT_ORDER_DELIVERY — cho phép so sánh/kết hợp hai nghiệp vụ trên cùng bộ chiều.")
bullet("DIM_PRODUCT chỉ áp dụng cho FACT_SALES (mức đơn hàng không gắn với một sản phẩm "
       "duy nhất).")
para("Role-playing dimension: ", bold=True)
bullet("DIM_DATE đóng hai vai trò trong cùng một fact qua hai khoá order_date_key (ngày "
       "đặt) và ship_date_key (ngày giao). Trong Tableau sẽ nối DIM_DATE hai lần.")

# ---- 6.4 Facts & Measures ----
h("6.4. Xác định Facts và Measures", 2)
para("FACT_SALES — base measures (lưu tại grain mặt hàng):", bold=True)
table(["Measure", "Ý nghĩa", "Cách tính", "Đơn vị", "Khả năng tổng hợp"],
      [["quantity", "Số lượng bán", "nguồn", "cái", "Additive"],
       ["sales", "Doanh thu gộp", "quantity × unit_price", "USD", "Additive"],
       ["discount", "Tiền chiết khấu", "nguồn", "USD", "Additive"],
       ["net_sales", "Doanh thu ròng", "sales − discount", "USD", "Additive"],
       ["profit", "Lợi nhuận", "nguồn (Order Profit Per Order)", "USD", "Additive"],
       ["unit_price", "Đơn giá", "nguồn", "USD", "Non-additive (giá)"],
       ["profit_ratio", "Tỷ suất LN dòng", "profit / net_sales", "%", "Non-additive (tỷ lệ)"]],
      widths=[1.0, 1.3, 1.6, 0.7, 1.4])
para("FACT_ORDER_DELIVERY — base measures (lưu tại grain đơn hàng):", bold=True)
table(["Measure", "Ý nghĩa", "Cách tính", "Đơn vị", "Khả năng tổng hợp"],
      [["days_ship_scheduled", "Số ngày giao cam kết", "nguồn", "ngày", "Semi-additive (TB theo đơn)"],
       ["days_ship_real", "Số ngày giao thực tế", "nguồn", "ngày", "Semi-additive"],
       ["shipping_delay", "Độ trễ", "real − scheduled", "ngày", "Semi-additive"],
       ["late_delivery_flag", "Cờ giao trễ (0/1)", "nguồn Late_delivery_risk", "cờ", "Additive (đếm số trễ)"],
       ["on_time_flag", "Cờ đúng hạn (0/1)", "1 − late_delivery_flag", "cờ", "Additive"],
       ["order_sales", "Doanh thu đơn", "Σ sales trong đơn", "USD", "Additive"],
       ["order_profit", "Lợi nhuận đơn", "Σ profit trong đơn", "USD", "Additive"]],
      widths=[1.2, 1.2, 1.3, 0.7, 1.5])
para("Base measures vs Calculated measures/KPI:", bold=True)
bullet("chỉ lưu các đại lượng CÓ THỂ CỘNG (quantity, sales, profit, late_delivery_flag…) "
       "trong fact table.", "Lưu trực tiếp (base): ")
bullet("các KPI dạng tỷ lệ/trung bình KHÔNG lưu sẵn mà tính ở tầng OLAP/BI từ base measures, "
       "ví dụ Biên LN% = SUM(profit)/SUM(sales); Tỷ lệ giao trễ% = SUM(late_delivery_flag)/"
       "COUNT(order_id); AOV = SUM(order_net_sales)/COUNTD(order_id). Lý do: tỷ lệ không "
       "cộng được nên nếu lưu sẵn sẽ sai khi roll-up.", "Tính ở OLAP (calculated): ")
para("Kiểm chứng khả năng tổng hợp: các base measures cộng được bảo đảm tính lại đúng mọi "
     "KPI khi roll-up/drill-down/slice/dice — ví dụ tỷ lệ giao trễ ở bất kỳ mức nào (thị "
     "trường, tháng) đều = Σ(late_delivery_flag)/Σ(count) tại mức đó, luôn chính xác.",
     italic=True, size=9, color=GREY)

# ---- 6.5 Loại Fact Table ----
h("6.5. Xác định loại Fact Table", 2)
table(["Fact table", "Loại", "Lý do"],
      [["FACT_SALES", "Transaction fact", "Ghi nhận từng giao dịch bán (mặt hàng) tại thời điểm phát sinh; grain nguyên tử."],
       ["FACT_ORDER_DELIVERY", "Accumulating snapshot", "Theo vòng đời đơn hàng với các mốc ngày đặt → ngày giao và các đại lượng độ trễ tích luỹ giữa các mốc."]],
      widths=[2.0, 1.6, 2.7])

# ---- 6.6 Sơ đồ ----
h("6.6. Lược đồ Star Schema", 2)
doc.add_picture(os.path.join(HERE, "star_schema.png"), width=Inches(6.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
para("Hình 1. Lược đồ sao hai fact của DSS chuỗi cung ứng DataCo (thể hiện PK, FK, surrogate "
     "key, measures, quan hệ, grain, conformed & role-playing dimension).",
     italic=True, size=9, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)

# ---- 6.7 Kiểm chứng ----
h("6.7. Kiểm chứng thiết kế", 2)
para("Mô hình trả lời được các câu hỏi phân tích yêu cầu (đã điều chỉnh theo dữ liệu DataCo — "
     "bộ dữ liệu không có thông tin nhân viên giao hàng, nhiều lần giao hay lý do thất bại "
     "chi tiết, nên các câu hỏi liên quan được thay bằng chỉ số tương đương gần nhất):")
table(["Câu hỏi phân tích", "Fact dùng", "Dimension / drill", "Measure / KPI"],
      [["1. Quý gần nhất, khu vực nào tỷ lệ đúng hạn thấp nhất?", "Order_Delivery", "Dim_Date (lọc quý), Dim_Location→region", "AVG(on_time_flag)"],
       ["2. Phương thức nào giao chậm hơn cam kết nhiều nhất?", "Order_Delivery", "Dim_Delivery→shipping_mode", "AVG(shipping_delay)"],
       ["3. Tỷ lệ giao đúng hạn theo khu vực (thay cho 'thành công lần đầu')?", "Order_Delivery", "Dim_Location→region", "AVG(on_time_flag)"],
       ["4. Yếu tố nào gây trễ nhiều nhất (thay cho 'lý do thất bại')?", "Order_Delivery", "Dim_Delivery (mode, status)", "SUM(late_delivery_flag)"],
       ["5. Khu vực nào tỷ lệ giao trễ cao nhất (thay cho 'nhân viên')?", "Order_Delivery", "Dim_Location→region/state", "AVG(late_delivery_flag)"],
       ["6. Doanh thu Year→Quarter→Month theo phương thức?", "Sales", "Dim_Date (drill), Dim_Delivery→shipping_mode", "SUM(sales)"],
       ["7. Nhóm khách hàng nào doanh thu lớn nhất?", "Sales", "Dim_Customer→segment", "SUM(sales)"],
       ["8. Khu vực nào nhiều đơn nhưng đúng hạn thấp?", "Order_Delivery", "Dim_Location→region", "COUNT(order_id) vs AVG(on_time_flag)"]],
      widths=[2.4, 1.0, 1.9, 1.3])
para("Mọi câu hỏi đều drill-down được tới mức yêu cầu nhờ phân cấp của DIM_DATE và "
     "DIM_LOCATION; grain, dimensions và measures do đó đã đủ, không cần điều chỉnh.",
     italic=True, size=9, color=GREY)

# ============================================================ 7. ETL
def code(text, size=8.5):
    """Khối mã (SQL/Python) font Consolas, nền xám nhạt."""
    pp = doc.add_paragraph()
    r = pp.add_run(text)
    r.font.name = "Consolas"; r.font.size = Pt(size)
    pPr = pp._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), "F2F2F2")
    pPr.append(shd)
    return pp


h("7. Truy vấn dữ liệu & ETL vào Star Schema", 1)
para("Mục này làm rõ hai yêu cầu: (1) thiết kế ETL và truy vấn để đưa dữ liệu vào các bảng "
     "fact/dimension; (2) chứng minh dữ liệu được tổ chức hợp lý để phân tích. Cấu trúc: "
     "7.1 luồng ETL; 7.2 biến đổi và nạp từng bảng dimension; 7.3 biến đổi và nạp hai bảng "
     "fact; 7.4 mã SQL tương đương; 7.5 chứng minh tổ chức dữ liệu hợp lý (kiểm tra toàn "
     "vẹn, đối soát, quy tắc nghiệp vụ, truy vấn phân tích mẫu); 7.6 nhận xét và hạn chế.")

# ---------------------------------------------------------------- 7.1
h("7.1. Luồng ETL (Extract – Transform – Load)", 2)
para("Toàn bộ quy trình nằm trong etl_phanA.py (pandas), chạy một lệnh: python3 etl_phanA.py.")
table(["Bước", "Việc làm", "Đầu vào → đầu ra"],
      [["1. Extract", "Đọc CSV thô bằng pd.read_csv(encoding=\"latin-1\").",
        "DataCoSupplyChainDataset.csv → 180.519 dòng × 53 cột"],
       ["2. Làm sạch", "Bỏ 7 cột rác, bỏ dòng trùng, trim chuỗi, điền giá trị trống, chuẩn hoá "
        "ngày (M/D/YYYY H:M), tạo shipping_delay, rồi loại đơn CANCELED/SUSPECTED_FRAUD (Mục 3.6).",
        f"{vn(S['raw_rows'])} → {vn(S['clean_rows'])} dòng sạch"],
       ["3. Transform dimension", "Tách thuộc tính thành 5 bảng dimension (drop_duplicates theo "
        "khoá), sinh khoá thay thế cho location và delivery, sinh bảng lịch dim_date.",
        "dim_customer, dim_product, dim_date, dim_location, dim_delivery"],
       ["4. Transform fact bán hàng", "Đổi ngày thành khoá YYYYMMDD; gắn location_key, "
        "delivery_key bằng merge (left join) trên các cột tự nhiên.", f"fact_sales ({vn(S['clean_rows'])} dòng)"],
       ["5. Transform fact giao hàng", "Gộp fact_sales về grain đơn hàng (group by order_id) "
        "với quy tắc gộp riêng cho từng cột (Mục 7.3).", f"fact_order_delivery ({vn(S['n_orders'])} dòng)"],
       ["6. Load", "Ghi 7 tệp CSV vào tableau_data/ và kpi_summary.json; Tableau đọc trực "
        "tiếp bằng Text file (Mục 8.1).", "7 CSV + 1 JSON"]],
      widths=[1.3, 3.2, 1.8])
para("")
para("Thứ tự thực hiện bắt buộc: dimension trước, fact sau, vì fact cần khoá của dimension "
     "để gắn khoá ngoại.", italic=True, size=9, color=GREY)

# ---------------------------------------------------------------- 7.2
h("7.2. Biến đổi và nạp các bảng dimension", 2)
table(["Dimension", "Cột nguồn (thô)", "Khoá & cách sinh", "Quy tắc loại trùng", "Số dòng"],
      [["dim_customer", "Customer Id, Fname, Lname, Segment, City, State, Country, Zipcode",
        "customer_key = Customer Id (khoá tự nhiên dùng làm khoá chính)",
        "drop_duplicates(Customer Id)", f"{S['tables']['dim_customer'][0]:,}"],
       ["dim_product", "Product Card Id, Product Name, Category Id/Name, Department Id/Name, "
        "Product Price", "product_key = Product Card Id",
        "drop_duplicates(Product Card Id)", f"{S['tables']['dim_product'][0]:,}"],
       ["dim_date", "order date + shipping date (hợp hai cột)",
        "date_key = YYYYMMDD (số nguyên, dễ sắp xếp)",
        "Hợp ngày đặt và ngày giao, chuẩn hoá về 00:00, loại trùng",
        f"{S['tables']['dim_date'][0]:,}"],
       ["dim_location", "Market, Order Region, Order Country, Order State, Order City",
        "location_key = số thứ tự 1..N (khoá thay thế)",
        "drop_duplicates trên cả 5 cột", f"{S['tables']['dim_location'][0]:,}"],
       ["dim_delivery", "Shipping Mode, Delivery Status, Type (payment), Order Status",
        "delivery_key = số thứ tự 1..N (khoá thay thế)",
        "drop_duplicates trên cả 4 cột", f"{S['tables']['dim_delivery'][0]:,}"]],
      widths=[1.0, 1.9, 1.6, 1.3, 0.5])
para("")
bullet("dim_date là bảng lịch sinh thêm year, quarter, month, month_name, day, day_of_week, "
       "weekday_name, is_weekend từ ngày gốc; liên tục 1.133 ngày từ 01/01/2015 đến "
       "06/02/2018 (kết thúc sau ngày đặt cuối 31/01/2018 vì chứa cả ngày giao). Một bảng "
       "dùng cho hai vai trò (ngày đặt, ngày giao).", "Date: ")
bullet("dim_location lưu địa điểm GIAO HÀNG (Order Region/Country/State/City) với phân cấp "
       "Market → Region → Country → State → City; dim_customer lưu địa chỉ KHÁCH HÀNG — hai "
       "khái niệm địa lý khác nhau nên tách riêng.", "Location vs Customer: ")
bullet("dim_delivery là chiều hỗn hợp (junk dimension): gom bốn thuộc tính có ít giá trị "
       "(phương thức giao, trạng thái giao, hình thức thanh toán, trạng thái đơn) thành 56 "
       "tổ hợp thực tế thay vì tạo bốn bảng nhỏ, giúp fact chỉ cần một khoá ngoại.",
       "Delivery: ")
bullet("trong dữ liệu thô, không khách hàng nào và không sản phẩm nào có hơn một phiên bản "
       "thuộc tính (0/20.261 khách hàng, 0/118 sản phẩm) nên việc giữ một dòng cho mỗi khoá "
       "(slowly changing dimension kiểu 1) không làm mất thông tin lịch sử.",
       "Lịch sử thay đổi: ")
code("# Ví dụ: sinh khoá thay thế cho dim_location (etl_phanA.py)\n"
     "loc_cols = [\"Market\", \"Order Region\", \"Order Country\", \"Order State\", \"Order City\"]\n"
     "dim_location = df[loc_cols].drop_duplicates().reset_index(drop=True)\n"
     "dim_location.insert(0, \"location_key\", dim_location.index + 1)")

# ---------------------------------------------------------------- 7.3
h("7.3. Biến đổi và nạp hai bảng fact", 2)
para("FACT_SALES (grain: một mặt hàng trong đơn)", bold=True)
table(["Nhóm cột", "Cột đích", "Nguồn / phép biến đổi"],
      [["Khoá chính, degenerate", "order_item_id, order_id",
        "Order Item Id (PK); Order Id giữ làm degenerate dimension để gom các dòng cùng đơn"],
       ["Khoá ngoại → dimension", "customer_key, product_key",
        "Customer Id, Product Card Id (trùng khoá của dimension)"],
       ["Khoá ngoại → dim_date (role-playing)", "order_date_key, ship_date_key",
        "order/shipping date → strftime(\"%Y%m%d\")"],
       ["Khoá ngoại → dim_location, dim_delivery", "location_key, delivery_key",
        "merge(how=\"left\") với dimension trên các cột tự nhiên (5 cột và 4 cột)"],
       ["Measure cộng được", "quantity, discount, sales, net_sales, profit",
        "Order Item Quantity / Discount / Sales / Order Item Total / Order Profit Per Order"],
       ["Measure không cộng được", "unit_price, discount_rate, profit_ratio",
        "Lưu để tham khảo; KPI tỷ lệ luôn tính lại ở Tableau từ measure cộng được"],
       ["Thời gian giao", "days_ship_real, days_ship_scheduled, shipping_delay, "
        "late_delivery_risk", "Cột gốc + shipping_delay = real − scheduled"]],
      widths=[1.6, 2.0, 2.7])
para("")
para("FACT_ORDER_DELIVERY (grain: một đơn hàng) — gộp từ FACT_SALES bằng group by order_id", bold=True)
table(["Cột đích", "Phép gộp", "Lý do"],
      [["customer_key, order_date_key, ship_date_key, location_key, delivery_key, "
        "days_ship_scheduled, days_ship_real, shipping_delay", "first",
        "Thuộc tính của CẢ đơn nên giống nhau trên mọi mặt hàng; đã kiểm tra 0 đơn có hơn một "
        "giá trị (Mục 7.5, nhóm C)"],
       ["late_delivery_flag", "max", "Đơn trễ nếu có cờ trễ; thực tế cờ giống nhau trong đơn"],
       ["on_time_flag", "1 − late_delivery_flag", "Cờ đối ngẫu để tính tỷ lệ đúng hạn"],
       ["order_item_count", "count(order_item_id)", "Số mặt hàng trong đơn"],
       ["order_quantity, order_sales, order_net_sales, order_profit", "sum",
        "Cộng các measure cộng được lên cấp đơn; tổng toàn bảng bằng đúng fact_sales"]],
      widths=[2.5, 1.2, 2.6])
para("")
code("# Gộp về grain đơn hàng (etl_phanA.py)\n"
     "grp = fact_sales.groupby(\"order_id\")\n"
     "fact_order_delivery = grp.agg(\n"
     "    location_key=(\"location_key\", \"first\"),\n"
     "    late_delivery_flag=(\"late_delivery_risk\", \"max\"),\n"
     "    order_item_count=(\"order_item_id\", \"count\"),\n"
     "    order_sales=(\"sales\", \"sum\"), ...).reset_index()")

# ---------------------------------------------------------------- 7.4
h("7.4. Mã SQL tương đương", 2)
para("Ngoài ETL bằng Python, nhóm viết lại các bước tạo dimension và fact bằng SQL để có thể "
     "chạy trong cơ sở dữ liệu quan hệ. Giả định có bảng stg_orders là dữ liệu thô đã làm "
     "sạch và đã loại đơn CANCELED/SUSPECTED_FRAUD ở Mục 3, kèm order_date_key và ship_date_key dạng YYYYMMDD; dim_customer và "
     "dim_product tạo bằng SELECT DISTINCT trên các cột tương ứng.")
code("""-- Nạp stg_orders: chỉ lấy đơn hợp lệ (đã làm sạch ở Mục 3)
--   WHERE order_status NOT IN ('CANCELED', 'SUSPECTED_FRAUD')

-- Dimension: khoá thay thế bằng ROW_NUMBER()
CREATE TABLE dim_location AS
SELECT ROW_NUMBER() OVER (ORDER BY market, region, country, state, city) AS location_key,
       market, region, country, state, city
FROM (SELECT DISTINCT market, region, country, state, city FROM stg_orders);

CREATE TABLE dim_delivery AS
SELECT ROW_NUMBER() OVER (ORDER BY shipping_mode, delivery_status, payment_type, order_status)
       AS delivery_key, shipping_mode, delivery_status, payment_type, order_status
FROM (SELECT DISTINCT shipping_mode, delivery_status, payment_type, order_status FROM stg_orders);

-- Fact bán hàng: gắn khoá ngoại bằng JOIN trên cột tự nhiên
CREATE TABLE fact_sales AS
SELECT s.order_item_id, s.order_id, s.customer_id AS customer_key, s.product_card_id AS product_key,
       s.order_date_key, s.ship_date_key, l.location_key, d.delivery_key,
       s.quantity, s.unit_price, s.discount, s.discount_rate, s.sales, s.net_sales, s.profit,
       s.profit_ratio, s.days_real AS days_ship_real, s.days_sched AS days_ship_scheduled,
       s.days_real - s.days_sched AS shipping_delay, s.late_risk AS late_delivery_risk
FROM stg_orders s
JOIN dim_location l ON l.market=s.market AND l.region=s.region AND l.country=s.country
                   AND l.state=s.state AND l.city=s.city
JOIN dim_delivery d ON d.shipping_mode=s.shipping_mode AND d.delivery_status=s.delivery_status
                   AND d.payment_type=s.payment_type AND d.order_status=s.order_status;

-- Fact giao hàng: gộp về một dòng cho mỗi đơn
CREATE TABLE fact_order_delivery AS
SELECT order_id, MIN(customer_key) AS customer_key, MIN(order_date_key) AS order_date_key,
       MIN(ship_date_key) AS ship_date_key, MIN(location_key) AS location_key,
       MIN(delivery_key) AS delivery_key, MIN(days_ship_scheduled) AS days_ship_scheduled,
       MIN(days_ship_real) AS days_ship_real, MIN(shipping_delay) AS shipping_delay,
       MAX(late_delivery_risk) AS late_delivery_flag, 1 - MAX(late_delivery_risk) AS on_time_flag,
       COUNT(*) AS order_item_count, SUM(quantity) AS order_quantity, SUM(sales) AS order_sales,
       SUM(net_sales) AS order_net_sales, SUM(profit) AS order_profit
FROM fact_sales GROUP BY order_id;""", size=8)
para("")
para("Kết quả chạy thử: các câu lệnh trên đã được chạy trên SQLite từ CSV thô (đã loại đơn huỷ/nghi gian lận) và cho cùng "
     "kết quả với ETL Python — 3.758 dòng dim_location, 56 dòng dim_delivery, 172.765 dòng "
     "fact_sales, 62.897 dòng fact_order_delivery; thuộc tính địa lý và giao hàng của cả "
     "172.765 dòng khớp 100%, và chỉ số cấp đơn của cả 62.897 đơn khớp (doanh thu, số mặt "
     "hàng, cờ trễ, độ trễ). Số khoá thay thế có thể khác (SQL đánh số theo thứ tự chữ cái, "
     "Python theo thứ tự xuất hiện) nhưng ý nghĩa không đổi.", italic=True, size=9, color=GREY)

# ---------------------------------------------------------------- 7.5
h("7.5. Chứng minh dữ liệu được tổ chức hợp lý để phân tích", 2)
para("Nhóm chạy bộ kiểm tra tự động verify_star_schema.py (chỉ dùng thư viện chuẩn Python: "
     "nạp 7 CSV vào SQLite, đối chiếu với CSV thô). Kết quả: 58/58 phép kiểm tra đạt (gồm 4 phép kiểm tra phần lọc đơn huỷ/nghi gian lận). Cách "
     "chạy lại: python3 verify_star_schema.py.")
para("A. Toàn vẹn khoá", bold=True)
table(["Kiểm tra", "Phương pháp", "Kết quả"],
      [["Khoá chính duy nhất", "COUNT(*) so với COUNT(DISTINCT khoá) trên cả 7 bảng",
        "7/7 bảng không trùng khoá"],
       ["Khoá ngoại hợp lệ", "Đếm dòng fact có khoá rỗng hoặc không tồn tại trong dimension, "
        "cho 11 quan hệ (6 của fact_sales, 5 của fact_order_delivery)",
        "0 dòng mồ côi, 0 khoá rỗng"],
       ["Dimension không thừa", "Dòng dimension không được fact_sales tham chiếu",
        "0 dòng thừa ở khách hàng, sản phẩm, địa điểm, giao hàng"],
       ["Bảng lịch liên tục", "Số dòng so với số ngày giữa ngày nhỏ nhất và lớn nhất",
        "1.133 dòng = 1.133 ngày, không thiếu ngày"]],
      widths=[1.5, 3.3, 1.5])
para("")
para("B. Đối soát với dữ liệu thô (không mất, không nhân đôi dữ liệu)", bold=True)
table(["Chỉ tiêu", "CSV thô (đã loại đơn huỷ/gian lận)", "fact_sales", "fact_order_delivery"],
      [["Số dòng", vn(S['clean_rows']), vn(S['clean_rows']),
        f"{vn(S['n_orders'])} đơn (Σ order_item_count = {vn(S['clean_rows'])})"],
       ["Dòng bị loại (CANCELED + SUSPECTED_FRAUD)",
        f"{vn(S['excluded']['rows'])} = {vn(S['raw_rows'])} − {vn(S['clean_rows'])}",
        "0 dòng còn trạng thái này", "—"],
       ["Số đơn khác nhau", vn(S['n_orders']), vn(S['n_orders']), vn(S['n_orders'])],
       ["Số khách hàng / sản phẩm", f"{vn(S['n_customers'])} / {vn(S['n_products'])}",
        f"dim_customer {vn(S['n_customers'])} / dim_product {vn(S['n_products'])}", "—"],
       ["Σ doanh thu (sales)", vn2(S['total_sales']), vn2(S['total_sales']), vn2(S['total_sales'])],
       ["Σ lợi nhuận (profit)", vn2(S['total_profit']), vn2(S['total_profit']), vn2(S['total_profit'])],
       ["Σ doanh thu ròng (net_sales)", "—", vn2(S['net_sales']), vn2(S['net_sales'])],
       ["Σ số lượng (quantity)", "—", vn(S['total_quantity']), vn(S['total_quantity'])]],
      widths=[1.7, 1.2, 1.7, 1.7])
para("")
para("C. Grain và quy tắc nghiệp vụ", bold=True)
table(["Kiểm tra", "Kết quả"],
      [["order_item_id duy nhất (grain mặt hàng); order_id duy nhất trong fact_order_delivery "
        "(grain đơn)", "Đạt"],
       ["9 thuộc tính cấp đơn (khách hàng, ngày đặt, ngày giao, địa điểm, giao hàng, ngày "
        "thực tế, ngày kế hoạch, độ trễ, cờ trễ) chỉ có một giá trị trong mỗi đơn → dùng "
        "« first » khi gộp không sai lệch", "0 đơn vi phạm"],
       ["sales = quantity × unit_price (sai số ≤ 0,05 USD)", "0 dòng vi phạm"],
       ["net_sales = sales − discount (sai số ≤ 0,05 USD)", "0 dòng vi phạm"],
       ["shipping_delay = days_ship_real − days_ship_scheduled", "0 dòng vi phạm"],
       ["Ngày giao ≥ ngày đặt", "0 dòng vi phạm"],
       ["Trạng thái « Late delivery » ⇔ late_delivery_flag = 1 (36.048 đơn)", "0 đơn lệch"],
       ["Measure sales, profit, quantity không có giá trị trống", "0 dòng trống"],
       ["Phân cấp: mỗi category thuộc đúng một department; mỗi region thuộc một market; mỗi "
        "(state, city) thuộc một country; khoá tự nhiên của dim_location và product_name "
        "không trùng", "0 vi phạm → drill-down cho kết quả đúng"]],
      widths=[5.0, 1.3])
para("")
para("Ví dụ hai truy vấn kiểm tra có thể chạy lại trực tiếp:", size=10)
code("""-- Khoá ngoại mồ côi (kỳ vọng 0)
SELECT COUNT(*) FROM fact_sales f
LEFT JOIN dim_location l ON f.location_key = l.location_key
WHERE l.location_key IS NULL;

-- Đối soát tổng doanh thu giữa hai fact (kỳ vọng hai số bằng nhau)
SELECT (SELECT SUM(sales) FROM fact_sales)             AS tong_fact_sales,
       (SELECT SUM(order_sales) FROM fact_order_delivery) AS tong_fact_order_delivery;""", size=8)
para("")
para("D. Truy vấn phân tích mẫu: mô hình trả lời được câu hỏi nghiệp vụ chỉ với một bước join",
     bold=True)
para("D1. Doanh thu và tỷ lệ giao trễ theo thị trường (fact_order_delivery + dim_location):",
     size=10)
code("""SELECT l.market,
       SUM(d.order_sales)        AS doanh_thu,
       COUNT(*)                  AS so_don,
       AVG(d.late_delivery_flag) AS ty_le_tre
FROM   fact_order_delivery d
JOIN   dim_location l ON d.location_key = l.location_key
GROUP  BY l.market
ORDER  BY doanh_thu DESC;""", size=8.5)
table(["market", "doanh_thu (USD)", "so_don", "ty_le_tre"],
      [["Europe", "10.405.371", "17.763", "57,42%"],
       ["LATAM", "9.824.330", "16.414", "56,90%"],
       ["Pacific Asia", "7.942.352", "16.853", "57,68%"],
       ["USCA", "4.836.414", "8.178", "57,52%"],
       ["Africa", "2.205.963", "3.689", "56,55%"],
       ["Tổng", "35.214.430", "62.897", "57,31%"]],
      widths=[1.6, 1.8, 1.2, 1.2])
para("Nhận xét: tổng doanh thu và tổng số đơn khớp với Mục 7.5-B; tỷ lệ trễ gần như bằng "
     "nhau giữa các thị trường (56,5–57,7%), tức vấn đề giao trễ không nằm ở khu vực mà ở "
     "phương thức giao (D2).", italic=True, size=9, color=GREY)
para("D2. Tỷ lệ trễ và độ trễ theo phương thức giao (fact_order_delivery + dim_delivery):", size=10)
code("""SELECT dl.shipping_mode, COUNT(*) AS so_don,
       AVG(d.late_delivery_flag) AS ty_le_tre, AVG(d.shipping_delay) AS tre_tb_ngay
FROM   fact_order_delivery d
JOIN   dim_delivery dl ON d.delivery_key = dl.delivery_key
GROUP  BY dl.shipping_mode ORDER BY ty_le_tre DESC;""", size=8.5)
table(["shipping_mode", "so_don", "ty_le_tre", "tre_tb_ngay"],
      [["First Class", "9.602", "100,0%", "1,00"],
       ["Second Class", "12.256", "80,0%", "2,00"],
       ["Same Day", "3.407", "48,4%", "0,48"],
       ["Standard Class", "37.632", "39,8%", "−0,01"]],
      widths=[1.8, 1.4, 1.4, 1.4])
para("")
para("D3. Doanh thu theo Năm → Quý (fact_sales + dim_date), kiểm tra roll-up và drill-down:",
     size=10)
code("""SELECT t.year, t.quarter, SUM(f.sales) AS doanh_thu
FROM   fact_sales f
JOIN   dim_date t ON f.order_date_key = t.date_key
GROUP  BY t.year, t.quarter ORDER BY t.year, t.quarter;""", size=8.5)
table(["Năm", "Q1", "Q2", "Q3", "Q4", "Tổng năm"],
      [["2015", "2.898.279", "2.948.575", "2.964.290", "3.006.623", "11.817.768"],
       ["2016", "2.924.098", "2.895.627", "2.970.055", "3.004.672", "11.794.452"],
       ["2017", "2.948.355", "3.024.342", "3.203.333", "2.109.467", "11.285.497"],
       ["2018", "316.713", "—", "—", "—", "316.713"]],
      widths=[0.8, 1.1, 1.1, 1.1, 1.1, 1.2])
para("Nhận xét: tổng bốn quý bằng tổng năm (ví dụ 2016: 11.794.452; lệch tối đa 1 USD do làm tròn từng quý) nên roll-up từ quý "
     "lên năm và drill-down từ năm xuống quý nhất quán; Q4/2017 giảm mạnh do thay đổi cấu "
     "trúc dữ liệu từ 10/2017 (Mục 3.5, Insight 1).", italic=True, size=9, color=GREY)
para("E. Bằng chứng cho việc tách hai bảng fact", bold=True)
bullet("nếu nối fact_sales với fact_order_delivery theo order_id rồi cộng order_sales (đại "
       "lượng cấp đơn bị lặp trên từng mặt hàng), tổng là 122.806.314 USD, gấp 3,49 lần doanh "
       "thu thật 35.214.430 USD.", "Cộng sai doanh thu: ")
bullet("đếm « đơn trễ » trên fact_sales cho 98.977 trường hợp thay vì 36.048 đơn thật (gấp "
       "2,75 lần). Tỷ lệ phần trăm ít sai lệch (≤ 0,5 điểm % theo phương thức) nhưng số đếm "
       "sai nhiều nên các KPI về số lượng đơn phải lấy từ fact_order_delivery.",
       "Đếm sai số đơn: ")
para("Kết luận 7.5: dữ liệu sau ETL đầy đủ (đối soát khớp 100% với dữ liệu thô sau khi loại đơn huỷ/nghi gian lận), nhất quán "
     "(khoá, grain, quy tắc nghiệp vụ, phân cấp đều đạt) và có cấu trúc phù hợp phân tích: "
     "mọi câu hỏi nghiệp vụ trả lời bằng một bước join từ fact tới dimension, drill-down và "
     "roll-up cho kết quả chính xác, và việc tách hai fact ngăn được việc cộng và đếm trùng.",
     bold=True)

# ---------------------------------------------------------------- 7.6
h("7.6. Nhận xét và hạn chế của mô hình dữ liệu", 2)
bullet(f"đơn CANCELED và SUSPECTED_FRAUD ({vn(S['excluded']['orders'])} đơn, {vn(S['excluded']['rows'])} dòng) "
       "đã bị loại ngay trong ETL (Mục 3.6) nên doanh thu, lợi nhuận và tỷ lệ giao trễ không còn tính "
       "các đơn này, đồng thời không còn dòng « Shipping canceled » nào bị tính là đúng hạn. Đánh đổi: "
       f"mô hình thấp hơn CSV thô {vn2(S['excluded']['sales'] / S['excluded']['before']['total_sales'] * 100, 1)}% doanh thu "
       "và không phân tích được tỷ lệ huỷ/gian lận (phải quay về CSV thô). Các đơn đang xử lý "
       "(PENDING, PENDING_PAYMENT, PROCESSING, ON_HOLD, PAYMENT_REVIEW) vẫn được giữ và tính vào doanh thu; "
       "muốn loại thêm chỉ cần bổ sung trạng thái vào EXCLUDED_STATUS rồi chạy lại ETL.",
       "Đơn huỷ và nghi gian lận: ")
bullet("dữ liệu từ 10/2017 có cấu trúc khác (Mục 3.5) nên các phép so sánh theo năm hoặc quý "
       "gồm giai đoạn này cần ghi chú; ETL giữ nguyên dữ liệu gốc.", "Thay đổi cuối kỳ: ")
bullet("dim_customer và dim_product dùng kiểu 1 (ghi đè). Dữ liệu nguồn không có thay đổi "
       "thuộc tính nên chưa mất thông tin; nếu sau này thuộc tính thay đổi theo thời gian "
       "cần chuyển sang kiểu 2 (thêm cột hiệu lực).", "Lịch sử thuộc tính: ")
bullet("các tệp CSV là kết quả cuối; để chạy định kỳ cần đưa SQL ở Mục 7.4 vào cơ sở dữ liệu "
       "hoặc lập lịch etl_phanA.py.", "Vận hành: ")

# ============================================================ 8. TABLEAU
h("8. Hướng dẫn thiết kế dashboard trên Tableau", 1)
h("8.1. Ứng dụng sử dụng & kết nối dữ liệu", 2)
bullet("dùng Tableau Desktop (bản Student miễn phí 1 năm qua Tableau for Students) hoặc Tableau "
       "Public (miễn phí; workbook được lưu công khai trên Tableau Public cloud — chấp nhận được "
       "vì DataCo là dữ liệu công khai trên Kaggle). Không cần Tableau Prep vì ETL đã làm bằng "
       "Python; không cần Power BI.", "Ứng dụng: ")
bullet("Connect → Text file → chọn các CSV trong tableau_data/. Kiểm tra kiểu dữ liệu: "
       "full_date = Date; các cột *_key, order_id, order_item_id, year, month = Dimension "
       "(chuột phải → Convert to Dimension) để Tableau không cộng dồn chúng như measure.",
       "Kết nối: ")
bullet("tạo HAI data source vì hai fact khác grain (nếu nối chung sẽ đếm trùng ~2,75 lần/đơn):",
       "Mô hình: ")
bullet("fact_sales ở giữa, Relationship tới dim_customer, dim_product, dim_location, "
       "dim_delivery (khớp theo *_key); dim_date nối hai lần: order_date_key và ship_date_key "
       "— đổi tên hai bảng logic thành Order Date và Ship Date (role-playing).",
       "Data source 1 « DataCo Sales »: ")
bullet("fact_order_delivery ở giữa, Relationship tới dim_customer, dim_location, "
       "dim_delivery, và dim_date (order_date_key; thêm lần hai cho ship_date_key nếu cần).",
       "Data source 2 « DataCo Delivery »: ")
bullet("Tableau không cho nối hai fact vào chung một mô hình khi chúng cùng chia sẻ nhiều "
       "dimension (sẽ tạo vòng lặp), nên cách tách hai data source là đúng. Hai nguồn dùng "
       "chung tên trường (market, shipping_mode, segment, year) để các Filter Action trên "
       "dashboard lọc xuyên hai nguồn.", "Lưu ý: ")
bullet("khi kéo order_date_key và ship_date_key cùng vào một quan hệ với dim_date, Tableau tạo "
       "điều kiện AND (cả hai khóa phải bằng date_key), nên chỉ giữ các đơn giao cùng ngày đặt "
       "(4.839/172.765 dòng; doanh thu chỉ ~$0,98 triệu thay vì $35,21 triệu) và sinh giá trị Null "
       "ở trường ngày. Cách đúng: mỗi quan hệ chỉ dùng MỘT khóa. Với dashboard này, nối dim_date "
       "theo order_date_key; chỉ thêm một bản dim_date thứ hai cho ship_date_key khi cần phân tích "
       "theo ngày giao.", "Lỗi thường gặp: ")

h("8.2. Dựng biểu đồ chi tiết trong Tableau", 2)
para("Bước 1 — Tạo Calculated Field (Analysis → Create Calculated Field), sau đó chuột phải "
     "vào field → Default Properties → Number Format để đặt Percentage (1 chữ số thập phân) "
     "hoặc Currency ($, 0 chữ số thập phân):")
table(["Tên field", "Data source", "Công thức"],
      [["Profit Margin %", "Sales", "SUM([profit]) / SUM([sales])"],
       ["Loss Flag", "Sales", "IF [profit] < 0 THEN 1 ELSE 0 END"],
       ["Loss Rate %", "Sales", "SUM([Loss Flag]) / COUNT([order_item_id])"],
       ["AOV", "Sales", "SUM([net_sales]) / COUNTD([order_id])"],
       ["Late Rate %", "Delivery", "SUM([late_delivery_flag]) / COUNT([order_id])"],
       ["Avg Delay (ngày)", "Delivery", "AVG([shipping_delay])"]],
      widths=[1.5, 1.1, 3.2])
para("")
para("Bước 2 — Dựng từng Sheet (Worksheet). Bảng dưới là tóm tắt; thao tác từng sheet và "
     "phân tích thiết kế nằm ở Mục 8.3. Cách chung: kéo field từ khung Data bên trái vào "
     "Columns / Rows / Marks (Color, Size, Label, Tooltip), hoặc chọn loại biểu đồ ở Show Me; "
     "đặt tên sheet theo cột đầu tiên của bảng dưới:")
table(["Sheet (nguồn)", "Loại biểu đồ", "Kéo thả & thiết lập", "Định dạng"],
      [["K1–K7 Thẻ KPI (Sales/ Delivery)", "Text (BAN)",
        "Mỗi KPI một sheet: kéo measure vào Text; Fit → Entire View; ẩn tiêu đề cột "
        "(chuột phải → Hide Field Labels). K1 SUM(sales), K2 SUM(profit), K3 Profit Margin %, "
        "K4 Late Rate %, K5 Avg Delay, K6 Loss Rate %, K7 AOV.",
        "Font 28 đậm; tiêu đề sheet = tên KPI; Late Rate % và Loss Rate % tô cam đỏ (#D55E00)."],
       ["S1 Xu hướng thời gian (Sales)", "Dual-axis line",
        "Columns: full_date (Order Date) → chuột phải chọn MONTH loại liên tục (Continuous). "
        "Rows: SUM(sales). Kéo SUM(profit) thả vào mép phải trục → Dual Axis.",
        "Không Synchronize Axis (hai thang đo khác nhau); Profit màu cam. Ghi chú: 01/2018 là "
        "tháng dở dang."],
       ["S2 Doanh thu theo thị trường (Sales)", "Bar ngang",
        "Rows: market; Columns: SUM(sales). Color: Profit Margin %. Label: SUM(sales).",
        "Sort giảm dần; dùng dải màu tuần tự (sequential) xanh nhạt → xanh đậm #DEEBF7 → #08519C: biên LN càng cao càng đậm."],
       ["S3 Doanh thu & biên LN theo ngành (Sales)", "Bar ngang",
        "Rows: department_name; Columns: SUM(sales); Color: Profit Margin %; Label: SUM(sales).",
        "Sort giảm dần; tooltip thêm SUM(profit)."],
       ["S4 Giao trễ theo phương thức (Delivery)", "Bar ngang + Reference Line",
        "Rows: shipping_mode; Columns: Late Rate %; Color: Avg Delay (độ trễ TB, mức 2); "
        "Tooltip: COUNTD(order_id). Analytics → Reference Line (Constant 0,5731).",
        "Dải cam tuần tự #FDE0C5 → #D55E00, càng trễ lâu càng đậm; Stepped 4; sort giảm dần; nhãn %."],
       ["S5 Phân khúc khách hàng (Sales)", "Treemap",
        "Marks: Square; Size: SUM(sales); Color: segment; Label: segment + SUM(sales) hiển "
        "thị % of total (Quick Table Calculation → Percent of Total).",
        "Chỉ 3 phân khúc nên giữ màu phân biệt rõ."],
       ["S6 Tỷ lệ đơn lỗ theo ngành (Sales)", "Bar ngang + Reference Line",
        "Rows: department_name; Columns: Loss Rate %; Color: COUNT(order_item_id) (cỡ mẫu, "
        "mức 2); Reference Line = 0,1869.",
        "Xanh xám Stepped 3 (nhạt = ít dữ liệu). Kết quả đồng đều ~18–19% ở mọi ngành lớn và "
        "mọi mức chiết khấu → lỗ mang tính hệ thống, không do chiết khấu."],
       ["S7 Top 10 sản phẩm (Sales)", "Bar ngang",
        "Rows: product_name; Columns: SUM(sales). Kéo product_name vào Filters → tab Top → "
        "By field: Top 10 by SUM(sales). Color: Profit Margin %.",
        "Thay thế bảng chi tiết trên dashboard chính."],
       ["D2 Pivot ngành × năm (Sales)", "Heatmap (Highlight table)",
        "Rows: department_name; Columns: YEAR(full_date) (Discrete); SUM(sales) vào Text và "
        "Color (mức 2).", "Sequential xanh Stepped 4; dùng cho thao tác Pivot (Mục 8.4)."],
       ["D1 Chi tiết sản phẩm (Sales)", "Text table",
        "Rows: product_name, department_name; Measure Values: SUM(sales), SUM(profit), "
        "Profit Margin %.", "Đặt ở Dashboard 3."]],
      widths=[1.3, 0.9, 2.6, 1.3])
para("")
para("Về bảng chi tiết: dashboard tổng quan KHÔNG nên chứa bảng nhiều dòng — người ra quyết "
     "định cần đọc được trong vài giây. Vì vậy bảng sản phẩm (D1) và bảng pivot (D2) được đặt "
     "ở Dashboard 3 (màn hình soi sâu), còn Dashboard 1 và 2 chỉ chứa thẻ KPI và biểu đồ. "
     "Người xem chuyển giữa ba dashboard bằng nút Navigation.", italic=True)

from chart_sections import add_chart_sections
add_chart_sections(h, para, bullet, table)

h("8.4. Hướng dẫn chi tiết các thao tác OLAP", 2)
para("Chuẩn bị phân cấp (làm một lần trong Data source 1):", bold=True)
bullet("kéo full_date của Order Date vào Columns; Tableau tạo YEAR(full_date). Bấm dấu (+) "
       "trên pill để khoan sâu Year → Quarter → Month → Day; bấm (−) để cuộn lên.",
       "Thời gian: ")
bullet("trong khung Data, kéo region thả lên market → đặt tên « Geography » → kéo thêm "
       "country, state, city vào theo thứ tự.", "Địa lý: ")
bullet("kéo category_name thả lên department_name → đặt tên « Product Hierarchy » → thêm "
       "product_name.", "Sản phẩm: ")
para("Các thao tác (số liệu trong ngoặc để đối chiếu khi demo):", bold=True)
bullet("(1) Mở sheet S2. (2) Kéo market vào Filters → chọn Europe → OK. (3) Chuột phải pill "
       "trong Filters → Show Filter → chọn Single Value (dropdown). Kết quả: doanh thu "
       "$10,41 triệu, biên LN 10,88%, ngành lớn nhất Fan Shop $4,54 triệu. Cách khác: nhấp một "
       "thanh trên S2 rồi chọn biểu tượng Use as Filter. Slice = cố định MỘT chiều.", "Slice: ")
bullet("(1) Giữ lọc market = Europe. (2) Kéo thêm YEAR(full_date) và department_name vào "
       "Filters, chọn 2017 và Fan Shop. Kết quả: $1,82 triệu, biên LN 11,06%. Để xem hộp "
       "chéo nhiều chiều: tạo sheet mới, Rows = department_name, Columns = market, Marks = "
       "Square, Color = Profit Margin % (Highlight table). Dice = lọc đồng thời "
       "NHIỀU chiều.", "Dice: ")
bullet("(1) Trên S1 bấm (+) ở YEAR → QUARTER → MONTH → DAY; mỗi lần bấm mức chi tiết tăng một "
       "cấp. Đối chiếu tổng theo năm: 2015 $11,82 triệu; 2016 $11,79 triệu; 2017 $11,29 triệu "
       "(−4,3% so với 2016, nhưng cùng kỳ 01–09 thì +4,4%: dữ liệu từ 10/2017 đổi cấu trúc, xem S1); 2018 chỉ có tháng 1 ($0,32 triệu). (2) Trên S2 dùng Geography: bấm "
       "(+) ở market để xuống region → country. Roll-up: bấm (−) trên pill để quay về cấp cao.",
       "Drill-down / Roll-up: ")
bullet("(1) Dùng sheet D2 (Pivot_Dept_Year): Rows = department_name; Columns = YEAR(full_date); "
       "kéo SUM(sales) vào Text và Color → có bảng chéo (cross-tab) dạng heatmap. (2) Nhấn Swap Rows and Columns "
       "(Ctrl+W) để xoay: năm xuống hàng, ngành lên cột. (3) Kéo market từ Filters sang "
       "Columns để thêm chiều thứ hai và so sánh. Lưu ý đây là pivot theo nghĩa OLAP (xoay "
       "góc nhìn), khác lệnh Pivot trong Data Source (chuyển cột thành hàng) — không cần dùng.",
       "Pivot: ")
bullet("Dashboard → Actions → Add Action → Filter: Source = các sheet S2–S3, Target = tất cả "
       "sheet, Run action on Select. Bấm một thị trường thì cả KPI và biểu đồ (cả nguồn Sales "
       "lẫn Delivery) đổi theo vì dùng chung tên trường market. Thêm Highlight Action cho "
       "shipping_mode.", "Tương tác: ")
para("Kịch bản demo 3 phút: Slice Europe → Dice thêm 2017 và Fan Shop → Drill-down theo "
     "tháng thấy xu hướng (và chỗ gãy 10/2017) → Pivot ngành × năm → bỏ lọc, sang Dashboard "
     "2, bấm First Class trên S4 để cho thấy chỉ số giao trễ.", italic=True)

h("8.5. Bố cục dashboard", 2)
para("Báo cáo gồm 3 dashboard; tất cả các sheet K1–K7, S1–S7, D1, D2 đều được đặt vào một "
     "dashboard (không có sheet nào bị bỏ sót). Mỗi dashboard: Dashboard → New Dashboard, "
     "Size: Fixed 1300×800, đặt tên theo bảng dưới rồi kéo sheet vào theo bố cục.")
para("Dashboard 1 — Hiệu quả kinh doanh", bold=True)
table(["Vùng", "Nội dung"],
      [["Hàng 1", "Tiêu đề « DataCo — Hiệu quả kinh doanh » + 4 thẻ KPI K1 Doanh thu, K2 Lợi "
                  "nhuận, K3 Biên LN, K7 AOV (Horizontal container)"],
       ["Hàng 2 (toàn chiều rộng)", "S1 Xu hướng doanh thu & lợi nhuận theo tháng"],
       ["Hàng 3 — trái (50%)", "S2 Doanh thu theo thị trường"],
       ["Hàng 3 — phải (50%)", "S3 Doanh thu & biên LN theo ngành hàng"],
       ["Cạnh phải", "Bộ lọc: Year, Market, Department (Show Filter)"]],
      widths=[1.7, 4.1])
para("")
para("Dashboard 2 — Giao hàng & đơn bán lỗ", bold=True)
table(["Vùng", "Nội dung"],
      [["Hàng 1", "Tiêu đề « DataCo — Giao hàng & đơn lỗ » + 3 thẻ KPI K4 Tỷ lệ giao trễ, "
                  "K5 Độ trễ TB, K6 Tỷ lệ đơn lỗ"],
       ["Hàng 2 — trái (50%)", "S4 Tỷ lệ giao trễ & độ trễ theo phương thức"],
       ["Hàng 2 — phải (50%)", "S6 Tỷ lệ dòng đơn lỗ theo ngành"],
       ["Cạnh phải", "Bộ lọc: Year, Market, Shipping Mode"]],
      widths=[1.7, 4.1])
para("")
para("Dashboard 3 — Khách hàng, sản phẩm & chi tiết", bold=True)
table(["Vùng", "Nội dung"],
      [["Hàng 1", "Tiêu đề « DataCo — Khách hàng & sản phẩm »"],
       ["Hàng 2 — trái (40%)", "S5 Cơ cấu doanh thu theo phân khúc (treemap)"],
       ["Hàng 2 — phải (60%)", "S7 Top 10 sản phẩm"],
       ["Hàng 3 — trái (50%)", "D2 Heatmap ngành × năm"],
       ["Hàng 3 — phải (50%)", "D1 Bảng chi tiết sản phẩm"],
       ["Cạnh phải", "Bộ lọc: Year, Segment, Department"]],
      widths=[1.7, 4.1])
para("")
bullet("thêm nút Navigation (Objects → Navigation) ở đầu mỗi dashboard để chuyển giữa 3 "
       "màn hình; mỗi dashboard có tiêu đề nêu chủ đề.", "Điều hướng: ")
bullet("bộ lọc nhanh chỉ có tác dụng trong cùng một data source. Ở Dashboard 2, S4 dùng "
       "nguồn Delivery còn S6 dùng nguồn Sales nên cần liên kết trường dùng chung: Data → Edit "
       "Blend Relationships → liên kết market, shipping_mode và year; sau đó chuột phải bộ "
       "lọc → Apply to Worksheets → All Using Related Data Sources. Nếu không muốn cấu hình, "
       "dùng Filter Action (Mục 8.4) vì hai nguồn trùng tên trường. Nên thử lọc một giá trị "
       "để xác nhận cả hai sheet cùng đổi.", "Lọc xuyên hai nguồn: ")
bullet("tối đa 2–3 màu chủ đạo; cam đỏ #D55E00 chỉ dành cho chỉ số cảnh báo (trễ, lỗ); một "
       "dải xanh tuần tự duy nhất cho biên LN/doanh thu ở S2, S3, S7, D2; bảng màu Okabe-Ito an "
       "toàn cho người mù màu; mỗi biểu đồ có tiêu đề nêu thông điệp.", "Quy tắc màu: ")

h("8.6. Kiểm tra mức độ đầy đủ của dashboard", 2)
table(["Yêu cầu đề bài", "Đáp ứng bởi", "Kết luận"],
      [["≥ 5 KPI trên dashboard", "7 thẻ KPI: K1, K2, K3, K7 (Dashboard 1); K4, K5, K6 "
        "(Dashboard 2) — Mục 5", "Đủ"],
       ["KPI thể hiện bằng biểu đồ/ bảng phù hợp", "Doanh thu/LN/biên LN → S1, S2, S3; giao "
        "trễ & độ trễ → S4; đơn lỗ → S6; AOV → K7 (so theo phân khúc ở tooltip S5)", "Đủ"],
       ["OLAP: slice/dice, drill-down, pivot", "Mục 8.4: Filter market, Highlight table, "
        "hierarchy Date/Geography/Product, cross-tab ngành × năm (D2)", "Đủ"],
       ["≥ 3 insight có bằng chứng trên dashboard", "7 insight (Mục 9), chia theo 3 dashboard; "
        "mọi sheet S1–S7 và D2 đều là bằng chứng cho ít nhất một insight", "Đủ"],
       ["Dễ đọc, phục vụ ra quyết định", "Mỗi dashboard một chủ đề, tối đa 4–5 visual; bảng "
        "chi tiết tách sang Dashboard 3", "Đủ"]],
      widths=[1.8, 3.3, 0.7])
para("")
para("Hạn chế cần nêu khi trình bày: (a) biên lợi nhuận hầu như đồng đều giữa các thị trường "
     "và ngành (≈ 10–11%) nên màu theo biên LN ít phân biệt; điểm khác biệt thực sự nằm ở quy "
     "mô doanh thu và tỷ lệ trễ theo phương thức giao hàng; (b) dữ liệu 2018 chỉ có tháng 1; "
     "(c) từ 10/2017 dữ liệu đổi cấu trúc (dữ liệu thô đúng 68–69 đơn/ngày, mỗi đơn một mặt hàng, "
     "chỉ 8–20 sản phẩm) nên xu hướng và so sánh theo năm giai đoạn này cần đọc thận trọng; "
     "(d) bản đồ không dùng vì tên quốc gia (tiếng Tây Ban Nha) khó nhận diện địa lý.",
     italic=True)


# ============================================================ 9. INSIGHT
h("9. Phân tích & khuyến nghị", 1)
para("Dựa trên 3 dashboard, nhóm rút ra 7 insight; mỗi insight có bằng chứng trực tiếp "
     "trên một hoặc nhiều biểu đồ và dẫn tới một quyết định cụ thể. Các tỷ lệ giao trễ tính "
     "trên 62.897 đơn (fact_order_delivery).")
table(["Insight", "Dashboard", "Bằng chứng", "Quyết định hỗ trợ"],
      [["1. Doanh thu ổn định; cuối kỳ gãy do dữ liệu", "1", "S1, D2, K1",
        "Không dùng 10/2017–01/2018 để đánh giá hay dự báo; kiểm tra lại nguồn dữ liệu"],
       ["2. Doanh thu tập trung vào vài thị trường/ngành", "1", "S2, S3",
        "Phân bổ nguồn lực; quản lý rủi ro tập trung"],
       ["3. Biên lợi nhuận đồng đều ~10–11%", "1", "K3, S2, S3",
        "Tăng lợi nhuận bằng quy mô và giảm đơn lỗ, không bằng đổi thị trường"],
       ["4. Giao trễ tập trung ở First/Second Class", "2", "S4, K4, K5",
        "Rà soát SLA, nhà vận chuyển hai phương thức này"],
       ["5. Đơn bán lỗ lan đều, không do chiết khấu", "2", "K6, S6",
        "Rà soát giá vốn/giá bán toàn công ty thay vì chỉ siết chiết khấu"],
       ["6. Consumer chiếm hơn một nửa doanh thu", "3", "S5",
        "Ưu tiên marketing theo quy mô phân khúc"],
       ["7. 10 sản phẩm tạo 90% doanh thu", "3", "S7, D2",
        "Quản lý tồn kho, nhà cung cấp, khuyến mãi theo sản phẩm chủ lực"]],
      widths=[2.1, 0.7, 1.0, 2.5])
para("")
para("Dashboard 1 — Hiệu quả kinh doanh", bold=True)
bullet("doanh thu quanh $1,0 triệu/tháng suốt 01/2015–09/2017 (cùng kỳ 01–09: 2017 cao hơn "
       "2016 4,4%). Doanh thu rơi từ 10/2017 xuống $0,32 triệu (01/2018) nhưng đồng thời "
       "dữ liệu đổi cấu trúc (dữ liệu thô đúng 68–69 đơn/ngày, mỗi đơn một mặt hàng, Fan Shop/Golf/Footwear "
       "gần như biến mất) nên chưa đủ cơ sở kết luận kinh doanh suy giảm.", "Insight 1 — ")
dep = S['sales_by_department']; top_dep = list(dep.items())[0]
mk = S['sales_by_market']
bullet(f"Europe ({money(mk['Europe'])}) + LATAM ({money(mk['LATAM'])}) chiếm 57,4% doanh thu; "
       f"ngành {top_dep[0]} một mình chiếm {money(top_dep[1])} (46,6%) và bốn ngành lớn chiếm "
       "91,6%. Africa chỉ 6,3% — vừa là rủi ro phụ thuộc, vừa là dư địa tăng trưởng.",
       "Insight 2 — ")
bullet(f"biên lợi nhuận toàn công ty chỉ {S['profit_margin_pct']}% và gần như không đổi giữa "
       "các thị trường (10,40%–11,15%) hay ngành lớn (10,15%–11,08%). Đòn bẩy lợi nhuận "
       "nằm ở quy mô và ở việc giảm đơn lỗ (Insight 5).", "Insight 3 — ")
para("Dashboard 2 — Giao hàng & đơn bán lỗ", bold=True)
bullet(f"giao trễ ở mức báo động {S['late_delivery_rate_order_pct']}% tổng số đơn. Nghịch "
       "lý: phương thức nhanh/đắt tiền lại trễ nhiều nhất — First Class 100% (trễ TB 1,0 "
       "ngày), Second Class 80,0% (trễ TB 2,0 ngày), trong khi Standard Class chỉ 39,8%. Tỷ lệ "
       "trễ ở mọi thị trường đều ~56,5–57,7% nên nguyên nhân nằm ở phương thức giao chứ không "
       "ở địa phương.", "Insight 4 — ")
bullet(f"{S['loss_order_line_pct']}% dòng đơn ({S['loss_order_lines']:,} dòng) bị LỖ dù doanh "
       "thu dương, tổng lỗ khoảng $3,71 triệu — gần bằng toàn bộ lợi nhuận ròng "
       f"{money(S['total_profit'])}. Nếu các dòng này về mức hoà vốn, lợi nhuận tăng lên "
       "khoảng $7,51 triệu (gần gấp đôi). Tỷ lệ lỗ đồng đều ở mọi ngành lớn (18,7–19,0%) và "
       "mọi mức chiết khấu 0–25% (18–20%); chiết khấu trung bình của dòng lỗ (10,20%) xấp xỉ "
       "dòng có lãi (10,15%) → chiết khấu KHÔNG phải nguyên nhân chính.", "Insight 5 — ")
para("Dashboard 3 — Khách hàng, sản phẩm & chi tiết", bold=True)
bullet("Consumer đóng góp $18,26 triệu (51,9%), Corporate 30,4%, Home Office 17,7%. Giá trị đơn "
       "trung bình và biên LN của ba phân khúc gần như bằng nhau ($497,6–505,9; 10,67–10,86%) "
       "nên khác biệt chỉ ở số đơn.", "Insight 6 — ")
bullet("10 sản phẩm trên tổng 118 tạo 90,0% doanh thu; riêng Field & Stream Sportsman 16 Gun "
       "Fire Safe chiếm 18,8% và hai sản phẩm đầu chiếm 30,9%. Heatmap D2 cho thấy sự phụ "
       "thuộc vào Fan Shop và chỗ gãy của năm 2017.", "Insight 7 — ")
para("Khuyến nghị ra quyết định:")
bullet("rà soát nhà vận chuyển và quy trình của First/Second Class, đặt SLA và cảnh báo riêng "
       "cho hai phương thức này vì đang trễ nặng nhất (Insight 4).", "Vận hành: ")
bullet("điều tra cơ cấu giá vốn/giá bán của các dòng lỗ và đặt cảnh báo khi profit_ratio âm; "
       "chỉ siết ngưỡng chiết khấu không đủ vì tỷ lệ lỗ không đổi theo chiết khấu (Insight 5).",
       "Giá & lợi nhuận: ")
bullet("tập trung marketing/tồn kho vào Europe, LATAM, Fan Shop, phân khúc Consumer và nhóm "
       "sản phẩm chủ lực; đồng thời lập kế hoạch giảm rủi ro phụ thuộc, khảo sát vì sao "
       "Africa còn nhỏ để tìm dư địa tăng trưởng (Insight 2, 6, 7).", "Kinh doanh: ")
bullet("xác minh với đơn vị cung cấp dữ liệu lý do cấu trúc dữ liệu thay đổi từ 10/2017 trước "
       "khi dùng 4 tháng cuối cho dự báo hoặc đánh giá xu hướng (Insight 1).", "Dữ liệu: ")

# ============================================================ LƯU
out = os.path.join(HERE, "PhanA_BaoCao.docx")
doc.save(out)
print("saved", out)
